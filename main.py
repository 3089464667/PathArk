import os
import time
import sqlite3
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import PhotoImage
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image, ImageTk
import pylnk3
import win32api
import win32con
import win32ui
import win32gui
import win32com.client
import subprocess
import pystray
from pystray import MenuItem as item
from PIL import Image as PILImage

RECENT_PATH = r'C:\Users\Lenovo\AppData\Roaming\Microsoft\Windows\Recent'
DB_PATH = 'access_freq.db'
TOP_N = 100

class AccessDB:
    def __init__(self, db_path=DB_PATH):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.create_table()
    def create_table(self):
        self.conn.execute('''CREATE TABLE IF NOT EXISTS access (
            target TEXT PRIMARY KEY,
            name TEXT,
            is_folder INTEGER,
            freq INTEGER DEFAULT 0,
            last_atime REAL DEFAULT 0
        )''')
        self.conn.commit()
    def add_or_update(self, target, name, is_folder):
        cur = self.conn.cursor()
        atime = 0
        try:
            atime = os.path.getatime(target)
        except Exception:
            atime = 0
        cur.execute('SELECT freq FROM access WHERE target=?', (target,))
        row = cur.fetchone()
        if row:
            cur.execute('UPDATE access SET freq=freq+1, last_atime=? WHERE target=?', (atime, target))
        else:
            cur.execute('INSERT INTO access (target, name, is_folder, freq, last_atime) VALUES (?, ?, ?, 1, ?)', (target, name, is_folder, atime))
        self.conn.commit()
    def remove(self, target):
        cur = self.conn.cursor()
        cur.execute('DELETE FROM access WHERE target=?', (target,))
        self.conn.commit()
    def exists(self, target):
        cur = self.conn.cursor()
        cur.execute('SELECT 1 FROM access WHERE target=?', (target,))
        return cur.fetchone() is not None
    def get_all_targets(self):
        cur = self.conn.cursor()
        cur.execute('SELECT target FROM access')
        return [row[0] for row in cur.fetchall()]
    def get_top(self, n=TOP_N):
        cur = self.conn.cursor()
        cur.execute('SELECT target, name, is_folder, freq FROM access ORDER BY freq DESC LIMIT ?', (n,))
        return cur.fetchall()
    def close(self):
        self.conn.close()
    def init_from_recent(self, recent_path=RECENT_PATH, n=TOP_N):
        # 获取Recent目录下所有.lnk文件，按修改时间排序，取前n个
        lnk_files = [os.path.join(recent_path, f) for f in os.listdir(recent_path) if f.lower().endswith('.lnk')]
        lnk_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        count = 0
        for lnk in lnk_files:
            target = get_lnk_target(lnk)
            if target and os.path.exists(target):
                name = os.path.basename(target)
                isdir = int(is_folder(target))
                cur = self.conn.cursor()
                cur.execute('SELECT 1 FROM access WHERE target=?', (target,))
                if not cur.fetchone():
                    cur.execute('INSERT INTO access (target, name, is_folder, freq) VALUES (?, ?, ?, 1)', (target, name, isdir))
                    self.conn.commit()
                    count += 1
                if count >= n:
                    break

def get_lnk_target(lnk_path):
    try:
        lnk = pylnk3.parse(lnk_path)
        return lnk.path
    except Exception:
        try:
            shell = win32com.client.Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(lnk_path)
            return shortcut.Targetpath
        except Exception:
            return None

def is_folder(path):
    return os.path.isdir(path)

def get_icon_image(path, size=(32, 32)):
    try:
        ilist = win32gui.SHGetFileInfo(path, 0, win32con.SHGFI_ICON | win32con.SHGFI_LARGEICON)
        hicon = ilist[0]
        if hicon:
            icon = win32ui.CreateBitmapFromHandle(hicon)
            bmpinfo = icon.GetInfo()
            bmpstr = icon.GetBitmapBits(True)
            img = Image.frombuffer('RGBA', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRA', 0, 1)
            img = img.resize(size, Image.LANCZOS)
            return ImageTk.PhotoImage(img)
    except Exception:
        pass
    return None

class RecentHandler(FileSystemEventHandler):
    def __init__(self, db: AccessDB):
        self.db = db
    def on_modified(self, event):
        if event.is_directory:
            if os.path.exists(event.src_path):
                self.db.add_or_update(event.src_path, os.path.basename(event.src_path), 1)
            return
        if event.src_path.lower().endswith('.lnk'):
            target = get_lnk_target(event.src_path)
            if target and os.path.exists(target):
                self.db.add_or_update(target, os.path.basename(target), int(is_folder(target)))
        else:
            if os.path.exists(event.src_path):
                self.db.add_or_update(event.src_path, os.path.basename(event.src_path), int(is_folder(event.src_path)))
    def on_created(self, event):
        self.on_modified(event)

class App:
    def __init__(self, root, db: AccessDB):
        self.root = root
        self.db = db
        self.light_theme = {
            "bg": "#FFFFFF",
            "fg": "#222222",
            "button_bg": "#F1F1F1"
        }
        self.dark_theme = {
            "bg": "#181A1B",
            "fg": "#EEEEEE",
            "button_bg": "#232323"
        }
        self.current_theme = self.light_theme

        self.root.title('PathArk')
        self.root.geometry('700x700')
        self.root.configure(bg=self.current_theme["bg"])
        self.style = ttk.Style()
        self.style.theme_use('clam')
        search_frame = tk.Frame(root, bg=self.current_theme["bg"])
        search_frame.pack(fill='x', padx=20, pady=(18, 0))
        self.search_var = tk.StringVar()
        self.search_canvas = tk.Canvas(search_frame, width=380, height=44, bg=self.current_theme["bg"], highlightthickness=0)
        self.search_canvas.pack(side='left', padx=(0, 10))
        self._draw_search_box(bg=self.current_theme["bg"])
        self.search_entry = tk.Entry(
            search_frame,
            textvariable=self.search_var,
            font=('微软雅黑', 13),
            bd=0,
            relief='flat',
            bg=self.current_theme["bg"],
            fg=self.current_theme["fg"]
        )
        self.search_entry.place(in_=self.search_canvas, x=18, y=8, width=340, height=28)
        self.search_entry.bind('<KeyRelease>', self.on_search)
        self.search_entry.bind('<FocusIn>', self._on_search_focus_in)
        self.search_entry.bind('<FocusOut>', self._on_search_focus_out)
        self.search_entry.insert(0, '🔍 搜索文件或文件夹...')
        self.searching = False
        self.theme_button = tk.Button(
            search_frame,
            text="切换主题",
            command=self.toggle_theme,
            bg=self.current_theme["button_bg"],
            fg=self.current_theme["fg"],
            relief='flat'
        )
        self.theme_button.pack(side='right', padx=(10, 0), pady=0)
        self.tree = ttk.Treeview(root, columns=('名称', '类型', '次数'), show='tree headings', selectmode='browse')
        self.tree.heading('#0', text='')
        self.tree.column('#0', width=48, anchor='center', stretch=False)
        self.tree.heading('名称', text='名称')
        self.tree.heading('类型', text='类型')
        self.tree.heading('次数', text='访问次数')
        self.tree.column('名称', width=320, anchor='w')
        self.tree.column('类型', width=100, anchor='center')
        self.tree.column('次数', width=100, anchor='center')
        self.tree.pack(fill='both', expand=True, padx=20, pady=20)
        self.tree.bind('<Double-1>', self.open_selected)
        self.icons = {}
        self.filtered_items = None

        # 统一收集需要主题的控件
        self.widgets = [
            self.root,
            search_frame,
            self.search_canvas,
            self.search_entry,
            self.theme_button
        ]

        # 图标初始化
        self.init_default_icons()
        self.refresh()
        self.root.after(5000, self.refresh)
        self.apply_theme(self.current_theme)

    def init_default_icons(self):
        # 加载file.png和folder.png，先缩放为较大尺寸再缩小为32x32，保证清晰
        from PIL import Image as PILImage, ImageDraw
        icon_final_size = (32, 32)
        icon_src_size = (96, 96)  # 源图先缩放到较大尺寸
        def make_circle_icon(img_path, fallback_color):
            try:
                img = PILImage.open(img_path).convert('RGBA')
                img = img.resize(icon_src_size, PILImage.LANCZOS)
                # 创建圆形蒙版
                mask = PILImage.new('L', icon_src_size, 0)
                draw = ImageDraw.Draw(mask)
                draw.ellipse((0, 0, icon_src_size[0], icon_src_size[1]), fill=255)
                # 应用蒙版
                circle_img = PILImage.new('RGBA', icon_src_size, (0, 0, 0, 0))
                circle_img.paste(img, (0, 0), mask)
                # 最后缩放到目标尺寸，保证边缘平滑
                circle_img = circle_img.resize(icon_final_size, PILImage.LANCZOS)
                return ImageTk.PhotoImage(circle_img)
            except Exception:
                # 兜底：纯色圆形
                circle_img = PILImage.new('RGBA', icon_src_size, (0, 0, 0, 0))
                draw = ImageDraw.Draw(circle_img)
                draw.ellipse((0, 0, icon_src_size[0], icon_src_size[1]), fill=fallback_color)
                circle_img = circle_img.resize(icon_final_size, PILImage.LANCZOS)
                return ImageTk.PhotoImage(circle_img)
        self.icons['file'] = make_circle_icon('file.png', (200, 200, 200, 255))
        self.icons['folder'] = make_circle_icon('folder.png', (180, 210, 240, 255))

    def apply_theme(self, theme):
        self.root.configure(bg=theme["bg"])
        for widget in self.widgets:
            # 只配置有bg/fg属性的控件
            try:
                widget.configure(bg=theme["bg"])
            except Exception:
                pass
            try:
                widget.configure(fg=theme["fg"])
            except Exception:
                pass
        try:
            self.theme_button.configure(bg=theme["button_bg"], fg=theme["fg"])
        except Exception:
            pass
        # 搜索框Entry
        try:
            self.search_entry.configure(bg=theme["bg"], fg=theme["fg"])
        except Exception:
            pass
        # Treeview样式
        self.style.configure('Treeview',
            font=('微软雅黑', 12),
            rowheight=40,
            background=theme["bg"],
            fieldbackground=theme["bg"],
            foreground=theme["fg"]
        )
        self.style.configure('Treeview.Heading',
            font=('微软雅黑', 14, 'bold'),
            background=theme["bg"],
            foreground=theme["fg"]
        )
        self._draw_search_box(bg=theme["bg"])

    def toggle_theme(self):
        self.current_theme = self.dark_theme if self.current_theme == self.light_theme else self.light_theme
        self.apply_theme(self.current_theme)

    def _draw_search_box(self, bg='#ffffff'):
        # 绘制圆角矩形，优化右侧样式
        self.search_canvas.delete('all')
        r = 22
        w, h = 380, 44
        # 主体圆角矩形
        self.search_canvas.create_rectangle(r, 0, w - r, h, fill=bg, outline=bg)
        self.search_canvas.create_rectangle(0, r, w, h - r, fill=bg, outline=bg)
        self.search_canvas.create_oval(0, 0, r * 2, r * 2, fill=bg, outline=bg)
        self.search_canvas.create_oval(w - r * 2, 0, w, r * 2, fill=bg, outline=bg)
        self.search_canvas.create_oval(0, h - r * 2, r * 2, h, fill=bg, outline=bg)
        self.search_canvas.create_oval(w - r * 2, h - r * 2, w, h, fill=bg, outline=bg)
        # 只在下方绘制一层柔和阴影，不突出右侧
        self.search_canvas.create_oval(r, h - 8, w - r, h + 8, fill='#e0e6ef', outline='#e0e6ef')

    def _on_search_focus_in(self, event):
        if self.search_entry.get().startswith('🔍'):
            self.search_entry.delete(0, 'end')
        # 动画过渡到高亮色
        self._animate_search_box('#eaf3fc')

    def _on_search_focus_out(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, '🔍 搜索文件或文件夹...')
        # 动画过渡回白色
        self._animate_search_box('#ffffff')

    def _animate_search_box(self, target_bg):
        # 简单动画过渡
        import threading
        import time as _time
        color_map = {
            '#ffffff': (255, 255, 255),
            '#eaf3fc': (234, 243, 252)
        }
        start = self.search_canvas.itemcget(1, 'fill')
        if start not in color_map:
            start = '#ffffff'
        start_rgb = color_map[start]
        end_rgb = color_map[target_bg]
        steps = 8
        def animate():
            for i in range(1, steps + 1):
                r = int(start_rgb[0] + (end_rgb[0] - start_rgb[0]) * i / steps)
                g = int(start_rgb[1] + (end_rgb[1] - start_rgb[1]) * i / steps)
                b = int(start_rgb[2] + (end_rgb[2] - start_rgb[2]) * i / steps)
                color = f'#{r:02x}{g:02x}{b:02x}'
                self.root.after(0, lambda c=color: self._draw_search_box(bg=c))
                _time.sleep(0.015)
        threading.Thread(target=animate, daemon=True).start()

    def on_search(self, event=None):
        query = self.search_var.get().strip().lower()
        if not query or query.startswith('🔍'):
            self.filtered_items = None
        else:
            all_items = self.db.get_top(TOP_N)
            self.filtered_items = [item for item in all_items if query in item[1].lower()]
        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        # 支持搜索过滤
        if self.filtered_items is not None:
            top_items = self.filtered_items
        else:
            top_items = self.db.get_top(TOP_N)
        for target, name, is_folder, freq in top_items:
            icon = self.get_icon(target, is_folder)
            # 使用#0列显示图标，text为文件名
            self.tree.insert('', 'end', text='', image=icon, values=(name, '文件夹' if is_folder else '文件', freq), tags=(target,))
        self.root.after(10000, self.refresh)
    def get_icon(self, path, is_folder):
        # 优先使用统一风格的图标
        key = (path, is_folder)
        if key in self.icons:
            return self.icons[key]
        if is_folder:
            self.icons[key] = self.icons['folder']
            return self.icons['folder']
        else:
            self.icons[key] = self.icons['file']
            return self.icons['file']
    def open_selected(self, event):
        item = self.tree.selection()
        if not item:
            return
        target = self.tree.item(item[0], 'tags')[0]
        if os.path.exists(target):
            try:
                # 更新访问时间和修改时间为当前时间
                now = time.time()
                os.utime(target, (now, now))
                # 仅当数据库有记录时才+1
                if self.db.exists(target):
                    self.db.add_or_update(target, os.path.basename(target), int(is_folder(target)))
                os.startfile(target)
            except Exception as e:
                messagebox.showerror('打开失败', f'无法打开：{target}\n{e}')
        else:
            messagebox.showwarning('文件不存在', f'目标已不存在：{target}')

def start_watcher(db):
    event_handler = RecentHandler(db)
    observer = Observer()
    observer.schedule(event_handler, RECENT_PATH, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def sync_recent_with_db(db: AccessDB, interval=10):
    while True:
        # 1. 获取Recent目录下所有文件和文件夹
        entries = [os.path.join(RECENT_PATH, f) for f in os.listdir(RECENT_PATH)]
        recent_targets = set()
        for entry in entries:
            if os.path.isdir(entry):
                # 目录直接加入
                target = entry
            elif entry.lower().endswith('.lnk'):
                target = get_lnk_target(entry)
                if not target:
                    continue
            else:
                # 其他文件直接加入
                target = entry
            if target and os.path.exists(target):
                recent_targets.add(target)
                if not db.exists(target):
                    db.add_or_update(target, os.path.basename(target), int(is_folder(target)))
        # 2. 移除数据库中不在Recent的条目
        all_db_targets = set(db.get_all_targets())
        for target in all_db_targets - recent_targets:
            db.remove(target)
        time.sleep(interval)

def create_tray(app, root):
    # 创建一个简单的托盘图标
    icon_img = PILImage.new('RGBA', (32, 32), (70, 130, 180, 255))
    def on_show(icon=None, item=None):
        root.after(0, root.deiconify)
    def on_exit(icon, item):
        icon.stop()
        root.after(0, root.destroy)
    menu = (item('显示主界面', on_show), item('退出', on_exit))
    icon = pystray.Icon('file_rapid', icon_img, 'PathArk', menu)
    icon.on_activate = on_show  
    threading.Thread(target=icon.run, daemon=True).start()
    return icon

def main():
    db = AccessDB()
    db.init_from_recent()  # 首次运行时初始化数据库
    t = threading.Thread(target=start_watcher, args=(db,), daemon=True)
    t.start()
    # 修改为每10秒自动扫描一次Recent目录并同步数据库
    t2 = threading.Thread(target=sync_recent_with_db, args=(db, 10), daemon=True)
    t2.start()
    root = tk.Tk()
    app = App(root, db)
    tray_icon = create_tray(app, root)
    def on_closing():
        root.withdraw()  # 隐藏窗口到托盘
    root.protocol('WM_DELETE_WINDOW', on_closing)
    root.mainloop()
    db.close()

if __name__ == '__main__':
    main()