import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import datetime
import pytz
import time
import threading
import winsound
import os
import shutil
import ctypes

# 引入 Pillow 來處理圖片縮放與填滿
try:
    from PIL import Image, ImageTk, ImageOps
except ImportError:
    messagebox.showerror("缺少套件", "請先在終端機執行: pip install pillow")

# Windows DPI 清晰化設定
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class CustomCircularButton(tk.Canvas):
    def __init__(self, parent, text, color_normal, color_active, command=None, radius=35, font=("Microsoft JhengHei", 10, "bold"), **kwargs):
        self.radius, self.color_normal, self.color_active = radius, color_normal, color_active
        self.command, self.is_pressed = command, False
        size = radius * 2
        super().__init__(parent, width=size, height=size, bg=parent['bg'], highlightthickness=0, cursor="hand2", **kwargs)
        self.create_oval(4, 4, size-4, size-4, fill=color_normal, outline="#2d3436", width=2, tags="btn_body")
        self.create_text(radius, radius, text=text, fill="white", font=font, tags="btn_text", width=radius*1.5, justify="center")
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)

    def config_visuals(self, text, color, active_color):
        self.color_normal, self.color_active = color, active_color
        self.itemconfigure("btn_body", fill=color)
        self.itemconfigure("btn_text", text=text)

    def _on_press(self, event):
        self.is_pressed = True
        self.itemconfigure("btn_body", fill=self.color_active)
        self.move("btn_text", 1, 1)

    def _on_release(self, event):
        if self.is_pressed:
            self.itemconfigure("btn_body", fill=self.color_normal)
            self.move("btn_text", -1, -1)
            self.is_pressed = False
            if self.command: self.command()

class TaipeiTimeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("台北時間 - 專注計時器")
        self.root.geometry("520x980")
        
        self.music_folder = "sounds"
        if not os.path.exists(self.music_folder): os.makedirs(self.music_folder)
        
        self.taipei_tz = pytz.timezone('Asia/Taipei')
        self.active_tasks = {"alarm": [], "timer": []}
        self.lap_records, self.timer_counter = [], 0
        self.stopwatch = {"running": False, "elapsed": 0.0, "start": 0.0}
        self.last_lap_split, self.current_mode = 0.0, "碼表"
        self._last_display_val = self._last_date_str = ""
        
        self.is_ringing = False
        self.is_previewing = False
        self.is_muted = tk.BooleanVar(value=False)
        self.selected_sound = tk.StringVar(value="預設嗶嗶聲")
        self.alarm_freq_var = tk.IntVar(value=60)
        
        # 背景圖相關 (新增 Pillow 支援)
        self.bg_image = None 
        self.original_bg_image = None # 儲存原始圖片以便縮放
        self.resize_timer = None # 用於防抖動縮放
        self.bg_label = tk.Label(self.root) # 最底層背景
        self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        self.dark_mode = True
        self.theme_color = "#00d2d3" 
        self.theme_color_act = "#26de81"
        self.colors = {}
        self._load_theme()

        self.root.configure(bg=self.colors["casing"])
        self._set_styles()
        self._build_ui()
        self._bind_events()
        self.switch_mode("碼表")
        self.update_master_loop()

    def _mci_work(self, command):
        ctypes.windll.winmm.mciSendStringW(command, None, 0, None)

    def _play_alarm_logic(self):
        if self.is_muted.get(): return
        selection = self.selected_sound.get()
        use_mci = False
        if selection != "預設嗶嗶聲":
            file_path = os.path.abspath(os.path.join(self.music_folder, selection))
            if os.path.exists(file_path):
                self._mci_work(f'open "{file_path}" type mpegvideo alias AlarmSound')
                self._mci_work('play AlarmSound repeat')
                use_mci = True
        while self.is_ringing:
            if use_mci: time.sleep(0.5)
            else:
                for _ in range(4):
                    if not self.is_ringing: break
                    freq = 400 + (self.alarm_freq_var.get() * 20)
                    try: winsound.Beep(freq, 120); time.sleep(0.08)
                    except: pass
                time.sleep(0.6)
        if use_mci:
            self._mci_work('stop AlarmSound'); self._mci_work('close AlarmSound')

    def start_alarm_sound(self):
        if not self.is_ringing:
            self.is_ringing = True
            threading.Thread(target=self._play_alarm_logic, daemon=True).start()

    def stop_alarm_sound(self):
        self.is_ringing = False

    def preview_sound(self):
        if self.is_ringing: return
        if self.is_previewing:
            self._mci_work('stop PreviewSound')
            self._mci_work('close PreviewSound')
            self.is_previewing = False
            if hasattr(self, 'preview_btn'): self.preview_btn.config(text="▶ 試聽", bg="#4b6584")
            return
        sel = self.selected_sound.get()
        if sel == "預設嗶嗶聲":
            winsound.Beep(600, 500)
        else:
            file_path = os.path.abspath(os.path.join(self.music_folder, sel))
            if os.path.exists(file_path):
                self._mci_work('close PreviewSound')
                self._mci_work(f'open "{file_path}" type mpegvideo alias PreviewSound')
                self._mci_work('play PreviewSound')
                self.is_previewing = True
                if hasattr(self, 'preview_btn'): self.preview_btn.config(text="■ 停止", bg="#eb4d4b")
            else:
                messagebox.showerror("錯誤", "找不到音效檔案")

    def _load_theme(self):
        if self.dark_mode:
            self.colors.update({
                "casing": "#1e272e", "panel": "#2d3436", "lcd_bg": "#000000", 
                "lcd_text": self.theme_color, "lcd_active": self.theme_color_act,
                "btn_go": "#26de81", "btn_go_act": "#20bf6b",
                "btn_stop": "#eb4d4b", "btn_stop_act": "#ff7675",
                "btn_lap": "#f7b731", "btn_lap_act": "#fa8231",
                "btn_reset": "#778ca3", "btn_reset_act": "#a5b1c2",
                "status_bg": "#2d3436", "dash_bg": "#111111", "text_main": "#ffffff"
            })
        else:
            self.colors.update({
                "casing": "#dfe6e9", "panel": "#b2bec3", "lcd_bg": "#95a5a6", 
                "lcd_text": "#2d3436", "lcd_active": self.theme_color,
                "btn_go": "#00b894", "btn_go_act": "#55efc4",
                "btn_stop": "#d63031", "btn_stop_act": "#ff7675",
                "btn_lap": "#e17055", "btn_lap_act": "#fab1a0",
                "btn_reset": "#636e72", "btn_reset_act": "#b2bec3",
                "status_bg": "#bdc3c7", "dash_bg": "#ffffff", "text_main": "#2d3436"
            })

    def _set_styles(self):
        style = ttk.Style()
        style.theme_use('default')
        trough_clr = "#111111" if self.dark_mode else "#dfe6e9"
        style.configure("Timer.Horizontal.TProgressbar", troughcolor=trough_clr, background="#f7d794", thickness=10, borderwidth=0)
        style.configure("Alarm.Horizontal.TProgressbar", troughcolor=trough_clr, background=self.theme_color_act, thickness=10, borderwidth=0)
        style.configure("Treeview", background="#2d3436" if self.dark_mode else "#ecf0f1", 
                        foreground="white" if self.dark_mode else "black", 
                        fieldbackground="#2d3436" if self.dark_mode else "#ecf0f1", rowheight=30)
        style.configure("Treeview.Heading", background="#4b6584", foreground="white", font=("Microsoft JhengHei", 9, "bold"))

    def _build_ui(self):
        self.status_bar = tk.Frame(self.root, bg=self.colors["status_bg"], height=35)
        self.status_bar.pack(fill="x", side="top")
        self.status_lbl = tk.Label(self.status_bar, text="🕒 台北時間", font=("Microsoft JhengHei", 9, "bold"), bg=self.colors["status_bg"], fg=self.colors["lcd_text"])
        self.status_lbl.pack(side="left", padx=20)
        self.date_lbl = tk.Label(self.status_bar, text="", font=("Courier", 10, "bold"), bg=self.colors["status_bg"], fg=self.colors["text_main"])
        self.date_lbl.pack(side="right", padx=20)

        self.dash_frame = tk.Frame(self.root, bg=self.colors["dash_bg"], height=200)
        self.dash_frame.pack(fill="x", padx=20, pady=(15, 10))
        self.dash_frame.pack_propagate(False)
        tk.Label(self.dash_frame, text="任務狀態監控", font=("Microsoft JhengHei", 8, "bold"), bg=self.colors["dash_bg"], fg="#57606f").place(x=10, y=8)
        tk.Button(self.dash_frame, text="全部清空", font=("Microsoft JhengHei", 8, "bold"), bg="#3d3d3d", fg="#ff7675", bd=0, command=self.clear_all_tasks, padx=10).place(x=390, y=6)

        self.task_canvas = tk.Canvas(self.dash_frame, bg=self.colors["dash_bg"], highlightthickness=0)
        self.task_scrollbar = tk.Scrollbar(self.dash_frame, orient="vertical", command=self.task_canvas.yview)
        self.task_container = tk.Frame(self.task_canvas, bg=self.colors["dash_bg"])
        self.task_canvas.create_window((0, 0), window=self.task_container, anchor="nw", width=380)
        self.task_canvas.configure(yscrollcommand=self.task_scrollbar.set, yscrollincrement=5)
        self.task_canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=(35,10))
        self.task_scrollbar.pack(side="right", fill="y")

        self.lcd_frame = tk.Frame(self.root, bg=self.colors["lcd_bg"], bd=5, relief="ridge")
        self.lcd_frame.pack(pady=15, padx=30, fill="x")
        self.main_display = tk.Label(self.lcd_frame, text="00:00:00.00", font=("Courier", 48, "bold"), bg=self.colors["lcd_bg"], fg=self.colors["lcd_text"])
        self.main_display.pack(pady=30)

        self.mode_frame = tk.Frame(self.root, bg=self.colors["casing"])
        self.mode_frame.pack(pady=10)
        self.mode_btns = {}
        for m in ["鬧鐘", "碼表", "計時器"]:
            btn = tk.Button(self.mode_frame, text=m, width=8, font=("Microsoft JhengHei", 9, "bold"), command=lambda x=m: self.switch_mode(x), pady=5)
            btn.grid(row=0, column=len(self.mode_btns), padx=2); self.mode_btns[m] = btn
        
        self.style_btn = tk.Button(self.mode_frame, text="🎨", width=4, bg="#636e72", fg="white", font=("Microsoft JhengHei", 9), command=self.show_style_picker, pady=5)
        self.style_btn.grid(row=0, column=3, padx=2)
        self.theme_btn = tk.Button(self.mode_frame, text="🌙 夜間", width=7, bg="#4b6584", fg="white", font=("Microsoft JhengHei", 9, "bold"), command=self.toggle_theme, pady=5)
        self.theme_btn.grid(row=0, column=4, padx=2)
        
        # 背景按鈕
        self.bg_btn = tk.Button(self.mode_frame, text="🖼️ 背景", width=7, bg="#778ca3", fg="white", font=("Microsoft JhengHei", 9, "bold"), command=self.import_background, pady=5)
        self.bg_btn.grid(row=0, column=5, padx=2)

        self.op_panel = tk.Frame(self.root, bg=self.colors["panel"])
        self.op_panel.pack(pady=(10, 20), padx=25, fill="both", expand=True)
        self.content_frame = tk.Frame(self.op_panel, bg=self.colors["panel"])
        self.content_frame.pack(fill="both", expand=True, pady=15)

    # --- 新增/修改：背景處理邏輯 ---
    def import_background(self):
        """匯入圖片並設定為自動縮放背景"""
        f_path = filedialog.askopenfilename(filetypes=[("圖片檔案", "*.png *.jpg *.jpeg *.gif *.bmp")])
        if f_path:
            try:
                # 使用 Pillow 讀取原圖
                self.original_bg_image = Image.open(f_path)
                self._resize_background() # 呼叫縮放函式
                messagebox.showinfo("成功", "背景已更新並會自動填滿視窗！")
            except Exception as e:
                messagebox.showerror("錯誤", f"無法載入圖片：\n{e}\n提示：請確保已安裝 Pillow 模組")

    def _resize_background(self):
        """處理背景圖縮放與填滿"""
        if not self.original_bg_image:
            return
            
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        
        # 避免視窗還沒渲染完就縮放導致錯誤
        if w < 10 or h < 10:
            w, h = 520, 980
            
        try:
            # ImageOps.fit 可以等比例縮放並裁切掉多餘部分，完美填滿視窗 (類似 CSS object-fit: cover)
            resized_image = ImageOps.fit(self.original_bg_image, (w, h), Image.Resampling.LANCZOS)
            self.bg_image = ImageTk.PhotoImage(resized_image)
            self.bg_label.config(image=self.bg_image)
        except Exception as e:
            print(f"背景縮放錯誤: {e}")

    def _on_window_resize(self, event):
        """視窗大小改變時觸發"""
        if event.widget == self.root:
            # 使用防抖動(Debounce)，避免拖拉視窗時卡頓
            if self.resize_timer:
                self.root.after_cancel(self.resize_timer)
            self.resize_timer = self.root.after(100, self._resize_background)
    # -------------------------------

    def show_style_picker(self):
        picker = tk.Toplevel(self.root)
        picker.title("主題色")
        picker.geometry("200x280")
        picker.attributes("-topmost", True)
        picker.configure(bg=self.colors["casing"])
        
        styles = [
            ("翡翠綠", "#00d2d3", "#26de81"),
            ("寶石藍", "#0984e3", "#74b9ff"),
            ("夕陽紅", "#d63031", "#ff7675"),
            ("魅惑紫", "#6c5ce7", "#a29bfe"),
            ("琥珀金", "#f39c12", "#f1c40f")
        ]
        for name, c1, c2 in styles:
            tk.Button(picker, text=name, bg=c1, fg="white", font=("Microsoft JhengHei", 9), width=15,
                      command=lambda a=c1, b=c2: self.apply_style(a, b, picker)).pack(pady=5)

    def apply_style(self, color, active_color, window):
        self.theme_color = color
        self.theme_color_act = active_color
        window.destroy()
        self._load_theme()
        self._set_styles()
        self.refresh_ui_colors()

    def refresh_ui_colors(self):
        # 如果沒有背景圖，才刷 root 背景色
        if not self.bg_image:
            self.root.configure(bg=self.colors["casing"])
        self.status_bar.configure(bg=self.colors["status_bg"])
        self.status_lbl.configure(bg=self.colors["status_bg"], fg=self.colors["lcd_text"])
        self.date_lbl.configure(bg=self.colors["status_bg"], fg=self.colors["text_main"])
        self.dash_frame.configure(bg=self.colors["dash_bg"])
        self.lcd_frame.configure(bg=self.colors["lcd_bg"])
        self.main_display.configure(bg=self.colors["lcd_bg"], fg=self.colors["lcd_text"])
        self.op_panel.configure(bg=self.colors["panel"])
        self.content_frame.configure(bg=self.colors["panel"])
        self.mode_frame.configure(bg=self.colors["casing"])
        self.switch_mode(self.current_mode)
        self.refresh_dashboard()

    def _bind_events(self):
        self.root.bind("<Return>", lambda e: self.add_task())
        self.root.bind("<space>", lambda e: self.toggle_stopwatch() if self.current_mode == "碼表" else None)
        self.status_bar.bind("<Double-1>", self.toggle_pin)
        self.task_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        # 新增：綁定視窗縮放事件，讓背景圖自動適應
        self.root.bind("<Configure>", self._on_window_resize)

    def _on_mousewheel(self, event):
        delta = -1 if (event.num == 5 or event.delta < 0) else 1
        self.task_canvas.yview_scroll(-1 * delta, "units")

    def toggle_pin(self, event=None):
        is_top = not self.root.attributes("-topmost")
        self.root.attributes("-topmost", is_top)
        self.status_lbl.config(text="🕒 台北時間 (置頂)" if is_top else "🕒 台北時間", fg="#f7b731" if is_top else self.colors["lcd_text"])

    def switch_mode(self, mode):
        self.current_mode = mode
        if self.is_previewing: self.preview_sound() 
        for widget in self.content_frame.winfo_children(): widget.destroy()
        for name, btn in self.mode_btns.items():
            btn.config(bg="#4b6584" if mode == name else self.colors["panel"], fg="white" if mode == name else ("white" if self.dark_mode else "black"))
        if mode == "碼表": self._setup_stopwatch_ui()
        else: self._setup_alarm_timer_ui(mode)

    def _setup_stopwatch_ui(self):
        f = tk.Frame(self.content_frame, bg=self.colors["panel"]); f.pack(expand=True, fill="both")
        self.sw_btn = CustomCircularButton(f, text="開始", color_normal=self.colors["btn_go"], color_active=self.colors["btn_go_act"], command=self.toggle_stopwatch)
        self.sw_btn.place(x=35, y=15)
        self.lap_btn = CustomCircularButton(f, text="計圈", color_normal=self.colors["btn_lap"], color_active=self.colors["btn_lap_act"], command=self.record_lap)
        self.lap_btn.place(x=35, y=105)
        CustomCircularButton(f, text="重設", color_normal=self.colors["btn_reset"], color_active=self.colors["btn_reset_act"], command=self.reset_stopwatch).place(x=35, y=195)
        self.lap_tree = ttk.Treeview(f, columns=("Rank", "Lap", "Time", "Trend"), show='headings')
        for col, head, w in [("Rank", "排名", 50), ("Lap", "序號", 50), ("Time", "單圈耗時", 120), ("Trend", "進退步", 70)]:
            self.lap_tree.heading(col, text=head); self.lap_tree.column(col, width=w, anchor="center")
        self.lap_tree.place(x=145, y=5, width=310, height=270)
        for tag, color in [('up', '#26de81'), ('down', '#eb4d4b')]: self.lap_tree.tag_configure(tag, foreground=color)
        self.lap_tree.tag_configure('best', background='#f7b731', foreground='black')
        self._refresh_lap_display(); self._update_sw_visuals()

    def _setup_alarm_timer_ui(self, mode):
        header = tk.Frame(self.content_frame, bg=self.colors["panel"]); header.pack(fill="x", pady=(5, 10))
        tk.Label(header, text=f"── {mode} ──", bg=self.colors["panel"], fg=self.colors["text_main"], font=("Microsoft JhengHei", 11, "bold")).pack(side="left", padx=25)
        ctrl = tk.Frame(header, bg=self.colors["panel"]); ctrl.pack(side="right", padx=25)
        tk.Button(ctrl, text="重設", font=("Microsoft JhengHei", 9), bg="#778ca3", fg="white", bd=0, command=self.clear_inputs, padx=8).pack(side="right", padx=5)
        if mode == "鬧鐘":
            tk.Button(ctrl, text="🕒 同步", font=("Microsoft JhengHei", 9), bg="#4b6584", fg="white", bd=0, command=self.sync_time, padx=8).pack(side="right", padx=5)

        sound_area = tk.LabelFrame(self.content_frame, text=" 提示音 ", bg=self.colors["panel"], fg=self.colors["text_main"], font=("Microsoft JhengHei", 9, "bold"))
        sound_area.pack(fill="x", padx=25, pady=5)
        file_list = ["預設嗶嗶聲"] + [f for f in os.listdir(self.music_folder) if f.endswith(('.mp3', '.wav'))]
        self.sound_combo = ttk.Combobox(sound_area, textvariable=self.selected_sound, values=file_list, state="readonly", width=18)
        self.sound_combo.pack(side="left", padx=10, pady=10)
        btn_box = tk.Frame(sound_area, bg=self.colors["panel"]); btn_box.pack(side="right", padx=10)
        self.preview_btn = tk.Button(btn_box, text="▶ 試聽", font=("Arial", 8), bg="#4b6584", fg="white", command=self.preview_sound)
        self.preview_btn.pack(side="left", padx=2)
        tk.Button(btn_box, text="➕", font=("Arial", 8), bg="#26de81", command=self.import_audio).pack(side="left", padx=2)
        tk.Button(btn_box, text="🗑️", font=("Arial", 8), bg="#eb4d4b", fg="white", command=self.delete_audio).pack(side="left", padx=2)
        tk.Checkbutton(btn_box, text="靜音", variable=self.is_muted, bg=self.colors["panel"], fg=self.colors["text_main"], selectcolor="#333" if self.dark_mode else "white", font=("Microsoft JhengHei", 8)).pack(side="left", padx=5)

        self.spins = []
        h_limit = 99 if mode == "計時器" else 23
        s_frame = tk.Frame(self.content_frame, bg=self.colors["panel"]); s_frame.pack(pady=10)
        for i, (l, limit) in enumerate([("時", h_limit), ("分", 59), ("秒", 59)]):
            f = tk.Frame(s_frame, bg=self.colors["panel"]); f.grid(row=0, column=i, padx=12)
            sb = tk.Spinbox(f, from_=0, to=limit, format="%02.0f", width=4, font=("Courier", 26, "bold"), justify="center", bd=0)
            sb.pack(pady=5); sb.delete(0, "end"); sb.insert(0, "00")
            sb.bind("<FocusOut>", lambda e, lim=limit: self._normalize_spin_input(e, lim))
            sb.bind("<KeyRelease>", lambda e, idx=i: self._auto_tab_gentle(e, idx))
            sb.bind("<FocusIn>", lambda e: e.widget.selection_range(0, "end"))
            tk.Label(f, text=l, bg=self.colors["panel"], fg=self.colors["text_main"], font=("Microsoft JhengHei", 10, "bold")).pack()
            self.spins.append(sb)
            
        q_frame = tk.Frame(self.content_frame, bg=self.colors["panel"]); q_frame.pack(pady=5)
        btn_clr = "#34495e" if mode == "鬧鐘" else "#4b7cf3"
        for val in ([5, 10, 30] if mode == "鬧鐘" else [1, 5, 10, 30]):
            tk.Button(q_frame, text=f"+{val}分", width=8, bg=btn_clr, fg="white", font=("Microsoft JhengHei", 9, "bold"), command=lambda x=val: self.quick_add(x), pady=3).pack(side="left", padx=4)
        CustomCircularButton(self.content_frame, text="執行", color_normal=self.colors["btn_go"], color_active=self.colors["btn_go_act"], command=self.add_task, radius=40).pack(pady=15)

    def _normalize_spin_input(self, event, limit):
        widget = event.widget
        try:
            val = int(widget.get() or 0)
            val = max(0, min(val, limit))
            widget.delete(0, "end"); widget.insert(0, f"{val:02d}")
        except:
            widget.delete(0, "end"); widget.insert(0, "00")

    def import_audio(self):
        f_path = filedialog.askopenfilename(filetypes=[("音訊檔案", "*.mp3 *.wav")])
        if f_path:
            f_name = os.path.basename(f_path)
            shutil.copy(f_path, os.path.join(self.music_folder, f_name))
            self._refresh_sound_list(); messagebox.showinfo("成功", f"已匯入：{f_name}")

    def delete_audio(self):
        sel = self.selected_sound.get()
        if sel == "預設嗶嗶聲": return
        if messagebox.askyesno("刪除", f"確定刪除 {sel}？"):
            try: os.remove(os.path.join(self.music_folder, sel))
            except: pass
            self.selected_sound.set("預設嗶嗶聲"); self._refresh_sound_list()

    def _refresh_sound_list(self):
        file_list = ["預設嗶嗶聲"] + [f for f in os.listdir(self.music_folder) if f.endswith(('.mp3', '.wav'))]
        self.sound_combo['values'] = file_list

    def quick_add(self, delta_minutes):
        try:
            h = int(self.spins[0].get() or 0); m = int(self.spins[1].get() or 0); s = int(self.spins[2].get() or 0)
            if self.current_mode == "計時器":
                total = min((h * 3600) + (m * 60) + s + (delta_minutes * 60), (99*3600+59*60+59))
                new_h, rem = divmod(total, 3600); new_m, new_s = divmod(rem, 60)
            else:
                total_m = (h * 60) + m + delta_minutes
                new_h, new_m = divmod(total_m, 60); new_h %= 24; new_s = s
            for i, v in enumerate([new_h, new_m, new_s]):
                self.spins[i].delete(0, "end"); self.spins[i].insert(0, f"{int(v):02d}")
        except: pass

    def _auto_tab_gentle(self, event, idx):
        if event.keysym == "BackSpace" and not self.spins[idx].get() and idx > 0:
            self.spins[idx-1].focus_set(); self.spins[idx-1].selection_range(0, "end")
        elif event.char.isdigit() and len(self.spins[idx].get()) >= 2 and idx < 2:
            self.root.after(100, lambda: self._execute_jump(idx))

    def _execute_jump(self, idx):
        if self.root.focus_get() == self.spins[idx]:
            self.spins[idx+1].focus_set(); self.spins[idx+1].selection_range(0, "end")

    def toggle_stopwatch(self):
        self.stopwatch["running"] = not self.stopwatch["running"]
        self.main_display.config(fg=self.colors["lcd_active"] if self.stopwatch["running"] else self.colors["lcd_text"])
        if self.stopwatch["running"]: self.stopwatch["start"] = time.time() - self.stopwatch["elapsed"]
        self._update_sw_visuals()

    def record_lap(self):
        if self.stopwatch["elapsed"] > 0:
            total = self.stopwatch["elapsed"]; dur = total - self.last_lap_split; self.last_lap_split = total
            trend, tag = "--", ""
            if self.lap_records:
                diff = dur - self.lap_records[-1]['dur']
                if diff < -0.01: trend, tag = f"▲{abs(diff):.2f}", "up"
                elif diff > 0.01: trend, tag = f"▼{diff:.2f}", "down"
            self.lap_records.append({'dur': dur, 'trend': trend, 'tag': tag}); self._refresh_lap_display()

    def _refresh_lap_display(self):
        self.lap_tree.delete(*self.lap_tree.get_children())
        if not self.lap_records: return
        best_idx = min(range(len(self.lap_records)), key=lambda i: self.lap_records[i]['dur'])
        for i in reversed(range(len(self.lap_records))):
            rec = self.lap_records[i]; tags = [rec['tag']] if rec['tag'] else []
            if i == best_idx and len(self.lap_records) > 1: tags.append('best')
            self.lap_tree.insert("", "end", values=(f"#{i+1}", f"L{i+1}", self.format_time_precision(rec['dur']), rec['trend']), tags=tuple(tags))

    def reset_stopwatch(self):
        self.stopwatch = {"running": False, "elapsed": 0.0, "start": 0.0}
        self.lap_records, self.last_lap_split = [], 0.0
        self.main_display.config(fg=self.colors["lcd_text"]); self._update_sw_visuals(); self._refresh_lap_display()

    def _update_sw_visuals(self):
        if hasattr(self, 'sw_btn'):
            txt, clr = ("停止", self.colors["btn_stop"]) if self.stopwatch["running"] else ("開始", self.colors["btn_go"])
            self.sw_btn.config_visuals(txt, clr, clr)

    def clear_inputs(self):
        for sb in self.spins: sb.delete(0, "end"); sb.insert(0, "00")

    def sync_time(self):
        now = datetime.datetime.now(self.taipei_tz)
        for i, v in enumerate([now.hour, now.minute, now.second]):
            self.spins[i].delete(0, "end"); self.spins[i].insert(0, f"{v:02d}")

    def add_task(self):
        try:
            h_lim = 99 if self.current_mode == "計時器" else 23
            h = max(0, min(int(self.spins[0].get() or 0), h_lim))
            m = max(0, min(int(self.spins[1].get() or 0), 59))
            s = max(0, min(int(self.spins[2].get() or 0), 59))
            for i, val in enumerate([h, m, s]):
                self.spins[i].delete(0, "end"); self.spins[i].insert(0, f"{val:02d}")
            now_ts = time.time()
            if self.current_mode == "鬧鐘":
                target = datetime.datetime.now(self.taipei_tz).replace(hour=h, minute=m, second=s, microsecond=0)
                if target.timestamp() <= now_ts: target += datetime.timedelta(days=1)
                self.active_tasks["alarm"].append({"time_str": target.strftime("%H:%M:%S"), "loop": False, "start_ts": now_ts, "target_ts": target.timestamp(), "widgets": {}, "triggered": False})
            else:
                sec = h*3600 + m*60 + s
                if sec <= 0: return
                self.timer_counter += 1
                self.active_tasks["timer"].append({"id": self.timer_counter, "total": sec, "loop": False, "end": now_ts + sec, "widgets": {}, "triggered": False})
            self.refresh_dashboard()
        except: pass

    def refresh_dashboard(self):
        for w in self.task_container.winfo_children(): w.destroy()
        for ttype, tasks in self.active_tasks.items():
            for task in tasks: self._create_task_row(ttype, task)
        self.task_canvas.config(scrollregion=self.task_canvas.bbox("all"))

    def _create_task_row(self, t_type, task):
        row_bg = "#262626" if self.dark_mode else "#ecf0f1"
        item = tk.Frame(self.task_container, bg=row_bg); item.pack(fill="x", pady=4, padx=8)
        loop_btn = tk.Button(item, text="🔁" if task["loop"] else "🔜", font=("Arial", 9), bg="#3d3d3d", fg="white", bd=0, width=4, command=lambda: self.toggle_loop(task))
        loop_btn.pack(side="left", padx=5)
        name = f"鬧鐘 {task['time_str']}" if t_type == "alarm" else f"計時 T{task['id']}"
        clr = self.theme_color_act if t_type == "alarm" else ("#f7d794" if self.dark_mode else "#d35400")
        lbl = tk.Label(item, text=name, font=("Microsoft JhengHei", 9), bg=row_bg, fg=clr, width=14, anchor="w"); lbl.pack(side="left")
        pbar = ttk.Progressbar(item, orient="horizontal", length=110, style="Alarm.Horizontal.TProgressbar" if t_type == "alarm" else "Timer.Horizontal.TProgressbar"); pbar.pack(side="left", padx=5)
        task["widgets"] = {"pbar": pbar, "loop_btn": loop_btn, "label": lbl}
        tk.Button(item, text="✕", font=("Arial", 9, "bold"), bg="#3d3d3d", fg="#ff7675", bd=0, command=lambda: self.remove_task(t_type, task), padx=10).pack(side="right", padx=2)

    def toggle_loop(self, task):
        task["loop"] = not task["loop"]
        if "loop_btn" in task["widgets"]: task["widgets"]["loop_btn"].config(text="🔁" if task["loop"] else "🔜")

    def remove_task(self, t_type, task):
        if task in self.active_tasks[t_type]: self.active_tasks[t_type].remove(task); self.refresh_dashboard()

    def clear_all_tasks(self):
        if messagebox.askyesno("確認", "是否清空所有監測中的任務？"):
            self.active_tasks = {"alarm": [], "timer": []}; self.refresh_dashboard()

    def notify_event(self, title, msg, color=None):
        if color is None: color = self.theme_color_act
        self.start_alarm_sound()
        top = tk.Toplevel(self.root); top.title(f"🕒 {title}"); top.geometry("320x160"); top.attributes("-topmost", True); top.configure(bg=self.colors["casing"])
        def on_close(): self.stop_alarm_sound(); top.destroy()
        tk.Label(top, text=f"🕒 {title}", font=("Microsoft JhengHei", 14, "bold"), bg=self.colors["casing"], fg=color).pack(pady=15)
        tk.Label(top, text=msg, font=("Microsoft JhengHei", 10), bg=self.colors["casing"], fg=self.colors["text_main"]).pack(pady=5)
        tk.Button(top, text="了解", bg="#4b6584", fg="white", font=("Microsoft JhengHei", 9, "bold"), command=on_close, padx=20, pady=5).pack(pady=15)
        top.protocol("WM_DELETE_WINDOW", on_close)

    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self._load_theme()
        self._set_styles()
        self.theme_btn.config(text="🌙 夜間" if self.dark_mode else "☀️ 日間", bg="#4b6584" if self.dark_mode else "#778ca3")
        self.refresh_ui_colors()

    def update_master_loop(self):
        now_dt, curr_ts = datetime.datetime.now(self.taipei_tz), time.time()
        date_str = now_dt.strftime("%Y-%m-%d %a").upper()
        if date_str != self._last_date_str: self.date_lbl.config(text=date_str); self._last_date_str = date_str
        for ttype in ["alarm", "timer"]:
            for task in self.active_tasks[ttype][:]:
                rem = (task["target_ts"] if ttype == "alarm" else task["end"]) - curr_ts
                if "widgets" in task and task["widgets"]:
                    if ttype == "timer":
                        h, rem_m = divmod(max(0, int(rem)), 3600); m, s = divmod(rem_m, 60)
                        task["widgets"]["label"].config(text=f"T{task['id']} {h:02d}:{m:02d}:{s:02d}" if h > 0 else f"T{task['id']} {m:02d}:{s:02d}")
                    start_ts = task.get("start_ts", curr_ts - task.get("total", 0))
                    total_dur = (task["target_ts"] - start_ts) if ttype == "alarm" else task["total"]
                    if total_dur > 0:
                        p_val = (1 - rem / total_dur) if ttype == "alarm" else (rem / task["total"])
                        task["widgets"]["pbar"]['value'] = max(0, min(100, p_val * 100))
                if rem <= 0 and not task.get("triggered", False):
                    task["triggered"] = True
                    if task["loop"]:
                        if ttype == "alarm": task["start_ts"], task["target_ts"] = curr_ts, task["target_ts"] + 86400
                        else: task["end"] = curr_ts + task["total"]
                        task["triggered"] = False
                    else: self.remove_task(ttype, task)
                    self.notify_event("提醒", f"{task.get('time_str', '計時器')} 已到")
        new_val = self.format_time_precision(self.stopwatch["elapsed"]) if self.current_mode == "碼表" else now_dt.strftime("%H:%M:%S")
        if self.stopwatch["running"]: self.stopwatch["elapsed"] = time.time() - self.stopwatch["start"]
        if new_val != self._last_display_val: self.main_display.config(text=new_val); self._last_display_val = new_val
        self.root.after(20, self.update_master_loop)

    def format_time_precision(self, ts):
        m, s = divmod(ts, 60); h, m = divmod(m, 60)
        return f"{int(h):02}:{int(m):02}:{int(s):02}.{int((ts % 1)*100):02}"

if __name__ == "__main__":
    root = tk.Tk(); app = TaipeiTimeApp(root); root.mainloop()