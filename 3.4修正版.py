import tkinter as tk
from tkinter import messagebox, ttk
import datetime
import pytz
import time

# Windows DPI 清晰化設定
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

class IndustrialCircularButton(tk.Canvas):
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

class HourglassIndustrialClock:
    def __init__(self, root):
        self.root = root
        self.root.title("工業計時儀 Pro V8.9 - 台北標準時")
        self.root.geometry("520x960")
        
        self.taipei_tz = pytz.timezone('Asia/Taipei')
        self.active_tasks = {"alarm": [], "timer": []}
        self.lap_records, self.timer_counter = [], 0
        self.stopwatch = {"running": False, "elapsed": 0.0, "start": 0.0}
        self.last_lap_split, self.current_mode = 0.0, "碼表"
        self._last_display_val = self._last_date_str = ""
        
        self.colors = {
            "casing": "#1e272e", "panel": "#d2dae2", "lcd_bg": "#000000", 
            "lcd_text": "#00d2d3", "lcd_active": "#26de81",
            "btn_go": "#26de81", "btn_go_act": "#20bf6b",
            "btn_stop": "#eb4d4b", "btn_stop_act": "#ff7675",
            "btn_lap": "#f7b731", "btn_lap_act": "#fa8231",
            "btn_reset": "#778ca3", "btn_reset_act": "#a5b1c2",
            "status_bg": "#2d3436"
        }

        self.root.configure(bg=self.colors["casing"])
        self._set_styles()
        self._build_ui()
        self._bind_events()
        self.switch_mode("碼表")
        self.update_master_loop()

    def _set_styles(self):
        style = ttk.Style()
        style.theme_use('default')
        style.configure("Timer.Horizontal.TProgressbar", troughcolor="#111111", background="#f7d794", thickness=10, borderwidth=0)
        style.configure("Alarm.Horizontal.TProgressbar", troughcolor="#111111", background="#26de81", thickness=10, borderwidth=0)
        style.configure("Treeview", background="#2d3436", foreground="white", fieldbackground="#2d3436", rowheight=30)
        style.configure("Treeview.Heading", background="#4b6584", foreground="white", font=("Microsoft JhengHei", 9, "bold"))

    def _build_ui(self):
        # 狀態欄與日期
        self.status_bar = tk.Frame(self.root, bg=self.colors["status_bg"], height=35)
        self.status_bar.pack(fill="x", side="top")
        self.status_lbl = tk.Label(self.status_bar, text="🕒 TAIPEI STD TIME", font=("Microsoft JhengHei", 9, "bold"), bg=self.colors["status_bg"], fg="#00d2d3")
        self.status_lbl.pack(side="left", padx=20)
        self.date_lbl = tk.Label(self.status_bar, text="", font=("Courier", 10, "bold"), bg=self.colors["status_bg"], fg="#ffffff")
        self.date_lbl.pack(side="right", padx=20)

        # 任務監控看板
        self.dash_frame = tk.Frame(self.root, bg="#111111", height=200)
        self.dash_frame.pack(fill="x", padx=20, pady=(15, 10))
        self.dash_frame.pack_propagate(False)
        tk.Label(self.dash_frame, text="TASK MONITORING", font=("Arial", 8, "bold"), bg="#111111", fg="#57606f").place(x=10, y=8)
        tk.Button(self.dash_frame, text="CLEAR ALL", font=("Arial", 8, "bold"), bg="#3d3d3d", fg="#ff7675", bd=0, command=self.clear_all_tasks, padx=10).place(x=390, y=6)

        self.task_canvas = tk.Canvas(self.dash_frame, bg="#111111", highlightthickness=0)
        self.task_scrollbar = tk.Scrollbar(self.dash_frame, orient="vertical", command=self.task_canvas.yview)
        self.task_container = tk.Frame(self.task_canvas, bg="#111111")
        self.task_canvas.create_window((0, 0), window=self.task_container, anchor="nw", width=380)
        self.task_canvas.configure(yscrollcommand=self.task_scrollbar.set, yscrollincrement=5)
        self.task_canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=(35,10))
        self.task_scrollbar.pack(side="right", fill="y")

        # LCD 主顯示
        self.lcd_frame = tk.Frame(self.root, bg=self.colors["lcd_bg"], bd=5, relief="ridge")
        self.lcd_frame.pack(pady=15, padx=30, fill="x")
        self.main_display = tk.Label(self.lcd_frame, text="00:00:00.00", font=("Courier", 48, "bold"), bg=self.colors["lcd_bg"], fg=self.colors["lcd_text"])
        self.main_display.pack(pady=30)

        # 模式選單
        self.mode_frame = tk.Frame(self.root, bg=self.colors["casing"])
        self.mode_frame.pack(pady=10)
        self.mode_btns = {}
        for m in ["鬧鐘", "碼表", "計時器"]:
            btn = tk.Button(self.mode_frame, text=m, width=14, font=("Microsoft JhengHei", 9, "bold"), command=lambda x=m: self.switch_mode(x), pady=5)
            btn.grid(row=0, column=len(self.mode_btns), padx=5)
            self.mode_btns[m] = btn

        # 操作面板
        self.op_panel = tk.Frame(self.root, bg=self.colors["panel"])
        self.op_panel.pack(pady=(10, 20), padx=25, fill="both", expand=True)
        self.content_frame = tk.Frame(self.op_panel, bg=self.colors["panel"])
        self.content_frame.pack(fill="both", expand=True, pady=15)

    def _bind_events(self):
        self.root.bind("<Return>", lambda e: self.add_task())
        self.root.bind("<space>", lambda e: self.toggle_stopwatch() if self.current_mode == "碼表" else None)
        self.status_bar.bind("<Double-1>", self.toggle_pin)
        self.task_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        delta = -1 if (event.num == 5 or event.delta < 0) else 1
        self.task_canvas.yview_scroll(-1 * delta, "units")

    def toggle_pin(self, event=None):
        is_top = not self.root.attributes("-topmost")
        self.root.attributes("-topmost", is_top)
        self.status_lbl.config(text="🕒 SYSTEM PINNED" if is_top else "🕒 TAIPEI STD TIME", fg="#f7b731" if is_top else "#00d2d3")

    def switch_mode(self, mode):
        self.current_mode = mode
        for widget in self.content_frame.winfo_children(): widget.destroy()
        for name, btn in self.mode_btns.items():
            btn.config(bg="#4b6584" if mode == name else "#d2dae2", fg="white" if mode == name else "black")
        
        if mode == "碼表": self._setup_stopwatch_ui()
        else: self._setup_alarm_timer_ui(mode)

    def _setup_stopwatch_ui(self):
        f = tk.Frame(self.content_frame, bg=self.colors["panel"]); f.pack(expand=True, fill="both")
        self.sw_btn = IndustrialCircularButton(f, text="START", color_normal=self.colors["btn_go"], color_active=self.colors["btn_go_act"], command=self.toggle_stopwatch)
        self.sw_btn.place(x=35, y=15)
        self.lap_btn = IndustrialCircularButton(f, text="LAP", color_normal=self.colors["btn_lap"], color_active=self.colors["btn_lap_act"], command=self.record_lap)
        self.lap_btn.place(x=35, y=105)
        IndustrialCircularButton(f, text="RESET", color_normal=self.colors["btn_reset"], color_active=self.colors["btn_reset_act"], command=self.reset_stopwatch).place(x=35, y=195)
        
        self.lap_tree = ttk.Treeview(f, columns=("Rank", "Lap", "Time", "Trend"), show='headings')
        for col, head, w in [("Rank", "排名", 50), ("Lap", "序號", 50), ("Time", "單圈耗時", 120), ("Trend", "進退步", 70)]:
            self.lap_tree.heading(col, text=head); self.lap_tree.column(col, width=w, anchor="center")
        self.lap_tree.place(x=145, y=5, width=310, height=270)
        for tag, color in [('up', '#26de81'), ('down', '#eb4d4b')]: self.lap_tree.tag_configure(tag, foreground=color)
        self.lap_tree.tag_configure('best', background='#f7b731', foreground='black')
        self._refresh_lap_display()
        self._update_sw_visuals()

    def _setup_alarm_timer_ui(self, mode):
        header = tk.Frame(self.content_frame, bg=self.colors["panel"]); header.pack(fill="x", pady=(5, 15))
        tk.Label(header, text=f"── 配置 {mode} ──", bg=self.colors["panel"], font=("Microsoft JhengHei", 11, "bold")).pack(side="left", padx=25)
        
        ctrl = tk.Frame(header, bg=self.colors["panel"]); ctrl.pack(side="right", padx=25)
        tk.Button(ctrl, text="清空", font=("Microsoft JhengHei", 9), bg="#778ca3", fg="white", bd=0, command=self.clear_inputs, padx=8).pack(side="right", padx=5)
        if mode == "鬧鐘":
            tk.Button(ctrl, text="🕒 同步", font=("Microsoft JhengHei", 9), bg="#4b6584", fg="white", bd=0, command=self.sync_time, padx=8).pack(side="right", padx=5)

        self.spins = []
        s_frame = tk.Frame(self.content_frame, bg=self.colors["panel"]); s_frame.pack(pady=15)
        for i, l in enumerate(["H", "M", "S"]):
            f = tk.Frame(s_frame, bg=self.colors["panel"]); f.grid(row=0, column=i, padx=12)
            sb = tk.Spinbox(f, from_=0, to=99, format="%02.0f", width=4, font=("Courier", 26, "bold"), justify="center", bd=0)
            sb.pack(pady=5); sb.delete(0, "end"); sb.insert(0, "00")
            sb.bind("<KeyRelease>", lambda e, idx=i: self._auto_tab_gentle(e, idx))
            sb.bind("<FocusIn>", lambda e: e.widget.selection_range(0, "end"))
            tk.Label(f, text=l, bg=self.colors["panel"], font=("Courier", 10, "bold")).pack()
            self.spins.append(sb)
            
        q_frame = tk.Frame(self.content_frame, bg=self.colors["panel"]); q_frame.pack(pady=10)
        for val in [1, 5, 10, 30]:
            tk.Button(q_frame, text=f"+{val}m", width=8, bg="#4b7cf3", fg="white", font=("Arial", 9, "bold"), command=lambda x=val: self.quick_add(x), pady=3).grid(row=0, column=[1,5,10,30].index(val), padx=4)
        IndustrialCircularButton(self.content_frame, text="EXECUTE", color_normal=self.colors["btn_go"], color_active=self.colors["btn_go_act"], command=self.add_task, radius=40).pack(pady=25)

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
        if self.stopwatch["running"]:
            self.stopwatch["start"] = time.time() - self.stopwatch["elapsed"]
            self.main_display.config(fg=self.colors["lcd_active"])
        else:
            self.main_display.config(fg=self.colors["lcd_text"])
        self._update_sw_visuals()

    def record_lap(self):
        if self.stopwatch["elapsed"] > 0:
            total = self.stopwatch["elapsed"]; dur = total - self.last_lap_split; self.last_lap_split = total
            trend, tag = "--", ""
            if self.lap_records:
                diff = dur - self.lap_records[-1]['dur']
                if diff < -0.01: trend, tag = f"▲{abs(diff):.2f}", "up"
                elif diff > 0.01: trend, tag = f"▼{diff:.2f}", "down"
                else: trend = "穩定"
            self.lap_records.append({'dur': dur, 'trend': trend, 'tag': tag})
            self._refresh_lap_display()

    def _refresh_lap_display(self):
        self.lap_tree.delete(*self.lap_tree.get_children())
        if not self.lap_records: return
        best_idx = min(range(len(self.lap_records)), key=lambda i: self.lap_records[i]['dur'])
        for i in reversed(range(len(self.lap_records))):
            rec = self.lap_records[i]
            tags = [rec['tag']] if rec['tag'] else []
            if i == best_idx and len(self.lap_records) > 1: tags.append('best')
            self.lap_tree.insert("", "end", values=(f"#{i+1}", f"L{i+1}", self.format_time_precision(rec['dur']), rec['trend']), tags=tuple(tags))

    def reset_stopwatch(self):
        self.stopwatch = {"running": False, "elapsed": 0.0, "start": 0.0}
        self.lap_records, self.last_lap_split = [], 0.0
        self.main_display.config(fg=self.colors["lcd_text"])
        self._update_sw_visuals(); self._refresh_lap_display()

    def _update_sw_visuals(self):
        if hasattr(self, 'sw_btn'):
            txt, clr = ("STOP", self.colors["btn_stop"]) if self.stopwatch["running"] else ("START", self.colors["btn_go"])
            self.sw_btn.config_visuals(txt, clr, clr)

    def clear_inputs(self):
        for sb in self.spins: sb.delete(0, "end"); sb.insert(0, "00")

    def sync_time(self):
        now = datetime.datetime.now(self.taipei_tz)
        for i, v in enumerate([now.hour, now.minute, now.second]):
            self.spins[i].delete(0, "end"); self.spins[i].insert(0, f"{v:02d}")

    def quick_add(self, delta):
        try:
            m_raw = int(self.spins[1].get() or 0) + delta
            h_raw = int(self.spins[0].get() or 0) + (m_raw // 60)
            # 確保快速加時不會爆表
            h, m = min(h_raw, 23), min(m_raw % 60, 59)
            self.spins[1].delete(0, "end"); self.spins[1].insert(0, f"{m:02d}")
            self.spins[0].delete(0, "end"); self.spins[0].insert(0, f"{h:02d}")
        except ValueError: pass

    def add_task(self):
        """核心優化：整合規格約束 (Clamping)"""
        try:
            # 讀取並約束數值：H <= 23, M <= 59, S <= 59
            h = max(0, min(int(self.spins[0].get() or 0), 23))
            m = max(0, min(int(self.spins[1].get() or 0), 59))
            s = max(0, min(int(self.spins[2].get() or 0), 59))

            # 同步填回介面 (視覺回饋)
            for i, val in enumerate([h, m, s]):
                self.spins[i].delete(0, "end"); self.spins[i].insert(0, f"{val:02d}")

            now_ts = time.time()
            if self.current_mode == "鬧鐘":
                target = datetime.datetime.now(self.taipei_tz).replace(hour=h, minute=m, second=s, microsecond=0)
                if target.timestamp() <= now_ts: target += datetime.timedelta(days=1)
                task = {"time_str": target.strftime("%H:%M:%S"), "loop": False, "start_ts": now_ts, "target_ts": target.timestamp(), "widgets": {}, "triggered": False}
                self.active_tasks["alarm"].append(task)
            else:
                sec = h*3600 + m*60 + s
                if sec <= 0: return
                self.timer_counter += 1
                task = {"id": self.timer_counter, "total": sec, "loop": False, "end": now_ts + sec, "widgets": {}, "triggered": False}
                self.active_tasks["timer"].append(task)
            
            self.refresh_dashboard()
            if self.spins: self.spins[0].focus_set()
        except (ValueError, tk.TclError): pass

    def refresh_dashboard(self):
        for w in self.task_container.winfo_children(): w.destroy()
        for ttype, tasks in self.active_tasks.items():
            for task in tasks: self._create_task_row(ttype, task)
        self.task_container.update_idletasks()
        self.task_canvas.config(scrollregion=self.task_canvas.bbox("all"))

    def _create_task_row(self, t_type, task):
        item = tk.Frame(self.task_container, bg="#1a1a1a"); item.pack(fill="x", pady=4, padx=8)
        loop_btn = tk.Button(item, text="🔁" if task["loop"] else "🔜", font=("Arial", 9), bg="#3d3d3d", fg="white", bd=0, width=4, command=lambda: self.toggle_loop(task))
        loop_btn.pack(side="left", padx=5)
        name = f"ALM {task['time_str']}" if t_type == "alarm" else f"T{task['id']} 00:00"
        color = "#26de81" if t_type == "alarm" else "#f7d794"
        lbl = tk.Label(item, text=name, font=("Courier", 10), bg="#1a1a1a", fg=color, width=12, anchor="w"); lbl.pack(side="left")
        style = "Alarm.Horizontal.TProgressbar" if t_type == "alarm" else "Timer.Horizontal.TProgressbar"
        pbar = ttk.Progressbar(item, orient="horizontal", length=130, style=style); pbar.pack(side="left", padx=10)
        task["widgets"] = {"pbar": pbar, "loop_btn": loop_btn, "label": lbl}
        tk.Button(item, text="X", font=("Arial", 8, "bold"), bg="#3d3d3d", fg="#ff7675", bd=0, command=lambda: self.remove_task(t_type, task), padx=10).pack(side="right", padx=5)

    def toggle_loop(self, task):
        task["loop"] = not task["loop"]
        if "loop_btn" in task["widgets"]: task["widgets"]["loop_btn"].config(text="🔁" if task["loop"] else "🔜")

    def remove_task(self, t_type, task):
        if task in self.active_tasks[t_type]:
            self.active_tasks[t_type].remove(task); self.refresh_dashboard()

    def clear_all_tasks(self):
        if messagebox.askyesno("CONFIRM", "是否清空所有監測任務？"):
            self.active_tasks = {"alarm": [], "timer": []}; self.refresh_dashboard()

    def notify_event(self, title, msg, color="#26de81"):
        top = tk.Toplevel(self.root)
        top.title(f"🕒 {title}"); top.geometry("320x160"); top.configure(bg=self.colors["casing"]); top.attributes("-topmost", True)
        tk.Label(top, text=f"🕒 {title}", font=("Arial", 14, "bold"), bg=self.colors["casing"], fg=color).pack(pady=15)
        tk.Label(top, text=msg, font=("Microsoft JhengHei", 10), bg=self.colors["casing"], fg="white").pack(pady=5)
        tk.Button(top, text="了解", bg="#4b6584", fg="white", font=("Microsoft JhengHei", 9, "bold"), command=top.destroy, padx=20, pady=5).pack(pady=15)

    def update_master_loop(self):
        now_dt, curr_ts = datetime.datetime.now(self.taipei_tz), time.time()
        
        # 狀態更新
        date_str = now_dt.strftime("%Y-%m-%d %a").upper()
        if date_str != self._last_date_str:
            self.date_lbl.config(text=date_str); self._last_date_str = date_str
        
        # 任務更新
        for ttype, tasks in self.active_tasks.items():
            for task in tasks[:]:
                rem = (task["target_ts"] if ttype == "alarm" else task["end"]) - curr_ts
                if task["widgets"]:
                    if ttype == "timer":
                        m, s = divmod(max(0, int(rem)), 60)
                        task["widgets"]["label"].config(text=f"T{task['id']} {m:02d}:{s:02d}")
                    p_val = (1 - rem / (task["target_ts"] - task["start_ts"])) if ttype == "alarm" else (rem / task["total"])
                    task["widgets"]["pbar"]['value'] = max(0, min(100, p_val * 100))
                
                if rem <= 0 and not task["triggered"]:
                    task["triggered"] = True
                    if task["loop"]:
                        if ttype == "alarm": task["start_ts"], task["target_ts"] = curr_ts, task["target_ts"] + 86400
                        else: task["end"] = curr_ts + task["total"]
                        task["triggered"] = False
                    else:
                        self.active_tasks[ttype].remove(task); self.refresh_dashboard()
                    self.notify_event("提醒", f"{task.get('time_str', '計時器')} 已到", color="#26de81" if ttype == "alarm" else "#f7b731")
        
        # LCD 更新
        new_val = self.format_time_precision(self.stopwatch["elapsed"]) if self.current_mode == "碼表" else now_dt.strftime("%H:%M:%S")
        if self.stopwatch["running"]: self.stopwatch["elapsed"] = time.time() - self.stopwatch["start"]
        if new_val != self._last_display_val:
            self.main_display.config(text=new_val); self._last_display_val = new_val

        self.root.after(20, self.update_master_loop)

    def format_time_precision(self, ts):
        m, s = divmod(ts, 60); h, m = divmod(m, 60)
        return f"{int(h):02}:{int(m):02}:{int(s):02}.{int((ts % 1)*100):02}"

if __name__ == "__main__":
    root = tk.Tk()
    app = HourglassIndustrialClock(root)
    root.mainloop()