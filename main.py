import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import sqlite3
import threading
import time
import datetime
import win32print
import platform
import calendar

# =========================================================
# [전역 설정] 기본값 (DB에서 로드됨)
# =========================================================
current_settings = {
    "cost_bw_a4": 50,      # A4 흑백 단가
    "cost_color_a4": 200,  # A4 컬러 단가
    "mult_a3_bw": 2.0,     # A3 흑백 배수
    "mult_a3_color": 2.0   # A3 컬러 배수
}

# =========================================================
# 1. 데이터베이스 및 설정 관리
# =========================================================
def init_db():
    conn = sqlite3.connect("print_log.db")
    cursor = conn.cursor()
    
    # 로그 테이블
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            job_id INTEGER,
            printer_name TEXT,
            computer_name TEXT,
            user_name TEXT,
            document_name TEXT,
            pages INTEGER,
            paper_size TEXT,
            is_color INTEGER,
            unit_cost INTEGER,
            cost INTEGER,
            print_time TEXT,
            status TEXT
        )
    ''')
    
    # 설정 테이블
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value REAL
        )
    ''')
    
    # 기본 설정값 초기화
    defaults = {
        "cost_bw_a4": 50,
        "cost_color_a4": 200,
        "mult_a3_bw": 2.0,
        "mult_a3_color": 2.0
    }
    
    for key, val in defaults.items():
        cursor.execute("SELECT value FROM settings WHERE key=?", (key,))
        if not cursor.fetchone():
            cursor.execute("INSERT INTO settings (key, value) VALUES (?, ?)", (key, val))
    
    conn.commit()
    conn.close()
    load_settings()

def load_settings():
    global current_settings
    conn = sqlite3.connect("print_log.db")
    cursor = conn.cursor()
    cursor.execute("SELECT key, value FROM settings")
    rows = cursor.fetchall()
    for key, value in rows:
        current_settings[key] = value
    conn.close()

def update_setting(key, value):
    conn = sqlite3.connect("print_log.db")
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, value))
    conn.commit()
    conn.close()
    # 설정 저장 후 로드까지 수행
    load_settings()

# [NEW] 모든 과거 데이터의 비용을 현재 설정으로 재계산하는 함수
def recalculate_db_costs():
    conn = sqlite3.connect("print_log.db")
    cursor = conn.cursor()
    
    # 1. 모든 로그 가져오기
    cursor.execute("SELECT id, pages, paper_size, is_color FROM logs")
    rows = cursor.fetchall()
    
    # 2. 현재 설정값 준비
    c_bw_a4 = current_settings['cost_bw_a4']
    c_col_a4 = current_settings['cost_color_a4']
    m_a3_bw = current_settings['mult_a3_bw']
    m_a3_col = current_settings['mult_a3_color']
    
    updates = []
    
    for row in rows:
        rid, pages, size, is_color = row
        
        # 비용 계산 로직 (monitor_loop와 동일하게 적용)
        base_cost = c_col_a4 if is_color else c_bw_a4
        multiplier = 1.0
        
        if size == "A3":
            multiplier = m_a3_col if is_color else m_a3_bw
            
        unit_cost = int(base_cost * multiplier)
        cost = int(pages * unit_cost)
        
        # 업데이트할 데이터 (새 단가, 새 총비용, ID)
        updates.append((unit_cost, cost, rid))
        
    # 3. DB 일괄 업데이트
    if updates:
        cursor.executemany("UPDATE logs SET unit_cost=?, cost=? WHERE id=?", updates)
        conn.commit()
        print(f"[시스템] {len(updates)}개의 과거 로그 재계산 완료.")
    
    conn.close()

# =========================================================
# 2. 감시 엔진
# =========================================================
def monitor_loop(app_instance):
    processed_jobs = set()
    print("[엔진] 감시 시작... (A3/A4, 컬러/흑백 구분)")
    
    my_computer_name = platform.node()

    while True:
        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            printers = win32print.EnumPrinters(flags)

            for printer in printers:
                printer_name = printer[2]
                phandle = None
                try:
                    phandle = win32print.OpenPrinter(printer_name)
                    jobs = win32print.EnumJobs(phandle, 0, 100, 2)

                    for job in jobs:
                        p_job_id = job['JobId']
                        unique_id = f"{printer_name}_{p_job_id}"

                        if unique_id in processed_jobs: continue

                        p_pages = job.get('TotalPages', 0)
                        if p_pages == 0: continue

                        p_user = job.get('pUserName', 'Guest')
                        p_doc = job.get('pDocument', 'Unknown Document')
                        
                        # 컬러 감지
                        p_is_color = 0 
                        try:
                            devmode = job.get('pDevMode')
                            if devmode and getattr(devmode, 'Color', 1) == 2:
                                p_is_color = 1
                        except: pass
                        if p_is_color == 0 and 'color' in p_doc.lower():
                            p_is_color = 1

                        # 용지 감지
                        p_paper_size = "A4"
                        try:
                            if devmode:
                                paper_id = getattr(devmode, 'PaperSize', 9)
                                if paper_id == 8: p_paper_size = "A3"
                                elif paper_id == 9: p_paper_size = "A4"
                                else: p_paper_size = "Etc"
                        except: pass

                        # 비용 계산 (실시간)
                        # *주의: 여기서 계산해서 넣지만, 나중에 설정을 바꾸면 recalculate_db_costs()가 덮어씀
                        base_cost = current_settings["cost_color_a4"] if p_is_color else current_settings["cost_bw_a4"]
                        multiplier = 1.0
                        if p_paper_size == "A3":
                            multiplier = current_settings["mult_a3_color"] if p_is_color else current_settings["mult_a3_bw"]
                        
                        unit_cost = int(base_cost * multiplier)
                        p_cost = p_pages * unit_cost
                        p_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        conn = sqlite3.connect("print_log.db")
                        cursor = conn.cursor()
                        cursor.execute('''
                            INSERT INTO logs (job_id, printer_name, computer_name, user_name, document_name, pages, paper_size, is_color, unit_cost, cost, print_time, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (p_job_id, printer_name, my_computer_name, p_user, p_doc, p_pages, p_paper_size, p_is_color, unit_cost, p_cost, p_time, "PRINTED"))
                        conn.commit()
                        conn.close()

                        processed_jobs.add(unique_id)
                        
                        color_str = "컬러" if p_is_color else "흑백"
                        print(f"[감지] {p_doc} | {p_paper_size} {color_str} {p_pages}장 | {p_cost}원")
                        
                        app_instance.event_generate("<<NewLog>>", when="tail")

                except: pass
                finally:
                    if phandle: win32print.ClosePrinter(phandle)
            
            time.sleep(1)

        except Exception as e:
            print(f"[에러] {e}")
            time.sleep(5)

# =========================================================
# 3. GUI 애플리케이션
# =========================================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("통합 프린트 비용 관리 시스템 v2.1")
        self.geometry("1100x750")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # 사이드바
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(5, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Print Manager", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.btn_dash = ctk.CTkButton(self.sidebar_frame, text="종합 대시보드", command=self.show_dashboard)
        self.btn_dash.grid(row=1, column=0, padx=20, pady=10)
        
        self.btn_history = ctk.CTkButton(self.sidebar_frame, text="상세 출력 이력", command=self.show_history)
        self.btn_history.grid(row=2, column=0, padx=20, pady=10)

        self.btn_settings = ctk.CTkButton(self.sidebar_frame, text="단가/배수 설정", command=self.show_settings)
        self.btn_settings.grid(row=3, column=0, padx=20, pady=10)

        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)

        self.setup_treeview_style()
        self.show_dashboard()
        
        self.bind("<<NewLog>>", self.on_new_log)
        
        self.monitor_thread = threading.Thread(target=monitor_loop, args=(self,), daemon=True)
        self.monitor_thread.start()

    def setup_treeview_style(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=30)
        style.configure("Treeview.Heading", background="#3a3a3a", foreground="white", font=('Arial', 10, 'bold'))
        style.map("Treeview", background=[('selected', '#1f6aa5')])

    def clear_main_frame(self):
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def on_new_log(self, event):
        if hasattr(self, 'current_page') and self.current_page == 'dashboard':
            self.refresh_dashboard_stats()
        elif hasattr(self, 'current_page') and self.current_page == 'history':
            self.show_history()

    # --- 대시보드 ---
    def show_dashboard(self):
        self.current_page = 'dashboard'
        self.clear_main_frame()
        
        # 필터
        filter_frame = ctk.CTkFrame(self.main_frame)
        filter_frame.pack(fill="x", pady=(0, 20), ipady=5)
        
        ctk.CTkLabel(filter_frame, text="기간 조회:", font=("Arial", 14, "bold")).pack(side="left", padx=20)
        
        self.entry_start = ctk.CTkEntry(filter_frame, width=100, placeholder_text="YYYY-MM-DD")
        self.entry_start.pack(side="left", padx=5)
        ctk.CTkLabel(filter_frame, text="~").pack(side="left")
        self.entry_end = ctk.CTkEntry(filter_frame, width=100, placeholder_text="YYYY-MM-DD")
        self.entry_end.pack(side="left", padx=5)
        
        ctk.CTkButton(filter_frame, text="검색", width=60, command=self.refresh_dashboard_stats).pack(side="left", padx=10)
        ctk.CTkButton(filter_frame, text="오늘", width=60, fg_color="#555555", command=lambda: self.set_date_filter("today")).pack(side="left", padx=5)
        ctk.CTkButton(filter_frame, text="이번달", width=60, fg_color="#555555", command=lambda: self.set_date_filter("month")).pack(side="left", padx=5)

        self.stats_container = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.stats_container.pack(fill="x", expand=False)
        
        self.set_date_filter("today")

    def set_date_filter(self, mode):
        today = datetime.datetime.now()
        if mode == "today":
            str_date = today.strftime("%Y-%m-%d")
            self.entry_start.delete(0, 'end'); self.entry_start.insert(0, str_date)
            self.entry_end.delete(0, 'end'); self.entry_end.insert(0, str_date)
        elif mode == "month":
            start_date = today.replace(day=1).strftime("%Y-%m-%d")
            end_date = today.strftime("%Y-%m-%d")
            self.entry_start.delete(0, 'end'); self.entry_start.insert(0, start_date)
            self.entry_end.delete(0, 'end'); self.entry_end.insert(0, end_date)
        
        self.refresh_dashboard_stats()

    def refresh_dashboard_stats(self):
        for widget in self.stats_container.winfo_children(): widget.destroy()

        start_date = self.entry_start.get()
        end_date = self.entry_end.get()
        query_start = f"{start_date} 00:00:00"
        query_end = f"{end_date} 23:59:59"

        conn = sqlite3.connect("print_log.db")
        cursor = conn.cursor()
        
        stats = {
            "A4_BW": {"cnt":0, "cost":0}, "A3_BW": {"cnt":0, "cost":0},
            "A4_Col": {"cnt":0, "cost":0}, "A3_Col": {"cnt":0, "cost":0}
        }
        
        cursor.execute("SELECT pages, paper_size, is_color, cost FROM logs WHERE print_time BETWEEN ? AND ?", (query_start, query_end))
        rows = cursor.fetchall()
        
        total_pages = 0
        total_cost = 0
        
        for p, size, is_col, cost in rows:
            total_pages += p
            total_cost += cost
            key = "A4_BW"
            if size == "A3": key = "A3_Col" if is_col else "A3_BW"
            else: key = "A4_Col" if is_col else "A4_BW"
            stats[key]["cnt"] += p
            stats[key]["cost"] += cost
        conn.close()

        total_frame = ctk.CTkFrame(self.stats_container, fg_color="#1f6aa5")
        total_frame.grid(row=0, column=0, columnspan=4, sticky="ew", padx=5, pady=10)
        ctk.CTkLabel(total_frame, text=f"기간 총 비용 ({start_date} ~ {end_date})", font=("Arial", 16, "bold"), text_color="white").pack(pady=(10,0))
        ctk.CTkLabel(total_frame, text=f"{total_cost:,} 원", font=("Arial", 36, "bold"), text_color="white").pack(pady=(5,10))
        ctk.CTkLabel(total_frame, text=f"총 {total_pages:,} 장", font=("Arial", 14), text_color="white").pack(pady=(0,10))

        def create_card(col, title, count, cost, color_theme):
            f = ctk.CTkFrame(self.stats_container, border_width=2, border_color=color_theme)
            f.grid(row=1, column=col, sticky="ew", padx=5, pady=5)
            ctk.CTkLabel(f, text=title, font=("Arial", 14, "bold")).pack(pady=(10,5))
            ctk.CTkLabel(f, text=f"{count:,} 장", font=("Arial", 20, "bold")).pack()
            ctk.CTkLabel(f, text=f"{cost:,} 원", font=("Arial", 16), text_color=color_theme).pack(pady=(5,10))

        self.stats_container.grid_columnconfigure(0, weight=1)
        self.stats_container.grid_columnconfigure(1, weight=1)
        self.stats_container.grid_columnconfigure(2, weight=1)
        self.stats_container.grid_columnconfigure(3, weight=1)

        create_card(0, "A4 흑백", stats["A4_BW"]["cnt"], stats["A4_BW"]["cost"], "gray")
        create_card(1, "A3 흑백", stats["A3_BW"]["cnt"], stats["A3_BW"]["cost"], "gray")
        create_card(2, "A4 컬러", stats["A4_Col"]["cnt"], stats["A4_Col"]["cost"], "#E04F5F")
        create_card(3, "A3 컬러", stats["A3_Col"]["cnt"], stats["A3_Col"]["cost"], "#E04F5F")

    # --- 출력 이력 ---
    def show_history(self):
        self.current_page = 'history'
        self.clear_main_frame()
        
        ctk.CTkLabel(self.main_frame, text="상세 출력 이력", font=("Arial", 20, "bold")).pack(pady=(0, 10), anchor="w")

        table_frame = ctk.CTkFrame(self.main_frame)
        table_frame.pack(fill="both", expand=True)

        columns = ("time", "printer", "user", "doc", "spec", "pages", "unit", "cost")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Treeview")

        tree.heading("time", text="시간")
        tree.heading("printer", text="프린터(드라이버)")
        tree.heading("user", text="사용자")
        tree.heading("doc", text="문서명")
        tree.heading("spec", text="사양")
        tree.heading("pages", text="매수")
        tree.heading("unit", text="적용단가")
        tree.heading("cost", text="총비용")

        tree.column("time", width=130, anchor="center")
        tree.column("printer", width=150, anchor="w")
        tree.column("user", width=80, anchor="center")
        tree.column("doc", width=200, anchor="w")
        tree.column("spec", width=80, anchor="center")
        tree.column("pages", width=50, anchor="center")
        tree.column("unit", width=60, anchor="e")
        tree.column("cost", width=80, anchor="e")

        scrollbar = ctk.CTkScrollbar(table_frame, command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(side="left", fill="both", expand=True)

        conn = sqlite3.connect("print_log.db")
        cursor = conn.cursor()
        cursor.execute("SELECT print_time, printer_name, user_name, document_name, paper_size, is_color, pages, unit_cost, cost FROM logs ORDER BY id DESC LIMIT 100")
        rows = cursor.fetchall()
        
        for row in rows:
            color_txt = "컬러" if row[5] else "흑백"
            spec_txt = f"{row[4]} / {color_txt}"
            tree.insert("", "end", values=(row[0], row[1], row[2], row[3], spec_txt, f"{row[6]}장", f"@{row[7]}", f"{row[8]:,}원"))
            
        conn.close()

    # --- 설정 ---
    def show_settings(self):
        self.current_page = 'settings'
        self.clear_main_frame()
        
        ctk.CTkLabel(self.main_frame, text="단가 및 배수 설정", font=("Arial", 24, "bold")).pack(pady=20)
        
        form_frame = ctk.CTkFrame(self.main_frame)
        form_frame.pack(pady=10, padx=50, fill="x")

        # 1. 기본 단가
        ctk.CTkLabel(form_frame, text="[기본 단가 (A4)]", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(10,5))
        
        ctk.CTkLabel(form_frame, text="A4 흑백 (원):").grid(row=1, column=0, padx=20, pady=10)
        e_bw = ctk.CTkEntry(form_frame)
        e_bw.insert(0, str(int(current_settings['cost_bw_a4'])))
        e_bw.grid(row=1, column=1, padx=20, pady=10)

        ctk.CTkLabel(form_frame, text="A4 컬러 (원):").grid(row=2, column=0, padx=20, pady=10)
        e_col = ctk.CTkEntry(form_frame)
        e_col.insert(0, str(int(current_settings['cost_color_a4'])))
        e_col.grid(row=2, column=1, padx=20, pady=10)

        # 2. A3 배수
        ctk.CTkLabel(form_frame, text="[A3 요금 배수 (A4 대비)]", font=("Arial", 14, "bold")).grid(row=3, column=0, columnspan=2, pady=(20,5))

        ctk.CTkLabel(form_frame, text="A3 흑백 배수:").grid(row=4, column=0, padx=20, pady=10)
        combo_bw_mult = ctk.CTkComboBox(form_frame, values=["1.0", "1.5", "2.0", "2.5", "3.0"])
        combo_bw_mult.set(str(current_settings['mult_a3_bw']))
        combo_bw_mult.grid(row=4, column=1, padx=20, pady=10)

        ctk.CTkLabel(form_frame, text="A3 컬러 배수:").grid(row=5, column=0, padx=20, pady=10)
        combo_col_mult = ctk.CTkComboBox(form_frame, values=["1.0", "1.5", "2.0", "2.5", "3.0"])
        combo_col_mult.set(str(current_settings['mult_a3_color']))
        combo_col_mult.grid(row=5, column=1, padx=20, pady=10)

        def save():
            try:
                # 1. 설정 저장
                update_setting('cost_bw_a4', float(e_bw.get()))
                update_setting('cost_color_a4', float(e_col.get()))
                update_setting('mult_a3_bw', float(combo_bw_mult.get()))
                update_setting('mult_a3_color', float(combo_col_mult.get()))
                
                # 2. [핵심] 과거 데이터 전체 재계산 호출
                recalculate_db_costs()
                
                messagebox.showinfo("성공", "설정이 저장되고\n모든 과거 데이터에 새 단가가 적용되었습니다.")
            except ValueError:
                messagebox.showerror("오류", "올바른 숫자를 입력해주세요.")

        ctk.CTkButton(self.main_frame, text="설정 저장 (전체 데이터 반영)", command=save, height=40, fg_color="green").pack(pady=20)

if __name__ == "__main__":
    init_db()
    app = App()
    app.mainloop()