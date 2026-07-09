import os
import re
import sys
import threading
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import matplotlib
matplotlib.use("TkAgg")
matplotlib.rcParams['axes.unicode_minus'] = False
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

try:
    from matplotlib.backends.backend_pdf import PdfPages
except ImportError:
    pass

import matplotlib.font_manager as fm
try:
    font_location = "C:/Windows/Fonts/malgun.ttf"
    font_name = fm.FontProperties(fname=font_location).get_name()
    matplotlib.rc('font', family=font_name)
except:
    pass

class AutoClosingProgressPop(tk.Toplevel):
    def __init__(self, parent, title="Processing"):
        super().__init__(parent)
        self.title(title)
        self.geometry("380x130")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        self.update_idletasks()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (190)
        y = (hs / 2) - (65)
        self.geometry(f'380x130+{int(x)}+{int(y)}')
        
        self.lbl = tk.Label(self, text="작업을 처리 중입니다...", font=("맑은 고딕", 10))
        self.lbl.pack(pady=15)
        
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=5)
        self.update()
        
    def update_progress(self, current, total, text=""):
        percent = int((current / total) * 100) if total > 0 else 0
        if percent > 100: percent = 100
        self.progress['value'] = percent
        self.lbl.config(text=f"{text} ({percent}%)")
        self.update_idletasks()
        self.update()
        if percent >= 100:
            self.after(200, self.destroy)

class DataAnalysisApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reliability Data Analyzer v11.0 - [Plus Premium Engine]")
        self.geometry("1450x950")
        self.center_window(self, 1450, 950)
        
        self.raw_files_data = {}  
        self.parameter_list = []
        self.selected_parameters = []
        self.lot_groups = {}
        self.lot_display_names = {}
        
        self.custom_colors = {}   
        self.deleted_units = {}    
        self.undo_stack = []       
        
        self.is_delta_mode = tk.BooleanVar(value=False)
        self.init_upload_menu()
        
    def center_window(self, win, w, h):
        win.update_idletasks()
        ws = win.winfo_screenwidth()
        hs = win.winfo_screenheight()
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        win.geometry(f'{w}x{h}+{int(x)}+{int(y)}')

    def init_upload_menu(self):
        for widget in self.winfo_children(): widget.destroy()
        f = tk.Frame(self, pady=100)
        f.pack(expand=True, fill=tk.BOTH)
        tk.Label(f, text="Smart Reliability Data Analyzer", font=("맑은 고딕", 18, "bold")).pack(pady=10)
        tk.Button(f, text="파일 일괄 선택 (Lot/Read-out 혼합 가능)", font=("맑은 고딕", 12, "bold"), 
                  bg="#2b579a", fg="white", padx=20, pady=10, command=self.handle_file_upload).pack(pady=20)

    def parse_filename_info(self, filename):
        name_we = os.path.splitext(filename)[0]
        lot_match = re.search(r'(lot[\s_\-]*[a-zA-Z0-9]+)', name_we, re.IGNORECASE)
        lot_str = lot_match.group(1).upper().replace(" ", "").replace("_","").replace("-","") if lot_match else "UNKNOWN_LOT"
        
        ro_match = re.search(r'(\d+\s*(?:hr|cyc|min|sec|day|wk|month|r|t|step|st))', name_we, re.IGNORECASE)
        if ro_match:
            ro_str = ro_match.group(1).lower().replace(" ", "")
            ro_num = int(re.findall(r'\d+', ro_str)[0])
        else:
            nums = re.findall(r'\d+', name_we)
            ro_num = int(nums[-1]) if nums else 99999
            ro_str = f"{ro_num}_ReadOut" if nums else "0HR"
            
        rem = name_we
        if lot_match: rem = rem.replace(lot_match.group(1), "")
        if ro_match: rem = rem.replace(ro_match.group(1), "")
        rem = re.sub(r'[\s_\-]+', ' ', rem).strip()
        tokens = [t for t in rem.split(' ') if t and not t.isdigit()]
        test_item = tokens[0].upper() if tokens else "RELIABILITY_TEST"
        
        group_key = f"{test_item}_{lot_str}"
        return group_key, ro_str, ro_num, test_item, lot_str

    def full_load_dataframe(self, path):
        if path.endswith('.csv'):
            try: return pd.read_csv(path, header=None, engine='c', on_bad_lines='skip')
            except: return pd.read_csv(path, header=None, engine='python')
        else:
            try: return pd.read_excel(path, header=None)
            except: return pd.read_excel(path, header=None, engine='openpyxl')

    def handle_file_upload(self):
        files = filedialog.askopenfilenames(title="파일 선택", filetypes=[("Data Files", "*.csv *.xlsx *.xls")])
        if not files: return
        
        m = tk.Toplevel(self); m.title("Mode"); m.transient(self); m.grab_set()
        self.center_window(m, 300, 120)
        
        tk.Label(m, text="데이터 유형 선택", font=("맑은 고딕", 10, "bold")).pack(pady=10)
        f = tk.Frame(m); f.pack()
        tk.Button(f, text="Discrete", width=10, command=lambda: self.start_proc(files, "Discrete", m)).pack(side=tk.LEFT, padx=5)
        tk.Button(f, text="Module", width=10, command=lambda: self.start_proc(files, "Module", m)).pack(side=tk.LEFT, padx=5)

    def start_proc(self, files, mode, win):
        self.data_mode = mode
        win.destroy()
        
        pb = AutoClosingProgressPop(self, "데이터 연산 처리 중")
        def target_thread():
            try:
                self.process_files(files, pb)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"분석 오류 발생:\n{str(e)}"))
                if pb.winfo_exists(): pb.destroy()
        threading.Thread(target=target_thread, daemon=True).start()

    def process_files(self, files, pb):
        temp_data, all_p, self.lot_groups = {}, set(), {}
        total_files = len(files)
        
        for idx, path in enumerate(files):
            fname = os.path.basename(path)
            self.after(0, pb.update_progress, idx, total_files, f"정밀 마스터 스캔 중 ({idx+1}/{total_files})")
            
            group_key, ro, ro_n, test_item, lot_id = self.parse_filename_info(fname)
            df = self.full_load_dataframe(path)
            if df.empty: continue
            
            # --- [실제 원본 대응 2단계 자동 검색 가동] ---
            p_name_row_idx = None
            unit_row_idx = None
            test_no_row_idx = None
            start_col_idx = None
            
            # 가이드 1: 상단 헤더 라인 찾기 (T1, T2 지시어 및 Unit 검색)
            for i in range(min(50, len(df))):
                row_list_clean = [str(x).strip().lower().replace(" ", "") for x in df.iloc[i].tolist()]
                
                if "unit" in row_list_clean:
                    unit_row_idx = i
                
                if any(re.match(r'^t1$', x) for x in row_list_clean):
                    p_name_row_idx = i + 1  # T1 바로 다음 줄이 이름 행
                    for c_idx, c_val in enumerate(row_list_clean):
                        if "t1" in c_val:
                            start_col_idx = c_idx
                            break

            # 가이드 2: 실제 시료 데이터 시작 위치 찾기 (0열이 순수 숫자인 행 역추적)
            # 19번 행 근처의 가짜 숫자를 거르기 위해 30번 행 이후부터 검사하여 안정성 확보
            for i in range(30, len(df)):
                val_0 = str(df.iloc[i, 0]).strip()
                if val_0.isdigit() and int(val_0) >= 1:
                    test_no_row_idx = i
                    break

            # 예외 예방 백업 디폴트 좌표 (A, B, C 시트 크로스 체크 매칭값)
            if p_name_row_idx is None: p_name_row_idx = 19
            if unit_row_idx is None: unit_row_idx = 26
            if test_no_row_idx is None: test_no_row_idx = 46
            if start_col_idx is None: start_col_idx = 7

            if test_no_row_idx >= len(df): continue
            
            # 시료 유효 넘버링 슬라이싱 추출 (문자열 미세 노이즈 전처리)
            units = []
            data_end_row = test_no_row_idx
            for i in range(test_no_row_idx, len(df)):
                v0 = str(df.iloc[i, 0]).strip().replace(".0", "")
                if v0.isdigit():
                    units.append(v0)
                    data_end_row = i + 1
                else:
                    # 중간에 빈칸이나 요약 정보가 나오면 시료 종료로 인식
                    if len(units) > 0 and v0 == "":
                        break
            
            num_samples = len(units)
            if num_samples == 0: continue
            
            p_dict = {}
            cont_prefix = ""
            max_cols = df.shape[1]
            
            # 파라미터 컬럼 분석 루프 가동
            for col_idx in range(start_col_idx, max_cols):
                if col_idx >= df.shape[1] or p_name_row_idx >= len(df) or unit_row_idx >= len(df):
                    continue
                
                p_name_raw = str(df.iloc[p_name_row_idx, col_idx]).strip()
                if pd.isna(df.iloc[p_name_row_idx, col_idx]) or p_name_raw == "" or p_name_raw.lower() in ["nan", "item", "parameter", "test", "'", "color"]: 
                    continue
                
                if self.data_mode == "Module":
                    if "scan" in p_name_raw.lower(): continue
                    if p_name_raw.upper().startswith("CONT_"):
                        cont_prefix = p_name_raw.upper().split('_')[1]
                        continue
                
                # 중복 방지 조합 명명법 (이름 행 기준 아래 3번째 줄 조건 결합)
                sub_name_idx = p_name_row_idx + 3 
                p_name_final = p_name_raw
                if sub_name_idx < len(df):
                    sub_val = str(df.iloc[sub_name_idx, col_idx]).strip()
                    if pd.notna(df.iloc[sub_name_idx, col_idx]) and sub_val != "" and sub_val.lower() != "nan" and sub_val != "0" and sub_val != "'":
                        sub_val_clean = re.sub(r'[\s]+', '', sub_val)
                        if sub_val_clean and not sub_val_clean.replace('.','').isdigit():
                            p_name_final = f"{p_name_raw}_{sub_val_clean}"
                
                if self.data_mode == "Module" and cont_prefix:
                    p_name_final = f"{cont_prefix}_{p_name_final}"
                
                # 단위 예외 처리 (원인 2번 해결: 작은따옴표만 있거나 공백인 경우 대시 치환)
                unit_val = str(df.iloc[unit_row_idx, col_idx]).strip().replace("'", "")
                if pd.isna(df.iloc[unit_row_idx, col_idx]) or unit_val == "" or unit_val.lower() in ["nan", "unit"]: 
                    unit_val = "-"  
                
                # 데이터 값 수집 및 타입 컨버팅 (비수치 기호는 자동으로 NaN 처리)
                raw_vals = df.iloc[test_no_row_idx : data_end_row, col_idx].tolist()
                vals = pd.to_numeric(raw_vals, errors='coerce').tolist()
                
                # 데이터가 전부 공란인 열만 아니라면 매칭 사전에 탑재
                if not all(v is None or np.isnan(v) for v in vals):
                    p_dict[p_name_final] = {'unit': unit_val, 'values': vals, 'units_map': units}
                    all_p.add(p_name_final)
                
            if p_dict:
                if group_key in self.lot_groups and any(temp_data[f]['ro'] == ro for f in self.lot_groups[group_key]):
                    ro = f"{ro}_{idx}"
                temp_data[fname] = {'lot_key': group_key, 'ro': ro, 'ro_num': ro_n, 'params': p_dict, 'test_item': test_item, 'lot_id': lot_id}
                if group_key not in self.lot_groups: self.lot_groups[group_key] = []
                self.lot_groups[group_key].append(fname)

        if not all_p: 
            raise ValueError("업로드하신 실제 CSV 파일 구조 내에서 유효한 시료 데이터 매칭에 실패했습니다.\n파일 손상 여부나 데이터 양식을 확인해주세요.")
        
        self.parameter_list = sorted(list(all_p))
        self.raw_files_data = temp_data
        
        for g_key in self.lot_groups:
            self.lot_groups[g_key].sort(key=lambda x: self.raw_files_data[x]['ro_num'])
            f_meta = self.raw_files_data[self.lot_groups[g_key][0]]
            self.lot_display_names[g_key] = f"{f_meta['test_item']}_{f_meta['lot_id']}"
            
        self.after(0, pb.update_progress, total_files, total_files, "파싱 완료")
        self.after(250, self.init_analysis_menu)

    def init_analysis_menu(self):
        for widget in self.winfo_children(): widget.destroy()
        
        t = tk.Frame(self, bg="#f4f4f4", pady=10, padx=10); t.pack(fill=tk.X)
        ctrl_f = tk.LabelFrame(t, text="Analysis Control Panel", font=("맑은 고딕", 9, "bold"), bg="#f4f4f4", padx=10)
        ctrl_f.pack(side=tk.RIGHT, padx=10)
        
        tk.Checkbutton(ctrl_f, text="Delta Mode (%)", variable=self.is_delta_mode, bg="#f4f4f4", 
                       command=lambda: self.run_with_progress_pop("모드 변환 및 연산 중", self.execute_ui_rendering)).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_f, text="↩ 되돌리기 (Undo)", font=("맑은 고딕", 9, "bold"), bg="#7f8c8d", fg="white", 
                  command=lambda: self.run_with_progress_pop("작업 되돌리는 중", self.perform_undo)).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_f, text="📄 가로 포맷 PDF 리포트 저장", font=("맑은 고딕", 9, "bold"), bg="#c0392b", fg="white", command=self.export_to_pdf).pack(side=tk.LEFT, padx=5)

        tk.Label(t, text="Parameter Selector (다중 선택 가능):", font=("맑은 고딕", 11, "bold"), bg="#f4f4f4").pack(anchor="w")
        lf = tk.Frame(t); lf.pack(fill=tk.X, pady=5)
        
        self.param_listbox = tk.Listbox(lf, selectmode=tk.EXTENDED, height=6, font=("Consolas", 10))
        self.param_listbox.pack(fill=tk.X, side=tk.LEFT, expand=True)
        sb = ttk.Scrollbar(lf, command=self.param_listbox.yview); sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.param_listbox.config(yscrollcommand=sb.set)
        
        self.param_listbox.insert(tk.END, "★ 전체 선택")
        for p in self.parameter_list: self.param_listbox.insert(tk.END, p)
        self.param_listbox.selection_set(0)

        btn_f = tk.Frame(t, bg="#f4f4f4"); btn_f.pack(fill=tk.X)
        tk.Button(btn_f, text="그래프 그리기", bg="#107c41", fg="white", font=("맑은 고딕", 10, "bold"), 
                  command=lambda: self.run_with_progress_pop("그래프 화면 렌더링 중", self.update_selections_and_render)).pack(side=tk.LEFT, padx=5)
        
        c = tk.Frame(self); c.pack(fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(c, highlightthickness=0)
        self.scrollable_frame = tk.Frame(self.canvas)
        sb_v = ttk.Scrollbar(c, orient="vertical", command=self.canvas.yview); sb_v.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=sb_v.set)
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas_window, width=e.width))
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_wheel)
        
        self.run_with_progress_pop("초기 그래프 로딩 중", self.update_selections_and_render)

    def run_with_progress_pop(self, title_text, target_func):
        pb = AutoClosingProgressPop(self, title_text)
        def run():
            self.after(50, lambda: pb.update_progress(30, 100, "데이터 정렬 중..."))
            self.after(150, lambda: pb.update_progress(60, 100, "행렬 매칭 연산 중..."))
            target_func()
            self.after(250, lambda: pb.update_progress(100, 100, "완료되었습니다!"))
        threading.Thread(target=run, daemon=True).start()

    def _on_mouse_wheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_selections_and_render(self):
        selections = self.param_listbox.curselection()
        sel_items = [self.param_listbox.get(i) for i in selections]
        if "★ 전체 선택" in sel_items: self.selected_parameters = self.parameter_list.copy()
        else: self.selected_parameters = [v for v in sel_items if v != "★ 전체 선택"]
        if not self.selected_parameters: return
        self.execute_ui_rendering()

    def build_chart_data_structures(self, target_group_key):
        lot_files = self.lot_groups[target_group_key]
        base_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
        
        line_plots_meta = []
        box_plots_meta = []

        for param in self.selected_parameters:
            unit_str = ""
            for fn in lot_files:
                if param in self.raw_files_data[fn]['params']:
                    unit_str = self.raw_files_data[fn]['params'][param]['unit']; break
            if not unit_str: continue

            master_map = {} 
            all_samples_set = set()

            for filename in lot_files:
                if param not in self.raw_files_data[filename]['params']: continue
                p_info = self.raw_files_data[filename]['params'][param]
                ro_lbl = self.raw_files_data[filename]['ro']
                
                for s_id, val in zip(p_info['units_map'], p_info['values']):
                    if s_id == "": continue
                    all_samples_set.add(s_id)
                    if s_id not in master_map: master_map[s_id] = {}
                    master_map[s_id][ro_lbl] = val

            try: all_samples = sorted(list(all_samples_set), key=lambda x: float(re.findall(r'\d+\.?\d*', x)[0]) if re.findall(r'\d+\.?\d*', x) else x)
            except: all_samples = sorted(list(all_samples_set))

            if self.is_delta_mode.get() and len(lot_files) > 0:
                ref_ro = self.raw_files_data[lot_files[0]]['ro']
                for s_id in all_samples:
                    ref_val = master_map[s_id].get(ref_ro, None)
                    if ref_val is not None and not np.isnan(ref_val) and ref_val != 0:
                        for ro_lbl in master_map[s_id]:
                            v = master_map[s_id][ro_lbl]
                            if v is not None and not np.isnan(v):
                                master_map[s_id][ro_lbl] = 100.0 * (v - ref_val) / ref_val
                    else:
                        for ro_lbl in master_map[s_id]: master_map[s_id][ro_lbl] = np.nan

            del_set = self.deleted_units.get((target_group_key, param), set())
            lines_dataset = []

            for f_idx, filename in enumerate(lot_files):
                ro_lbl = self.raw_files_data[filename]['ro']
                px, py, pc, pm = [], [], [], []
                
                for s_id in all_samples:
                    if s_id in del_set: continue
                    val = master_map[s_id].get(ro_lbl, np.nan)
                    if pd.isna(val) or np.isinf(val): continue
                    
                    px.append(s_id)
                    py.append(float(val))
                    
                    c_key = (target_group_key, param, s_id)
                    if c_key in self.custom_colors:
                        pc.append(self.custom_colors[c_key])
                        pm.append('^')
                    else:
                        pc.append(base_colors[f_idx % len(base_colors)])
                        pm.append('o')
                
                if px:
                    lines_dataset.append((px, py, pc, pm, ro_lbl, base_colors[f_idx % len(base_colors)]))

            if lines_dataset:
                display_unit = "%" if self.is_delta_mode.get() else unit_str
                line_plots_meta.append({
                    'param': param, 'title': f"{param} ({display_unit})", 'dataset': lines_dataset, 'all_samples': [s for s in all_samples if s not in del_set]
                })

        for param in self.selected_parameters:
            b_data, a_labels, b_cols, stats = [], [], [], []
            del_set = self.deleted_units.get((target_group_key, param), set())

            for f_idx, fn in enumerate(lot_files):
                if param in self.raw_files_data[fn]['params']:
                    p_info = self.raw_files_data[fn]['params'][param]
                    vals = [uy for ux, py_val in zip(p_info['units_map'], p_info['values']) for uy in [py_val] if uy is not None and not np.isnan(uy) and ux not in del_set]
                    if vals:
                        b_data.append(vals)
                        a_labels.append(self.raw_files_data[fn]['ro'])
                        b_cols.append(base_colors[f_idx % len(base_colors)])
                        stats.append(f"[{self.raw_files_data[fn]['ro']}]\nAvg:{np.mean(vals):.1f}\nStd:{np.std(vals):.1f}")
            
            if b_data: box_plots_meta.append({'title': f"{param} Dist", 'b_data': b_data, 'a_labels': a_labels, 'b_cols': b_cols, 'stats': stats})

        return line_plots_meta, box_plots_meta

    def execute_ui_rendering(self):
        for w in self.scrollable_frame.winfo_children(): w.destroy()
        
        for g_key in sorted(self.lot_groups.keys()):
            line_meta, box_meta = self.build_chart_data_structures(g_key)
            if not line_meta and not box_meta: continue
            
            header_f = tk.Frame(self.scrollable_frame, bg="#eaf2f8", pady=6)
            header_f.pack(fill=tk.X, padx=10, pady=5)
            tk.Label(header_f, text=f"■ [{self.lot_display_names[g_key]}] Chart Group", font=("맑은 고딕", 13, "bold"), fg="#1e3799", bg="#eaf2f8").pack(side=tk.LEFT, padx=10)
            
            rename_f = tk.Frame(header_f, bg="#eaf2f8")
            rename_f.pack(side=tk.RIGHT, padx=15)
            ent = tk.Entry(rename_f, width=25, font=("맑은 고딕", 9))
            ent.insert(0, self.lot_display_names[g_key]); ent.pack(side=tk.LEFT, padx=5)
            tk.Button(rename_f, text="제목 변경", font=("맑은 고딕", 8, "bold"), bg="#546e7a", fg="white", 
                      command=lambda k=g_key, e=ent: self.update_lot_name(k, e.get())).pack(side=tk.LEFT)
            
            if line_meta:
                grid_frame = tk.Frame(self.scrollable_frame)
                grid_frame.pack(fill=tk.X, padx=15, pady=5)
                cols = 3 if self.data_mode == "Module" else 1
                for c in range(cols): grid_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(line_meta):
                    fig, ax = plt.subplots(figsize=(4.2 if cols==3 else 13.0, 3.4))
                    
                    for px, py, pc, pm, ro_lbl, b_col in m['dataset']:
                        ax.plot(px, py, color=b_col, alpha=0.7, linewidth=1.5, zorder=1, label=ro_lbl)
                        for xi, yi, ci, mi in zip(px, py, pc, pm):
                            sc = ax.scatter(xi, yi, color=ci, marker=mi, s=55 if mi=='^' else 35, zorder=3, picker=3)
                            sc.__dict__['metadata'] = {'group_key': g_key, 'param': m['param'], 'unit': xi, 'ro': ro_lbl}
                    
                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=9, weight='bold')
                    ax.set_xticks(range(len(m['all_samples'])))
                    ax.set_xticklabels(m['all_samples'], rotation=15, fontsize=7)
                    ax.grid(True, linestyle=":", alpha=0.5)
                    
                    handles, labels = ax.get_legend_handles_labels()
                    by_label = dict(zip(labels, handles))
                    if by_label: ax.legend(by_label.values(), by_label.keys(), loc="best", fontsize=7, framealpha=0.8)
                    
                    plt.tight_layout()
                    cell = tk.Frame(grid_frame, bd=1, relief=tk.RIDGE, bg="white")
                    cell.grid(row=idx//cols, column=idx%cols, padx=4, pady=4, sticky="nsew")
                    canvas_obj = FigureCanvasTkAgg(fig, master=cell)
                    canvas_obj.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                    fig.canvas.mpl_connect('pick_event', self.on_chart_point_clicked)
                    plt.close(fig)

            if box_meta:
                box_frame = tk.Frame(self.scrollable_frame, bg="#f9f9f9")
                box_frame.pack(fill=tk.X, padx=15, pady=10)
                for c in range(4): box_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(box_meta):
                    fig, ax = plt.subplots(figsize=(3.1, 2.5))
                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                    ax.set_xticklabels(m['a_labels'], fontsize=8)
                    for patch, color in zip(bp['boxes'], m['b_cols']):
                        patch.set_facecolor(color); patch.set_alpha(0.6)
                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=9, weight='bold')
                    ax.grid(True, alpha=0.3)
                    plt.tight_layout()
                    
                    bb = tk.Frame(box_frame, bd=1, relief=tk.GROOVE, bg="white")
                    bb.grid(row=idx//4, column=idx%4, padx=5, pady=5, sticky="nsew")
                    FigureCanvasTkAgg(fig, master=bb).get_tk_widget().pack()
                    
                    sf = tk.Frame(bb, bg="#fafafa"); sf.pack(fill=tk.X)
                    for s in m['stats']: tk.Label(sf, text=s, font=("맑은 고딕", 7), bg="#fafafa", justify=tk.LEFT).pack(anchor="w", padx=5)
                    plt.close(fig)

    def on_chart_point_clicked(self, event):
        scatter = event.artist
        if 'metadata' not in scatter.__dict__: return
        
        ind = event.ind
        if len(ind) == 0: return
        
        meta = scatter.__dict__['metadata']
        g_key, param, unit_id, ro_info = meta['group_key'], meta['param'], meta['unit'], meta['ro']
        
        m = tk.Toplevel(self); m.title("Data Editor"); m.geometry("450x180")
        self.center_window(m, 450, 180); m.transient(self); m.grab_set()
        
        tk.Label(m, text=f"선택 시료 번호: {unit_id} ({ro_info})", font=("맑은 고딕", 11, "bold")).pack(pady=10)
        color_section = tk.LabelFrame(m, text="변경할 색상 선택", font=("맑은 고딕", 9))
        color_section.pack(fill=tk.X, padx=15, pady=5)
        
        distinct_palette = ["#FF0000", "#0026ff", "#00b321", "#9400d3", "#ff8c00"]
        btn_frame = tk.Frame(color_section)
        btn_frame.pack(pady=5)
        
        for hex_code in distinct_palette:
            btn = tk.Button(btn_frame, bg=hex_code, activebackground=hex_code, width=5, height=2, bd=2, relief=tk.RAISED,
                            command=lambda k=g_key, p=param, u=unit_id, c=hex_code: [m.destroy(), self.run_with_progress_pop("마커 색상 변경 중", lambda: self.apply_point_color(k, p, u, c))])
            btn.pack(side=tk.LEFT, padx=8)
            
        action_f = tk.Frame(m); action_f.pack(pady=10)
        tk.Button(action_f, text="🗑️ 해당 시료 데이터 삭제", bg="#2c3e50", fg="white", font=("맑은 고딕", 9, "bold"),
                  command=lambda: [m.destroy(), self.run_with_progress_pop("시료 데이터 제외 중", lambda: self.delete_target_unit(g_key, param, unit_id))]).pack(side=tk.LEFT, padx=10)
        tk.Button(action_f, text="창 닫기", command=m.destroy).pack(side=tk.LEFT, padx=10)

    def apply_point_color(self, g_key, param, unit_id, chosen_color):
        c_key = (g_key, param, unit_id)
        self.undo_stack.append(('color', c_key, self.custom_colors.get(c_key, None)))
        self.custom_colors[c_key] = chosen_color
        self.execute_ui_rendering()

    def delete_target_unit(self, g_key, param, unit_id):
        key = (g_key, param)
        if key not in self.deleted_units: self.deleted_units[key] = set()
        self.deleted_units[key].add(unit_id)
        self.undo_stack.append(('delete', key, unit_id))
        self.execute_ui_rendering()

    def perform_undo(self):
        if not self.undo_stack:
            messagebox.showinfo("Undo", "되돌릴 작업 히스토리가 없습니다.")
            return
        action = self.undo_stack.pop()
        if action[0] == 'color':
            if action[2] is None: self.custom_colors.pop(action[1], None)
            else: self.custom_colors[action[1]] = action[2]
        elif action[0] == 'delete': self.deleted_units[action[1]].discard(action[2])
        self.execute_ui_rendering()

    def update_lot_name(self, g_key, new_name):
        if not new_name.strip(): return
        self.lot_display_names[g_key] = new_name.strip()
        self.run_with_progress_pop("그룹 타이틀 변경 중", self.execute_ui_rendering)

    def export_to_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 리포트 파일", "*.pdf")])
        if not path: return
        
        if os.path.exists(path):
            try:
                with open(path, 'a'): pass
            except IOError:
                messagebox.showerror("수정 불가", "동일한 이름의 PDF 리포트 파일이 열려 있습니다.\n닫고 다시 시도하세요.")
                return

        pb = AutoClosingProgressPop(self, "PDF 리포트 저장 중")
        
        def pdf_thread():
            try:
                with PdfPages(path) as pdf:
                    groups = sorted(self.lot_groups.keys())
                    total_steps = len(groups)
                    
                    for step, g_key in enumerate(groups):
                        self.after(0, pb.update_progress, int((step/total_steps)*100), 100, f"[{g_key}] 컴파일 중...")
                        line_meta, box_meta = self.build_chart_data_structures(g_key)
                        
                        if line_meta:
                            cols, rows, items_per_page = (3, 3, 9) if self.data_mode == "Module" else (1, 3, 3)
                            for i in range(0, len(line_meta), items_per_page):
                                chunk = line_meta[i:i+items_per_page]
                                fig, axes = plt.subplots(rows, cols, figsize=(11, 8.5), squeeze=False)
                                
                                for idx, m in enumerate(chunk):
                                    r, c = idx // cols, idx % cols
                                    ax = axes[r, c]
                                    
                                    for px, py, pc, pm, ro_lbl, b_col in m['dataset']:
                                        ax.plot(px, py, color=b_col, alpha=0.5, label=ro_lbl)
                                        for xi, yi, ci, mi in zip(px, py, pc, pm):
                                            ax.scatter(xi, yi, color=ci, marker=mi, s=40 if mi=='^' else 20)
                                    
                                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=8, weight='bold')
                                    ax.set_xticks(range(len(m['all_samples'])))
                                    ax.set_xticklabels(m['all_samples'], rotation=15, fontsize=6)
                                    ax.grid(True, linestyle=":", alpha=0.4)
                                    
                                    handles, labels = ax.get_legend_handles_labels()
                                    by_label = dict(zip(labels, handles))
                                    if by_label: ax.legend(by_label.values(), by_label.keys(), loc="best", fontsize=5)
                                
                                for idx in range(len(chunk), rows * cols): 
                                    axes[idx // cols, idx % cols].axis('off')
                                    
                                plt.tight_layout()
                                pdf.savefig(fig, dpi=200)
                                plt.close(fig)
                                
                        if box_meta:
                            b_cols, b_items_per_page = 4, 8
                            for i in range(0, len(box_meta), b_items_per_page):
                                chunk = box_meta[i:i+b_items_per_page]
                                b_rows = int(np.ceil(len(chunk) / b_cols))
                                
                                fig, axes = plt.subplots(b_rows, b_cols, figsize=(11, 8.5), squeeze=False)
                                for idx, m in enumerate(chunk):
                                    r, c = idx // b_cols, idx % b_cols
                                    ax = axes[r, c]
                                    
                                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                                    ax.set_xticklabels(m['a_labels'], fontsize=7)
                                    for patch, color in zip(bp['boxes'], m['b_cols']):
                                        patch.set_facecolor(color); patch.set_alpha(0.5)
                                        
                                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=8, weight='bold')
                                    ax.grid(True, alpha=0.3)
                                    
                                    stat_str = "\n".join([s.replace('\n', ' ') for s in m['stats']])
                                    ax.text(0.05, -0.4, stat_str, transform=ax.transAxes, fontsize=5, verticalalignment='top')
                                
                                for idx in range(len(chunk), b_rows * b_cols): axes[idx // b_cols, idx % b_cols].axis('off')
                                plt.tight_layout()
                                pdf.savefig(fig, dpi=200)
                                plt.close(fig)
                                
                self.after(0, pb.update_progress, 100, 100, "완료")
                self.after(300, lambda: messagebox.showinfo("Success", "매칭 PDF 저장이 완료되었습니다."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("PDF Export Error", f"PDF 컴파일 에러:\n{str(e)}"))
                if pb.winfo_exists(): pb.destroy()

        threading.Thread(target=pdf_thread, daemon=True).start()

if __name__ == "__main__":
    DataAnalysisApp().mainloop()
