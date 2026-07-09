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
        self.title("Reliability Data Analyzer v19.0 - [Pure Filename Engine]")
        self.geometry("1450x950")
        self.center_window(self, 1450, 950)
        
        self.reset_internal_states()
        self.init_upload_menu()
        
    def reset_internal_states(self):
        self.raw_files_data = {}  
        self.parameter_list = []
        self.selected_parameters = []
        self.lot_groups = {}
        self.lot_display_names = {}
        self.custom_colors = {}   
        self.deleted_units = {}    
        self.undo_stack = []       
        self.is_delta_mode = tk.BooleanVar(value=False)
        self.individual_limits = {}

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
        # [핵심 변경]: 내부 데이터를 일절 보지 않고 오직 겉으로 보이는 'filename' 자체만 뜯어서 분석합니다.
        name_we = os.path.splitext(filename)[0].upper()
        tokens = [t.strip() for t in re.split(r'[\s_\-]+', name_we) if t.strip()]
        
        rel_keywords = [
            "D-H3TRB", "H3TRB", "HTGB+", "HTGB-", "HTGB", "HTRB", "THBT", "THB", "THU", 
            "UHAST", "HAST", "TST", "PCT", "HTOL", "LTSL", "HTSL", "HTFS", "HTFB",
            "IOL", "VIB", "DGS", "DRB", "HTS", "LTS", "MS", "PC", "PRECON", "TC"
        ]
        
        # 1. 순수 파일명 토큰에서 신뢰성 이름 추출
        test_item = "RELIABILITY_TEST"
        for kw in rel_keywords:
            if kw in tokens:
                test_item = kw
                break
        if test_item == "RELIABILITY_TEST" and tokens:
            test_item = tokens[0]
            
        # 2. 순수 파일명 토큰에서 Read-out 시간대 추출
        ro_str = "0HR"
        ro_num = 0
        ro_match = re.search(r'(\d+\s*(?:HR|CYC|MIN|SEC|DAY|WK|MONTH|R|T|STEP|ST))', name_we, re.IGNORECASE)
        if ro_match:
            ro_str = ro_match.group(1).upper().replace(" ", "")
            ro_num = int(re.findall(r'\d+', ro_str)[0])
        else:
            for t in tokens:
                digits = re.findall(r'\d+', t)
                if digits and t != test_item:
                    ro_num = int(digits[0])
                    ro_str = f"{ro_num}HR"
                    break
                    
        # 3. 신뢰성 이름과 시간대를 제외한 나머지 토큰을 LOT 번호로 지정 ('3T' 버그 원천 봉쇄)
        lot_tokens = [t for t in tokens if t != test_item and t != ro_str]
        lot_str = "UNKNOWN_LOT"
        for lt in lot_tokens:
            if "LOT" in lt:
                lot_str = lt
                break
        if lot_str == "UNKNOWN_LOT" and lot_tokens:
            lot_str = "_".join(lot_tokens)
            
        group_key = f"{test_item}_{lot_str}"
        return group_key, ro_str.lower(), ro_num, test_item, lot_str

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
            self.after(0, pb.update_progress, idx, total_files, f"파일 파싱 중 ({idx+1}/{total_files})")
            
            group_key, ro, ro_n, test_item, lot_id = self.parse_filename_info(fname)
            df = self.full_load_dataframe(path)
            if df.empty: continue
            
            p_name_row_idx = None
            unit_row_idx = None
            test_no_row_idx = None
            start_col_idx = None
            
            for i in range(min(60, len(df))):
                col0_str = str(df.iloc[i, 0]).strip().replace(" ", "").lower()
                if "parameter" in col0_str:
                    p_name_row_idx = i + 1  
                    for c_idx in range(df.shape[1]):
                        if re.match(r'^t1$', str(df.iloc[i, c_idx]).strip(), re.IGNORECASE):
                            start_col_idx = c_idx
                            break
                if "unit" in col0_str:
                    unit_row_idx = i
                if "testno" in col0_str:
                    test_no_row_idx = i

            if p_name_row_idx is None: p_name_row_idx = 19
            if unit_row_idx is None: unit_row_idx = 26
            if test_no_row_idx is None: test_no_row_idx = 46
            if start_col_idx is None: start_col_idx = 7

            if test_no_row_idx >= len(df): continue
            
            units = []
            data_row_positions = []
            for i in range(test_no_row_idx + 1, len(df)):
                v0 = str(df.iloc[i, 0]).strip().replace(".0", "")
                if v0.isdigit():
                    units.append(v0)
                    data_row_positions.append(i)
                else:
                    if len(units) > 0 and v0 == "":
                        break
                        
            if len(units) == 0: continue
            
            p_dict = {}
            cont_prefix = ""
            max_cols = df.shape[1]
            
            for col_idx in range(start_col_idx, max_cols):
                if col_idx >= df.shape[1] or p_name_row_idx >= len(df) or unit_row_idx >= len(df):
                    continue
                
                unit_val = str(df.iloc[unit_row_idx, col_idx]).strip().replace("'", "")
                if pd.isna(df.iloc[unit_row_idx, col_idx]) or unit_val == "" or unit_val.lower() in ["nan", "unit"]: 
                    continue
                
                p_name_raw = str(df.iloc[p_name_row_idx, col_idx]).strip()
                if pd.isna(df.iloc[p_name_row_idx, col_idx]) or p_name_raw == "" or p_name_raw.lower() in ["nan", "item", "parameter", "test", "'", "color"]: 
                    continue
                
                if self.data_mode == "Module":
                    if "scan" in p_name_raw.lower(): continue
                    if p_name_raw.upper().startswith("CONT_"):
                        cont_prefix = p_name_raw.upper().split('_')[1]
                        continue
                
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
                
                vals = []
                for r_pos in data_row_positions:
                    if r_pos < len(df):
                        vals.append(df.iloc[r_pos, col_idx])
                    else:
                        vals.append(np.nan)
                vals = pd.to_numeric(vals, errors='coerce').tolist()
                
                if not all(v is None or np.isnan(v) for v in vals):
                    p_dict[p_name_final] = {'unit': unit_val, 'values': vals, 'units_map': units}
                    all_p.add(p_name_final)
            
            if p_dict:
                unique_fname_key = f"{fname}_{idx}"
                temp_data[unique_fname_key] = {'lot_key': group_key, 'ro': ro, 'ro_num': ro_n, 'params': p_dict, 'test_item': test_item, 'lot_id': lot_id}
                if group_key not in self.lot_groups: 
                    self.lot_groups[group_key] = []
                self.lot_groups[group_key].append(unique_fname_key)

        if not all_p: 
            raise ValueError("단위가 존재하는 유효한 파라미터 데이터를 추출하지 못했습니다.")
        
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
        tk.Button(ctrl_f, text="🏠 처음 화면으로 돌아가기", font=("맑은 고딕", 9, "bold"), bg="#2ecc71", fg="white", command=self.return_to_home_screen).pack(side=tk.LEFT, padx=15)

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

    def build_chart_data_structures(self, target_group_key):
        lot_files = self.lot_groups[target_group_key]
        base_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf', '#e31a1c', '#33a02c', '#fdbf6f', '#cab2d6', '#6a3d9a']
        
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
                if self.is_delta_mode.get() and f_idx == 0: continue
                    
                px, py, pc, pm = [], [], [], []
                for s_id in all_samples:
                    if s_id in del_set: continue
                    val = master_map[s_id].get(ro_lbl, np.nan)
                    if pd.isna(val) or np.isinf(val): continue
                    
                    px.append(s_id)
                    py.append(float(val))
                    
                    c_key = (target_group_key, param, s_id, ro_lbl)
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
            b_data, a_labels, b_cols, stats_data = [], [], [], []
            del_set = self.deleted_units.get((target_group_key, param), set())

            for f_idx, fn in enumerate(lot_files):
                if param in self.raw_files_data[fn]['params']:
                    p_info = self.raw_files_data[fn]['params'][param]
                    vals = [uy for ux, py_val in zip(p_info['units_map'], p_info['values']) for uy in [py_val] if uy is not None and not np.isnan(uy) and ux not in del_set]
                    if vals:
                        b_data.append(vals)
                        a_labels.append(self.raw_files_data[fn]['ro'])
                        b_cols.append(base_colors[f_idx % len(base_colors)])
                        
                        stats_data.append({
                            'ro': self.raw_files_data[fn]['ro'],
                            'min': np.min(vals),
                            'max': np.max(vals),
                            'avg': np.mean(vals),
                            'std': np.std(vals)
                        })
            
            if b_data: box_plots_meta.append({'title': f"{param} Dist", 'b_data': b_data, 'a_labels': a_labels, 'b_cols': b_cols, 'stats_data': stats_data})

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
            tk.Button(rename_f, text="제목 변경", font=("맑은 고딕", 8, "bold"), bg="#546e7a", fg="white", command=lambda k=g_key, e=ent: self.update_lot_name(k, e.get())).pack(side=tk.LEFT)
            
            if line_meta:
                grid_frame = tk.Frame(self.scrollable_frame)
                grid_frame.pack(fill=tk.X, padx=15, pady=5)
                cols = 3 if self.data_mode == "Module" else 1
                for c in range(cols): grid_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(line_meta):
                    cell = tk.Frame(grid_frame, bd=1, relief=tk.RIDGE, bg="white")
                    cell.grid(row=idx//cols, column=idx%cols, padx=4, pady=4, sticky="nsew")
                    
                    # [요청 규칙 1]: 개별 조절 UI 폼 상단 추가
                    input_f = tk.Frame(cell, bg="white")
                    input_f.pack(fill=tk.X, padx=5, pady=2)
                    tk.Label(input_f, text=f"Y축 범위:", bg="white", font=("맑은 고딕", 8)).pack(side=tk.LEFT)
                    
                    chart_id = f"{g_key}_{m['param']}_line"
                    cur_lim = self.individual_limits.get(chart_id, {"min": "", "max": ""})
                    
                    ent_min = tk.Entry(input_f, width=5, font=("Consolas", 8))
                    ent_min.insert(0, cur_lim["min"]); ent_min.pack(side=tk.LEFT, padx=2)
                    tk.Label(input_f, text="~", bg="white").pack(side=tk.LEFT)
                    ent_max = tk.Entry(input_f, width=5, font=("Consolas", 8))
                    ent_max.insert(0, cur_lim["max"]); ent_max.pack(side=tk.LEFT, padx=2)
                    
                    fig, ax = plt.subplots(figsize=(4.2 if cols==3 else 13.0, 3.4))
                    
                    if cur_lim["min"] != "": ax.set_ylim(bottom=float(cur_lim["min"]))
                    if cur_lim["max"] != "": ax.set_ylim(top=float(cur_lim["max"]))
                    
                    tk.Button(input_f, text="적용", font=("맑은 고딕", 7, "bold"), bg="#34495e", fg="white",
                              command=lambda cid=chart_id, emin=ent_min, emax=ent_max: self.apply_individual_y_limit(cid, emin.get(), emax.get())).pack(side=tk.LEFT, padx=5)
                    
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
                    canvas_obj = FigureCanvasTkAgg(fig, master=cell)
                    canvas_obj.get_tk_widget().pack(fill=tk.BOTH, expand=True)
                    fig.canvas.mpl_connect('pick_event', self.on_chart_point_clicked)
                    plt.close(fig)

            if box_meta:
                box_frame = tk.Frame(self.scrollable_frame, bg="#f9f9f9")
                box_frame.pack(fill=tk.X, padx=15, pady=10)
                for c in range(4): box_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(box_meta):
                    bb = tk.Frame(box_frame, bd=1, relief=tk.GROOVE, bg="white")
                    bb.grid(row=idx//4, column=idx%4, padx=5, pady=5, sticky="nsew")
                    
                    # 개별 조절 UI 폼 박스 연동
                    input_f = tk.Frame(bb, bg="white")
                    input_f.pack(fill=tk.X, padx=5, pady=2)
                    tk.Label(input_f, text=f"Y축 범위:", bg="white", font=("맑은 고딕", 8)).pack(side=tk.LEFT)
                    
                    chart_id = f"{g_key}_{m['title']}_box"
                    cur_lim = self.individual_limits.get(chart_id, {"min": "", "max": ""})
                    
                    ent_min = tk.Entry(input_f, width=5, font=("Consolas", 8))
                    ent_min.insert(0, cur_lim["min"]); ent_min.pack(side=tk.LEFT, padx=2)
                    tk.Label(input_f, text="~", bg="white").pack(side=tk.LEFT)
                    ent_max = tk.Entry(input_f, width=5, font=("Consolas", 8))
                    ent_max.insert(0, cur_lim["max"]); ent_max.pack(side=tk.LEFT, padx=2)
                    
                    fig, (ax, ax_stat) = plt.subplots(2, 1, figsize=(3.3, 4.3), gridspec_kw={'height_ratios': [2.2, 1.5]})
                    
                    if cur_lim["min"] != "": ax.set_ylim(bottom=float(cur_lim["min"]))
                    if cur_lim["max"] != "": ax.set_ylim(top=float(cur_lim["max"]))
                    
                    tk.Button(input_f, text="적용", font=("맑은 고딕", 7, "bold"), bg="#34495e", fg="white",
                              command=lambda cid=chart_id, emin=ent_min, emax=ent_max: self.apply_individual_y_limit(cid, emin.get(), emax.get())).pack(side=tk.LEFT, padx=5)
                    
                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                    ax.set_xticklabels(m['a_labels'], fontsize=8)
                    for patch, color in zip(bp['boxes'], m['b_cols']):
                        patch.set_facecolor(color); patch.set_alpha(0.6)
                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=9, weight='bold')
                    ax.grid(True, alpha=0.3)
                    
                    ax_stat.axis('off')
                    
                    # [요청 규칙 1]: 세로방향 1set(Min->Max->AVG->STD) 수직 배치 정렬
                    for b_idx, s in enumerate(m['stats_data']):
                        x_pos = b_idx + 1
                        stat_text = f"[{s['ro']}]\nMin:{s['min']:.2f}\nMax:{s['max']:.2f}\nAVG:{s['avg']:.2f}\nSTD:{s['std']:.2f}"
                        ax.text(x_pos, -0.15, stat_text, transform=ax.get_xaxis_transform(),
                                fontsize=7, fontfamily="Consolas", ha='center', va='top',
                                bbox=dict(boxstyle='square,pad=0.2', facecolor='#fafafa', edgecolor='#dddddd', alpha=0.8))
                    
                    plt.tight_layout()
                    FigureCanvasTkAgg(fig, master=bb).get_tk_widget().pack()
                    plt.close(fig)

    def apply_individual_y_limit(self, chart_id, ymin, ymax):
        self.individual_limits[chart_id] = {"min": ymin.strip(), "max": ymax.strip()}
        self.execute_ui_rendering()

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
        color_section = tk.LabelFrame(m, text="변경할 색상 선택 (해당 타임 마커만 단독 변경)", font=("맑은 고딕", 9))
        color_section.pack(fill=tk.X, padx=15, pady=5)
        
        distinct_palette = ["#FF0000", "#0026ff", "#00b321", "#9400d3", "#ff8c00"]
        btn_frame = tk.Frame(color_section)
        btn_frame.pack(pady=5)
        
        for hex_code in distinct_palette:
            btn = tk.Button(btn_frame, bg=hex_code, activebackground=hex_code, width=5, height=2, bd=2, relief=tk.RAISED,
                            command=lambda k=g_key, p=param, u=unit_id, r=ro_info, c=hex_code: [m.destroy(), self.run_with_progress_pop("마커 색상 변경 중", lambda: self.apply_point_color(k, p, u, r, c))])
            btn.pack(side=tk.LEFT, padx=8)
            
        action_f = tk.Frame(m); action_f.pack(pady=10)
        tk.Button(action_f, text="🗑️ 해당 시료 데이터 삭제", bg="#2c3e50", fg="white", font=("맑은 고딕", 9, "bold"),
                  command=lambda: [m.destroy(), self.run_with_progress_pop("시료 데이터 제외 중", lambda: self.delete_target_unit(g_key, param, unit_id))]).pack(side=tk.LEFT, padx=10)
        tk.Button(action_f, text="창 닫기", command=m.destroy).pack(side=tk.LEFT, padx=10)

    def apply_point_color(self, g_key, param, unit_id, ro_info, chosen_color):
        c_key = (g_key, param, unit_id, ro_info)
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

    def return_to_home_screen(self):
        if messagebox.askyesno("화면 초기화", "분석을 종료하고 처음 파일 선택 화면으로 이동하시겠습니까?"):
            self.reset_internal_states()
            self.init_upload_menu()

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
                                    
                                    chart_id = f"{g_key}_{m['param']}_line"
                                    lim = self.individual_limits.get(chart_id, {"min": "", "max": ""})
                                    if lim["min"] != "": ax.set_ylim(bottom=float(lim["min"]))
                                    if lim["max"] != "": ax.set_ylim(top=float(lim["max"]))
                                    
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
                                
                                for idx in range(len(chunk), rows * cols): axes[idx // cols, idx % cols].axis('off')
                                plt.tight_layout()
                                pdf.savefig(fig, dpi=200)
                                plt.close(fig)
                                
                        if box_meta:
                            # [요청 규칙 2]: 무조건 4개 * 3줄 = 12개 픽스 페이지 그리드 컴파일
                            b_cols, b_rows, b_items_per_page = 4, 3, 12
                            for i in range(0, len(box_meta), b_items_per_page):
                                chunk = box_meta[i:i+b_items_per_page]
                                fig, axes = plt.subplots(b_rows, b_cols, figsize=(11, 8.5), squeeze=False)
                                
                                for idx, m in enumerate(chunk):
                                    r, c = idx // b_cols, idx % b_cols
                                    ax = axes[r, c]
                                    
                                    chart_id = f"{g_key}_{m['title']}_box"
                                    lim = self.individual_limits.get(chart_id, {"min": "", "max": ""})
                                    if lim["min"] != "": ax.set_ylim(bottom=float(lim["min"]))
                                    if lim["max"] != "": ax.set_ylim(top=float(lim["max"]))
                                    
                                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                                    ax.set_xticklabels(m['a_labels'], fontsize=7)
                                    for patch, color in zip(bp['boxes'], m['b_cols']):
                                        patch.set_facecolor(color); patch.set_alpha(0.5)
                                        
                                    ax.set_title(f"[{self.lot_display_names[g_key]}] {m['title']}", fontsize=8, weight='bold')
                                    ax.grid(True, alpha=0.3)
                                    
                                    for b_idx, s in enumerate(m['stats_data']):
                                        x_pos = b_idx + 1
                                        stat_str_pdf = f"Min:{s['min']:.1f}\nMax:{s['max']:.1f}\nAVG:{s['avg']:.1f}\nSTD:{s['std']:.1f}"
                                        ax.text(x_pos, -0.18, stat_str_pdf, transform=ax.get_xaxis_transform(),
                                                fontsize=5, fontfamily="Consolas", ha='center', va='top')
                                
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
