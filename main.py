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
import matplotlib.ticker as ticker
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
        if not self.winfo_exists(): return
        percent = int((current / total) * 100) if total > 0 else 0
        if percent > 100: percent = 100
        self.progress['value'] = percent
        self.lbl.config(text=f"{text} ({percent}%)")
        self.update_idletasks()
        self.update()
        if percent >= 100:
            self.after(200, self.safe_destroy)

    def safe_destroy(self):
        try:
            if self.winfo_exists():
                self.grab_release()
                self.destroy()
        except:
            pass

class DataAnalysisApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reliability Data Analyzer v21.0 - [Plus Absolute Fixed Suite]")
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
        # 전체를 대문자로 변환하여 일관성 확보
        name_we = os.path.splitext(filename)[0].upper()
        
        # [최종 수정] 1순위 강제 격리 추출 엔진 가동: Read-out(시간대/사이클) 정보 원천 차단 필터
        ro_match = re.search(r'(\d+)\s*(HR|CYC|MIN|SEC|DAY|WK|MONTH|R|T|STEP|ST)', name_we, re.IGNORECASE)
        if ro_match:
            ro_num = int(ro_match.group(1))
            ro_str = ro_match.group(0).lower()
            # 파일명 전체 문자열에서 검출된 시간대 정보 문자열을 흔적도 없이 삭제 (LOT 이름 훼손 원인 차단)
            name_we = name_we.replace(ro_match.group(0), "")
        else:
            # 보조 숫자 스캔 필터링
            digits = re.findall(r'\d+', name_we)
            if digits:
                ro_num = int(digits[0])
                ro_str = f"{ro_num}hr"
                name_we = name_we.replace(digits[0], "")
            else:
                ro_num = 0
                ro_str = "0hr"

        # 타임 정보가 삭제된 상태의 파일명에서 토큰 재분리
        tokens = [t.strip() for t in name_we.split('+') if t.strip()]
        
        rel_keywords = [
            "D-H3TRB", "H3TRB", "HTGB+", "HTGB-", "HTGB", "HTRB", "THBT", "THB", "THU", 
            "UHAST", "HAST", "TST", "PCT", "HTOL", "LTSL", "HTSL", "HTFS", "HTFB",
            "IOL", "VIB", "DGS", "DRB", "HTS", "LTS", "MS", "PC", "PRECON", "TC"
        ]
        
        test_item = "RELIABILITY_TEST"
        for t in tokens:
            if t in rel_keywords:
                test_item = t
                break
                
        # 신뢰성 항목 명칭을 제외한 순수 알짜배기 찌꺼기만 결합하여 LOT 번호 생성
        lot_tokens = [t for t in tokens if t != test_item]
        if lot_tokens:
            lot_str = "+".join(lot_tokens)
        else:
            lot_str = "UNKNOWN_LOT"
            
        # 이제 동일 Lot 데이터는 무조건 단 하나의 일치하는 group_key를 선점하게 됩니다.
        group_key = f"{test_item}_{lot_str}"
        return group_key, ro_str, ro_num, test_item, lot_str

    def full_load_dataframe(self, path):
        path_lower = path.lower()
        if path_lower.endswith('.csv') or '.csv' in path_lower:
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
        
        pb = AutoClosingProgressPop(self, "데이터 구조 분석 및 자동 매칭 연산 중")
        def target_thread():
            try:
                self.process_files(files, pb)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"분석 오류 발생:\n{str(e)}"))
                self.after(0, pb.safe_destroy)
        threading.Thread(target=target_thread, daemon=True).start()

    def process_files(self, files, pb):
        temp_data, all_p, self.lot_groups = {}, set(), {}
        total_files = len(files)
        
        for idx, path in enumerate(files):
            fname = os.path.basename(path)
            self.after(0, pb.update_progress, idx, total_files, f"격자 규격 주소 동기화 중 ({idx+1}/{total_files})")
            
            group_key, ro, ro_n, test_item, lot_id = self.parse_filename_info(fname)
            df = self.full_load_dataframe(path)
            if df.empty: continue
            
            base_row_idx = None
            test_no_row_idx = None
            start_col_idx = None
            
            for i in range(min(120, len(df))):
                found_t1 = False
                for c_idx in range(df.shape[1]):
                    if re.match(r'^t1$', str(df.iloc[i, c_idx]).strip(), re.IGNORECASE):
                        start_col_idx = c_idx
                        base_row_idx = i  
                        found_t1 = True
                        break
                if found_t1: break

            if base_row_idx is None:
                for i in range(min(120, len(df))):
                    col0_str = str(df.iloc[i, 0]).strip().replace(" ", "").lower()
                    if "parameter" in col0_str: base_row_idx = i
            
            # [수정] 오프셋 규칙 절대 고정: N(기준행), N+1(파라미터명), N+4(테스트 조건), N+7(단위행)
            if base_row_idx is not None:
                p_name_row_idx = base_row_idx + 1
                cond_row_idx = base_row_idx + 4  
                unit_row_idx = base_row_idx + 7
                test_no_row_idx = base_row_idx + 27 if base_row_idx + 27 < len(df) else base_row_idx + 10
            else:
                p_name_row_idx, cond_row_idx, unit_row_idx, test_no_row_idx = 20, 23, 26, 46
            
            for i in range(min(120, len(df))):
                col0_str = str(df.iloc[i, 0]).strip().replace(" ", "").lower()
                if "unit" in col0_str: unit_row_idx = i
                if "testno" in col0_str: test_no_row_idx = i

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
                    if len(units) > 0 and v0 == "": break
                        
            if len(units) == 0: continue
            
            p_dict = {}
            max_cols = df.shape[1]
            current_cont_prefix = "" 
            
            for col_idx in range(start_col_idx, max_cols):
                if col_idx >= df.shape[1] or p_name_row_idx >= len(df) or unit_row_idx >= len(df): continue
                
                # 단위 데이터 추출
                unit_val = str(df.iloc[unit_row_idx, col_idx]).strip().replace("'", "")
                if pd.isna(df.iloc[unit_row_idx, col_idx]) or unit_val == "" or unit_val.lower() in ["nan", "unit"]: 
                    unit_val = ""
                
                p_name_raw = str(df.iloc[p_name_row_idx, col_idx]).strip()
                if pd.isna(df.iloc[p_name_row_idx, col_idx]) or p_name_raw == "" or p_name_raw.lower() in ["nan", "item", "parameter", "'"]: 
                    continue
                
                # [수정] Module 전용 CONT_ 접두사 영속 상속 상태 머신 가동
                if self.data_mode == "Module":
                    if p_name_raw.upper().startswith("CONT_"):
                        tokens_cont = p_name_raw.upper().split('_')
                        if len(tokens_cont) > 1:
                            current_cont_prefix = tokens_cont[1]
                        continue 
                
                # [수정] 단위가 누락된 파라미터 컬럼 항목 필터링 차단
                if not unit_val:
                    continue

                cond_val = ""
                if cond_row_idx < len(df):
                    cond_val = str(df.iloc[cond_row_idx, col_idx]).strip()
                    if pd.isna(df.iloc[cond_row_idx, col_idx]) or cond_val.lower() in ["nan", "'"]: cond_val = ""

                # 유일 해시 결합 매칭 키 생성 규칙 통일
                p_name_key = p_name_raw.upper()
                if self.data_mode == "Module" and current_cont_prefix:
                    p_name_key = f"{current_cont_prefix}_{p_name_key}"
                
                if cond_val:
                    p_name_final = f"{p_name_key}_{cond_val.upper()}"
                else:
                    p_name_final = p_name_key
                
                vals = []
                for r_pos in data_row_positions:
                    if r_pos < len(df): vals.append(df.iloc[r_pos, col_idx])
                    else: vals.append(np.nan)
                vals = pd.to_numeric(vals, errors='coerce').tolist()
                
                if not all(v is None or np.isnan(v) for v in vals):
                    display_raw_title = f"{current_cont_prefix}_{p_name_raw}" if (self.data_mode == "Module" and current_cont_prefix) else p_name_raw
                    p_dict[p_name_final] = {
                        'raw_name': display_raw_title, 'cond': cond_val, 'unit': unit_val, 'values': vals, 'units_map': units
                    }
                    all_p.add(p_name_final)
            
            if p_dict:
                unique_fname_key = f"{fname}_{idx}"
                temp_data[unique_fname_key] = {'lot_key': group_key, 'ro': ro, 'ro_num': ro_n, 'params': p_dict, 'test_item': test_item, 'lot_id': lot_id}
                if group_key not in self.lot_groups: 
                    self.lot_groups[group_key] = []
                self.lot_groups[group_key].append(unique_fname_key)

        if not all_p: 
            raise ValueError("단위를 포함하는 유효 파라미터가 감출되지 않았습니다.")
        
        self.parameter_list = sorted(list(all_p))
        self.raw_files_data = temp_data
        
        for g_key in self.lot_groups:
            self.lot_groups[g_key].sort(key=lambda x: self.raw_files_data[x]['ro_num'])
            f_meta = self.raw_files_data[self.lot_groups[g_key][0]]
            self.lot_display_names[g_key] = f"{f_meta['test_item']}_{f_meta['lot_id']}"
            
        self.after(0, pb.update_progress, total_files, total_files, "파싱 정렬 완료")
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

    def run_with_progress_pop(self, title_text, target_func):
        pb = AutoClosingProgressPop(self, title_text)
        def run():
            self.after(30, lambda: pb.update_progress(50, 100, "행렬 렌더러 정렬 중..."))
            target_func()
            self.after(50, lambda: pb.update_progress(100, 100, "완료되었습니다!"))
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
            unit_str, raw_name, cond_str = "", "", ""
            for fn in lot_files:
                if param in self.raw_files_data[fn]['params']:
                    meta = self.raw_files_data[fn]['params'][param]
                    unit_str = meta['unit']
                    raw_name = meta['raw_name']
                    cond_str = meta['cond']
                    break
            if not raw_name: continue

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
                # [수정] 차트 제목 포맷 가이드라인 절대 준수 ('parameter_테스트 조건')
                title_formatter = f"{raw_name}_{cond_str} ({display_unit})" if cond_str else f"{raw_name} ({display_unit})"
                line_plots_meta.append({
                    'param': param, 'title': title_formatter, 'dataset': lines_dataset, 'all_samples': [s for s in all_samples if s not in del_set]
                })

        for param in self.selected_parameters:
            b_data, a_labels, b_cols, stats_data = [], [], [], []
            del_set = self.deleted_units.get((target_group_key, param), set())
            raw_name, cond_str = "", ""
            
            for fn in lot_files:
                if param in self.raw_files_data[fn]['params']:
                    raw_name = self.raw_files_data[fn]['params'][param]['raw_name']
                    cond_str = self.raw_files_data[fn]['params'][param]['cond']
                    break

            for f_idx, fn in enumerate(lot_files):
                if param in self.raw_files_data[fn]['params']:
                    p_info = self.raw_files_data[fn]['params'][param]
                    vals = [uy for ux, py_val in zip(p_info['units_map'], p_info['values']) for uy in [py_val] if uy is not None and not np.isnan(uy) and ux not in del_set]
                    if vals:
                        b_data.append(vals)
                        a_labels.append(self.raw_files_data[fn]['ro'])
                        b_cols.append(base_colors[f_idx % len(base_colors)])
                        
                        stats_data.append({
                            'ro': self.raw_files_data[fn]['ro'], 'min': np.min(vals), 'max': np.max(vals), 'avg': np.mean(vals), 'std': np.std(vals)
                        })
            
            if b_data: 
                box_title = f"{raw_name}_{cond_str} Dist" if cond_str else f"{raw_name} Dist"
                box_plots_meta.append({'title': box_title, 'b_data': b_data, 'a_labels': a_labels, 'b_cols': b_cols, 'stats_data': stats_data, 'chart_key': param})

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
                      command=lambda k=g_key, e=ent: self.run_with_progress_pop("그룹 타이틀 변경 중", lambda: self.update_lot_name(k, e.get()))).pack(side=tk.LEFT)
            
            if line_meta:
                grid_frame = tk.Frame(self.scrollable_frame)
                grid_frame.pack(fill=tk.X, padx=15, pady=5)
                
                cols = 1 if self.data_mode == "Discrete" else 3
                for c in range(cols): grid_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(line_meta):
                    cell = tk.Frame(grid_frame, bd=1, relief=tk.RIDGE, bg="white")
                    cell.grid(row=idx//cols, column=idx%cols, padx=4, pady=4, sticky="nsew")
                    
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
                    
                    fig_w = 13.2 if self.data_mode == "Discrete" else 4.2
                    fig, ax = plt.subplots(figsize=(fig_w, 3.4))
                    
                    if cur_lim["min"] != "": ax.set_ylim(bottom=float(cur_lim["min"]))
                    if cur_lim["max"] != "": ax.set_ylim(top=float(cur_lim["max"]))
                    
                    tk.Button(input_f, text="적용", font=("맑은 고딕", 7, "bold"), bg="#34495e", fg="white",
                              command=lambda cid=chart_id, emin=ent_min, emax=ent_max: self.run_with_progress_pop("Y축 스케일 범위 연산 및 재적용 중", lambda: self.apply_individual_y_limit(cid, emin.get(), emax.get()))).pack(side=tk.LEFT, padx=5)
                    
                    for px, py, pc, pm, ro_lbl, b_col in m['dataset']:
                        ax.plot(px, py, color=b_col, alpha=0.7, linewidth=1.5, zorder=1, label=ro_lbl)
                        for xi, yi, ci, mi in zip(px, py, pc, pm):
                            sc = ax.scatter(xi, yi, color=ci, marker=mi, s=55 if mi=='^' else 35, zorder=3, picker=3)
                            sc.__dict__['metadata'] = {'group_key': g_key, 'param': m['param'], 'unit': xi, 'ro': ro_lbl}
                    
                    # [수정] 순수 텍스트 타이틀 표기 규칙 적용
                    ax.set_title(m['title'], fontsize=10, weight='bold')
                    
                    # [수정] Discrete 분석 모드 축 레이블 처리 및 홀수 보조 격자선(Minor Grid) 구현 완료
                    if self.data_mode == "Discrete":
                        tick_positions, tick_labels = [], []
                        for s_idx, s_name in enumerate(m['all_samples']):
                            try:
                                num_val = int(re.findall(r'\d+', s_name)[0])
                                if num_val % 2 == 0:
                                    tick_positions.append(s_idx)
                                    tick_labels.append(s_name)
                            except:
                                if s_idx % 2 == 1:
                                    tick_positions.append(s_idx)
                                    tick_labels.append(s_name)
                        ax.set_xticks(tick_positions)
                        ax.set_xticklabels(tick_labels, rotation=0, fontsize=7)
                        
                        # 홀수 자리를 정밀 지탱하는 마이너 로케이터 그리드 결합
                        ax.xaxis.set_minor_locator(ticker.MultipleLocator(1))
                        ax.grid(True, which='minor', color='#e5e5e5', linestyle='--', alpha=0.7)
                    else:
                        ax.set_xticks(range(len(m['all_samples'])))
                        ax.set_xticklabels(m['all_samples'], rotation=15, fontsize=7)
                        
                    ax.grid(True, which='major', linestyle=":", alpha=0.6)
                    
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
                
                box_cols = 4
                for c in range(box_cols): box_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(box_meta):
                    bb = tk.Frame(box_frame, bd=1, relief=tk.GROOVE, bg="white")
                    bb.grid(row=idx//box_cols, column=idx%box_cols, padx=5, pady=5, sticky="nsew")
                    
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
                    
                    fig, ax = plt.subplots(figsize=(3.3, 2.6))
                    
                    if cur_lim["min"] != "": ax.set_ylim(bottom=float(cur_lim["min"]))
                    if cur_lim["max"] != "": ax.set_ylim(top=float(cur_lim["max"]))
                    
                    tk.Button(input_f, text="적용", font=("맑은 고딕", 7, "bold"), bg="#34495e", fg="white",
                              command=lambda cid=chart_id, emin=ent_min, emax=ent_max: self.run_with_progress_pop("Y축 스케일 범위 연산 및 재적용 중", lambda: self.apply_individual_y_limit(cid, emin.get(), emax.get()))).pack(side=tk.LEFT, padx=5)
                    
                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                    ax.set_xticklabels(m['a_labels'], fontsize=8)
                    for patch, color in zip(bp['boxes'], m['b_cols']):
                        patch.set_facecolor(color); patch.set_alpha(0.6)
                    ax.set_title(m['title'], fontsize=10, weight='bold')
                    ax.grid(True, alpha=0.3)
                    
                    plt.tight_layout()
                    FigureCanvasTkAgg(fig, master=bb).get_tk_widget().pack(fill=tk.BOTH, expand=True)
                    plt.close(fig)
                    
                    stat_table_frame = tk.Frame(bb, bg="#fafafa", bd=1, relief=tk.SOLID)
                    stat_table_frame.pack(fill=tk.X, padx=4, pady=4)
                    
                    for s_idx, s in enumerate(m['stats_data']):
                        lbl_text = f"[{s['ro']}] Min:{s['min']:.2f} | Max:{s['max']:.2f} | AVG:{s['avg']:.2f} | STD:{s['std']:.2f}"
                        tk.Label(stat_table_frame, text=lbl_text, font=("Consolas", 7), bg="#fafafa", anchor="w").pack(fill=tk.X, padx=2)

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
        self.execute_ui_rendering()

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

        pb = AutoClosingProgressPop(self, "PDF 리포트 가로 포맷 저장 중")
        
        def pdf_thread():
            try:
                with PdfPages(path) as pdf:
                    groups = sorted(self.lot_groups.keys())
                    total_steps = len(groups)
                    
                    for step, g_key in enumerate(groups):
                        self.after(0, pb.update_progress, int((step/total_steps)*100), 100, f"[{g_key}] 가로 벡터 데이터 변환 중...")
                        line_meta, box_meta = self.build_chart_data_structures(g_key)
                        
                        if line_meta:
                            cols, rows, items_per_page = (1, 3, 3) if self.data_mode == "Discrete" else (3, 3, 9)
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
                                    
                                    ax.set_title(m['title'], fontsize=8, weight='bold')
                                    
                                    if self.data_mode == "Discrete":
                                        tick_pos, tick_lbl = [], []
                                        for s_idx, s_name in enumerate(m['all_samples']):
                                            try:
                                                if int(re.findall(r'\d+', s_name)[0]) % 2 == 0:
                                                    tick_pos.append(s_idx); tick_lbl.append(s_name)
                                            except:
                                                if s_idx % 2 == 1: tick_pos.append(s_idx); tick_lbl.append(s_name)
                                        ax.set_xticks(tick_pos)
                                        ax.set_xticklabels(tick_lbl, fontsize=6)
                                        ax.xaxis.set_minor_locator(ticker.MultipleLocator(1))
                                    else:
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
                                        
                                    ax.set_title(m['title'], fontsize=8, weight='bold')
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
                self.after(0, pb.safe_destroy)

        threading.Thread(target=pdf_thread, daemon=True).start()

if __name__ == "__main__":
    DataAnalysisApp().mainloop()
