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

class ProgressBarPop(tk.Toplevel):
    def __init__(self, parent, title="Processing"):
        super().__init__(parent)
        self.title(title)
        self.geometry("350x120")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # 화면 중앙 정렬
        self.update_idletasks()
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws / 2) - (175)
        y = (hs / 2) - (60)
        self.geometry(f'350x120+{int(x)}+{int(y)}')
        
        self.lbl = tk.Label(self, text="준비 중...", font=("Arial", 10))
        self.lbl.pack(pady=15)
        
        self.progress = ttk.Progressbar(self, orient="horizontal", length=280, mode="determinate")
        self.progress.pack(pady=5)
        
    def update_progress(self, current, total, text=""):
        percent = int((current / total) * 100)
        self.progress['value'] = percent
        self.lbl.config(text=f"{text} ({percent}%)")
        self.update()

class DataAnalysisApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Reliability Data Analyzer v5.0 - [All Layout & PDF Errors Fixed]")
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
        tk.Label(f, text="Smart Reliability Data Analyzer", font=("Arial", 18, "bold")).pack(pady=10)
        tk.Button(f, text="파일 일괄 선택 (Lot/Read-out 혼합 가능)", font=("Arial", 12, "bold"), 
                  bg="#2b579a", fg="white", padx=20, pady=10, command=self.handle_file_upload).pack(pady=20)

    def parse_filename_info(self, filename):
        # 파일명에서 Lot 형태 추출 정규식 고도화
        lot_match = re.search(r'(lot[\s_\-]*[a-zA-Z0-9]+)', filename, re.IGNORECASE)
        lot_str = lot_match.group(1).upper().replace(" ", "").replace("_","").replace("-","") if lot_match else "UNKNOWN_LOT"
        
        # 파일명에서 Read-out 형태(ex: 0hr, 168hr, 0cyc, 500cyc)를 추출하되 실패 시 파일명 전체 활용
        ro_match = re.search(r'(\d+\s*(?:hr|cyc|min|sec|day|wk|month))', filename, re.IGNORECASE)
        if ro_match:
            ro_str = ro_match.group(1).lower().replace(" ", "")
            ro_num = int(re.findall(r'\d+', ro_str)[0])
        else:
            # 매칭 안 될 경우 파일명에서 숫자 추출 시도
            nums = re.findall(r'\d+', filename)
            ro_str = filename.replace(".xlsx","").replace(".csv","").replace(".xls","")
            ro_num = int(nums[0]) if nums else 99999
            
        return lot_str, ro_str, ro_num

    def smart_read_csv_or_excel(self, path):
        if path.endswith('.csv'):
            try: return pd.read_csv(path, header=None, engine='c', on_bad_lines='skip')
            except:
                try: return pd.read_csv(path, header=None, engine='python', on_bad_lines='skip')
                except:
                    lines = []
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        for line in f: lines.append(line.strip().split(','))
                    return pd.DataFrame(lines)
        else:
            try: return pd.read_excel(path, header=None)
            except: return pd.read_excel(path, header=None, engine='openpyxl')

    def handle_file_upload(self):
        files = filedialog.askopenfilenames(title="파일 선택", filetypes=[("Data Files", "*.csv *.xlsx *.xls")])
        if not files: return
        
        m = tk.Toplevel(self); m.title("Mode"); m.transient(self); m.grab_set()
        self.center_window(m, 300, 120)
        
        tk.Label(m, text="데이터 유형 선택", font=("Arial", 10, "bold")).pack(pady=10)
        f = tk.Frame(m); f.pack()
        tk.Button(f, text="Discrete", width=10, command=lambda: self.start_proc(files, "Discrete", m)).pack(side=tk.LEFT, padx=5)
        tk.Button(f, text="Module", width=10, command=lambda: self.start_proc(files, "Module", m)).pack(side=tk.LEFT, padx=5)

    def start_proc(self, files, mode, win):
        self.data_mode = mode
        win.destroy()
        
        # 프로그레스 바 연동을 위해 별도 쓰레드로 분석 실행
        pb = ProgressBarPop(self, "데이터 읽는 중")
        def target_thread():
            try: self.process_files(files, pb)
            except Exception as e: messagebox.showerror("Error", str(e))
            finally: pb.destroy()
        threading.Thread(target=target_thread, daemon=True).start()

    def process_files(self, files, pb):
        df_s = self.smart_read_csv_or_excel(files[0])
        if df_s.empty: raise ValueError("첫 번째 파일이 비어있어 구조 파싱이 불가능합니다.")
            
        p_idx, u_idx, sample_start_row = None, None, None
        for i, r in df_s.iterrows():
            if r.empty or pd.isna(r.iloc[0]): continue
            val_first_col = str(r.iloc[0]).strip().lower().replace(" ", "")
            if "parameter" in val_first_col: p_idx = i
            if "unit" in val_first_col: u_idx = i
            if "sample" in val_first_col: sample_start_row = i

        if p_idx is None or u_idx is None or sample_start_row is None: 
            raise ValueError("파일 내에서 'Parameter', 'Unit' 혹은 'sample' 기준행을 찾을 수 없습니다.")

        temp_data, all_p, self.lot_groups = {}, set(), {}
        total_files = len(files)
        
        for idx, path in enumerate(files):
            fname = os.path.basename(path)
            # 프로그레스바 팝업 업데이트
            self.after(0, pb.update_progress, idx, total_files, f"파일 파싱 중: {fname[:20]}...")
            
            lot, ro, ro_n = self.parse_filename_info(fname)
            df = self.smart_read_csv_or_excel(path)
            if df.empty or sample_start_row + 1 >= len(df): continue
            if p_idx >= len(df) or u_idx >= len(df): continue
            
            units = df.iloc[sample_start_row + 1:, 0].dropna().astype(str).tolist()
            raw_p = df.iloc[p_idx, 1:].tolist()
            raw_u = df.iloc[u_idx, 1:].tolist()
            
            final_p, prefix = [], ""
            for p in raw_p:
                ps = str(p).strip() if not pd.isna(p) else ""
                if self.data_mode == "Module" and ps.lower().startswith("cont_"):
                    prefix = ps.split('_')[1]; final_p.append(ps)
                else:
                    final_p.append(f"{prefix}_{ps}" if prefix and self.data_mode == "Module" else ps)
            
            counts = {}; numbered_p = []
            for p in final_p:
                if not p: numbered_p.append(""); continue
                counts[p] = counts.get(p, 0) + 1; numbered_p.append(p)
            cur = {}
            for i_p, p in enumerate(numbered_p):
                if p and counts[p] > 1:
                    cur[p] = cur.get(p, 0) + 1; numbered_p[i_p] = f"{p}{cur[p]}"

            p_dict = {}
            for c_idx, pn in enumerate(numbered_p):
                if not pn or "cont_" in pn.lower(): continue
                un = str(raw_u[c_idx]).strip() if c_idx < len(raw_u) and not pd.isna(raw_u[c_idx]) else ""
                if not un: continue
                if c_idx + 1 >= df.shape[1]: continue
                
                vals = pd.to_numeric(df.iloc[sample_start_row + 1:, c_idx + 1], errors='coerce').tolist()
                vals = vals[:len(units)]
                
                if all(v is None or np.isnan(v) for v in vals): continue
                p_dict[pn] = {'unit': un, 'values': vals, 'units_map': units[:len(vals)]}
                all_p.add(pn)
            
            temp_data[fname] = {'lot': lot, 'ro': ro, 'ro_num': ro_n, 'params': p_dict}
            if lot not in self.lot_groups: self.lot_groups[lot] = []
            self.lot_groups[lot].append(fname)

        if not all_p: raise ValueError("가져온 데이터 세트에 유효한 파라미터가 없습니다.")
        
        self.parameter_list = sorted(list(all_p))
        self.raw_files_data = temp_data
        
        for l in self.lot_groups: 
            self.lot_groups[l].sort(key=lambda x: self.raw_files_data[x]['ro_num'])
            self.lot_display_names[l] = l
            
        self.after(0, self.init_analysis_menu)

    def init_analysis_menu(self):
        for widget in self.winfo_children(): widget.destroy()
        
        t = tk.Frame(self, bg="#f4f4f4", pady=10, padx=10); t.pack(fill=tk.X)
        ctrl_f = tk.LabelFrame(t, text="Analysis Control Panel", font=("Arial", 9, "bold"), bg="#f4f4f4", padx=10)
        ctrl_f.pack(side=tk.RIGHT, padx=10)
        
        tk.Checkbutton(ctrl_f, text="Delta Mode (%)", variable=self.is_delta_mode, bg="#f4f4f4", command=self.execute_ui_rendering).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_f, text="↩ 되돌리기 (Undo)", font=("Arial", 9, "bold"), bg="#7f8c8d", fg="white", command=self.perform_undo).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_f, text="📄 가로 포맷 PDF 리포트 저장", font=("Arial", 9, "bold"), bg="#c0392b", fg="white", command=self.export_to_pdf).pack(side=tk.LEFT, padx=5)

        tk.Label(t, text="Parameter Selector (다중 선택 가능):", font=("Arial", 11, "bold"), bg="#f4f4f4").pack(anchor="w")
        lf = tk.Frame(t); lf.pack(fill=tk.X, pady=5)
        
        self.param_listbox = tk.Listbox(lf, selectmode=tk.EXTENDED, height=6, font=("Consolas", 10))
        self.param_listbox.pack(fill=tk.X, side=tk.LEFT, expand=True)
        sb = ttk.Scrollbar(lf, command=self.param_listbox.yview); sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.param_listbox.config(yscrollcommand=sb.set)
        
        self.param_listbox.insert(tk.END, "★ 전체 선택")
        for p in self.parameter_list: self.param_listbox.insert(tk.END, p)
        self.param_listbox.selection_set(0)

        btn_f = tk.Frame(t, bg="#f4f4f4"); btn_f.pack(fill=tk.X)
        tk.Button(btn_f, text="그래프 그리기", bg="#107c41", fg="white", font=("Arial", 10, "bold"), command=self.update_selections_and_render).pack(side=tk.LEFT, padx=5)
        
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
        self.update_selections_and_render()

    def _on_mouse_wheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def update_selections_and_render(self):
        selections = self.param_listbox.curselection()
        sel_items = [self.param_listbox.get(i) for i in selections]
        if "★ 전체 선택" in sel_items: self.selected_parameters = self.parameter_list.copy()
        else: self.selected_parameters = [v for v in sel_items if v != "★ 전체 선택"]
        if not self.selected_parameters: return
        self.execute_ui_rendering()

    def build_chart_data_structures(self, target_lot):
        lot_files = self.lot_groups[target_lot]
        # 다양한 Read-out 구분을 위한 뚜렷한 색상 정의
        base_colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf', '#e31a1c', '#33a02c']
        
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
                    s_str = str(s_id)
                    all_samples_set.add(s_str)
                    if s_str not in master_map: master_map[s_str] = {}
                    master_map[s_str][ro_lbl] = val

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

            del_set = self.deleted_units.get((target_lot, param), set())
            lines_dataset = []

            # 모든 Read-out을 독립 선으로 추적하여 데이터 누락 방지
            for f_idx, filename in enumerate(lot_files):
                ro_lbl = self.raw_files_data[filename]['ro']
                px, py, pc, pm = [], [], [], []
                
                for s_id in all_samples:
                    if s_id in del_set: continue
                    val = master_map[s_id].get(ro_lbl, np.nan)
                    if pd.isna(val) or np.isinf(val): continue
                    
                    px.append(s_id)
                    py.append(float(val))
                    
                    c_key = (target_lot, param, s_id)
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
            del_set = self.deleted_units.get((target_lot, param), set())

            for f_idx, fn in enumerate(lot_files):
                if param in self.raw_files_data[fn]['params']:
                    p_info = self.raw_files_data[fn]['params'][param]
                    vals = [uy for ux, uy in zip(p_info['units_map'], p_info['values']) if uy is not None and not np.isnan(uy) and str(ux) not in del_set]
                    if vals:
                        b_data.append(vals)
                        a_labels.append(self.raw_files_data[fn]['ro'])
                        b_cols.append(base_colors[f_idx % len(base_colors)])
                        stats.append(f"[{self.raw_files_data[fn]['ro']}]\nAvg:{np.mean(vals):.1f}\nStd:{np.std(vals):.1f}")
            
            if b_data: box_plots_meta.append({'title': f"{param} Dist", 'b_data': b_data, 'a_labels': a_labels, 'b_cols': b_cols, 'stats': stats})

        return line_plots_meta, box_plots_meta

    def execute_ui_rendering(self):
        for w in self.scrollable_frame.winfo_children(): w.destroy()
        
        # Lot 별로 정렬하여 완전히 독립된 단독 그래프 그룹(Section)을 형성
        for lot_key in sorted(self.lot_groups.keys()):
            line_meta, box_meta = self.build_chart_data_structures(lot_key)
            if not line_meta and not box_meta: continue
            
            header_f = tk.Frame(self.scrollable_frame, bg="#eaf2f8", pady=6)
            header_f.pack(fill=tk.X, padx=10, pady=5)
            tk.Label(header_f, text=f"■ [{self.lot_display_names[lot_key]}] Lot Analysis", font=("Arial", 13, "bold"), fg="#1e3799", bg="#eaf2f8").pack(side=tk.LEFT, padx=10)
            
            rename_f = tk.Frame(header_f, bg="#eaf2f8")
            rename_f.pack(side=tk.RIGHT, padx=15)
            ent = tk.Entry(rename_f, width=15, font=("Arial", 9))
            ent.insert(0, self.lot_display_names[lot_key]); ent.pack(side=tk.LEFT, padx=5)
            tk.Button(rename_f, text="변경", font=("Arial", 8, "bold"), bg="#546e7a", fg="white", command=lambda l=lot_key, e=ent: self.update_lot_name(l, e.get())).pack(side=tk.LEFT)
            
            if line_meta:
                grid_frame = tk.Frame(self.scrollable_frame)
                grid_frame.pack(fill=tk.X, padx=15, pady=5)
                cols = 3 if self.data_mode == "Module" else 1
                for c in range(cols): grid_frame.grid_columnconfigure(c, weight=1)
                
                for idx, m in enumerate(line_meta):
                    fig, ax = plt.subplots(figsize=(4.2 if cols==3 else 13.0, 3.2))
                    
                    # 하나의 파라미터 그래프 안에 모든 Read-out 라인을 겹쳐서 플롯
                    for px, py, pc, pm, ro_lbl, b_col in m['dataset']:
                        ax.plot(px, py, color=b_col, alpha=0.6, zorder=1, label=ro_lbl)
                        for xi, yi, ci, mi in zip(px, py, pc, pm):
                            sc = ax.scatter(xi, yi, color=ci, marker=mi, s=55 if mi=='^' else 35, zorder=3, picker=True)
                            sc.__dict__['metadata'] = {'lot': lot_key, 'param': m['param'], 'unit': xi, 'ro': ro_lbl}
                    
                    ax.set_title(f"[{self.lot_display_names[lot_key]}] {m['title']}", fontsize=9, weight='bold')
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
                    ax.set_title(f"[{self.lot_display_names[lot_key]}] {m['title']}", fontsize=9, weight='bold')
                    ax.grid(True, alpha=0.3)
                    plt.tight_layout()
                    
                    bb = tk.Frame(box_frame, bd=1, relief=tk.GROOVE, bg="white")
                    bb.grid(row=idx//4, column=idx%4, padx=5, pady=5, sticky="nsew")
                    FigureCanvasTkAgg(fig, master=bb).get_tk_widget().pack()
                    
                    sf = tk.Frame(bb, bg="#fafafa"); sf.pack(fill=tk.X)
                    for s in m['stats']: tk.Label(sf, text=s, font=("Arial", 7), bg="#fafafa", justify=tk.LEFT).pack(anchor="w", padx=5)
                    plt.close(fig)

    def on_chart_point_clicked(self, event):
        scatter = event.artist
        if 'metadata' not in scatter.__dict__: return
        meta = scatter.__dict__['metadata']
        lot, param, unit_id, ro_info = meta['lot'], meta['param'], meta['unit'], meta['ro']
        
        m = tk.Toplevel(self); m.title("Data Editor"); m.geometry("450x180")
        self.center_window(m, 450, 180); m.transient(self); m.grab_set()
        
        tk.Label(m, text=f"선택 시료 번호: {unit_id} ({ro_info})", font=("Arial", 11, "bold")).pack(pady=10)
        color_section = tk.LabelFrame(m, text="변경할 색상 선택 (선택 시 세모 마커로 자동 변환)", font=("Arial", 9))
        color_section.pack(fill=tk.X, padx=15, pady=5)
        
        distinct_palette = ["#FF0000", "#0026ff", "#00b321", "#9400d3", "#ff8c00"]
        btn_frame = tk.Frame(color_section)
        btn_frame.pack(pady=5)
        
        for hex_code in distinct_palette:
            btn = tk.Button(btn_frame, bg=hex_code, activebackground=hex_code, width=5, height=2, bd=2, relief=tk.RAISED,
                            command=lambda l=lot, p=param, u=unit_id, c=hex_code: [m.destroy(), self.apply_point_color(l, p, u, c)])
            btn.pack(side=tk.LEFT, padx=8)
            
        action_f = tk.Frame(m); action_f.pack(pady=10)
        tk.Button(action_f, text="🗑️ 해당 시료 데이터 삭제 (Box연동)", bg="#2c3e50", fg="white", font=("Arial", 9, "bold"),
                  command=lambda: [m.destroy(), self.delete_target_unit(lot, param, unit_id)]).pack(side=tk.LEFT, padx=10)
        tk.Button(action_f, text="창 닫기", command=m.destroy).pack(side=tk.LEFT, padx=10)

    def apply_point_color(self, lot, param, unit_id, chosen_color):
        c_key = (lot, param, unit_id)
        self.undo_stack.append(('color', c_key, self.custom_colors.get(c_key, None)))
        self.custom_colors[c_key] = chosen_color
        self.execute_ui_rendering()

    def delete_target_unit(self, lot, param, unit_id):
        key = (lot, param)
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

    def update_lot_name(self, lot_key, new_name):
        if not new_name.strip(): return
        self.lot_display_names[lot_key] = new_name.strip()
        self.execute_ui_rendering()

    def export_to_pdf(self):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF 리포트 파일", "*.pdf")])
        if not path: return
        
        # [해결] 프로세스 점유에 따른 Permission Denied 예외 방어 검증 코드 추가
        if os.path.exists(path):
            try:
                with open(path, 'a'): pass
            except IOError:
                messagebox.showerror("수정 불가", "동일한 이름의 PDF 리포트 파일이 이미 열려 있습니다.\n파일을 닫은 다음 다시 시도해 주세요.")
                return

        pb = ProgressBarPop(self, "PDF 리포트 내보내는 중")
        
        def pdf_thread():
            try:
                with PdfPages(path) as pdf:
                    lots = sorted(self.lot_groups.keys())
                    total_steps = len(lots)
                    
                    for step, lot_key in enumerate(lots):
                        self.after(0, pb.update_progress, step, total_steps, f"[{lot_key}] 컴파일 중...")
                        line_meta, box_meta = self.build_chart_data_structures(lot_key)
                        
                        if line_meta:
                            cols = 3 if self.data_mode == "Module" else 1
                            items_per_page = 6 if cols == 3 else 3
                            for i in range(0, len(line_meta), items_per_page):
                                chunk = line_meta[i:i+items_per_page]
                                rows = int(np.ceil(len(chunk) / cols))
                                
                                fig, axes = plt.subplots(rows, cols, figsize=(11, 8.5), squeeze=False)
                                for idx, m in enumerate(chunk):
                                    r, c = idx // cols, idx % cols
                                    ax = axes[r, c]
                                    
                                    for px, py, pc, pm, ro_lbl, b_col in m['dataset']:
                                        ax.plot(px, py, color=b_col, alpha=0.5, label=ro_lbl)
                                        for xi, yi, ci, mi in zip(px, py, pc, pm):
                                            ax.scatter(xi, yi, color=ci, marker=mi, s=40 if mi=='^' else 20)
                                    
                                    ax.set_title(f"[{self.lot_display_names[lot_key]}] {m['title']}", fontsize=8, weight='bold')
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
                            cols, items_per_page = 4, 8
                            for i in range(0, len(box_meta), items_per_page):
                                chunk = box_meta[i:i+items_per_page]
                                rows = int(np.ceil(len(chunk) / cols))
                                
                                fig, axes = plt.subplots(rows, cols, figsize=(11, 8.5), squeeze=False)
                                for idx, m in enumerate(chunk):
                                    r, c = idx // cols, idx % cols
                                    ax = axes[r, c]
                                    
                                    bp = ax.boxplot(m['b_data'], patch_artist=True)
                                    ax.set_xticklabels(m['a_labels'], fontsize=7)
                                    for patch, color in zip(bp['boxes'], m['b_cols']):
                                        patch.set_facecolor(color); patch.set_alpha(0.5)
                                        
                                    ax.set_title(f"[{self.lot_display_names[lot_key]}] {m['title']}", fontsize=8, weight='bold')
                                    ax.grid(True, alpha=0.3)
                                    
                                    stat_str = "\n".join([s.replace('\n', ' ') for s in m['stats']])
                                    ax.text(0.05, -0.4, stat_str, transform=ax.transAxes, fontsize=5, verticalalignment='top')
                                
                                for idx in range(len(chunk), rows * cols): axes[idx // cols, idx % cols].axis('off')
                                plt.tight_layout()
                                pdf.savefig(fig, dpi=200)
                                plt.close(fig)
                                
                self.after(0, lambda: messagebox.showinfo("Success", "가로 포맷의 PDF 리포트 저장이 성공적으로 완료되었습니다."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("PDF Export Error", f"PDF 컴파일 에러:\n{str(e)}"))
            finally:
                self.after(0, pb.destroy)

        threading.Thread(target=pdf_thread, daemon=True).start()

if __name__ == "__main__":
    DataAnalysisApp().mainloop()
