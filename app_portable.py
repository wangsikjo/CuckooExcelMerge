# -*- coding: utf-8 -*-
import os, sys, json, traceback
from tkinter import (BOTH, LEFT, RIGHT, Y, END, SINGLE, EXTENDED, Tk, Frame, Label, Button, Listbox,
                     Scrollbar, filedialog, messagebox, Menu, StringVar, Radiobutton, IntVar)
DND_AVAILABLE = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False
try:
    import win32com.client as win32
    from win32com.client import constants
except Exception:
    win32 = None
    constants = None
APP_TITLE = "Cuckoo ExcelMergeToPdf"
MERGE_MODE_COPY = 1
MERGE_MODE_APPEND = 2
PDF_MODE_SINGLE = 1
PDF_MODE_PER_SHEET = 2
PDF_MODE_PER_FILE = 3
DEFAULTS = {"merge_mode": MERGE_MODE_COPY, "pdf_mode": PDF_MODE_SINGLE, "pdf_dir": None, "last_xlsx_dir": None}
def base_dir():
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
def config_path():
    d = os.path.join(base_dir(), "data"); os.makedirs(d, exist_ok=True)
    return os.path.join(d, "config.json")
def load_config():
    p = config_path()
    if os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f); return {**DEFAULTS, **data}
        except Exception:
            return DEFAULTS.copy()
    return DEFAULTS.copy()
def save_config(cfg):
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
def ensure_excel():
    if win32 is None:
        raise RuntimeError("pywin32가 누락되었습니다. exe로 실행하거나 'pip install pywin32' 필요.")
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        raise RuntimeError("Microsoft Excel이 설치되어 있지 않거나 실행할 수 없습니다.")
def list_sheets(files):
    excel = None; sheets = []
    try:
        excel = ensure_excel(); excel.DisplayAlerts = False; excel.Visible=False
        for p in files:
            wb=None
            try:
                wb=excel.Workbooks.Open(os.path.abspath(p), ReadOnly=True)
                for ws in wb.Worksheets: sheets.append((p, ws.Name))
            finally:
                if wb: wb.Close(SaveChanges=False)
    finally:
        if excel: excel.Quit()
    return sheets
def sanitize_sheet_name(base, sheet):
    nm = f"{base}_{sheet}".replace("/", "-").replace("\", "-")
    return nm[:31] if len(nm) > 31 else nm
def merge_copy_mode(selected_sheets, out_xlsx_path):
    excel=None; wb_out=None; first_sheet=None
    try:
        excel=ensure_excel(); excel.DisplayAlerts=False; excel.Visible=False; excel.ScreenUpdating=False
        wb_out=excel.Workbooks.Add(); first_sheet=wb_out.Worksheets(1)
        grouped={}
        for f,s in selected_sheets: grouped.setdefault(f,[]).append(s)
        for f, sheet_names in grouped.items():
            wb_in=None
            try:
                wb_in=excel.Workbooks.Open(os.path.abspath(f), ReadOnly=True)
                for s in sheet_names:
                    try: ws=wb_in.Worksheets(s)
                    except Exception: continue
                    ws.Copy(After=wb_out.Worksheets(wb_out.Worksheets.Count))
                    new_ws=wb_out.Worksheets(wb_out.Worksheets.Count)
                    base=os.path.splitext(os.path.basename(f))[0]
                    safe=sanitize_sheet_name(base, s)
                    try: new_ws.Name=safe
                    except Exception: pass
            finally:
                if wb_in: wb_in.Close(SaveChanges=False)
        try:
            if wb_out.Worksheets.Count>1 and first_sheet is not None: first_sheet.Delete()
        except Exception: pass
        wb_out.SaveAs(os.path.abspath(out_xlsx_path), FileFormat=51)
    finally:
        if wb_out: wb_out.Close(SaveChanges=False)
        if excel: excel.Quit()
def merge_append_mode(selected_sheets, out_xlsx_path):
    excel=None; wb_out=None
    try:
        excel=ensure_excel(); excel.DisplayAlerts=False; excel.Visible=False; excel.ScreenUpdating=False
        wb_out=excel.Workbooks.Add(); ws_out=wb_out.Worksheets(1); ws_out.Name="통합"
        current_row=1; first=True
        for f,s in selected_sheets:
            wb_in=None
            try:
                wb_in=excel.Workbooks.Open(os.path.abspath(f), ReadOnly=True)
                try: ws_in=wb_in.Worksheets(s)
                except Exception: continue
                used=ws_in.UsedRange; rows=used.Rows.Count; cols=used.Columns.Count
                if first:
                    try:
                        for c in range(1, cols+1): ws_out.Columns(c).ColumnWidth = ws_in.Columns(c).ColumnWidth
                    except Exception: pass
                    first=False
                target=ws_out.Range(ws_out.Cells(current_row,1), ws_out.Cells(current_row+rows-1, cols))
                target.Value=used.Value
                current_row += rows + 1
            finally:
                if wb_in: wb_in.Close(SaveChanges=False)
        wb_out.SaveAs(os.path.abspath(out_xlsx_path), FileFormat=51)
    finally:
        if wb_out: wb_out.Close(SaveChanges=False)
        if excel: excel.Quit()
def export_pdf_single(xlsx_path, pdf_path):
    excel=None; wb=None
    try:
        excel=ensure_excel(); excel.DisplayAlerts=False; excel.Visible=False; excel.ScreenUpdating=False
        wb=excel.Workbooks.Open(os.path.abspath(xlsx_path), ReadOnly=True)
        wb.Worksheets.Select()
        try:
            for sht in wb.Worksheets:
                sht.PageSetup.Zoom=False; sht.PageSetup.FitToPagesWide=1; sht.PageSetup.FitToPagesTall=False
        except Exception: pass
        wb.ExportAsFixedFormat(constants.xlTypePDF, os.path.abspath(pdf_path))
    finally:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()
def export_pdf_per_sheet(xlsx_path, out_dir):
    excel=None; wb=None; created=[]
    try:
        excel=ensure_excel(); excel.DisplayAlerts=False; excel.Visible=False; excel.ScreenUpdating=False
        wb=excel.Workbooks.Open(os.path.abspath(xlsx_path), ReadOnly=True)
        for sht in wb.Worksheets:
            try:
                sht.Select()
                try:
                    sht.PageSetup.Zoom=False; sht.PageSetup.FitToPagesWide=1; sht.PageSetup.FitToPagesTall=False
                except Exception: pass
                out_path=os.path.join(out_dir, f"{sht.Name}.pdf")
                wb.ExportAsFixedFormat(constants.xlTypePDF, os.path.abspath(out_path))
                created.append(out_path)
            except Exception: continue
    finally:
        if wb: wb.Close(SaveChanges=False)
        if excel: excel.Quit()
    return created
def export_pdf_per_file(selected_sheets, out_dir):
    excel=None; created=[]
    try:
        excel=ensure_excel(); excel.DisplayAlerts=False; excel.Visible=False; excel.ScreenUpdating=False
        grouped={}
        for f,s in selected_sheets: grouped.setdefault(f,[]).append(s)
        for f, sheet_names in grouped.items():
            wb_out=None; first_sheet=None; wb_in=None
            try:
                wb_out=excel.Workbooks.Add(); first_sheet=wb_out.Worksheets(1)
                wb_in=excel.Workbooks.Open(os.path.abspath(f), ReadOnly=True)
                for s in sheet_names:
                    try: ws=wb_in.Worksheets(s)
                    except Exception: continue
                    ws.Copy(After=wb_out.Worksheets(wb_out.Worksheets.Count))
                try:
                    if wb_out.Worksheets.Count>1 and first_sheet is not None: first_sheet.Delete()
                except Exception: pass
                for sht in wb_out.Worksheets:
                    try:
                        sht.PageSetup.Zoom=False; sht.PageSetup.FitToPagesWide=1; sht.PageSetup.FitToPagesTall=False
                    except Exception: pass
                out_pdf=os.path.join(out_dir, os.path.splitext(os.path.basename(f))[0] + ".pdf")
                wb_out.Worksheets.Select()
                wb_out.ExportAsFixedFormat(constants.xlTypePDF, os.path.abspath(out_pdf))
                created.append(out_pdf)
            finally:
                if wb_in: wb_in.Close(SaveChanges=False)
                if wb_out: wb_out.Close(SaveChanges=False)
    finally:
        if excel: excel.Quit()
    return created
class App:
    def __init__(self, master):
        self.master=master; master.title(APP_TITLE)
        self.cfg=load_config()
        self.files=[]; self.available=[]; self.selected=[]; self.output_xlsx=None
        self.pdf_dir=self.cfg.get("pdf_dir"); self.pdf_dir_var=StringVar(value=self.pdf_dir or "(미설정)")
        self.merge_mode=IntVar(value=self.cfg.get("merge_mode", MERGE_MODE_COPY))
        self.pdf_mode=IntVar(value=self.cfg.get("pdf_mode", PDF_MODE_SINGLE))
        Label(master, text="1) 엑셀 파일 추가 — 드래그앤드롭 가능").pack(anchor="w", padx=10, pady=(10,0))
        top=Frame(master); top.pack(fill=BOTH, expand=True, padx=10)
        self.listbox=Listbox(top, selectmode=SINGLE); self.listbox.pack(side=LEFT, fill=BOTH, expand=True)
        sb=Scrollbar(top, orient="vertical", command=self.listbox.yview); sb.pack(side=RIGHT, fill=Y)
        self.listbox.config(yscrollcommand=sb.set)
        if DND_AVAILABLE and isinstance(master, TkinterDnD.Tk):
            self.listbox.drop_target_register(DND_FILES); self.listbox.dnd_bind('<<Drop>>', self.on_drop_files)
        btns=Frame(master); btns.pack(anchor="w", padx=10, pady=5)
        Button(btns, text="엑셀 파일 추가", command=self.add_files).pack(side=LEFT, padx=(0,5))
        Button(btns, text="선택 항목 제거", command=self.remove_selected_file).pack(side=LEFT, padx=(0,5))
        Button(btns, text="전체 비우기", command=self.clear_all_files).pack(side=LEFT, padx=(0,5))
        Button(btns, text="시트 목록 불러오기", command=self.load_sheets).pack(side=LEFT)
        self.ctx_files=Menu(self.listbox, tearoff=0); self.ctx_files.add_command(label="선택 항목 제거", command=self.remove_selected_file)
        self.listbox.bind("<Button-3>", self.popup_files_ctx); self.listbox.bind("<Delete>", lambda e: self.remove_selected_file())
        sheet_frame=Frame(master); sheet_frame.pack(fill=BOTH, expand=True, padx=10, pady=(10,0))
        Label(sheet_frame, text="2) 시트 선택 및 순서 지정").grid(row=0, column=0, sticky="w")
        Label(sheet_frame, text="사용 가능(왼쪽) → 선택(오른쪽)").grid(row=0, column=1, sticky="w")
        self.lb_available=Listbox(sheet_frame, selectmode=EXTENDED); self.lb_available.grid(row=1, column=0, sticky="nsew", padx=(0,5))
        self.lb_selected=Listbox(sheet_frame, selectmode=EXTENDED); self.lb_selected.grid(row=1, column=1, sticky="nsew")
        sheet_frame.grid_columnconfigure(0, weight=1); sheet_frame.grid_columnconfigure(1, weight=1); sheet_frame.grid_rowconfigure(1, weight=1)
        ctrl=Frame(sheet_frame); ctrl.grid(row=1, column=2, padx=(8,0), sticky="ns")
        Button(ctrl, text="추가 ▶", command=self.add_selected_sheets).pack(pady=2, fill="x")
        Button(ctrl, text="◀ 제거", command=self.remove_selected_sheets).pack(pady=2, fill="x")
        Button(ctrl, text="위로 ↑", command=lambda: self.move_selected(-1)).pack(pady=10, fill="x")
        Button(ctrl, text="아래로 ↓", command=lambda: self.move_selected(1)).pack(pady=2, fill="x")
        Button(ctrl, text="모두 제거", command=self.clear_selected_sheets).pack(pady=(10,2), fill="x")
        mm=Frame(master); mm.pack(anchor="w", padx=10, pady=(10,0))
        Label(mm, text="3) 병합 모드: ").pack(side=LEFT)
        Radiobutton(mm, text="시트 복사 (서식 보존, 기본)", variable=self.merge_mode, value=MERGE_MODE_COPY).pack(side=LEFT)
        Radiobutton(mm, text="데이터 이어붙이기 (한 시트)", variable=self.merge_mode, value=MERGE_MODE_APPEND).pack(side=LEFT)
        af=Frame(master); af.pack(fill=BOTH, expand=False, padx=10, pady=(10,0))
        Label(af, text="4) PDF 기본 저장 폴더: ").pack(side=LEFT); Label(af, textvariable=self.pdf_dir_var).pack(side=LEFT)
        Button(af, text="변경…", command=self.choose_pdf_dir).pack(side=LEFT, padx=8)
        pm=Frame(master); pm.pack(anchor="w", padx=10, pady=(6,0))
        Label(pm, text="PDF 출력 방식: ").pack(side=LEFT)
        Radiobutton(pm, text="통합 PDF 1개", variable=self.pdf_mode, value=PDF_MODE_SINGLE).pack(side=LEFT)
        Radiobutton(pm, text="시트별 개별 PDF", variable=self.pdf_mode, value=PDF_MODE_PER_SHEET).pack(side=LEFT)
        Radiobutton(pm, text="원본 파일별 PDF", variable=self.pdf_mode, value=PDF_MODE_PER_FILE).pack(side=LEFT)
        Button(master, text="통합 엑셀 만들기", command=self.merge_only).pack(padx=10, pady=(10,4))
        Button(master, text="PDF 만들기", command=self.make_pdf).pack(padx=10, pady=(0,12))
        Button(master, text="한 번에: 통합 엑셀 + PDF", command=self.merge_and_pdf).pack(padx=10, pady=(0,14))
        if not DND_AVAILABLE:
            messagebox.showinfo("드래그앤드롭 안내","드래그앤드롭을 사용하려면 tkinterdnd2가 필요합니다.
지금도 '엑셀 파일 추가' 버튼으로 파일 선택은 가능합니다.")
    def on_drop_files(self, event):
        raw=event.data; paths=[]; buf=""; in_brace=False
        for ch in raw:
            if ch=="{": in_brace=True; buf=""
            elif ch=="}": in_brace=False; paths.append(buf); buf=""
            elif ch==" " and not in_brace:
                if buf: paths.append(buf); buf=""
            else: buf += ch
        if buf: paths.append(buf)
        valid_ext=(".xlsx",".xls",".xlsm")
        for p in paths:
            if os.path.isdir(p):
                for r,_,fs in os.walk(p):
                    for fn in fs:
                        if fn.lower().endswith(valid_ext): self._add_file(os.path.join(r,fn))
            else:
                if p.lower().endswith(valid_ext): self._add_file(p)
    def _add_file(self, p):
        if p not in self.files:
            self.files.append(p); self.listbox.insert(END, p)
    def add_files(self):
        initialdir=self.cfg.get("last_xlsx_dir") or base_dir()
        paths=filedialog.askopenfilenames(title="엑셀 파일 선택", filetypes=[("Excel Files","*.xlsx *.xls *.xlsm")], initialdir=initialdir)
        if paths:
            self.cfg["last_xlsx_dir"]=os.path.dirname(paths[0]); save_config(self.cfg)
            for p in paths: self._add_file(p)
    def remove_selected_file(self):
        sel=self.listbox.curselection()
        if not sel: return
        idx=sel[0]; self.listbox.delete(idx); del self.files[idx]
    def clear_all_files(self):
        self.listbox.delete(0, END); self.files=[]; self.available=[]; self.selected=[]
        try:
            self.lb_available.delete(0, END); self.lb_selected.delete(0, END)
        except Exception: pass
    def popup_files_ctx(self, event):
        try:
            self.listbox.selection_clear(0, END); idx=self.listbox.nearest(event.y)
            self.listbox.selection_set(idx); self.ctx_files.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx_files.grab_release()
    def load_sheets(self):
        if not self.files:
            messagebox.showwarning("안내","먼저 파일을 추가하세요."); return
        try:
            self.available=list_sheets(self.files); self.lb_available.delete(0, END)
            for f,s in self.available: self.lb_available.insert(END, f"{os.path.basename(f)} :: {s}")
            messagebox.showinfo("완료", f"시트 {len(self.available)}개를 불러왔습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"{e}

{traceback.format_exc()}")
    def add_selected_sheets(self):
        sel=self.lb_available.curselection()
        if not sel: return
        for idx in sel:
            pair=self.available[idx]
            if pair not in self.selected:
                self.selected.append(pair); self.lb_selected.insert(END, f"{os.path.basename(pair[0])} :: {pair[1]}")
    def remove_selected_sheets(self):
        sel=list(self.lb_selected.curselection())
        if not sel: return
        sel.sort(reverse=True)
        for i in sel:
            del self.selected[i]; self.lb_selected.delete(i)
    def clear_selected_sheets(self):
        self.selected=[]; self.lb_selected.delete(0, END)
    def move_selected(self, direction):
        sel=self.lb_selected.curselection()
        if not sel: return
        index=sel[0]; new_index=index+direction
        if new_index<0 or new_index>=len(self.selected): return
        self.selected[index], self.selected[new_index] = self.selected[new_index], self.selected[index]
        self.lb_selected.delete(0, END)
        for f,s in self.selected: self.lb_selected.insert(END, f"{os.path.basename(f)} :: {s}")
        self.lb_selected.selection_set(new_index)
    def choose_pdf_dir(self):
        d=filedialog.askdirectory(title="PDF 기본 저장 폴더 선택", initialdir=self.pdf_dir or base_dir())
        if d:
            self.pdf_dir=d; self.pdf_dir_var.set(self.pdf_dir); self.cfg["pdf_dir"]=d; save_config(self.cfg)
    def _ask_out_xlsx(self):
        initdir=self.cfg.get("last_xlsx_dir") or base_dir()
        return filedialog.asksaveasfilename(title="통합 엑셀 저장 위치", defaultextension=".xlsx",
            filetypes=[("Excel Workbook","*.xlsx")], initialdir=initdir)
    def _default_pdf_dir(self, base_path):
        if self.pdf_dir and os.path.isdir(self.pdf_dir): return self.pdf_dir
        return os.path.dirname(base_path)
    def merge_only(self):
        if not self.selected:
            messagebox.showwarning("안내","먼저 시트를 선택하세요."); return
        out_xlsx=self._ask_out_xlsx()
        if not out_xlsx: return
        try:
            if self.merge_mode.get()==MERGE_MODE_COPY: merge_copy_mode(self.selected,out_xlsx)
            else: merge_append_mode(self.selected,out_xlsx)
            self.output_xlsx=out_xlsx; self.cfg["last_xlsx_dir"]=os.path.dirname(out_xlsx); save_config(self.cfg)
            messagebox.showinfo("완료", f"통합 엑셀 생성 완료:
{out_xlsx}")
        except Exception as e:
            messagebox.showerror("오류", f"{e}

{traceback.format_exc()}")
    def make_pdf(self):
        if not self.output_xlsx or not os.path.exists(self.output_xlsx):
            messagebox.showwarning("안내","먼저 '통합 엑셀 만들기'를 실행하세요."); return
        try:
            if self.pdf_mode.get()==PDF_MODE_SINGLE:
                initialdir=self._default_pdf_dir(self.output_xlsx)
                initialfile=os.path.splitext(os.path.basename(self.output_xlsx))[0]+".pdf"
                pdf_path=filedialog.asksaveasfilename(title="PDF 저장 위치", defaultextension=".pdf",
                    filetypes=[("PDF","*.pdf")], initialdir=initialdir, initialfile=initialfile)
                if not pdf_path: return
                export_pdf_single(self.output_xlsx, pdf_path)
                messagebox.showinfo("완료", f"PDF 저장 성공:
{pdf_path}")
            elif self.pdf_mode.get()==PDF_MODE_PER_SHEET:
                out_dir=filedialog.askdirectory(title="시트별 PDF 저장 폴더 선택", initialdir=self._default_pdf_dir(self.output_xlsx))
                if not out_dir: return
                created=export_pdf_per_sheet(self.output_xlsx, out_dir)
                messagebox.showinfo("완료", f"시트별 PDF {len(created)}개 저장 완료:
{out_dir}")
            else:
                out_dir=filedialog.askdirectory(title="파일별 PDF 저장 폴더 선택", initialdir=self._default_pdf_dir(self.output_xlsx))
                if not out_dir: return
                created=export_pdf_per_file(self.selected, out_dir)
                messagebox.showinfo("완료", f"파일별 PDF {len(created)}개 저장 완료:
{out_dir}")
        except Exception as e:
            messagebox.showerror("오류", f"{e}

{traceback.format_exc()}")
    def merge_and_pdf(self):
        if not self.selected:
            messagebox.showwarning("안내","먼저 시트를 선택하세요."); return
        out_xlsx=self._ask_out_xlsx()
        if not out_xlsx: return
        try:
            if self.merge_mode.get()==MERGE_MODE_COPY: merge_copy_mode(self.selected,out_xlsx)
            else: merge_append_mode(self.selected,out_xlsx)
            self.output_xlsx=out_xlsx; self.cfg["last_xlsx_dir"]=os.path.dirname(out_xlsx); save_config(self.cfg)
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 병합 실패: {e}

{traceback.format_exc()}"); return
        self.make_pdf()
def main():
    root = TkinterDnD.Tk() if DND_AVAILABLE else Tk()
    root.title(APP_TITLE); root.geometry("940x700")
    App(root); root.mainloop()
if __name__ == "__main__":
    main()
