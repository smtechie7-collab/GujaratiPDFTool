"""
GujaratiAllInOneGUI_v2.py
All-in-one: Convert PDFs -> Searchable PDFs (Tesseract) + Search inside PDFs (PyMuPDF or OCR fallback)
Features:
 - Tabs: Convert / Search / Settings
 - Improved progress (file + page level) + status label
 - Results Table (Treeview) with Export CSV/XLSX
 - Dark mode toggle (simple)
 - One-click build batch file provided separately
Author: SM TECHIE (adapted)
"""

import os
import subprocess
import threading
import unicodedata
import tempfile
import shutil
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from PIL import Image, ImageTk
import pandas as pd
from rapidfuzz import fuzz
from PyPDF2 import PdfMerger

# -----------------------------
# DEFAULT CONFIG - edit if needed
# -----------------------------
DEFAULT_TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
DEFAULT_POPPLER_PATH = r"C:\Program Files\poppler-25.07.0\Library\bin"
DEFAULT_OCR_LANG = "guj+eng"
DEFAULT_DPI = 250
DEFAULT_FUZZY = 70
DEFAULT_OUTPUT_DIRNAME = "_searchable"

# -----------------------------
# Helpers
# -----------------------------
def normalize_text(s):
    if not s:
        return ""
    t = unicodedata.normalize("NFC", s)
    t = t.replace("\u200c", "").replace("\u200d", "")
    return t.strip()

def fuzzy_score(a, b):
    try:
        return fuzz.partial_ratio(a.lower(), b.lower())
    except Exception:
        return 0

def fuzzy_match(line, term, threshold):
    ln = normalize_text(line)
    tm = normalize_text(term)
    if tm in ln:
        return True
    return fuzzy_score(tm, ln) >= threshold

def run_tesseract_on_image(img_path, out_base, tesseract_cmd, lang):
    cmd = f'"{tesseract_cmd}" "{img_path}" "{out_base}" -l {lang} pdf'
    subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

def make_searchable_from_images(img_files, out_pdf_path, tesseract_cmd, lang):
    temp_pdfs = []
    for img in img_files:
        base = img.rsplit(".", 1)[0] + "_ocr"
        run_tesseract_on_image(img, base, tesseract_cmd, lang)
        temp_pdfs.append(base + ".pdf")
    merger = PdfMerger()
    for p in temp_pdfs:
        if os.path.exists(p):
            merger.append(p)
    merger.write(out_pdf_path)
    merger.close()

# -----------------------------
# GUI App
# -----------------------------
class GujaratiAllInOneGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gujarati PDF Converter & Search — SM TECHIE")
        self.root.geometry("1100x740")

        # vars
        self.input_folder = StringVar()
        self.output_folder = StringVar()
        self.tesseract_cmd = StringVar(value=DEFAULT_TESSERACT_CMD)
        self.poppler_path = StringVar(value=DEFAULT_POPPLER_PATH)
        self.ocr_lang = StringVar(value=DEFAULT_OCR_LANG)
        self.dpi = IntVar(value=DEFAULT_DPI)
        self.fuzzy = IntVar(value=DEFAULT_FUZZY)
        self.search_terms = StringVar()
        self.create_subfolder = BooleanVar(value=True)
        self.dark_mode = BooleanVar(value=False)

        # results list
        self.results = []

        self._build_ui()
        self.apply_style()  # apply initial style

    def _build_ui(self):
        # Notebook tabs
        nb = ttk.Notebook(self.root)
        nb.pack(fill=BOTH, expand=True, padx=8, pady=8)

        # --- Tab: Convert ---
        tab_convert = Frame(nb)
        nb.add(tab_convert, text="Convert -> Searchable PDFs")

        frm_cv_top = Frame(tab_convert, pady=6)
        frm_cv_top.pack(fill=X)
        Label(frm_cv_top, text="Input Folder:").grid(row=0, column=0, sticky=W)
        Entry(frm_cv_top, textvariable=self.input_folder, width=70).grid(row=0, column=1, padx=6)
        Button(frm_cv_top, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=6)

        Label(frm_cv_top, text="Output Folder:").grid(row=1, column=0, sticky=W)
        Entry(frm_cv_top, textvariable=self.output_folder, width=70).grid(row=1, column=1, padx=6)
        Button(frm_cv_top, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=6)

        Checkbutton(frm_cv_top, text="Create _searchable subfolder", variable=self.create_subfolder).grid(row=2, column=1, sticky=W, pady=4)

        btn_frame = Frame(tab_convert)
        btn_frame.pack(fill=X, pady=6)
        Button(btn_frame, text="Convert Folder → Searchable PDFs", bg="#1976D2", fg="white", command=self.start_convert).pack(side=LEFT, padx=6)
        Button(btn_frame, text="Convert PDFs → Images (for debug)", command=self.start_convert_images_only).pack(side=LEFT, padx=6)

        # --- Tab: Search ---
        tab_search = Frame(nb)
        nb.add(tab_search, text="Search Inside PDFs")

        frm_search_top = Frame(tab_search, pady=6)
        frm_search_top.pack(fill=X)
        Label(frm_search_top, text="Scan Folder (search will check PDFs inside):").grid(row=0, column=0, sticky=W)
        Entry(frm_search_top, textvariable=self.input_folder, width=70).grid(row=0, column=1, padx=6)
        Button(frm_search_top, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=6)

        Label(frm_search_top, text="Search Terms (comma-separated):").grid(row=1, column=0, sticky=W)
        Entry(frm_search_top, textvariable=self.search_terms, width=60).grid(row=1, column=1, padx=6)
        Button(frm_search_top, text="Start Search", bg="#2E7D32", fg="white", command=self.start_search_thread).grid(row=1, column=2, padx=6)

        # progress + status
        self.progress = ttk.Progressbar(tab_search, length=900, mode="determinate")
        self.progress.pack(pady=6)
        self.status_label = Label(tab_search, text="Ready", anchor="w")
        self.status_label.pack(fill=X, padx=8)

        # results treeview
        cols = ["PDF File", "Page", "Matched Term", "Matched Line", "Context (3 lines)", "Mode"]
        self.tree = ttk.Treeview(tab_search, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree.heading(c, text=c)
            width = 120 if c != "Context (3 lines)" else 400
            self.tree.column(c, width=width, anchor=W)
        self.tree.pack(fill=BOTH, expand=True, padx=8, pady=8)

        # results buttons
        frm_results = Frame(tab_search)
        frm_results.pack(fill=X, pady=6)
        Button(frm_results, text="Export CSV", command=self.export_csv).pack(side=LEFT, padx=6)
        Button(frm_results, text="Export Excel", command=self.export_excel).pack(side=LEFT, padx=6)
        Button(frm_results, text="Clear Results", command=self.clear_results).pack(side=LEFT, padx=6)
        Button(frm_results, text="Copy Selected Context", command=self.copy_selected_context).pack(side=LEFT, padx=6)

        # --- Tab: Settings ---
        tab_settings = Frame(nb)
        nb.add(tab_settings, text="Settings")

        frm_set = Frame(tab_settings, pady=10)
        frm_set.pack(fill=X)
        Label(frm_set, text="Tesseract Executable (tesseract.exe):").grid(row=0, column=0, sticky=W)
        Entry(frm_set, textvariable=self.tesseract_cmd, width=70).grid(row=0, column=1, padx=6)
        Button(frm_set, text="Locate", command=self.locate_tesseract).grid(row=0, column=2, padx=6)

        Label(frm_set, text="Poppler bin folder (pdftoppm.exe):").grid(row=1, column=0, sticky=W)
        Entry(frm_set, textvariable=self.poppler_path, width=70).grid(row=1, column=1, padx=6)
        Button(frm_set, text="Locate", command=self.locate_poppler).grid(row=1, column=2, padx=6)

        Label(frm_set, text="OCR Lang (Tesseract):").grid(row=2, column=0, sticky=W)
        Entry(frm_set, textvariable=self.ocr_lang, width=20).grid(row=2, column=1, sticky=W, padx=6)

        Label(frm_set, text="DPI for conversion:").grid(row=3, column=0, sticky=W)
        Entry(frm_set, textvariable=self.dpi, width=6).grid(row=3, column=1, sticky=W, padx=6)

        Label(frm_set, text="Fuzzy threshold (%):").grid(row=4, column=0, sticky=W)
        Entry(frm_set, textvariable=self.fuzzy, width=6).grid(row=4, column=1, sticky=W, padx=6)

        Checkbutton(frm_set, text="Dark mode", variable=self.dark_mode, command=self.toggle_dark).grid(row=5, column=1, sticky=W, pady=8)

        # bottom log area
        self.log = Text(self.root, height=6, bg="#f4f4f4")
        self.log.pack(fill=X, padx=8, pady=(0,8))

    # Utility UI helpers
    def browse_input(self):
        d = filedialog.askdirectory()
        if d:
            self.input_folder.set(d)

    def browse_output(self):
        d = filedialog.askdirectory()
        if d:
            self.output_folder.set(d)

    def locate_tesseract(self):
        p = filedialog.askopenfilename(title="Locate tesseract.exe", filetypes=[("exe","*.exe")])
        if p:
            self.tesseract_cmd.set(p)

    def locate_poppler(self):
        d = filedialog.askdirectory(title="Locate poppler bin folder")
        if d:
            self.poppler_path.set(d)

    def log_print(self, *args):
        s = " ".join(str(a) for a in args)
        self.log.insert(END, s + "\n")
        self.log.see(END)
        print(s)

    def clear_results(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.results = []
        self.progress["value"] = 0
        self.status_label.config(text="Ready")
        self.log_print("Results cleared.")

    def export_csv(self):
        if not self.results:
            messagebox.showinfo("Info", "No results to export.")
            return
        df = pd.DataFrame(self.results)
        out = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV","*.csv")], initialfile="search_results.csv")
        if out:
            df.to_csv(out, index=False, encoding="utf-8-sig")
            messagebox.showinfo("Saved", f"CSV saved: {out}")

    def export_excel(self):
        if not self.results:
            messagebox.showinfo("Info", "No results to export.")
            return
        df = pd.DataFrame(self.results)
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")], initialfile="search_results.xlsx")
        if out:
            df.to_excel(out, index=False)
            messagebox.showinfo("Saved", f"Excel saved: {out}")

    def copy_selected_context(self):
        sels = self.tree.selection()
        if not sels:
            messagebox.showinfo("Info", "No rows selected.")
            return
        lines = []
        for s in sels:
            vals = self.tree.item(s)["values"]
            ctx = vals[4] if len(vals) > 4 else ""
            lines.append(str(ctx))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n\n".join(lines))
        messagebox.showinfo("Copied", "Selected context copied to clipboard.")

    # ----------------- Conversion Flow -----------------
    def start_convert(self):
        thr = threading.Thread(target=self._convert_folder_thread, daemon=True)
        thr.start()

    def start_convert_images_only(self):
        thr = threading.Thread(target=self._convert_images_only_thread, daemon=True)
        thr.start()

    def _convert_images_only_thread(self):
        inp = self.input_folder.get().strip()
        out = self.output_folder.get().strip()
        poppler = self.poppler_path.get().strip()
        dpi = int(self.dpi.get() or DEFAULT_DPI)
        if not inp or not os.path.isdir(inp) or not out:
            messagebox.showerror("Error", "Select valid input and output folders.")
            return
        self.progress["value"] = 0
        pdfs = [f for f in os.listdir(inp) if f.lower().endswith(".pdf")]
        self.progress["maximum"] = len(pdfs)
        cnt = 0
        for f in pdfs:
            try:
                src = os.path.join(inp, f)
                name = os.path.splitext(f)[0]
                out_dir = os.path.join(out, name)
                os.makedirs(out_dir, exist_ok=True)
                pages = convert_from_path(src, dpi=dpi, poppler_path=poppler)
                for i, p in enumerate(pages, start=1):
                    out_file = os.path.join(out_dir, f"{name}_page_{i}.png")
                    p.save(out_file, "PNG")
                self.log_print(f"Converted {f} -> {len(pages)} images.")
            except Exception as e:
                self.log_print("Error converting", f, ":", e)
            cnt += 1
            self.progress["value"] = cnt
        self.log_print("Image conversion done.")

    def _convert_folder_thread(self):
        inp = self.input_folder.get().strip()
        out = self.output_folder.get().strip()
        poppler = self.poppler_path.get().strip()
        tess = self.tesseract_cmd.get().strip()
        lang = self.ocr_lang.get().strip()
        dpi = int(self.dpi.get() or DEFAULT_DPI)

        if not inp or not os.path.isdir(inp) or not out or not os.path.isdir(out):
            messagebox.showerror("Error", "Select valid input and output folders.")
            return
        if not tess or not os.path.isfile(tess):
            messagebox.showerror("Error", "Provide valid tesseract.exe path in Settings.")
            return

        pdfs = [f for f in os.listdir(inp) if f.lower().endswith(".pdf")]
        if not pdfs:
            messagebox.showinfo("Info", "No PDF files found in input folder.")
            return

        # prepare output base
        base_out = out
        if self.create_subfolder.get():
            base_out = os.path.join(out, DEFAULT_OUTPUT_DIRNAME)
            os.makedirs(base_out, exist_ok=True)

        self.progress["maximum"] = len(pdfs)
        c = 0
        for f in pdfs:
            c += 1
            self.status_label.config(text=f"Converting {f} ({c}/{len(pdfs)})")
            src = os.path.join(inp, f)
            name = os.path.splitext(f)[0]
            temp_dir = os.path.join(base_out, name + "_imgs")
            os.makedirs(temp_dir, exist_ok=True)
            try:
                # step1: PDF -> images
                pages = convert_from_path(src, dpi=dpi, poppler_path=poppler)
                img_files = []
                for i, p in enumerate(pages, start=1):
                    out_img = os.path.join(temp_dir, f"{name}_page_{i}.png")
                    p.save(out_img, "PNG")
                    img_files.append(out_img)
                # step2: OCR images -> per-page pdfs -> merge
                out_pdf_final = os.path.join(base_out, name + "_searchable.pdf")
                make_searchable_from_images(img_files, out_pdf_final, tess, lang)
                self.log_print(f"Converted: {f} -> {out_pdf_final}")
                # cleanup temp per-page PDFs and images if wanted (commented)
                # shutil.rmtree(temp_dir)
            except Exception as e:
                self.log_print("Error processing", f, ":", e)
            self.progress["value"] = c

        self.status_label.config(text="Conversion completed.")
        self.log_print("All PDFs converted to searchable PDFs.")

    # ----------------- Search Flow -----------------
    def start_search_thread(self):
        thr = threading.Thread(target=self._search_thread, daemon=True)
        thr.start()

    def _search_thread(self):
        folder = self.input_folder.get().strip()
        terms_raw = self.search_terms.get().strip()
        poppler = self.poppler_path.get().strip()
        tess = self.tesseract_cmd.get().strip()
        dpi = int(self.dpi.get() or DEFAULT_DPI)
        fuzzy_thr = int(self.fuzzy.get() or DEFAULT_FUZZY)
        lang = self.ocr_lang.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Error", "Select valid folder to search.")
            return
        if not terms_raw:
            messagebox.showerror("Error", "Enter search terms (comma separated).")
            return

        keywords = [k.strip() for k in terms_raw.split(",") if k.strip()]
        # reset
        self.clear_results()

        # collect files (pdf + images)
        file_list = []
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith((".pdf", ".png", ".jpg", ".jpeg")):
                    file_list.append(os.path.join(root, f))

        total = len(file_list)
        if total == 0:
            messagebox.showinfo("Info", "No searchable files found in folder.")
            return

        self.progress["maximum"] = total
        processed = 0
        for fp in file_list:
            processed += 1
            self.progress["value"] = processed
            self.status_label.config(text=f"Scanning: {os.path.basename(fp)} ({processed}/{total})")
            try:
                if fp.lower().endswith(".pdf"):
                    # try PyMuPDF text extraction first
                    try:
                        doc = fitz.open(fp)
                        if doc.page_count == 0:
                            doc.close()
                            continue
                        for pno in range(doc.page_count):
                            page_text = normalize_text(doc[pno].get_text("text") or "")
                            if page_text.strip():
                                self._search_text_in_doc(fp, pno+1, page_text, keywords, "searchable", fuzzy_thr)
                            else:
                                # fallback OCR page
                                page_txt = self._ocr_pdf_page(fp, pno+1, poppler, tess, lang, dpi)
                                self._search_text_in_doc(fp, pno+1, page_txt, keywords, "ocr", fuzzy_thr)
                        doc.close()
                    except Exception:
                        # fallback: convert each page to image and OCR
                        doc = fitz.open(fp)
                        for pno in range(doc.page_count):
                            page_txt = self._ocr_pdf_page(fp, pno+1, poppler, tess, lang, dpi)
                            self._search_text_in_doc(fp, pno+1, page_txt, keywords, "ocr", fuzzy_thr)
                        doc.close()
                else:
                    # image file
                    txt = self._ocr_image(fp, tess, lang)
                    self._search_text_in_doc(fp, None, txt, keywords, "image", fuzzy_thr)
            except Exception as e:
                self.log_print("Scan error:", fp, e)

        self.status_label.config(text="Search completed.")
        self.log_print("Search done. Matches:", len(self.results))

    def _search_text_in_doc(self, file_path, page_no, text, keywords, mode, fuzzy_threshold):
        if not text:
            return
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        for idx, line in enumerate(lines):
            for kw in keywords:
                if fuzzy_match(line, kw, fuzzy_threshold):
                    low = max(0, idx-1)
                    high = min(len(lines)-1, idx+1)
                    context = " | ".join(lines[low:high+1])
                    rec = {
                        "PDF File": os.path.basename(file_path),
                        "Page": page_no if page_no else "",
                        "Matched Term": kw,
                        "Matched Line": line,
                        "Context (3 lines)": context,
                        "Mode": mode
                    }
                    self.results.append(rec)
                    self.tree.insert("", END, values=[rec[c] for c in ["PDF File","Page","Matched Term","Matched Line","Context (3 lines)","Mode"]])

    def _ocr_image(self, img_path, tess, lang):
        try:
            if tess:
                pytesseract.pytesseract.tesseract_cmd = tess
            img = Image.open(img_path)
            txt = pytesseract.image_to_string(img, lang=lang)
            img.close()
            return normalize_text(txt)
        except Exception as e:
            return ""

    def _ocr_pdf_page(self, pdf_path, page_no, poppler, tess, lang, dpi):
        try:
            images = convert_from_path(pdf_path, first_page=page_no, last_page=page_no, dpi=dpi, poppler_path=poppler)
            if images:
                if tess:
                    pytesseract.pytesseract.tesseract_cmd = tess
                txt = pytesseract.image_to_string(images[0], lang=lang)
                images[0].close()
                return normalize_text(txt)
        except Exception as e:
            return ""
        return ""

    # ----------------- UI style / Dark mode -----------------
    def apply_style(self):
        style = ttk.Style(self.root)
        # platform default theme mapping
        try:
            style.theme_use('clam')
        except Exception:
            pass
        if self.dark_mode.get():
            style.configure(".", background="#2b2b2b", foreground="#e6e6e6", fieldbackground="#3a3a3a")
            self.log.config(bg="#1e1e1e", fg="#e6e6e6")
            self.tree.tag_configure('odd', background='#2e2e2e')
        else:
            style.configure(".", background="#f0f0f0", foreground="#000000", fieldbackground="#ffffff")
            self.log.config(bg="#f4f4f4", fg="#000000")

    def toggle_dark(self):
        self.apply_style()

# -----------------------------
# MAIN
# -----------------------------
if __name__ == "__main__":
    root = Tk()
    app = GujaratiAllInOneGUI(root)
    root.mainloop()
