import sys
import os
import io
import re
import base64
import tempfile
import hashlib
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from PIL import Image, ImageTk

# Optional high-performance libraries with graceful fallbacks
try:
    import pymupdf
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import tkinterweb
    HAS_TKINTERWEB = True
except ImportError:
    HAS_TKINTERWEB = False

try:
    import pygments
    from pygments.lexers import get_lexer_for_filename, get_lexer_by_name, TextLexer
    HAS_PYGMENTS = True
except ImportError:
    HAS_PYGMENTS = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import docx
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


def get_html_color(rgb_obj):
    if not rgb_obj or not hasattr(rgb_obj, 'rgb'):
        return None
    val = str(rgb_obj.rgb)
    if len(val) == 8 and val.startswith("FF"):
        return f"#{val[2:]}"
    elif len(val) == 6:
        return f"#{val}"
    return None


class UniversalNotebook:
    def __init__(self, root, file_path=None):
        self.root = root
        self.root.title("Buuggeyga (My Notebook Pro)")
        self.root.geometry("1250x820")
        self.root.minsize(900, 600)
        
        # Color Theme (VS Code Dark & Excel Green Accent)
        self.bg_dark = "#1e1e1e"
        self.bg_sidebar = "#252526"
        self.bg_toolbar = "#2d2d2d"
        self.fg_text = "#d4d4d4"
        self.accent_blue = "#007acc"
        self.accent_excel = "#217346"
        self.highlight_bg = "#264f78"
        
        self.root.configure(bg=self.bg_dark)
        self.font_family = "Consolas"
        self.font_size = 11
        
        self.current_file = None
        self.raw_bytes = None
        self.excel_wb = None
        self.excel_sheets = []
        self.current_sheet_name = None
        self.current_pdf_doc = None
        self.current_pdf_page = 0
        self.pdf_zoom = 1.2
        self.image_zoom = 1.0
        self.current_pil_image = None
        self.current_mode = "text" # "text", "excel", "word", "pdf", "html", "image", "hex"
        
        self._build_menu()
        self._build_toolbar()
        self._build_main_area()
        self._build_status_bar()
        self._build_search_bar()
        self._bind_shortcuts()
        
        if file_path and os.path.exists(file_path):
            self.load_file(file_path)
        else:
            self.new_file()

    def _build_menu(self):
        self.menu_bar = tk.Menu(self.root, bg=self.bg_sidebar, fg=self.fg_text)
        self.root.config(menu=self.menu_bar)
        
        # File menu (Faylka)
        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Faylka (File)", menu=self.file_menu)
        self.file_menu.add_command(label="Cusub (New)", command=self.new_file, accelerator="Ctrl+N")
        self.file_menu.add_command(label="Fur (Open)...", command=self.open_file_dialog, accelerator="Ctrl+O")
        self.file_menu.add_command(label="Keydi (Save)", command=self.save_file, accelerator="Ctrl+S")
        self.file_menu.add_command(label="Keydi Sida... (Save As...)", command=self.save_as_file)
        self.file_menu.add_separator()
        self.file_menu.add_command(label="🖨️ Dabac (Print)...", command=self.print_file, accelerator="Ctrl+P")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Ka bax (Exit)", command=self.root.quit)
        
        # View Mode Menu (Eeg)
        self.view_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Eeg (View Mode)", menu=self.view_menu)
        self.view_menu.add_command(label="📊 Excel Spreadsheet View", command=lambda: self.switch_mode("excel"))
        self.view_menu.add_command(label="📘 Word Document View", command=lambda: self.switch_mode("word"))
        self.view_menu.add_command(label="🌐 HTML Web Preview", command=lambda: self.switch_mode("html"))
        self.view_menu.add_command(label="📕 PDF Document Visualizer", command=lambda: self.switch_mode("pdf"))
        self.view_menu.add_command(label="🖼️ Sawir (Image View)", command=lambda: self.switch_mode("image"))
        self.view_menu.add_command(label="📝 Qoraal / Code", command=lambda: self.switch_mode("text"))
        self.view_menu.add_command(label="⚙️ Machine Code / Hex Inspector", command=lambda: self.switch_mode("hex"))
        
        # Tools & Zoom
        self.tools_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Aaladaha (Tools)", menu=self.tools_menu)
        self.tools_menu.add_command(label="Raadi (Find)...", command=self.toggle_search_bar, accelerator="Ctrl+F")
        self.tools_menu.add_command(label="Weyneey Font (Zoom In)", command=self.zoom_in, accelerator="Ctrl++")
        self.tools_menu.add_command(label="Yaree Font (Zoom Out)", command=self.zoom_out, accelerator="Ctrl+-")
        
        # Help
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Caawinaad (Help)", menu=self.help_menu)
        self.help_menu.add_command(label="Laga helay (About)", command=self.show_about)

    def _build_toolbar(self):
        self.toolbar = tk.Frame(self.root, bg=self.bg_toolbar, height=40)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        
        btn_opts = {"bg": "#3e3e42", "fg": "white", "activebackground": self.accent_blue, "activeforeground": "white", "bd": 0, "padx": 9, "pady": 4, "font": ("Segoe UI", 9, "bold")}
        
        tk.Button(self.toolbar, text="📄 Cusub", command=self.new_file, **btn_opts).pack(side=tk.LEFT, padx=2, pady=4)
        tk.Button(self.toolbar, text="📁 Fur (Open)", command=self.open_file_dialog, **btn_opts).pack(side=tk.LEFT, padx=2, pady=4)
        tk.Button(self.toolbar, text="💾 Keydi (Save)", command=self.save_file, **btn_opts).pack(side=tk.LEFT, padx=2, pady=4)
        
        # Print Button
        btn_print_opts = {"bg": "#217346", "fg": "white", "activebackground": "#0e4727", "activeforeground": "white", "bd": 0, "padx": 10, "pady": 4, "font": ("Segoe UI", 9, "bold")}
        tk.Button(self.toolbar, text="🖨️ Dabac (Print)", command=self.print_file, **btn_print_opts).pack(side=tk.LEFT, padx=5, pady=4)
        
        tk.Label(self.toolbar, text=" | Modagga:", bg=self.bg_toolbar, fg="#aaaaaa", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=4)
        
        self.btn_mode_excel = tk.Button(self.toolbar, text="📊 Excel", command=lambda: self.switch_mode("excel"), **btn_opts)
        self.btn_mode_excel.pack(side=tk.LEFT, padx=2, pady=4)

        self.btn_mode_word = tk.Button(self.toolbar, text="📘 Word", command=lambda: self.switch_mode("word"), **btn_opts)
        self.btn_mode_word.pack(side=tk.LEFT, padx=2, pady=4)

        self.btn_mode_html = tk.Button(self.toolbar, text="🌐 HTML", command=lambda: self.switch_mode("html"), **btn_opts)
        self.btn_mode_html.pack(side=tk.LEFT, padx=2, pady=4)
        
        self.btn_mode_pdf = tk.Button(self.toolbar, text="📕 PDF", command=lambda: self.switch_mode("pdf"), **btn_opts)
        self.btn_mode_pdf.pack(side=tk.LEFT, padx=2, pady=4)
        
        self.btn_mode_text = tk.Button(self.toolbar, text="📝 Code/Text", command=lambda: self.switch_mode("text"), **btn_opts)
        self.btn_mode_text.pack(side=tk.LEFT, padx=2, pady=4)
        
        self.btn_mode_hex = tk.Button(self.toolbar, text="⚙️ Machine Code", command=lambda: self.switch_mode("hex"), **btn_opts)
        self.btn_mode_hex.pack(side=tk.LEFT, padx=2, pady=4)

        # Search toggle button
        tk.Button(self.toolbar, text="🔍 Raadi", command=self.toggle_search_bar, **btn_opts).pack(side=tk.RIGHT, padx=5, pady=4)

    def _build_search_bar(self):
        self.search_frame = tk.Frame(self.root, bg="#2d2d30", height=35)
        
        tk.Label(self.search_frame, text="Raadi (Find):", bg="#2d2d30", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        self.search_entry = tk.Entry(self.search_frame, bg="#1e1e1e", fg="white", insertbackground="white", font=("Consolas", 10), width=25)
        self.search_entry.pack(side=tk.LEFT, padx=5, pady=4)
        self.search_entry.bind("<Return>", lambda event: self.find_next())
        
        tk.Button(self.search_frame, text="Raadi Kan Xiga", command=self.find_next, bg="#007acc", fg="white", bd=0, padx=8, pady=2, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=3)
        tk.Button(self.search_frame, text="❌", command=self.toggle_search_bar, bg="#3e3e42", fg="white", bd=0, padx=5, pady=2).pack(side=tk.RIGHT, padx=5)
        
        self.search_status = tk.Label(self.search_frame, text="", bg="#2d2d30", fg="#4ec9b0", font=("Segoe UI", 9))
        self.search_status.pack(side=tk.LEFT, padx=10)

    def toggle_search_bar(self):
        if self.search_frame.winfo_ismapped():
            self.search_frame.pack_forget()
            self.text_area.tag_remove("found", "1.0", tk.END)
        else:
            self.search_frame.pack(side=tk.TOP, fill=tk.X, before=self.main_container)
            self.search_entry.focus_set()

    def find_next(self):
        query = self.search_entry.get()
        if not query:
            return
        
        self.text_area.tag_remove("found", "1.0", tk.END)
        self.text_area.tag_config("found", background="#515c6b", foreground="#ffffff")
        
        idx = "1.0"
        count = 0
        first_pos = None
        
        while True:
            idx = self.text_area.search(query, idx, stopindex=tk.END, nocase=True)
            if not idx:
                break
            if not first_pos:
                first_pos = idx
            end_idx = f"{idx}+{len(query)}c"
            self.text_area.tag_add("found", idx, end_idx)
            count += 1
            idx = end_idx
            
        if first_pos:
            self.text_area.see(first_pos)
            self.text_area.mark_set(tk.INSERT, first_pos)
            self.search_status.config(text=f"{count} ayaa la helay")
        else:
            self.search_status.config(text="Lama helin")

    def _build_main_area(self):
        self.main_container = tk.Frame(self.root, bg=self.bg_dark)
        self.main_container.pack(expand=True, fill=tk.BOTH)
        
        # 1. Text & Code View Frame
        self.text_frame = tk.Frame(self.main_container, bg=self.bg_dark)
        
        self.line_numbers = tk.Text(self.text_frame, width=5, padx=3, takefocus=0, border=0,
                                    bg="#1e1e1e", fg="#858585", state='disabled', font=(self.font_family, self.font_size))
        self.line_numbers.pack(side=tk.LEFT, fill=tk.Y)
        
        self.text_area = scrolledtext.ScrolledText(
            self.text_frame, wrap=tk.NONE, font=(self.font_family, self.font_size),
            bg=self.bg_dark, fg=self.fg_text, insertbackground="white",
            selectbackground=self.highlight_bg, selectforeground="white", undo=True
        )
        self.text_area.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        
        self.xscrollbar = tk.Scrollbar(self.text_area, orient=tk.HORIZONTAL, command=self.text_area.xview)
        self.text_area.config(xscrollcommand=self.xscrollbar.set)
        self.xscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.text_area.bind("<KeyRelease>", self._on_text_change)
        self.text_area.bind("<MouseWheel>", self._on_text_scroll)

        # 2. Excel & Web View Frame
        self.html_frame = tk.Frame(self.main_container, bg="#ffffff")
        
        self.excel_toolbar = tk.Frame(self.html_frame, bg="#217346", height=38)
        
        tk.Label(self.excel_toolbar, text="📊 Excel Sheet:", bg="#217346", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=8)
        self.sheet_combo = ttk.Combobox(self.excel_toolbar, state="readonly", width=32)
        self.sheet_combo.pack(side=tk.LEFT, padx=5, pady=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_selected)
        
        tk.Button(self.excel_toolbar, text="🖨️ Dabac (Print)", command=self.print_file, bg="#0e4727", fg="white", bd=0, padx=10, pady=2, font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=8)
        tk.Button(self.excel_toolbar, text="📝 CSV / Raw Text View", command=lambda: self.switch_mode("text"), bg="#3e3e42", fg="white", bd=0, padx=8, pady=2).pack(side=tk.RIGHT, padx=5)
        
        if HAS_TKINTERWEB:
            self.html_viewer = tkinterweb.HtmlFrame(self.html_frame, horizontal_scrollbar=True, vertical_scrollbar=True)
            self.html_viewer.pack(expand=True, fill=tk.BOTH)
        else:
            self.html_fallback_label = tk.Label(self.html_frame, text="tkinterweb ma rakibna.", bg="#252526", fg="white")
            self.html_fallback_label.pack(expand=True, fill=tk.BOTH)

        # 3. PDF View Frame
        self.pdf_frame = tk.Frame(self.main_container, bg="#252526")
        
        self.pdf_control_bar = tk.Frame(self.pdf_frame, bg="#333333", height=35)
        self.pdf_control_bar.pack(side=tk.TOP, fill=tk.X)
        
        btn_pdf_style = {"bg": "#3e3e42", "fg": "white", "bd": 0, "padx": 8, "pady": 2, "font": ("Segoe UI", 9)}
        tk.Button(self.pdf_control_bar, text="◀ Isku soo celi", command=self.pdf_prev_page, **btn_pdf_style).pack(side=tk.LEFT, padx=5, pady=4)
        
        self.pdf_page_label = tk.Label(self.pdf_control_bar, text="Bogga 0 / 0", bg="#333333", fg="white", font=("Segoe UI", 9, "bold"))
        self.pdf_page_label.pack(side=tk.LEFT, padx=10)
        
        tk.Button(self.pdf_control_bar, text="Shoray ▶", command=self.pdf_next_page, **btn_pdf_style).pack(side=tk.LEFT, padx=5, pady=4)
        
        tk.Label(self.pdf_control_bar, text=" | ", bg="#333333", fg="#888888").pack(side=tk.LEFT, padx=5)
        tk.Button(self.pdf_control_bar, text="🔍 Zoom +", command=self.pdf_zoom_in, **btn_pdf_style).pack(side=tk.LEFT, padx=3)
        tk.Button(self.pdf_control_bar, text="🔍 Zoom -", command=self.pdf_zoom_out, **btn_pdf_style).pack(side=tk.LEFT, padx=3)
        tk.Button(self.pdf_control_bar, text="🖨️ Dabac (Print)", command=self.print_file, bg="#217346", fg="white", bd=0, padx=8, pady=2, font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        tk.Button(self.pdf_control_bar, text="📄 Qoraal ahaan u soo saar", command=self.pdf_extract_text, **btn_pdf_style).pack(side=tk.RIGHT, padx=5)

        self.pdf_canvas_frame = tk.Frame(self.pdf_frame, bg="#1e1e1e")
        self.pdf_canvas_frame.pack(expand=True, fill=tk.BOTH)
        
        self.pdf_canvas = tk.Canvas(self.pdf_canvas_frame, bg="#1e1e1e", highlightthickness=0)
        self.pdf_vscroll = tk.Scrollbar(self.pdf_canvas_frame, orient=tk.VERTICAL, command=self.pdf_canvas.yview)
        self.pdf_hscroll = tk.Scrollbar(self.pdf_canvas_frame, orient=tk.HORIZONTAL, command=self.pdf_canvas.xview)
        
        self.pdf_canvas.config(yscrollcommand=self.pdf_vscroll.set, xscrollcommand=self.pdf_hscroll.set)
        self.pdf_vscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.pdf_hscroll.pack(side=tk.BOTTOM, fill=tk.X)
        self.pdf_canvas.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        # 4. Image Frame
        self.image_frame = tk.Frame(self.main_container, bg="#1e1e1e")
        self.image_canvas = tk.Canvas(self.image_frame, bg="#1e1e1e", highlightthickness=0)
        self.image_canvas.pack(expand=True, fill=tk.BOTH)

        # Show initial text frame
        self.text_frame.pack(expand=True, fill=tk.BOTH)

    def _build_status_bar(self):
        self.status_bar = tk.Frame(self.root, bg="#217346", height=24)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_file_type = tk.Label(self.status_bar, text="Nooca: Qoraal (Text)", bg="#217346", fg="white", font=("Segoe UI", 9, "bold"))
        self.status_file_type.pack(side=tk.LEFT, padx=10)
        
        self.status_pos = tk.Label(self.status_bar, text="Sadarka: 1, Qoraalka: 1", bg="#217346", fg="white", font=("Segoe UI", 9))
        self.status_pos.pack(side=tk.RIGHT, padx=10)
        
        self.status_encoding = tk.Label(self.status_bar, text="UTF-8", bg="#217346", fg="white", font=("Segoe UI", 9))
        self.status_encoding.pack(side=tk.RIGHT, padx=10)

    def _bind_shortcuts(self):
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-N>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file_dialog())
        self.root.bind("<Control-O>", lambda e: self.open_file_dialog())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_file())
        self.root.bind("<Control-p>", lambda e: self.print_file())
        self.root.bind("<Control-P>", lambda e: self.print_file())
        self.root.bind("<Control-f>", lambda e: self.toggle_search_bar())
        self.root.bind("<Control-F>", lambda e: self.toggle_search_bar())
        self.root.bind("<Control-plus>", lambda e: self.zoom_in())
        self.root.bind("<Control-minus>", lambda e: self.zoom_out())

    def _on_text_change(self, event=None):
        self._update_line_numbers()
        self._update_status_bar()
        if HAS_PYGMENTS and self.current_mode == "text":
            self.apply_syntax_highlighting()

    def _on_text_scroll(self, event=None):
        self._update_line_numbers()

    def _update_line_numbers(self):
        self.line_numbers.config(state='normal')
        self.line_numbers.delete("1.0", tk.END)
        
        line_count = int(self.text_area.index('end-1c').split('.')[0])
        lines_str = "\n".join(str(i) for i in range(1, line_count + 1))
        
        self.line_numbers.insert("1.0", lines_str)
        self.line_numbers.config(state='disabled')
        self.line_numbers.yview_moveto(self.text_area.yview()[0])

    def _update_status_bar(self):
        cursor_pos = self.text_area.index(tk.INSERT)
        line, col = cursor_pos.split('.')
        self.status_pos.config(text=f"Sadarka: {line}, Sadar-yar: {int(col)+1}")

    def zoom_in(self):
        self.font_size = min(36, self.font_size + 1)
        self.text_area.config(font=(self.font_family, self.font_size))
        self.line_numbers.config(font=(self.font_family, self.font_size))
        self._update_line_numbers()

    def zoom_out(self):
        self.font_size = max(8, self.font_size - 1)
        self.text_area.config(font=(self.font_family, self.font_size))
        self.line_numbers.config(font=(self.font_family, self.font_size))
        self._update_line_numbers()

    def print_file(self):
        """Universal Print Handler v12.0 Final - Perfect Excel Print with 20-row pages, clean borders, bold font"""
        try:
            if self.current_mode == "pdf" and self.current_file and os.path.exists(self.current_file):
                webbrowser.open(self.current_file)
                return
                
            html = ["""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Print Document - Buuggeyga Pro</title>
<style>
    @page { size: A4 landscape; margin: 8mm 5mm; }
    @media print {
        body { margin: 0; padding: 0; background: #fff !important; color: #000 !important; }
        .no-print { display: none !important; }
        .excel-table { width: 100% !important; }
        .excel-table th, .excel-table td {
            font-size: 12px !important;
            font-weight: 700 !important;
            color: #000000 !important;
            padding: 3px 5px !important;
        }
        .excel-table td.no-border { border: none !important; background: transparent !important; }
        .page-block { page-break-after: always; margin-bottom: 0; }
        .page-block:last-child { page-break-after: auto; }
    }
    * { box-sizing: border-box; }
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 4px; color: #000; background: #fff; }
    .hdr { font-size: 15px; font-weight: bold; color: #1a5c32; margin-bottom: 8px;
           border-bottom: 2.5px solid #217346; padding-bottom: 3px; }
    .page-block { margin-bottom: 6px; }
    .excel-table { border-collapse: collapse; white-space: nowrap;
                   font-size: 12px; font-weight: 700;
                   font-family: 'Segoe UI', Arial, sans-serif;
                   width: 100%; color: #000; }
    .excel-table thead th {
        background-color: #e8e8e8;
        color: #000000;
        border: 1px solid #999;
        padding: 4px 6px;
        font-weight: 700;
        text-align: center;
        font-size: 11px;
    }
    .excel-table td {
        border: 1px solid #aaa;
        padding: 3px 6px;
        color: #000000;
        font-weight: 700;
        font-size: 12px;
    }
    .excel-table td.no-border {
        border: none !important;
        background: transparent !important;
        color: transparent !important;
    }
    .row-num {
        background-color: #efefef !important;
        text-align: center;
        font-weight: 700;
        font-size: 11px;
        width: 28px;
        min-width: 28px;
        color: #555 !important;
        border: 1px solid #999 !important;
    }
    .col-margin { width: 20px !important; min-width: 20px !important; padding: 2px !important; }
</style>
</head>
<body>"""]

            if self.current_mode == "excel" and self.excel_wb:
                ws = self.excel_wb[self.current_sheet_name if self.current_sheet_name else self.excel_wb.sheetnames[0]]
                sheet_title = self.current_sheet_name if self.current_sheet_name else "Excel Sheet"
                html.append(f'<div class="hdr">&#128202; Microsoft Excel Print View: {sheet_title}</div>')

                # ── Build image map ──────────────────────────────────────────────
                image_map = {}
                if hasattr(ws, '_images'):
                    for img in ws._images:
                        try:
                            anchor = img.anchor
                            r, c = None, None
                            if hasattr(anchor, '_from'):
                                c = anchor._from.col + 1
                                r = anchor._from.row + 1
                            if r and c:
                                if hasattr(img, '_data'):
                                    img_data = img._data()
                                elif hasattr(img, 'ref'):
                                    img_data = img.ref.read()
                                else:
                                    buf = io.BytesIO()
                                    img.image.save(buf, format='PNG')
                                    img_data = buf.getvalue()
                                b64 = base64.b64encode(img_data).decode('utf-8')
                                image_map[(r, c)] = f"data:image/png;base64,{b64}"
                        except Exception:
                            pass

                # ── Build merged-cell map ────────────────────────────────────────
                merged_map = {}
                skip_cells = set()
                for m in ws.merged_cells.ranges:
                    colspan = m.max_col - m.min_col + 1
                    rowspan = m.max_row - m.min_row + 1
                    merged_map[(m.min_row, m.min_col)] = (colspan, rowspan)
                    for mr in range(m.min_row, m.max_row + 1):
                        for mc in range(m.min_col, m.max_col + 1):
                            if (mr, mc) != (m.min_row, m.min_col):
                                skip_cells.add((mr, mc))

                max_r = ws.max_row
                max_c = ws.max_column

                # ── Collect content rows (skip fill-only and fully-empty rows) ────
                content_rows = []
                for r in range(1, max_r + 1):
                    row_has_content = False
                    for c_chk in range(1, max_c + 1):
                        if (r, c_chk) in skip_cells:
                            continue
                        chk_cell = ws.cell(r, c_chk)
                        if chk_cell.value is not None:
                            row_has_content = True
                            break
                    if any(img_r == r for (img_r, _) in image_map):
                        row_has_content = True
                    if row_has_content:
                        content_rows.append(r)

                # ── Split into pages of 20 rows ──────────────────────────────────
                ROWS_PER_PAGE = 20
                # Header rows (r < 12) always appear on every page
                header_rows = [r for r in content_rows if r < 12]
                data_rows   = [r for r in content_rows if r >= 12]
                chunks = [data_rows[i:i+ROWS_PER_PAGE] for i in range(0, max(1, len(data_rows)), ROWS_PER_PAGE)]

                def render_col_header():
                    out = ['<thead><tr><th class="row-num">#</th>']
                    for c in range(1, max_c + 1):
                        col_class = " class='col-margin'" if c in [1, 2] else ""
                        out.append(f'<th{col_class}>{openpyxl.utils.get_column_letter(c)}</th>')
                    out.append('</tr></thead>')
                    return "".join(out)

                def render_rows(rows_to_render):
                    out = []
                    for r in rows_to_render:
                        out.append(f'<tr><td class="row-num">{r}</td>')
                        for c in range(1, max_c + 1):
                            if (r, c) in skip_cells:
                                continue
                            cell = ws.cell(r, c)
                            val  = cell.value if cell.value is not None else ""
                            cs, rs_span = merged_map.get((r, c), (1, 1))
                            span_attrs = ""
                            if cs > 1:
                                span_attrs += f' colspan="{cs}"'
                            if rs_span > 1:
                                span_attrs += f' rowspan="{rs_span}"'

                            styles = []
                            if c in [1, 2] and cs == 1:
                                styles.append("width:20px;min-width:20px;")

                            # Background color from Excel fill
                            if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                                bg_color = get_html_color(cell.fill.start_color)
                                if bg_color and bg_color != "#000000":
                                    styles.append(f"background-color:{bg_color}")

                            # Header empty cell → transparent (no floating border)
                            is_empty_header = (r < 12 and val == "" and (r, c) not in image_map)
                            if is_empty_header:
                                styles = [s for s in styles if not s.startswith("background-color")]
                                styles.append("border:none!important;background:transparent!important;color:transparent!important;")

                            # Font styling from Excel
                            if cell.font:
                                if cell.font.bold:
                                    styles.append("font-weight:700")
                                if cell.font.color and cell.font.color.rgb:
                                    fg = get_html_color(cell.font.color)
                                    if fg and not is_empty_header:
                                        styles.append(f"color:{fg}")

                            # Text alignment from Excel
                            if cell.alignment and cell.alignment.horizontal:
                                align = cell.alignment.horizontal
                                if align in ('center', 'right', 'left'):
                                    styles.append(f"text-align:{align}")

                            # Value formatting
                            if isinstance(val, (int, float)):
                                try:
                                    fmt = cell.number_format or ""
                                    if any(x in fmt for x in ['$', 'USD', '#,##0']):
                                        val_str = f"${val:,.2f}"
                                    elif '%' in fmt:
                                        val_str = f"{val*100:.1f}%" if val < 10 else f"{val:.1f}%"
                                    elif '.' in fmt:
                                        val_str = f"{val:,.2f}"
                                    else:
                                        val_str = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
                                except Exception:
                                    val_str = str(val)
                            else:
                                val_str = str(val)

                            # Embedded image
                            if (r, c) in image_map:
                                img_src = image_map[(r, c)]
                                img_tag = f'<img src="{img_src}" style="max-height:80px;max-width:120px;vertical-align:middle;">'
                                val_str = f'{img_tag} {val_str}'.strip()

                            style_attr = f' style="{"; ".join(styles)}"' if styles else ''
                            cell_class = ' class="no-border"' if is_empty_header else ''
                            out.append(f'<td{span_attrs}{cell_class}{style_attr}>{val_str}</td>')
                        out.append('</tr>')
                    return "".join(out)

                # ── Emit each page-block ─────────────────────────────────────────
                for page_idx, chunk in enumerate(chunks):
                    html.append('<div class="page-block">')
                    html.append('<table class="excel-table">')
                    html.append(render_col_header())
                    html.append('<tbody>')
                    # Repeat header rows on every page
                    html.append(render_rows(header_rows))
                    html.append(render_rows(chunk))
                    html.append('</tbody></table>')
                    html.append('</div>')
            elif self.current_mode == "word" and HAS_DOCX and self.current_file:
                doc = docx.Document(self.current_file)
                html.append(f'<div class="hdr">&#128216; Word Document Print View: {os.path.basename(self.current_file)}</div>')
                for p in doc.paragraphs:
                    if p.text.strip():
                        html.append(f"<p style='font-size:13px;margin:4px 0;color:#000;'>{p.text}</p>")
                for t in doc.tables:
                    html.append('<table class="excel-table">')
                    for row in t.rows:
                        html.append("<tr>" + "".join(f"<td>{cell.text}</td>" for cell in row.cells) + "</tr>")
                    html.append("</table>")
            else:
                raw_text = self.text_area.get(1.0, tk.END)
                html.append(f"<pre style='font-family:Consolas,monospace;font-size:12px;color:#000;'>{raw_text}</pre>")

            html.append("""<script>
    window.onload = function() {
        window.print();
    };
</script>
</body>
</html>""")

            temp_print_file = os.path.join(os.getenv("TEMP"), "notebook_print_doc.html")
            with open(temp_print_file, "w", encoding="utf-8") as f:
                f.write("".join(html))
                
            webbrowser.open(temp_print_file)
        except Exception as e:
            messagebox.showerror("Dabacaad Khalad", f"Lama daabaci karo: {e}")

    def switch_mode(self, mode):
        self.current_mode = mode
        
        self.text_frame.pack_forget()
        self.pdf_frame.pack_forget()
        self.html_frame.pack_forget()
        self.image_frame.pack_forget()
        self.excel_toolbar.pack_forget()
        
        for b in [self.btn_mode_text, self.btn_mode_excel, self.btn_mode_word, self.btn_mode_html, self.btn_mode_pdf, self.btn_mode_hex]:
            b.config(bg="#3e3e42")
            
        if mode == "text":
            self.text_frame.pack(expand=True, fill=tk.BOTH)
            self.btn_mode_text.config(bg=self.accent_blue)
            self.status_file_type.config(text="Modagga: Qoraal / Code Editor", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self._update_line_numbers()
        elif mode == "excel":
            self.html_frame.pack(expand=True, fill=tk.BOTH)
            self.excel_toolbar.pack(side=tk.TOP, fill=tk.X, before=self.html_viewer if HAS_TKINTERWEB else self.html_fallback_label)
            self.btn_mode_excel.config(bg=self.accent_excel)
            self.status_file_type.config(text="Modagga: 📊 Excel Spreadsheet Table View", bg=self.accent_excel)
            self.status_bar.config(bg=self.accent_excel)
            self.render_excel_view()
        elif mode == "word":
            self.html_frame.pack(expand=True, fill=tk.BOTH)
            self.btn_mode_word.config(bg=self.accent_blue)
            self.status_file_type.config(text="Modagga: 📘 Word Document Formatted View", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self.render_word_view()
        elif mode == "html":
            self.html_frame.pack(expand=True, fill=tk.BOTH)
            self.btn_mode_html.config(bg=self.accent_blue)
            self.status_file_type.config(text="Modagga: 🌐 HTML Interactive Web Preview", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self.render_html_preview()
        elif mode == "pdf":
            self.pdf_frame.pack(expand=True, fill=tk.BOTH)
            self.btn_mode_pdf.config(bg=self.accent_blue)
            self.status_file_type.config(text="Modagga: 📕 PDF Document Visualizer", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self.render_pdf_page()
        elif mode == "image":
            self.image_frame.pack(expand=True, fill=tk.BOTH)
            self.status_file_type.config(text="Modagga: 🖼️ Image Viewer", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self.render_image_view()
        elif mode == "hex":
            self.text_frame.pack(expand=True, fill=tk.BOTH)
            self.btn_mode_hex.config(bg=self.accent_blue)
            self.status_file_type.config(text="Modagga: ⚙️ Machine Code & Hex Inspector", bg=self.accent_blue)
            self.status_bar.config(bg=self.accent_blue)
            self.render_hex_view()

    def open_file_dialog(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Kulliga faylalka (All Files)", "*.*"),
                ("Excel Files (*.xlsx, *.xls, *.csv)", "*.xlsx;*.xls;*.csv"),
                ("Word Documents (*.docx)", "*.docx"),
                ("PDF Documents (*.pdf)", "*.pdf"),
                ("HTML Files (*.html, *.htm)", "*.html;*.htm"),
                ("Images (*.png, *.jpg, *.gif)", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"),
                ("Python Scripts (*.py)", "*.py"),
                ("Executables & Binary (*.exe, *.dll, *.bin)", "*.exe;*.dll;*.bin")
            ]
        )
        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path):
        self.current_file = file_path
        self.root.title(f"Buuggeyga Pro - {os.path.basename(file_path)}")
        ext = os.path.splitext(file_path)[1].lower()
        
        try:
            with open(file_path, "rb") as f:
                self.raw_bytes = f.read()
        except Exception as e:
            messagebox.showerror("Khalad (Error)", f"Lama furi karo faylka: {e}")
            return

        # 1. EXCEL files (.xlsx, .xls, .csv)
        if ext in [".xlsx", ".xls", ".csv", ".tsv", ".ods"]:
            self.load_excel_file(file_path)
            return

        # 2. WORD documents (.docx)
        if ext in [".docx", ".doc"]:
            self.load_word_file(file_path)
            return

        # 3. PDF documents (.pdf)
        if ext == ".pdf" or self.raw_bytes.startswith(b"%PDF"):
            if HAS_PYMUPDF:
                try:
                    self.current_pdf_doc = pymupdf.open(file_path)
                    self.current_pdf_page = 0
                    self.switch_mode("pdf")
                    return
                except Exception as e:
                    pass

        # 4. HTML files (.html, .htm)
        if ext in [".html", ".htm"] or b"<html" in self.raw_bytes.lower():
            self._load_as_text()
            self.switch_mode("html")
            return

        # 5. Image files (.png, .jpg, .jpeg, .bmp, .gif, .webp)
        if ext in [".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".ico"]:
            try:
                self.current_pil_image = Image.open(file_path)
                self.switch_mode("image")
                return
            except Exception:
                pass

        # 6. ONLY Binary Executables & Code Machine Files (.exe, .dll, .bin, .dat, .sys, .o, .pyc, .class)
        if ext in [".exe", ".dll", ".bin", ".dat", ".sys", ".o", ".pyc", ".class"]:
            self.switch_mode("hex")
            return

        # 7. Default Text & Code Mode
        self._load_as_text()
        self.switch_mode("text")

    def _load_as_text(self):
        if not self.raw_bytes:
            content = ""
        else:
            try:
                content = self.raw_bytes.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    content = self.raw_bytes.decode('latin-1')
                except Exception:
                    content = str(self.raw_bytes)
        
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, content)
        self._update_line_numbers()
        self._update_status_bar()
        if HAS_PYGMENTS:
            self.apply_syntax_highlighting()

    # --- EXCEL SPREADSHEET HANDLING ---
    def load_excel_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        self.excel_sheets = []
        
        if ext in [".xlsx", ".xls", ".ods"] and HAS_OPENPYXL:
            try:
                self.excel_wb = openpyxl.load_workbook(file_path, data_only=True)
                self.excel_sheets = self.excel_wb.sheetnames
                self.sheet_combo['values'] = self.excel_sheets
                if self.excel_sheets:
                    self.sheet_combo.set(self.excel_sheets[0])
                    self.current_sheet_name = self.excel_sheets[0]
                self.switch_mode("excel")
                return
            except Exception:
                pass
                
        self._load_as_text()
        self.switch_mode("excel")

    def _on_sheet_selected(self, event=None):
        self.current_sheet_name = self.sheet_combo.get()
        self.render_excel_view()

    def render_excel_view(self):
        if not HAS_TKINTERWEB:
            return
            
        html = []
        html.append("""<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 5px; background-color: #ffffff; color: #000000; }
    .excel-scroll-wrapper { width: 100%; overflow-x: scroll; overflow-y: auto; background-color: #ffffff; display: block; }
    .excel-table { border-collapse: collapse; white-space: nowrap; font-size: 13px; background-color: #ffffff; min-width: 100%; border: 1px solid #b0b0b0; font-family: 'Segoe UI', Arial, sans-serif; }
    .excel-table th { background-color: #f2f2f2; color: #000000; border: 1px solid #b0b0b0; padding: 7px 9px; font-weight: bold; text-align: center; }
    .excel-table td { border: 1px solid #b0b0b0; padding: 5px 8px; color: #000000; font-weight: 700; font-size: 13px; font-family: 'Segoe UI', Arial, sans-serif; }
    .excel-table td.no-border { border: none !important; background: transparent !important; color: transparent !important; }
    .row-num { background-color: #efefef !important; color: #444444 !important; text-align: center; font-weight: 700; font-size: 11px; width: 30px; min-width: 30px; border: 1px solid #b0b0b0 !important; }
    .col-margin { width: 20px !important; min-width: 20px !important; max-width: 24px !important; padding: 2px !important; }
    .excel-hdr-banner { background-color: #217346; color: white; padding: 8px 14px; font-size: 15px; font-weight: bold; font-family: 'Segoe UI'; margin-bottom: 10px; border-radius: 2px; }
</style>
</head>
<body>""")

        sheet_title = self.current_sheet_name if self.current_sheet_name else "Spreadsheet"
        html.append(f'<div class="excel-hdr-banner">📊 Microsoft Excel Sheet: {sheet_title}</div>')
        html.append('<div class="excel-scroll-wrapper"><table class="excel-table">')

        if self.excel_wb and self.current_sheet_name and HAS_OPENPYXL:
            ws = self.excel_wb[self.current_sheet_name]
            
            image_map = {}
            if hasattr(ws, '_images'):
                for img_idx, img in enumerate(ws._images):
                    try:
                        anchor = img.anchor
                        r, c = None, None
                        if hasattr(anchor, '_from'):
                            c = anchor._from.col + 1
                            r = anchor._from.row + 1
                        if r and c:
                            if hasattr(img, '_data'):
                                img_data = img._data()
                            elif hasattr(img, 'ref'):
                                img_data = img.ref.read()
                            else:
                                buf = io.BytesIO()
                                img.image.save(buf, format='PNG')
                                img_data = buf.getvalue()
                            tmp_logo = os.path.join(tempfile.gettempdir(), f'excel_logo_{img_idx}.png')
                            with open(tmp_logo, 'wb') as f_logo:
                                f_logo.write(img_data)
                            image_map[(r, c)] = tmp_logo
                    except Exception:
                        pass
                        
            merged_map = {}
            skip_cells = set()
            for m in ws.merged_cells.ranges:
                colspan = m.max_col - m.min_col + 1
                rowspan = m.max_row - m.min_row + 1
                merged_map[(m.min_row, m.min_col)] = (colspan, rowspan)
                for r in range(m.min_row, m.max_row + 1):
                    for c in range(m.min_col, m.max_col + 1):
                        if (r, c) != (m.min_row, m.min_col):
                            skip_cells.add((r, c))

            max_r = ws.max_row
            max_c = ws.max_column

            html.append("<thead><tr><th class='row-num'>#</th>")
            for c in range(1, max_c + 1):
                col_letter = openpyxl.utils.get_column_letter(c)
                col_class = " class='col-margin'" if c in [1, 2] else ""
                html.append(f"<th{col_class}>{col_letter}</th>")
            html.append("</tr></thead><tbody>")

            for r in range(1, max_r + 1):
                # Skip rows that have no text content and no images (ignore fill-only rows)
                row_has_content = False
                for c_chk in range(1, max_c + 1):
                    if (r, c_chk) in skip_cells:
                        continue
                    chk_cell = ws.cell(r, c_chk)
                    if chk_cell.value is not None:
                        row_has_content = True
                        break
                if any(img_r == r for (img_r, _) in image_map):
                    row_has_content = True

                if not row_has_content:
                    continue

                html.append(f'<tr><td class="row-num">{r}</td>')

                for c in range(1, max_c + 1):
                    if (r, c) in skip_cells:
                        continue
                        
                    cell = ws.cell(r, c)
                    val = cell.value if cell.value is not None else ""
                    
                    cs, rs = merged_map.get((r, c), (1, 1))
                    span_attrs = ""
                    if cs > 1:
                        span_attrs += f' colspan="{cs}"'
                    if rs > 1:
                        span_attrs += f' rowspan="{rs}"'
                        
                    styles = []
                    if c in [1, 2] and cs == 1:
                        styles.append("width: 22px; min-width: 22px;")
                        
                    bg_color = None
                    if cell.fill and cell.fill.start_color and cell.fill.start_color.rgb:
                        bg_color = get_html_color(cell.fill.start_color)
                        if bg_color and bg_color != "#000000":
                            styles.append(f"background-color: {bg_color}")
                            
                    # Header cell rule (r < 12): If cell has NO value and NO image, make transparent!
                    is_empty_header = (r < 12 and val == "" and (r, c) not in image_map)
                    if is_empty_header:
                        styles = [s for s in styles if not s.startswith("background-color")]
                        styles.append("border:none!important;background:transparent!important;color:transparent!important;")

                    if cell.font:
                        if cell.font.bold:
                            styles.append("font-weight:700")
                        if cell.font.color and cell.font.color.rgb:
                            fg_color = get_html_color(cell.font.color)
                            if fg_color and not is_empty_header:
                                styles.append(f"color:{fg_color}")

                    # Text alignment from Excel
                    if cell.alignment and cell.alignment.horizontal:
                        align = cell.alignment.horizontal
                        if align in ('center', 'right', 'left'):
                            styles.append(f"text-align:{align}")

                    # Smart number formatting using cell's actual format
                    if isinstance(val, (int, float)):
                        try:
                            fmt = cell.number_format or ""
                            if any(x in fmt for x in ['$', 'USD', '#,##0']):
                                val_str = f"${val:,.2f}"
                            elif '%' in fmt:
                                val_str = f"{val*100:.1f}%" if val < 10 else f"{val:.1f}%"
                            elif '.' in fmt:
                                val_str = f"{val:,.2f}"
                            else:
                                val_str = f"{val:,.0f}" if val == int(val) else f"{val:,.2f}"
                        except Exception:
                            val_str = str(val)
                    else:
                        val_str = str(val)

                    if (r, c) in image_map:
                        logo_file_path = image_map[(r, c)].replace('\\', '/')
                        img_tag = f'<img src="file:///{logo_file_path}" style="max-height:90px;max-width:140px;vertical-align:middle;">'
                        val_str = f'{img_tag} {val_str}'.strip()
                        
                    style_attr = f' style="{";".join(styles)}"' if styles else ''
                    cell_class = ' class="no-border"' if is_empty_header else ''
                    html.append(f'<td{span_attrs}{cell_class}{style_attr}>{val_str}</td>')
                html.append("</tr>")
            html.append("</tbody>")
        else:
            raw_text = self.text_area.get(1.0, tk.END).strip()
            lines = raw_text.splitlines()
            if lines:
                html.append("<tbody>")
                for r_idx, l in enumerate(lines[:500]):
                    cols = l.split(",") if "," in l else l.split("\t")
                    html.append(f"<tr><td class='row-num'>{r_idx+1}</td>")
                    for c in cols:
                        html.append(f"<td>{c.strip()}</td>")
                    html.append("</tr>")
                html.append("</tbody>")

        html.append("</table></div></body></html>")
        self.html_viewer.load_html("".join(html))

    # --- WORD DOCUMENT HANDLING ---
    def load_word_file(self, file_path):
        self.switch_mode("word")

    def render_word_view(self):
        if not HAS_TKINTERWEB or not self.current_file:
            return
            
        html = []
        html.append("""<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 25px; padding: 25px; background-color: #ffffff; color: #000000; line-height: 1.7; font-size: 15px; }
    h1, h2, h3, h4 { color: #2b579a; border-bottom: 1px solid #cccccc; padding-bottom: 6px; margin-top: 20px; }
    p { margin-bottom: 12px; text-align: justify; }
    table { border-collapse: collapse; width: 100%; margin: 20px 0; background-color: #ffffff; }
    th, td { border: 1px solid #cccccc; padding: 8px 12px; }
    th { background-color: #f2f2f2; color: #000000; font-weight: bold; }
    tr:nth-child(even) { background-color: #f9f9f9; }
</style>
</head>
<body>""")

        if HAS_DOCX and self.current_file and self.current_file.lower().endswith(".docx"):
            try:
                doc = docx.Document(self.current_file)
                for p in doc.paragraphs:
                    if p.text.strip():
                        style_name = p.style.name.lower()
                        if "heading 1" in style_name:
                            html.append(f"<h1>{p.text}</h1>")
                        elif "heading 2" in style_name:
                            html.append(f"<h2>{p.text}</h2>")
                        elif "heading 3" in style_name:
                            html.append(f"<h3>{p.text}</h3>")
                        else:
                            html.append(f"<p>{p.text}</p>")

                for t in doc.tables:
                    html.append("<table>")
                    for r_idx, row in enumerate(t.rows):
                        html.append("<tr>")
                        tag = "th" if r_idx == 0 else "td"
                        for cell in row.cells:
                            html.append(f"<{tag}>{cell.text.strip()}</{tag}>")
                        html.append("</tr>")
                    html.append("</table>")
            except Exception as e:
                html.append(f"<p style='color:red;'>Lama furi karo Word document-ka: {e}</p>")
        else:
            html.append("<p>Sida Word document-ka loo furo waxaa loogu talagalay faylalka .docx</p>")

        html.append("</body></html>")
        self.html_viewer.load_html("".join(html))

    # --- PDF RENDERING ---
    def render_pdf_page(self):
        if not HAS_PYMUPDF or not self.current_file or not self.current_file.lower().endswith(".pdf"):
            return

        if not self.current_pdf_doc:
            try:
                self.current_pdf_doc = pymupdf.open(self.current_file)
                self.current_pdf_page = 0
            except Exception as e:
                messagebox.showerror("Khalad PDF", f"PDF lama rami karo: {e}")
                return

        total_pages = len(self.current_pdf_doc)
        if total_pages == 0:
            return

        self.current_pdf_page = max(0, min(self.current_pdf_page, total_pages - 1))
        self.pdf_page_label.config(text=f"Bogga {self.current_pdf_page + 1} / {total_pages}")
        
        page = self.current_pdf_doc[self.current_pdf_page]
        mat = pymupdf.Matrix(self.pdf_zoom, self.pdf_zoom)
        pix = page.get_pixmap(matrix=mat)
        
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.pdf_image_tk = ImageTk.PhotoImage(img)
        
        self.pdf_canvas.delete("all")
        self.pdf_canvas.config(scrollregion=(0, 0, pix.width, pix.height))
        self.pdf_canvas.create_image(0, 0, anchor=tk.NW, image=self.pdf_image_tk)

    def pdf_prev_page(self):
        if self.current_pdf_doc and self.current_pdf_page > 0:
            self.current_pdf_page -= 1
            self.render_pdf_page()

    def pdf_next_page(self):
        if self.current_pdf_doc and self.current_pdf_page < len(self.current_pdf_doc) - 1:
            self.current_pdf_page += 1
            self.render_pdf_page()

    def pdf_zoom_in(self):
        self.pdf_zoom = min(4.0, self.pdf_zoom + 0.2)
        self.render_pdf_page()

    def pdf_zoom_out(self):
        self.pdf_zoom = max(0.4, self.pdf_zoom - 0.2)
        self.render_pdf_page()

    def pdf_extract_text(self):
        if not self.current_pdf_doc:
            return
        extracted = []
        for i, page in enumerate(self.current_pdf_doc):
            extracted.append(f"--- BOGGA {i+1} ---\n" + page.get_text())
        
        full_text = "\n\n".join(extracted)
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, full_text)
        self.switch_mode("text")
        messagebox.showinfo("Qoraal PDF", "Qoraalkii PDF-ka waa lagu soo saaray Mode-ka Qoraalka!")

    # --- HTML WEB PREVIEW ---
    def render_html_preview(self):
        if not HAS_TKINTERWEB:
            return
        content = self.text_area.get(1.0, tk.END)
        if self.current_file and os.path.exists(self.current_file) and self.current_file.lower().endswith(('.html', '.htm')):
            self.html_viewer.load_website(self.current_file)
        else:
            self.html_viewer.load_html(content)

    # --- IMAGE RENDERING ---
    def render_image_view(self):
        if not self.current_pil_image:
            return
        w, h = self.current_pil_image.size
        scaled_w = int(w * self.image_zoom)
        scaled_h = int(h * self.image_zoom)
        resized_img = self.current_pil_image.resize((scaled_w, scaled_h), Image.Resampling.LANCZOS)
        self.tk_image = ImageTk.PhotoImage(resized_img)
        
        self.image_canvas.delete("all")
        self.image_canvas.create_image(self.image_canvas.winfo_width()//2, self.image_canvas.winfo_height()//2, anchor=tk.CENTER, image=self.tk_image)

    # --- MACHINE CODE & HEX INSPECTOR ---
    def render_hex_view(self):
        if not self.raw_bytes:
            if self.current_file and os.path.exists(self.current_file):
                with open(self.current_file, "rb") as f:
                    self.raw_bytes = f.read()
            else:
                self.raw_bytes = self.text_area.get(1.0, tk.END).encode('utf-8')

        lines = []
        lines.append("=" * 80)
        lines.append("⚙️ BUUGGEYGA MACHINE LANGUAGE & HEX INSPECTOR")
        lines.append("=" * 80)
        
        file_size = len(self.raw_bytes)
        md5_hash = hashlib.md5(self.raw_bytes).hexdigest()
        sha256_hash = hashlib.sha256(self.raw_bytes).hexdigest()
        
        header_desc = "Standard Binary / Data File"
        if self.raw_bytes.startswith(b"MZ"):
            header_desc = "Windows Executable Program (PE / DOS MZ Header)"
        elif self.raw_bytes.startswith(b"\x7fELF"):
            header_desc = "Linux Executable Binary (ELF Header)"
        elif self.raw_bytes.startswith(b"\xca\xfe\xba\xbe"):
            header_desc = "Java Bytecode Class File"
        elif self.raw_bytes.startswith(b"\x50\x4b\x03\x04"):
            header_desc = "ZIP Archive / Compressed Container"
        elif self.raw_bytes.startswith(b"%PDF"):
            header_desc = "Adobe PDF Document Container"
            
        lines.append(f"Mugga (Size): {file_size} Bytes")
        lines.append(f"Header: {header_desc}")
        lines.append(f"MD5 Hash: {md5_hash}")
        lines.append(f"SHA256:   {sha256_hash}")
        lines.append("=" * 80)
        lines.append("OFFSET   | 00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F | ASCII TEXT")
        lines.append("-" * 80)
        
        max_bytes = min(file_size, 65536)
        for i in range(0, max_bytes, 16):
            chunk = self.raw_bytes[i:i+16]
            hex_part = " ".join(f"{b:02X}" for b in chunk)
            hex_part = hex_part.ljust(47)
            ascii_part = "".join(chr(b) if 32 <= b <= 126 else "." for b in chunk)
            lines.append(f"{i:08X} | {hex_part} | {ascii_part}")
            
        if file_size > 65536:
            lines.append(f"\n... [{file_size - 65536} bytes oo dheeraad ah ayaan la muujin si uu xawligu u sareeyo] ...")

        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(1.0, "\n".join(lines))
        self._update_line_numbers()

    def apply_syntax_highlighting(self):
        if not HAS_PYGMENTS or not self.current_file:
            return
        try:
            lexer = get_lexer_for_filename(self.current_file)
        except Exception:
            return

        content = self.text_area.get(1.0, tk.END)
        tokens = pygments.lex(content, lexer)
        
        self.text_area.tag_config("Token.Keyword", foreground="#569cd6", font=(self.font_family, self.font_size, "bold"))
        self.text_area.tag_config("Token.String", foreground="#ce9178")
        self.text_area.tag_config("Token.Comment", foreground="#6a9955", font=(self.font_family, self.font_size, "italic"))
        self.text_area.tag_config("Token.Name.Function", foreground="#dcdcaa")
        self.text_area.tag_config("Token.Name.Class", foreground="#4ec9b0")
        self.text_area.tag_config("Token.Number", foreground="#b5cea8")

    def new_file(self):
        self.current_file = None
        self.raw_bytes = b""
        self.excel_wb = None
        self.excel_sheets = []
        self.current_pdf_doc = None
        self.current_pil_image = None
        self.text_area.delete(1.0, tk.END)
        self.root.title("Buuggeyga Pro - Aan la magacaabin (Untitled)")
        self.switch_mode("text")

    def save_as_file(self):
        self.current_file = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("HTML files", "*.html"), ("All files", "*.*")]
        )
        if not self.current_file:
            return
        self.save_file_content()

    def save_file(self):
        if not self.current_file:
            self.save_as_file()
            return
        self.save_file_content()

    def save_file_content(self):
        try:
            content = self.text_area.get(1.0, tk.END)
            if content.endswith('\n'):
                content = content[:-1]
                
            with open(self.current_file, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("Guul (Success)", "Faylka si guul leh ayaa loo keydiyay!")
            self.root.title(f"Buuggeyga Pro - {os.path.basename(self.current_file)}")
        except Exception as e:
            messagebox.showerror("Khalad (Error)", f"Lama keydin karin faylka: {e}")

    def show_about(self):
        msg = (
            "Buuggeyga (My Notebook Pro) v11.0 Clean Title & Natural Print Edition\n\n"
            "• 🧹 Clean Title Cells: Stripped unneeded empty border boxes & empty sky-blue fills from rows 1-11\n"
            "• 📄 Natural Print Flow: Removed forced artificial page breaks so pages fill up naturally (~8-9 pages)\n"
            "• 🔤 Crisp Dark Bold Typography (13px Bold Black Text)\n"
            "• 🖼️ Automatic Excel Embedded Logo Extraction (CEELKA BIYAHA ZEYD KM13)\n"
            "• 📐 Slim Front Margin Columns (A & B 22px Alignment)\n"
            "• 📊 100% Perfect Merged Cell Colspan/Rowspan Renderer\n"
            "• 🎨 Full OpenPyXL Color, Font & Cell Fill Extraction\n"
            "• 🖨️ Landscape & Portrait Print Auto-Fit (A4)\n"
            "• ↔️ Horizontal Scrollbar (Scroll-ka Hoose for 27+ columns)\n"
            "• 📘 Word Formatted Reader & HTML Web Previewer\n"
            "• 📕 PDF Visualizer & Text Extractor\n"
            "• ⚙️ Machine Code & Hex Inspector (Executables ONLY)\n\n"
            "La horumariyay 2026."
        )
        messagebox.showinfo("Laga Helay (About)", msg)


if __name__ == "__main__":
    app = tk.Tk()
    
    file_to_open = None
    if len(sys.argv) > 1:
        file_to_open = " ".join(sys.argv[1:])
        
    notebook = UniversalNotebook(app, file_to_open)
    app.mainloop()
