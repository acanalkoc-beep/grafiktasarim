"""Grafik Oluşturucu v2 — spreadsheet, analiz, proje ve çoklu eksen sürümü."""
from __future__ import annotations

import copy
import csv
import json
import math
import os
import subprocess
import tempfile
import tkinter as tk
import tkinter.font as tkfont
from dataclasses import asdict
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.ticker import LogLocator, NullFormatter

from grafik_olusturucu_origin_advanced_v3_2_1_stable import (
    LINESTYLE_OPTIONS, MARKER_OPTIONS, PALETTES, SETTINGS, XYSeries,
    PlotCreatorApp, _read_csv_rows, _to_float, add_legend,
    apply_axis_limits, apply_comma_tick_formatter, apply_scientific_style,
    optional_float, style_labels,
)

TR_MARKERS = {
    "Daire (o)":"Circle (o)", "Kare (s)":"Square (s)", "Yukarı Üçgen (^)":"Triangle Up (^)",
    "Aşağı Üçgen (v)":"Triangle Down (v)", "Sol Üçgen (<)":"Triangle Left (<)", "Sağ Üçgen (>)":"Triangle Right (>)",
    "Elmas (D)":"Diamond (D)", "İnce Elmas (d)":"Thin Diamond (d)", "Beşgen (p)":"Pentagon (p)",
    "Altıgen 1 (h)":"Hexagon 1 (h)", "Altıgen 2 (H)":"Hexagon 2 (H)", "Artı (+)":"Plus (+)",
    "Dolu Artı (P)":"Filled Plus (P)", "Çarpı (x)":"Cross (x)", "Dolu Çarpı (X)":"Filled Cross (X)",
    "Yıldız (*)":"Star (*)", "Nokta (.)":"Point (.)", "Piksel (, )":"Pixel (,)",
    "Dikey Çizgi (|)":"Vertical Line (|)", "Yatay Çizgi (_ )":"Horizontal Line (_)", "Yok":"None",
}
TR_LINES={"Düz":"Solid","Kesikli":"Dashed","Kesik noktalı":"Dash-dot","Noktalı":"Dotted","Yok":"None"}
TR_MODES={"Çizgi + Sembol":"Line + Symbol","Yalnız Çizgi":"Line Only","Yalnız Sembol":"Symbol Only"}
TR_FILLS={"Açık":"Open","Dolu":"Filled"}
TR_COLORS={"Otomatik / Palet":"Auto / Palette","Siyah":"Black","Kırmızı":"Red","Mavi":"Blue","Yeşil":"Green","Turuncu":"Orange","Mor":"Purple","Camgöbeği":"Cyan","Macenta":"Magenta","Gri":"Gray","Kahverengi":"Brown"}
TR_LEGEND={"En uygun":"best","Sağ üst":"upper right","Sol üst":"upper left","Sağ alt":"lower right","Sol alt":"lower left","Sağ orta":"center right","Sol orta":"center left","Üst orta":"upper center","Alt orta":"lower center","Orta":"center"}


class GrafikOlusturucuV2(PlotCreatorApp):
    def __init__(self):
        self.language = "en"
        self.dark_mode = False
        self.sheet_headers: list[str] = []
        self.sheet_rows: list[list[str]] = []
        self.sheet_sort_reverse: dict[int, bool] = {}
        self.sheet_active_column = 0
        self.sheet_active_row = 0
        self.sheet_selection_start = None
        self.sheet_selected_cells = set()
        self.sheet_selection_overlays = []
        self.fit_series: set[int] = set()
        self.area_series: set[int] = set()
        self.area_values = {}
        self.analysis_history = []
        self.analysis_artists = []
        self.polynomial_fits = {}
        self.recent_history_file=Path(__file__).with_name(".scientific_graph_studio_recent.json")
        try:self.recent_projects=json.loads(self.recent_history_file.read_text(encoding="utf-8"))
        except Exception:self.recent_projects=[]
        self.axis_break = None
        self.annotation_artists = []
        self.annotation_mode = None
        self.annotation_clicks = []
        self.annotation_text = ""
        super().__init__()
        self.plot_font_family = tk.StringVar(value="Arial")
        self.plot_font_size = tk.IntVar(value=11)
        self.plot_font_italic = tk.BooleanVar(value=False)
        self.grid_line_width = tk.DoubleVar(value=0.75)
        self.plot_background = tk.StringVar(value="#FFFFFF")
        self.title("Scientific Graph Studio (Beta)")
        self.set_language(self.language)
        self.protocol("WM_DELETE_WINDOW",self.close_application)

    def _setup_ttk(self):
        self.style = ttk.Style(self)
        try: self.style.theme_use("clam")
        except tk.TclError: pass
        self.apply_theme()

    def apply_theme(self):
        dark = self.dark_mode
        bg, panel, fg, muted, accent = (("#151A22","#202733","#EDF2F7","#98A6B8","#4F8CFF") if dark else
                                        ("#F3F6FA","#FFFFFF","#182230","#6B7787","#2563EB"))
        self.configure(bg=bg)
        for style_name in ("TFrame","TLabelframe","TNotebook","TNotebook.Tab"):
            self.style.configure(style_name, background=bg, foreground=fg)
        self.style.configure("TLabelframe", background=panel, bordercolor="#3A4657" if dark else "#D8E0EA")
        self.style.configure("TLabelframe.Label", background=panel, foreground=fg, font=("Arial",9,"bold"))
        self.style.configure("TLabel", background=bg, foreground=fg)
        self.style.configure("Title.TLabel", font=("Arial",16,"bold"), background=bg, foreground=fg)
        self.style.configure("Section.TLabel", font=("Arial",10,"bold"), background=bg, foreground=fg)
        self.style.configure("Muted.TLabel", background=bg, foreground=muted)
        self.style.configure("TButton", padding=(8,5), background=panel, foreground=fg)
        self.style.configure("Compact.TButton", padding=(6,3), font=("Arial",8))
        self.style.configure("Accent.TButton", padding=(9,5), background=accent, foreground="white", font=("Arial",9,"bold"))
        self.style.map("Accent.TButton", background=[("active", "#1D4ED8")])
        self.style.configure("Square.TButton", padding=(12,12), background=accent, foreground="white", font=("Arial",10,"bold"))
        self.style.configure("Treeview", background=panel, fieldbackground=panel, foreground=fg, rowheight=25)
        self.style.configure("Treeview.Heading", background="#303A49" if dark else "#E7EDF5", foreground=fg)
        self.style.configure("TEntry", fieldbackground=panel, foreground=fg)
        self.style.configure("TCombobox", fieldbackground=panel, foreground=fg)
        if hasattr(self, "control_canvas"): self.control_canvas.configure(bg=bg)

    def _build_ui(self):
        self._build_menu()
        root = ttk.Frame(self, padding=8)
        root.pack(fill="both", expand=True)
        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 7))
        self.app_title_var = tk.StringVar(value="Scientific Graph Studio (Beta)")
        ttk.Label(header, textvariable=self.app_title_var, style="Title.TLabel").pack(side="left")
        self.preview_full_btn = ttk.Button(header, text="⛶ Önizlemeyi Büyüt", command=self.open_fullscreen_preview, style="Compact.TButton")
        self.preview_full_btn.pack(side="right")

        self.main_pane = ttk.Panedwindow(root, orient="horizontal")
        self.main_pane.pack(fill="both", expand=True)
        self.controls_host = ttk.Frame(self.main_pane, padding=(2, 2, 6, 2))
        preview_host = ttk.Frame(self.main_pane, padding=(6, 2, 2, 2))
        self.main_pane.add(self.controls_host, weight=1)
        self.main_pane.add(preview_host, weight=1)
        self.main_pane.bind("<Configure>", lambda _e: self.after_idle(self._set_equal_panes))

        self.control_canvas = tk.Canvas(self.controls_host, highlightthickness=0)
        self.controls = ttk.Frame(self.control_canvas, padding=(3, 3, 8, 8))
        self.control_window = self.control_canvas.create_window((0, 0), window=self.controls, anchor="nw")
        self.controls.bind("<Configure>", lambda _e: self.control_canvas.configure(scrollregion=self.control_canvas.bbox("all")))
        self.control_canvas.bind("<Configure>", lambda e: self.control_canvas.itemconfigure(self.control_window, width=e.width))
        self.control_canvas.pack(side="left", fill="both", expand=True)
        self._bind_mousewheel(self.control_canvas)

        actions = ttk.Frame(self.controls)
        actions.pack(fill="x", pady=(0, 7))
        self.load_btn = ttk.Button(actions, text="Veri Yükle", command=self.choose_csv, style="Accent.TButton")
        self.load_btn.pack(side="left")
        self.data_btn = ttk.Button(actions, text="Veri Tablosu", command=self.open_data_window, style="Compact.TButton")
        self.data_btn.pack(side="left", padx=5)
        self.graph_settings_btn = ttk.Button(actions, text="Grafik Ayarları", command=self.open_plot_settings_dialog, style="Compact.TButton")
        self.graph_settings_btn.pack(side="left")
        self.excel_file_label = ttk.Label(self.controls, text="Veri yüklenmedi", style="Muted.TLabel")
        self.excel_file_label.pack(fill="x", pady=(0, 6))

        type_row = ttk.Frame(self.controls)
        type_row.pack(fill="x", pady=(0, 6))
        self.plot_type_label = ttk.Label(type_row, text="Grafik türü:")
        self.plot_type_label.pack(side="left")
        self.plot_type = tk.StringVar(value="line")
        self.plot_family = tk.StringVar(value="line")
        self.line_radio = ttk.Radiobutton(type_row, text="Çizgi", variable=self.plot_family, value="line",command=self.update_plot_type_choices)
        self.line_radio.pack(side="left", padx=8)
        self.bar_radio = ttk.Radiobutton(type_row, text="Sütun", variable=self.plot_family, value="bar",command=self.update_plot_type_choices)
        self.bar_radio.pack(side="left")
        self.extra_plot_combo=ttk.Combobox(type_row,textvariable=self.plot_type,values=("line","area"),state="readonly",width=11)
        self.extra_plot_combo.pack(side="left",padx=6)
        self.plot_type.trace_add("write", lambda *_: self.update_series_panel_for_plot_type())

        self.first_row_header = tk.BooleanVar(value=True)
        self.sort_by_x = tk.BooleanVar(value=False)
        self._build_compact_series_panel()
        self._init_plot_variables()
        self.draw_btn = ttk.Button(self.controls, text="Grafik Oluştur", command=self.draw_selected_plot, style="Square.TButton")
        self.draw_btn.pack(anchor="center", pady=8, ipadx=5, ipady=5)

        preview_top = ttk.Frame(preview_host)
        preview_top.pack(fill="x", pady=(0, 5))
        self.preview_label_var = tk.StringVar(value="Önizleme")
        ttk.Label(preview_top, textvariable=self.preview_label_var, style="Section.TLabel").pack(side="left")
        self.text_tool_btn = ttk.Button(preview_top, text="T+", width=4, command=self.add_text_annotation, style="Compact.TButton")
        self.text_tool_btn.pack(side="right")
        self.arrow_tool_btn = ttk.Button(preview_top, text="➜", width=4, command=self.add_arrow_annotation, style="Compact.TButton")
        self.arrow_tool_btn.pack(side="right", padx=4)
        self.shape_tool_btn=ttk.Menubutton(preview_top,text=self.tr("Şekil","Shape"),style="Compact.TButton")
        shape_menu=tk.Menu(self.shape_tool_btn,tearoff=False);self.shape_tool_btn.configure(menu=shape_menu)
        shape_menu.add_command(label=self.tr("Çember","Circle"),command=lambda:self.add_shape_annotation("circle"))
        shape_menu.add_command(label=self.tr("Dikdörtgen","Rectangle"),command=lambda:self.add_shape_annotation("rectangle"))
        shape_menu.add_command(label=self.tr("Elips","Ellipse"),command=lambda:self.add_shape_annotation("ellipse"))
        self.shape_tool_btn.pack(side="right",padx=4)
        self.delete_annotation_btn = ttk.Button(preview_top, text="⌫", width=4, command=self.delete_last_annotation, style="Compact.TButton")
        self.delete_annotation_btn.pack(side="right")
        self.preview_frame = ttk.Frame(preview_host)
        self.preview_frame.pack(fill="both", expand=True)
        self._show_placeholder()
        self.update_series_panel_for_plot_type()
        self.after(150, self._set_equal_panes)

    def _build_menu(self):
        menu = tk.Menu(self)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label=self.tr("Proje Aç…", "Open Project…"), command=self.open_project)
        file_menu.add_command(label=self.tr("Projeyi Kaydet…", "Save Project…"), command=self.save_project)
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Grafiği Dışa Aktar…", "Export Graph…"), command=self.export_current)
        file_menu.add_command(label=self.tr("Yazdır…", "Print…"), command=self.print_current)
        self.recent_menu=tk.Menu(file_menu,tearoff=False)
        file_menu.add_cascade(label=self.tr("Son Projeler", "Recent Projects"),menu=self.recent_menu)
        self.refresh_recent_menu()
        file_menu.add_separator()
        file_menu.add_command(label=self.tr("Çıkış", "Exit"), command=self.destroy)
        menu.add_cascade(label=self.tr("Dosya", "File"), menu=file_menu)

        analysis = tk.Menu(menu, tearoff=False)
        analysis.add_command(label=self.tr("Doğrusal Regresyon", "Linear Regression"), command=self.linear_fit)
        analysis.add_command(label=self.tr("Polinom Uyumu…", "Polynomial Fit…"), command=self.polynomial_fit)
        analysis.add_command(label=self.tr("Uyum Raporu (Eğim, R², F)", "Fit Report (Slope, R², F)"), command=self.fit_report)
        analysis.add_command(label=self.tr("F Testi", "F Test"), command=self.f_test)
        analysis.add_command(label=self.tr("Tanımlayıcı İstatistikler", "Descriptive Statistics"), command=self.descriptive_statistics)
        analysis.add_command(label=self.tr("Eğri Altındaki Alan", "Area Under Curve"), command=self.area_under_curve)
        analysis.add_separator()
        analysis.add_command(label=self.tr("Son Analizi Geri Al", "Undo Last Analysis"), command=self.undo_analysis)
        analysis.add_command(label=self.tr("Analiz Sonuçları…", "Analysis Results…"), command=self.show_analysis_results)
        menu.add_cascade(label=self.tr("İstatistik", "Statistics"), menu=analysis)

        template = tk.Menu(menu, tearoff=False)
        template.add_command(label=self.tr("Şablonu Kaydet…", "Save Template…"), command=self.save_template)
        template.add_command(label=self.tr("Şablon Uygula…", "Apply Template…"), command=self.load_template)
        settings = tk.Menu(menu, tearoff=False)
        settings.add_command(label=self.tr("Grafik Ayarları…", "Graph Settings…"), command=self.open_plot_settings_dialog)
        settings.add_command(label=self.tr("Açık/Koyu Mod", "Light/Dark Mode"), command=self.toggle_dark_mode)
        language = tk.Menu(settings, tearoff=False)
        language.add_command(label="Türkçe", command=lambda: self.set_language("tr"))
        language.add_command(label="English", command=lambda: self.set_language("en"))
        settings.add_cascade(label=self.tr("Dil", "Language"), menu=language)
        settings.add_separator()
        settings.add_cascade(label=self.tr("Şablon", "Template"), menu=template)
        menu.add_cascade(label=self.tr("Ayarlar", "Settings"), menu=settings)

        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label=self.tr("Yardım ve Komutlar", "Help and Commands"), command=self.show_help)
        help_menu.add_command(label=self.tr("Hakkında", "About"), command=self.show_about)
        menu.add_cascade(label=self.tr("Hakkında", "About"), menu=help_menu)
        self.config(menu=menu)

    def tr(self, tr_text, en_text):
        return tr_text if self.language == "tr" else en_text

    def legend_location(self):
        return TR_LEGEND.get(self.legend_loc_var.get(), self.legend_loc_var.get())

    def refresh_recent_menu(self):
        if not hasattr(self,"recent_menu"):return
        self.recent_menu.delete(0,"end")
        for path in self.recent_projects[:10]:self.recent_menu.add_command(label=Path(path).name,command=lambda p=path:self.open_project_path(p))
        if not self.recent_projects:self.recent_menu.add_command(label=self.tr("Geçmiş boş","No recent projects"),state="disabled")

    def add_recent_project(self,path):
        path=str(Path(path).resolve());self.recent_projects=[p for p in self.recent_projects if p!=path];self.recent_projects.insert(0,path);self.refresh_recent_menu()
        try:self.recent_history_file.write_text(json.dumps(self.recent_projects[:10],ensure_ascii=False,indent=2),encoding="utf-8")
        except OSError:pass

    def close_application(self):
        try:self.recent_history_file.write_text(json.dumps(self.recent_projects[:10],ensure_ascii=False,indent=2),encoding="utf-8")
        except OSError:pass
        self.destroy()

    def update_plot_type_choices(self):
        if self.plot_family.get()=="line":
            self.extra_plot_combo.configure(values=("line","area"));self.plot_type.set("line")
        else:
            self.extra_plot_combo.configure(values=("bar","pie","3d_column","3d_bar"));self.plot_type.set("bar")

    def _bind_mousewheel(self, canvas):
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))
        canvas.bind_all("<Button-4>", lambda _e: canvas.yview_scroll(-1, "units"))
        canvas.bind_all("<Button-5>", lambda _e: canvas.yview_scroll(1, "units"))

    def _set_equal_panes(self):
        try:
            self.main_pane.sashpos(0, max(420, self.main_pane.winfo_width() // 2))
        except tk.TclError:
            pass

    def _init_plot_variables(self):
        def entry(default=""):
            widget = ttk.Entry(self.controls); widget.insert(0, default); return widget
        self.line_title, self.line_xlabel, self.line_ylabel = entry(), entry("X"), entry("Y")
        self.line_ymin, self.line_ymax, self.line_ystep = entry(), entry(), entry()
        self.line_xmin, self.line_xmax, self.line_xstep = entry(), entry(), entry()
        self.line_grid = tk.BooleanVar(value=False); self.line_legend = tk.BooleanVar(value=True)
        self.line_logx = tk.BooleanVar(value=False); self.line_logy = tk.BooleanVar(value=False)
        self.top_axis_var = tk.BooleanVar(value=True); self.right_axis_var = tk.BooleanVar(value=True)
        self.top_tick_var = tk.BooleanVar(value=True); self.right_tick_var = tk.BooleanVar(value=True)
        self.legend_loc_var = tk.StringVar(value="best")
        self.legend_cols = ttk.Spinbox(self.controls, from_=1, to=5); self.legend_cols.set("1")

    @staticmethod
    def _trapezoid(y, x):
        y=np.asarray(y,dtype=float); x=np.asarray(x,dtype=float)
        if len(x)<2:return 0.0
        return float(np.sum((x[1:]-x[:-1])*(y[1:]+y[:-1])*.5))

    @staticmethod
    def _betacf(a,b,x):
        qab=a+b; qap=a+1.; qam=a-1.; c=1.; d=1.-qab*x/qap
        if abs(d)<3e-14:d=3e-14
        d=1./d; h=d
        for m in range(1,201):
            m2=2*m; aa=m*(b-m)*x/((qam+m2)*(a+m2)); d=1.+aa*d; d=3e-14 if abs(d)<3e-14 else d
            c=1.+aa/c; c=3e-14 if abs(c)<3e-14 else c; d=1./d; h*=d*c
            aa=-(a+m)*(qab+m)*x/((a+m2)*(qap+m2)); d=1.+aa*d; d=3e-14 if abs(d)<3e-14 else d
            c=1.+aa/c; c=3e-14 if abs(c)<3e-14 else c; d=1./d; delta=d*c; h*=delta
            if abs(delta-1.)<3e-12:break
        return h

    @classmethod
    def _regularized_beta(cls,x,a,b):
        if x<=0:return 0.0
        if x>=1:return 1.0
        bt=math.exp(math.lgamma(a+b)-math.lgamma(a)-math.lgamma(b)+a*math.log(x)+b*math.log1p(-x))
        return bt*cls._betacf(a,b,x)/a if x<(a+1)/(a+b+2) else 1-bt*cls._betacf(b,a,1-x)/b

    @classmethod
    def _f_survival(cls,f_value,df1,df2):
        if not math.isfinite(f_value):return 0.0
        x=(df1*f_value)/(df1*f_value+df2)
        return max(0.0,min(1.0,1-cls._regularized_beta(x,df1/2,df2/2)))

    def _build_compact_series_panel(self):
        self.series_box = ttk.LabelFrame(self.controls, text="Seriler", padding=6)
        box=self.series_box; box.pack(fill="x", pady=4)
        self.series_tree = ttk.Treeview(box, columns=("name","n","axis"), show="headings", height=4, selectmode="browse")
        for col, label, width in (("name","Seri",160),("n","N",45),("axis","Y",55)):
            self.series_tree.heading(col,text=label); self.series_tree.column(col,width=width,anchor="center")
        self.series_tree.pack(fill="x")
        self.series_tree.bind("<<TreeviewSelect>>", self._on_series_select)
        editor = ttk.Frame(box); editor.pack(fill="x", pady=(6,0))
        self.series_name_var=tk.StringVar(); self.series_name_entry=ttk.Entry(editor,textvariable=self.series_name_var,width=17)
        self.series_name_entry.grid(row=0,column=0,columnspan=2,sticky="ew",padx=(0,5))
        self.series_axis_side_var=tk.StringVar(value="left")
        self.series_axis_side_combo=ttk.Combobox(editor,textvariable=self.series_axis_side_var,values=("left","right"),state="readonly",width=7)
        self.series_axis_side_combo.grid(row=0,column=2)
        self.series_visible_var=tk.BooleanVar(value=True); ttk.Checkbutton(editor,text="✓",variable=self.series_visible_var).grid(row=0,column=3)
        self.series_mode_var=tk.StringVar(value="Line + Symbol")
        self.series_mode_combo=ttk.Combobox(editor,textvariable=self.series_mode_var,values=("Line + Symbol","Line Only","Symbol Only"),state="readonly",width=15)
        ttk.Label(editor,text=self.tr("Çizim modu","Plot mode")).grid(row=1,column=0,sticky="w");self.series_mode_combo.grid(row=1,column=1,pady=3)
        self.series_marker_name_var=tk.StringVar(value="Circle (o)")
        self.series_marker_combo=ttk.Combobox(editor,textvariable=self.series_marker_name_var,values=list(MARKER_OPTIONS),state="readonly",width=15)
        ttk.Label(editor,text=self.tr("Sembol","Marker")).grid(row=2,column=0,sticky="w");self.series_marker_combo.grid(row=2,column=1,padx=3)
        self.series_fill_var=tk.StringVar(value="Open")
        self.series_fill_combo=ttk.Combobox(editor,textvariable=self.series_fill_var,values=("Open","Filled"),state="readonly",width=8)
        self.series_fill_combo.grid(row=2,column=2)
        self.series_marker_size_var=tk.StringVar(value=str(SETTINGS.marker_size)); self.series_marker_size_spin=ttk.Spinbox(editor,textvariable=self.series_marker_size_var,from_=.5,to=30,width=6)
        ttk.Label(editor,text=self.tr("Sembol boyutu","Marker size")).grid(row=3,column=0,sticky="w");self.series_marker_size_spin.grid(row=3,column=1)
        self.series_marker_edge_width_var=tk.StringVar(value=str(SETTINGS.marker_edge_width)); self.series_marker_edge_width_spin=ttk.Spinbox(editor,textvariable=self.series_marker_edge_width_var,from_=.1,to=8,width=6)
        ttk.Label(editor,text=self.tr("Sembol kalınlığı","Marker edge width")).grid(row=4,column=0,sticky="w");self.series_marker_edge_width_spin.grid(row=4,column=1)
        self.series_line_style_name_var=tk.StringVar(value="Solid"); self.series_line_style_combo=ttk.Combobox(editor,textvariable=self.series_line_style_name_var,values=list(LINESTYLE_OPTIONS),state="readonly",width=8)
        ttk.Label(editor,text=self.tr("Çizgi stili","Line style")).grid(row=5,column=0,sticky="w");self.series_line_style_combo.grid(row=5,column=1)
        self.series_line_width_var=tk.StringVar(value=str(SETTINGS.line_width)); self.series_line_width_spin=ttk.Spinbox(editor,textvariable=self.series_line_width_var,from_=.1,to=8,width=6)
        ttk.Label(editor,text=self.tr("Çizgi kalınlığı","Line width")).grid(row=6,column=0,sticky="w");self.series_line_width_spin.grid(row=6,column=1,pady=3)
        self.series_marker_every_var=tk.StringVar(value="1"); self.series_marker_every_spin=ttk.Spinbox(editor,textvariable=self.series_marker_every_var,from_=1,to=10000,width=6)
        ttk.Label(editor,text=self.tr("Sembol aralığı","Marker interval")).grid(row=7,column=0,sticky="w");self.series_marker_every_spin.grid(row=7,column=1)
        self.series_alpha_var=tk.StringVar(value="1.0"); self.series_alpha_spin=ttk.Spinbox(editor,textvariable=self.series_alpha_var,from_=.05,to=1,width=6)
        ttk.Label(editor,text=self.tr("Saydamlık","Opacity")).grid(row=8,column=0,sticky="w");self.series_alpha_spin.grid(row=8,column=1)
        self.series_color_var=tk.StringVar(value="Auto / Palette"); self.series_color_combo=ttk.Combobox(editor,textvariable=self.series_color_var,values=("Auto / Palette","Black","Red","Blue","Green","Orange","Purple","Cyan","Magenta","Gray","Brown"),state="readonly",width=15)
        ttk.Label(editor,text=self.tr("Renk","Color")).grid(row=9,column=0,sticky="w");self.series_color_combo.grid(row=9,column=1,columnspan=2,sticky="ew")
        self.series_color_preview=tk.Label(editor,text="  ",bg="white"); self.series_color_preview.grid(row=9,column=3)
        ttk.Button(editor,text="…",width=3,command=self.choose_series_custom_color,style="Compact.TButton").grid(row=9,column=4)
        self.apply_series_btn=ttk.Button(editor,text="Uygula",command=self.apply_selected_series_settings,style="Compact.TButton")
        self.apply_series_btn.grid(row=10,column=0,pady=(4,0))
        self.reset_series_btn=ttk.Button(editor,text="Sıfırla",command=self.reset_all_series_styles,style="Compact.TButton")
        self.reset_series_btn.grid(row=10,column=1,pady=(4,0))
        self.line_series_widgets=[self.series_mode_combo,self.series_marker_combo,self.series_fill_combo,
            self.series_marker_size_spin,self.series_marker_edge_width_spin,self.series_line_style_combo,
            self.series_line_width_spin,self.series_marker_every_spin]
        self.bar_series_frame=ttk.Frame(editor)
        self.bar_width_var=tk.StringVar(value=str(SETTINGS.bar_width))
        ttk.Label(self.bar_series_frame,text=self.tr("Sütun genişliği","Bar width")).pack(side="left")
        ttk.Spinbox(self.bar_series_frame,textvariable=self.bar_width_var,from_=.05,to=1.0,increment=.05,width=7).pack(side="left",padx=5)
        editor.columnconfigure(0,weight=1); editor.columnconfigure(1,weight=1)

    def update_series_panel_for_plot_type(self):
        if not hasattr(self,"line_series_widgets"):return
        if self.plot_family.get()=="bar":
            for widget in self.line_series_widgets: widget.grid_remove()
            self.bar_series_frame.grid(row=1,column=0,columnspan=4,sticky="ew",pady=5)
        else:
            self.bar_series_frame.grid_remove()
            for widget in self.line_series_widgets: widget.grid()

    def _refresh_series_tree(self):
        if not hasattr(self, "series_tree"): return
        self.series_tree.delete(*self.series_tree.get_children())
        for i,s in enumerate(self.loaded_xy_series):
            side=getattr(s,"y_axis_side","left")
            if self.language=="tr":side="Sol" if side=="left" else "Sağ"
            self.series_tree.insert("","end",iid=str(i),values=(s.name,len(s.x),side))

    def open_data_window(self):
        if hasattr(self, "sheet_window") and self.sheet_window.winfo_exists():
            self.sheet_window.lift(); return
        self.sheet_window = tk.Toplevel(self)
        self.sheet_window.title(self.tr("Veri Tablosu", "Data Table"))
        self.sheet_window.geometry("1000x620")
        self.sheet_window.minsize(650, 420)
        sheet_tab = ttk.Frame(self.sheet_window, padding=7)
        sheet_tab.pack(fill="both", expand=True)
        tools = ttk.Frame(sheet_tab)
        tools.pack(fill="x", pady=(0, 5))
        ttk.Button(tools, text=self.tr("Satır Ekle", "Add Row"), command=self.add_sheet_row).pack(side="left")
        ttk.Button(tools, text=self.tr("Satır Sil", "Delete Row"), command=self.delete_sheet_rows).pack(side="left", padx=4)
        ttk.Button(tools, text=self.tr("Sütun Ekle", "Add Column"), command=self.add_sheet_column).pack(side="left")
        ttk.Button(tools, text=self.tr("Sütun Sil", "Delete Column"), command=self.delete_sheet_column).pack(side="left", padx=4)
        ttk.Button(tools, text=self.tr("Yapıştır", "Paste"), command=self.paste_sheet_data).pack(side="left")
        ttk.Button(tools, text=self.tr("Transpoze", "Transpose"), command=self.transpose_data).pack(side="left", padx=4)
        ttk.Label(tools, text=self.tr("Filtre:", "Filter:")).pack(side="left", padx=(12, 3))
        self.sheet_filter = tk.StringVar()
        entry = ttk.Entry(tools, textvariable=self.sheet_filter, width=18)
        entry.pack(side="left", fill="x", expand=True)
        self.sheet_filter.trace_add("write", lambda *_: self.refresh_sheet())

        frame = ttk.Frame(sheet_tab)
        frame.pack(fill="both", expand=True)
        self.sheet = ttk.Treeview(frame, show="tree headings", selectmode="none", height=22)
        self.sheet.heading("#0",text="#");self.sheet.column("#0",width=48,minwidth=42,stretch=False,anchor="center")
        sy = ttk.Scrollbar(frame, orient="vertical", command=self.sheet.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=self.sheet.xview)
        self.sheet.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.sheet.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        self.sheet.bind("<Double-1>", self.edit_sheet_cell)
        self.sheet.bind("<Button-1>", self.begin_sheet_selection)
        self.sheet.bind("<B1-Motion>", self.extend_sheet_selection)
        self.sheet.bind("<ButtonRelease-1>", self.end_sheet_selection)
        self.sheet.bind("<Button-3>", self.sheet_context_menu)
        self.sheet.bind("<Control-v>", lambda _e: self.paste_sheet_data())
        self.sheet.bind("<Command-v>", lambda _e: self.paste_sheet_data())
        self.sheet.bind("<Delete>", lambda _e: self.clear_selected_cells())
        self.sheet.bind("<BackSpace>", lambda _e: self.clear_selected_cells())
        self.refresh_sheet()

    def choose_csv(self):
        path = filedialog.askopenfilename(title=self.tr("Veri dosyası seç", "Select data file"),
            filetypes=[(self.tr("CSV / Metin", "CSV / Text"), "*.csv *.txt"), (self.tr("Tüm dosyalar", "All files"), "*.*")])
        if not path:
            return
        try:
            rows = _read_csv_rows(path)
            if not rows:
                raise ValueError("Dosya boş.")
            width = max(map(len, rows))
            self.sheet_headers = [self.tr(f"Sütun {i+1}", f"Column {i+1}") for i in range(width)]
            data = rows
            self.sheet_rows = [list(r) + [""] * (width - len(r)) for r in data]
            self.excel_path = path
            self.excel_file_label.configure(text=Path(path).name)
            self.refresh_sheet()
            self.sync_series_from_sheet(preserve=False)
        except Exception as exc:
            messagebox.showerror(self.tr("Veri açılamadı", "Could not open data"), str(exc))

    def reload_excel_preview(self):
        if self.sheet_headers:
            self.sync_series_from_sheet(preserve=True)
        elif self.excel_path:
            super().reload_excel_preview()

    def refresh_sheet(self):
        if not hasattr(self, "sheet") or not self.sheet.winfo_exists():
            return
        cols = [f"c{i}" for i in range(len(self.sheet_headers))]
        self.sheet.configure(columns=cols)
        for i, col in enumerate(cols):
            self.sheet.heading(col, text=self.sheet_headers[i])
            self.sheet.column(col, width=105, minwidth=55, anchor="center")
        self._clear_selection_overlays()
        self.sheet.delete(*self.sheet.get_children())
        needle = self.sheet_filter.get().casefold() if hasattr(self, "sheet_filter") else ""
        for source_index, row in enumerate(self.sheet_rows):
            if needle and not any(needle in str(v).casefold() for v in row):
                continue
            self.sheet.insert("", "end", iid=str(source_index), text=str(source_index+1), values=row)
        self.sheet.after_idle(self._draw_cell_selection)

    def remember_sheet_cell(self, event):
        col = self.sheet.identify_column(event.x)
        if col:
            self.sheet_active_column = max(0, int(col[1:]) - 1)
        row = self.sheet.identify_row(event.y)
        if row:
            self.sheet_active_row = int(row)

    def _cell_from_xy(self,x,y):
        row=self.sheet.identify_row(y); col=self.sheet.identify_column(x)
        if not row or not col or self.sheet.identify_region(x,y)!="cell":return None
        return int(row),int(col[1:])-1

    def begin_sheet_selection(self,event):
        cell=self._cell_from_xy(event.x,event.y)
        if not cell:return
        self.sheet.focus_set(); self.sheet_selection_start=cell; self.sheet_active_row,self.sheet_active_column=cell
        self.sheet_selected_cells={cell}; self._draw_cell_selection()

    def extend_sheet_selection(self,event):
        cell=self._cell_from_xy(event.x,event.y)
        if not cell or self.sheet_selection_start is None:return
        r1,c1=self.sheet_selection_start; r2,c2=cell
        self.sheet_selected_cells={(r,c) for r in range(min(r1,r2),max(r1,r2)+1) for c in range(min(c1,c2),max(c1,c2)+1)}
        self.sheet_active_row,self.sheet_active_column=cell; self._draw_cell_selection()

    def end_sheet_selection(self,event):
        self.extend_sheet_selection(event)

    def _clear_selection_overlays(self):
        for overlay in self.sheet_selection_overlays:
            try: overlay.destroy()
            except tk.TclError: pass
        self.sheet_selection_overlays=[]

    def _draw_cell_selection(self):
        if not hasattr(self,"sheet") or not self.sheet.winfo_exists():return
        self._clear_selection_overlays()
        for row,col in self.sheet_selected_cells:
            bbox=self.sheet.bbox(str(row),f"#{col+1}")
            if not bbox:continue
            x,y,w,h=bbox; value=self.sheet_rows[row][col] if row<len(self.sheet_rows) and col<len(self.sheet_rows[row]) else ""
            overlay=tk.Label(self.sheet,text=value,bg="#DCEAFF",fg="#102A56",highlightbackground="#2563EB",highlightthickness=2,bd=0)
            overlay.place(x=x,y=y,width=w,height=h)
            overlay.bind("<Button-1>",lambda e,r=row,c=col:self._overlay_begin(e,r,c))
            overlay.bind("<Double-1>",lambda e,r=row,c=col:self.edit_selected_cell(r,c))
            overlay.bind("<B1-Motion>",self._overlay_motion)
            self.sheet_selection_overlays.append(overlay)

    def _overlay_begin(self,event,row,col):
        self.sheet.focus_set(); self.sheet_selection_start=(row,col); self.sheet_active_row=row; self.sheet_active_column=col
        self.sheet_selected_cells={(row,col)}; self._draw_cell_selection()

    def _overlay_motion(self,event):
        x=event.x_root-self.sheet.winfo_rootx(); y=event.y_root-self.sheet.winfo_rooty()
        class E:pass
        forwarded=E(); forwarded.x=x; forwarded.y=y
        self.extend_sheet_selection(forwarded)

    def clear_selected_cells(self):
        for row,col in self.sheet_selected_cells:
            if 0<=row<len(self.sheet_rows) and 0<=col<len(self.sheet_rows[row]):self.sheet_rows[row][col]=""
        self.refresh_sheet(); self.sync_series_from_sheet(preserve=True)

    def edit_selected_cell(self,row,col):
        value=simpledialog.askstring(self.tr("Hücreyi Düzenle","Edit Cell"),self.tr("Değer:","Value:"),initialvalue=self.sheet_rows[row][col],parent=self.sheet_window)
        if value is not None:
            self.sheet_rows[row][col]=value; self.refresh_sheet(); self.sync_series_from_sheet(preserve=True)

    def edit_sheet_cell(self, event):
        item = self.sheet.identify_row(event.y)
        col = self.sheet.identify_column(event.x)
        if not col:
            return
        ci = int(col[1:]) - 1
        if self.sheet.identify_region(event.x, event.y) == "heading":
            name = simpledialog.askstring(self.tr("Sütun Başlığı", "Column Header"),
                self.tr("Yeni başlık:", "New header:"), initialvalue=self.sheet_headers[ci], parent=self.sheet_window)
            if name is not None:
                self.sheet_headers[ci] = name.strip() or self.tr(f"Sütun {ci+1}", f"Column {ci+1}")
                self.refresh_sheet(); self.sync_series_from_sheet(preserve=True)
            return
        if not item:
            return
        index = int(item)
        x, y, w, h = self.sheet.bbox(item, col)
        edit = ttk.Entry(self.sheet)
        edit.insert(0, self.sheet_rows[index][ci])
        edit.place(x=x, y=y, width=w, height=h)
        edit.focus_set()
        edit.select_range(0, "end")
        def commit(_=None):
            if edit.winfo_exists():
                self.sheet_rows[index][ci] = edit.get()
                edit.destroy()
                self.refresh_sheet()
                self.sync_series_from_sheet(preserve=True)
        edit.bind("<Return>", commit)
        edit.bind("<FocusOut>", commit)
        edit.bind("<Escape>", lambda _e: edit.destroy())

    def add_sheet_row(self):
        self.sheet_rows.append([""] * max(1, len(self.sheet_headers)))
        self.refresh_sheet()
        if self.sheet_rows:
            self.sheet.see(str(len(self.sheet_rows) - 1))

    def add_sheet_column(self):
        name = simpledialog.askstring(self.tr("Sütun Ekle", "Add Column"), self.tr("Sütun başlığı:", "Column header:"), parent=self.sheet_window)
        if name is None: return
        self.sheet_headers.append(name.strip() or self.tr(f"Sütun {len(self.sheet_headers)+1}",f"Column {len(self.sheet_headers)+1}"))
        for row in self.sheet_rows: row.append("")
        self.refresh_sheet(); self.sync_series_from_sheet(preserve=True)

    def delete_sheet_column(self):
        ci = self.sheet_active_column
        if not self.sheet_headers or ci >= len(self.sheet_headers): return
        del self.sheet_headers[ci]
        for row in self.sheet_rows:
            if ci < len(row): del row[ci]
        self.sheet_active_column = max(0, ci-1)
        self.refresh_sheet(); self.sync_series_from_sheet(preserve=False)

    def paste_sheet_data(self):
        try: raw = self.clipboard_get()
        except tk.TclError: return
        pasted = [line.split("\t") for line in raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
        if pasted and pasted[-1] == [""]: pasted.pop()
        if not pasted: return
        start_row = self.sheet_active_row if self.sheet_rows else 0
        start_col = self.sheet_active_column
        needed_cols = start_col + max(map(len, pasted))
        while len(self.sheet_headers) < needed_cols:
            self.sheet_headers.append(self.tr(f"Sütun {len(self.sheet_headers)+1}",f"Column {len(self.sheet_headers)+1}"))
        while len(self.sheet_rows) < start_row + len(pasted):
            self.sheet_rows.append([""] * len(self.sheet_headers))
        for row in self.sheet_rows:
            row.extend([""] * (len(self.sheet_headers)-len(row)))
        for r, values in enumerate(pasted):
            for c, value in enumerate(values): self.sheet_rows[start_row+r][start_col+c] = value
        self.refresh_sheet(); self.sync_series_from_sheet(preserve=True)

    def delete_sheet_rows(self):
        iid = self.sheet_active_row
        if 0 <= iid < len(self.sheet_rows):
            del self.sheet_rows[iid]
        self.refresh_sheet()
        self.sync_series_from_sheet(preserve=True)

    def sort_sheet(self, column):
        reverse = self.sheet_sort_reverse.get(column, False)
        def key(row):
            value = row[column] if column < len(row) else ""
            number = _to_float(value)
            return (1, str(value).casefold()) if number is None else (0, number)
        self.sheet_rows.sort(key=key, reverse=reverse)
        self.sheet_sort_reverse[column] = not reverse
        self.refresh_sheet()
        self.sync_series_from_sheet(preserve=True)

    def sheet_context_menu(self, event):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="Transpoze", command=self.transpose_data)
        menu.add_command(label="Satır Ekle", command=self.add_sheet_row)
        menu.add_command(label="Seçili Satırı Sil", command=self.delete_sheet_rows)
        menu.tk_popup(event.x_root, event.y_root)

    def transpose_data(self):
        matrix = [self.sheet_headers] + self.sheet_rows
        if not matrix:
            return
        width = max(map(len, matrix))
        matrix = [r + [""] * (width - len(r)) for r in matrix]
        transposed = [list(r) for r in zip(*matrix)]
        self.sheet_headers = [str(v) or self.tr(f"Sütun {i+1}",f"Column {i+1}") for i, v in enumerate(transposed[0])]
        self.sheet_rows = transposed[1:]
        self.refresh_sheet()
        self.sync_series_from_sheet(preserve=False)

    def sync_series_from_sheet(self, preserve=True):
        old = self.loaded_xy_series
        result = []
        for pair, xcol in enumerate(range(0, len(self.sheet_headers) - 1, 2)):
            xs, ys = [], []
            for row in self.sheet_rows:
                x = _to_float(row[xcol] if xcol < len(row) else None)
                y = _to_float(row[xcol + 1] if xcol + 1 < len(row) else None)
                if x is not None and y is not None:
                    xs.append(x); ys.append(y)
            if not xs:
                continue
            s = XYSeries(self.tr(f"Seri {pair+1}", f"Series {pair+1}"), np.asarray(xs), np.asarray(ys),
                         self.sheet_headers[xcol], self.sheet_headers[xcol+1])
            s.y_axis_side = "left"
            if preserve and pair < len(old):
                for key in ("name", "plot_mode", "marker", "marker_fill", "line_style",
                            "line_width", "marker_size", "marker_edge_width", "marker_every",
                            "alpha", "visible", "color", "y_axis_side"):
                    if hasattr(old[pair], key): setattr(s, key, getattr(old[pair], key))
            result.append(s)
        self.loaded_xy_series = result
        self._refresh_series_tree()
        if result:
            self.series_tree.selection_set("0")
            self._load_series_into_editor(0)

    def _load_series_into_editor(self, idx):
        super()._load_series_into_editor(idx)
        if hasattr(self, "series_axis_side_var") and 0 <= idx < len(self.loaded_xy_series):
            self.series_axis_side_var.set(getattr(self.loaded_xy_series[idx], "y_axis_side", "left"))
        if self.language=="tr":
            reverse=lambda mapping,value: next((k for k,v in mapping.items() if v==value),value)
            self.series_mode_var.set(reverse(TR_MODES,self.series_mode_var.get()))
            self.series_marker_name_var.set(reverse(TR_MARKERS,self.series_marker_name_var.get()))
            self.series_fill_var.set(reverse(TR_FILLS,self.series_fill_var.get()))
            self.series_line_style_name_var.set(reverse(TR_LINES,self.series_line_style_name_var.get()))
            self.series_color_var.set(reverse(TR_COLORS,self.series_color_var.get()))
            self.series_axis_side_var.set("Sol" if getattr(self.loaded_xy_series[idx],"y_axis_side","left")=="left" else "Sağ")

    def apply_selected_series_settings(self):
        idx = self._selected_series_index()
        if self.plot_family.get()=="bar":
            try:
                SETTINGS.bar_width=float(self.bar_width_var.get().replace(",","."))
                if not .05<=SETTINGS.bar_width<=1: raise ValueError
            except ValueError:
                messagebox.showerror(self.tr("Geçersiz değer","Invalid value"),self.tr("Sütun genişliği 0,05–1 arasında olmalıdır.","Bar width must be between 0.05 and 1."));return
        localized=self.language=="tr"
        if localized:
            self.series_mode_var.set(TR_MODES.get(self.series_mode_var.get(),self.series_mode_var.get()))
            self.series_marker_name_var.set(TR_MARKERS.get(self.series_marker_name_var.get(),self.series_marker_name_var.get()))
            self.series_fill_var.set(TR_FILLS.get(self.series_fill_var.get(),self.series_fill_var.get()))
            self.series_line_style_name_var.set(TR_LINES.get(self.series_line_style_name_var.get(),self.series_line_style_name_var.get()))
            self.series_color_var.set(TR_COLORS.get(self.series_color_var.get(),self.series_color_var.get()))
            axis_side="left" if self.series_axis_side_var.get()=="Sol" else "right"
        else: axis_side=self.series_axis_side_var.get()
        super().apply_selected_series_settings()
        if idx is not None and idx < len(self.loaded_xy_series) and hasattr(self, "series_axis_side_var"):
            self.loaded_xy_series[idx].y_axis_side = axis_side
            self._load_series_into_editor(idx)
            if self.excel_path:
                self._redraw_line_preview()

    def _selected_series_or_warn(self):
        idx = self._selected_series_index()
        if idx is None or idx >= len(self.loaded_xy_series):
            messagebox.showwarning(self.tr("Seri seçilmedi", "No series selected"), self.tr("Seriler bölümünden bir seri seçin.", "Select a series in the Series panel."))
            return None
        return idx

    def linear_fit(self):
        idx = self._selected_series_or_warn()
        if idx is None: return
        self.fit_series.add(idx)
        s,fit,r2,fv,p=self._calculate_fit(idx);self.record_analysis("Linear Fit",s.name,{"slope":fit.slope,"intercept":fit.intercept,"R2":r2,"F":fv,"p":p})
        self._redraw_line_preview()

    def polynomial_fit(self):
        idx=self._selected_series_or_warn()
        if idx is None:return
        degree=simpledialog.askinteger(self.tr("Polinom Uyumu","Polynomial Fit"),self.tr("Derece (2–6):","Degree (2–6):"),minvalue=2,maxvalue=6,parent=self)
        if degree is None:return
        s=self.loaded_xy_series[idx];coeff=np.polyfit(s.x,s.y,degree);pred=np.polyval(coeff,s.x)
        ss_res=float(np.sum((s.y-pred)**2));ss_tot=float(np.sum((s.y-np.mean(s.y))**2));r2=1-ss_res/ss_tot if ss_tot else 1
        self.polynomial_fits[idx]=(degree,coeff,r2);self.record_analysis(f"Polynomial Fit d={degree}",s.name,{"degree":degree,"coefficients":coeff.tolist(),"R2":r2});self._redraw_line_preview();self.show_analysis_results()

    def record_analysis(self,kind,series,values):
        self.analysis_history.append({"analysis":kind,"series":series,"values":values})

    def undo_analysis(self):
        if not self.analysis_history:return
        item=self.analysis_history.pop();kind=item["analysis"]
        idx=next((i for i,s in enumerate(self.loaded_xy_series) if s.name==item["series"]),None)
        if idx is not None:
            if kind.startswith("Linear"):self.fit_series.discard(idx)
            if kind.startswith("Polynomial"):self.polynomial_fits.pop(idx,None)
            if kind.startswith("Area"):self.area_series.discard(idx);self.area_values.pop(idx,None)
        if self.loaded_xy_series:self._redraw_line_preview()

    def show_analysis_results(self):
        win=tk.Toplevel(self);win.title(self.tr("Analiz Sonuçları","Analysis Results"));win.geometry("780x430")
        tree=ttk.Treeview(win,columns=("analysis","series","details"),show="headings")
        for c,t,w in (("analysis",self.tr("Analiz","Analysis"),160),("series",self.tr("Seri","Series"),130),("details",self.tr("Ayrıntılar","Details"),450)):tree.heading(c,text=t);tree.column(c,width=w)
        for item in self.analysis_history:tree.insert("","end",values=(item["analysis"],item["series"],json.dumps(item["values"],ensure_ascii=False)))
        tree.pack(fill="both",expand=True,padx=8,pady=8)
        ttk.Button(win,text=self.tr("CSV Kaydet…","Save CSV…"),command=self.save_analysis_results).pack(pady=(0,8))

    def save_analysis_results(self):
        path=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")])
        if not path:return
        with open(path,"w",encoding="utf-8-sig",newline="") as f:
            writer=csv.writer(f);writer.writerow(["analysis","series","details"])
            for item in self.analysis_history:writer.writerow([item["analysis"],item["series"],json.dumps(item["values"],ensure_ascii=False)])

    def _calculate_fit(self, idx):
        s=self.loaded_xy_series[idx]
        if len(s.x)<3: raise ValueError(self.tr("Uyum için en az 3 nokta gerekir.","At least 3 points are required for fitting."))
        slope,intercept=np.polyfit(s.x,s.y,1); predicted=slope*s.x+intercept; residual=s.y-predicted
        ss_res=float(np.sum(residual**2)); ss_tot=float(np.sum((s.y-np.mean(s.y))**2)); r2=1-ss_res/ss_tot if ss_tot else 1.0
        stderr=math.sqrt(ss_res/(len(s.x)-2)/float(np.sum((s.x-np.mean(s.x))**2))) if len(s.x)>2 and np.sum((s.x-np.mean(s.x))**2)>0 else 0.0
        class Fit:pass
        fit=Fit(); fit.slope=float(slope); fit.intercept=float(intercept); fit.stderr=stderr; fit.rvalue=math.sqrt(max(0,r2))
        f_value=(r2/(1-r2))*(len(s.x)-2) if r2<1 else float("inf")
        f_p=self._f_survival(f_value,1,len(s.x)-2)
        return s,fit,r2,f_value,f_p

    def fit_report(self):
        idx=self._selected_series_or_warn()
        if idx is None:return
        try:
            s,fit,r2,f_value,f_p=self._calculate_fit(idx)
            report=self.tr(
                f"Seri: {s.name}\nN: {len(s.x)}\nEğim: {fit.slope:.8g}\nY kesişimi: {fit.intercept:.8g}\nR²: {r2:.8g}\nF(1, {len(s.x)-2}): {f_value:.8g}\np değeri: {f_p:.8g}\nEğim standart hatası: {fit.stderr:.8g}",
                f"Series: {s.name}\nN: {len(s.x)}\nSlope: {fit.slope:.8g}\nIntercept: {fit.intercept:.8g}\nR²: {r2:.8g}\nF(1, {len(s.x)-2}): {f_value:.8g}\np-value: {f_p:.8g}\nSlope standard error: {fit.stderr:.8g}")
            self.record_analysis("Fit Report",s.name,{"slope":fit.slope,"intercept":fit.intercept,"R2":r2,"F":f_value,"p":f_p,"significant":f_p<.05})
            messagebox.showinfo(self.tr("Doğrusal Uyum Raporu","Linear Fit Report"),report)
        except Exception as exc:messagebox.showerror(self.tr("Uyum Hatası","Fit Error"),str(exc))

    def f_test(self):
        idx=self._selected_series_or_warn()
        if idx is None:return
        try:
            s,_fit,_r2,f_value,f_p=self._calculate_fit(idx)
            confidence=(1-f_p)*100;status="PASS" if confidence>=90 else "FAIL"
            self.record_analysis("F Test",s.name,{"F":f_value,"p":f_p,"confidence_percent":confidence,"status":status})
            messagebox.showinfo(self.tr("Regresyon F Testi","Regression F Test"),self.tr(
                f"{s.name}\nF(1, {len(s.x)-2}) = {f_value:.8g}\np değeri = {f_p:.8g}\nGüven = %{confidence:.2f}\nSonuç: {status}\nModel {'anlamlıdır' if f_p<0.05 else 'anlamlı değildir'} (α=0,05).",
                f"{s.name}\nF(1, {len(s.x)-2}) = {f_value:.8g}\np-value = {f_p:.8g}\nConfidence = {confidence:.2f}%\nResult: {status}\nThe model is {'significant' if f_p<0.05 else 'not significant'} (α=0.05)."))
        except Exception as exc:messagebox.showerror(self.tr("F Testi Hatası","F Test Error"),str(exc))

    def area_under_curve(self):
        idx = self._selected_series_or_warn()
        if idx is None: return
        self.area_series.add(idx)
        s = self.loaded_xy_series[idx]
        order = np.argsort(s.x)
        area = self._trapezoid(s.y[order], s.x[order])
        self.area_values[idx]=area
        self.record_analysis("Area Under Curve",s.name,{"area":area,"method":"trapezoidal"})
        self._redraw_line_preview()
        messagebox.showinfo(self.tr("Eğri Altındaki Alan", "Area Under Curve"), self.tr(
            f"Seri: {s.name}\nİntegrasyon yöntemi: Trapez\nEğri altındaki alan = {area:.8g}",
            f"Series: {s.name}\nIntegration method: Trapezoidal\nArea under curve = {area:.8g}"))

    def descriptive_statistics(self):
        ci = self.sheet_active_column
        if ci >= len(self.sheet_headers):
            return
        vals = np.asarray([v for row in self.sheet_rows if (v := _to_float(row[ci] if ci < len(row) else None)) is not None])
        if not len(vals):
            messagebox.showinfo(self.tr("İstatistik", "Statistics"), self.tr("Seçili sütunda sayısal veri yok.", "The selected column has no numeric data."))
            return
        text = self.tr(
            f"Sütun: {self.sheet_headers[ci]}\nN: {len(vals)}\nOrtalama: {np.mean(vals):.8g}\nMedyan: {np.median(vals):.8g}\nStandart Sapma: {np.std(vals, ddof=1) if len(vals)>1 else 0:.8g}\nVaryans: {np.var(vals, ddof=1) if len(vals)>1 else 0:.8g}",
            f"Column: {self.sheet_headers[ci]}\nN: {len(vals)}\nMean: {np.mean(vals):.8g}\nMedian: {np.median(vals):.8g}\nStandard Deviation: {np.std(vals, ddof=1) if len(vals)>1 else 0:.8g}\nVariance: {np.var(vals, ddof=1) if len(vals)>1 else 0:.8g}")
        win = tk.Toplevel(self); win.title(self.tr("Tanımlayıcı İstatistikler", "Descriptive Statistics"))
        ttk.Label(win, text=text, padding=18, justify="left").pack()

    @staticmethod
    def _json_series(s):
        d = asdict(s); d["x"] = s.x.tolist(); d["y"] = s.y.tolist()
        d["y_axis_side"] = getattr(s, "y_axis_side", "left")
        return d

    def save_project(self):
        path = filedialog.asksaveasfilename(defaultextension=".myopj", filetypes=[(self.tr("Bilimsel Grafik Projesi", "Scientific Graph Project"), "*.myopj")])
        if not path: return
        data = {"format": "myopj", "version": 2, "headers": self.sheet_headers,
                "rows": self.sheet_rows, "series": [self._json_series(s) for s in self.loaded_xy_series],
                "settings": asdict(SETTINGS), "fit_series": sorted(self.fit_series),
                "area_series": sorted(self.area_series), "axis_break": self.axis_break,
                "plot": {"title": self.line_title.get(), "xlabel": self.line_xlabel.get(),
                         "ylabel": self.line_ylabel.get(), "legend": self.line_legend.get(),
                         "log_x": self.line_logx.get(), "log_y": self.line_logy.get(),
                         "top_axis": self.top_axis_var.get(), "top_ticks": self.top_tick_var.get(),
                         "right_axis": self.right_axis_var.get(), "right_ticks": self.right_tick_var.get(),
                         "legend_location": self.legend_location()}}
        Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.add_recent_project(path)
        messagebox.showinfo(self.tr("Proje", "Project"), self.tr(f"Proje kaydedildi:\n{path}", f"Project saved:\n{path}"))

    def open_project(self):
        path = filedialog.askopenfilename(filetypes=[(self.tr("Bilimsel Grafik Projesi", "Scientific Graph Project"), "*.myopj")])
        if not path: return
        self.open_project_path(path)

    def open_project_path(self,path):
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
            if data.get("format") != "myopj": raise ValueError(self.tr("Geçersiz proje dosyası.", "Invalid project file."))
            self.sheet_headers, self.sheet_rows = data["headers"], data["rows"]
            for key, value in data.get("settings", {}).items():
                if hasattr(SETTINGS, key): setattr(SETTINGS, key, value)
            self.sync_series_from_sheet(preserve=False)
            for s, saved in zip(self.loaded_xy_series, data.get("series", [])):
                for key, value in saved.items():
                    if key not in ("x", "y") and hasattr(s, key): setattr(s, key, value)
                    elif key == "y_axis_side": s.y_axis_side = value
            plot = data.get("plot", {})
            for entry, key in ((self.line_title,"title"),(self.line_xlabel,"xlabel"),(self.line_ylabel,"ylabel")):
                entry.delete(0,"end"); entry.insert(0, plot.get(key,""))
            for var, key in ((self.line_legend,"legend"),(self.line_logx,"log_x"),(self.line_logy,"log_y"),(self.top_axis_var,"top_axis"),(self.top_tick_var,"top_ticks"),(self.right_axis_var,"right_axis"),(self.right_tick_var,"right_ticks")):
                var.set(plot.get(key, False))
            self.legend_loc_var.set(plot.get("legend_location","best"))
            self.fit_series, self.area_series = set(data.get("fit_series",[])), set(data.get("area_series",[]))
            self.axis_break = data.get("axis_break")
            self.excel_path = path
            self.add_recent_project(path)
            self.refresh_sheet(); self._refresh_series_tree(); self._redraw_line_preview()
        except Exception as exc:
            messagebox.showerror(self.tr("Proje açılamadı", "Could not open project"), str(exc))

    def save_template(self):
        path = filedialog.asksaveasfilename(defaultextension=".otp", filetypes=[("Origin Template", "*.otp")])
        if path:
            Path(path).write_text(json.dumps(asdict(SETTINGS), ensure_ascii=False, indent=2), encoding="utf-8")

    def load_template(self):
        path = filedialog.askopenfilename(filetypes=[("Origin Template", "*.otp")])
        if not path: return
        try:
            for key, value in json.loads(Path(path).read_text(encoding="utf-8")).items():
                if hasattr(SETTINGS, key): setattr(SETTINGS, key, value)
            if self.loaded_xy_series: self._redraw_line_preview()
        except Exception as exc: messagebox.showerror(self.tr("Şablon", "Template"), str(exc))

    def draw_selected_plot(self):
        try:
            if not self.loaded_xy_series: raise ValueError(self.tr("Önce veri yükleyin.", "Load data first."))
            kind=self.plot_type.get()
            if kind in ("bar","pie","3d_column","3d_bar"): self._draw_bar_from_sheet()
            else: self._redraw_line_preview()
        except Exception as exc: messagebox.showerror(self.tr("Grafik oluşturulamadı", "Could not create graph"), str(exc))

    def _render(self, fig):
        self.annotation_artists=[]; self.annotation_mode=None; self.annotation_clicks=[]
        super()._render(fig)

    def export_current(self):
        if self.current_figure is None:messagebox.showinfo(self.tr("Dışa Aktar","Export"),self.tr("Önce grafik oluşturun.","Create a graph first."));return
        path=filedialog.asksaveasfilename(defaultextension=".png",filetypes=[("PNG","*.png"),("JPEG","*.jpg *.jpeg"),("TIFF","*.tif *.tiff"),("PDF","*.pdf"),("SVG","*.svg")])
        if not path:return
        suffix=Path(path).suffix.lower();kwargs={"bbox_inches":"tight","facecolor":self.plot_background.get()}
        if suffix in (".jpg",".jpeg"):kwargs.update(dpi=min(300,SETTINGS.dpi),pil_kwargs={"quality":88,"optimize":True,"progressive":True})
        elif suffix in (".tif",".tiff"):kwargs.update(dpi=min(300,SETTINGS.dpi),pil_kwargs={"compression":"tiff_lzw"})
        elif suffix==".png":kwargs["dpi"]=min(300,SETTINGS.dpi)
        try:self.current_figure.savefig(path,**kwargs);messagebox.showinfo(self.tr("Dışa Aktar","Export"),self.tr("Dosya kaydedildi.","File saved."))
        except Exception as exc:messagebox.showerror(self.tr("Kayıt Hatası","Save Error"),str(exc))

    def print_current(self):
        if self.current_figure is None:messagebox.showinfo(self.tr("Yazdır","Print"),self.tr("Önce grafik oluşturun.","Create a graph first."));return
        try:
            path=Path(tempfile.gettempdir())/"scientific_graph_studio_print.pdf"
            self.current_figure.savefig(path,bbox_inches="tight",facecolor=self.plot_background.get())
            if os.name=="nt":os.startfile(str(path),"print")
            else:subprocess.run(["lp",str(path)],check=True,capture_output=True)
            messagebox.showinfo(self.tr("Yazdır","Print"),self.tr("Grafik yazıcı kuyruğuna gönderildi.","Graph sent to the print queue."))
        except FileNotFoundError:messagebox.showerror(self.tr("Yazdırma Hatası","Print Error"),self.tr("Sistem yazdırma komutu bulunamadı.","System print command was not found."))
        except Exception as exc:messagebox.showerror(self.tr("Yazdırma Hatası","Print Error"),str(exc))

    def _draw_bar_from_sheet(self):
        import matplotlib.pyplot as plt
        kind=self.plot_type.get()
        if kind in ("3d_column","3d_bar"):
            fig=plt.figure(figsize=(SETTINGS.figure_width,SETTINGS.figure_height),dpi=110,layout="constrained");ax=fig.add_subplot(111,projection="3d")
        else: fig, ax = plt.subplots(figsize=(SETTINGS.figure_width, SETTINGS.figure_height), dpi=110, layout="constrained")
        if kind not in ("3d_column","3d_bar","pie"):apply_scientific_style(ax, SETTINGS.font_size, SETTINGS.axis_linewidth, False)
        ax.set_facecolor(self.plot_background.get()); fig.set_facecolor(self.plot_background.get())
        if kind not in ("3d_column","3d_bar","pie"):
            ax.set_axisbelow(True)
            ax.spines["top"].set_visible(self.top_axis_var.get()); ax.spines["top"].set_linewidth(SETTINGS.axis_linewidth)
            ax.spines["right"].set_visible(self.right_axis_var.get()); ax.spines["right"].set_linewidth(SETTINGS.axis_linewidth)
            ax.tick_params(top=self.top_axis_var.get() and self.top_tick_var.get(), right=self.right_axis_var.get() and self.right_tick_var.get())
            ax.grid(False,which="both")
        palette = PALETTES[SETTINGS.palette_name]
        series = [s for s in self.loaded_xy_series if s.visible]
        width = SETTINGS.bar_width / max(1, len(series)); base = np.arange(max(len(s.y) for s in series))
        if kind=="pie":
            s=series[0]; labels=[str(r[0]) if r else str(i+1) for i,r in enumerate(self.sheet_rows[:len(s.y)])]
            ax.pie(np.abs(s.y),labels=labels,autopct="%1.1f%%",colors=palette[:len(s.y)])
        for j,s in enumerate([] if kind=="pie" else series):
            pos = base[:len(s.y)] + (j-(len(series)-1)/2)*width
            if kind=="3d_column": ax.bar3d(pos-width*.45,np.full(len(pos),j)-.325,np.zeros(len(pos)),width*.9,.65,s.y,color=s.color or palette[j%len(palette)],label=s.name,shade=True)
            elif kind=="3d_bar": ax.bar3d(np.full(len(pos),j)-.325,pos-width*.45,np.zeros(len(pos)),.65,width*.9,s.y,color=s.color or palette[j%len(palette)],label=s.name,shade=True)
            else: ax.bar(pos, s.y, width=width*.9, color=s.color or palette[j%len(palette)], label=s.name)
        labels = [str(v) for v in self.sheet_rows[:len(base)] for v in ([v[0]] if v else [""])]
        if kind not in ("pie","3d_bar"):
            ax.set_xticks(base); ax.set_xticklabels(labels, rotation=0)
        style_labels(ax,self.line_title.get(),self.line_xlabel.get(),self.line_ylabel.get(),SETTINGS.font_size)
        if self.line_legend.get(): add_legend(ax,True,SETTINGS.font_size,self.legend_location(),int(self.legend_cols.get()))
        self.apply_plot_typography(fig)
        self._render(fig)

    def open_plot_settings_dialog(self):
        win=tk.Toplevel(self); win.title(self.tr("Grafik Ayarları","Graph Settings")); win.resizable(False,False)
        box=ttk.Frame(win,padding=14); box.pack(fill="both",expand=True)
        fields=(("Başlık","Title",self.line_title),("X ekseni","X axis",self.line_xlabel),("Y ekseni","Y axis",self.line_ylabel),
                ("X min","X min",self.line_xmin),("X max","X max",self.line_xmax),("X adım","X step",self.line_xstep),
                ("Y min","Y min",self.line_ymin),("Y max","Y max",self.line_ymax),("Y adım","Y step",self.line_ystep))
        for r,(tr,en,source) in enumerate(fields):
            ttk.Label(box,text=self.tr(tr,en)).grid(row=r,column=0,sticky="w",pady=3)
            e=ttk.Entry(box,width=24); e.insert(0,source.get()); e.grid(row=r,column=1,padx=(10,0),pady=3); e.source=source
        opts=ttk.LabelFrame(box,text=self.tr("Görünüm","Appearance"),padding=7); opts.grid(row=len(fields),column=0,columnspan=2,sticky="ew",pady=8)
        for i,(tr,en,var) in enumerate((("Açıklama kutusu","Legend",self.line_legend),("X logaritmik","Log X",self.line_logx),("Y logaritmik","Log Y",self.line_logy),("Üst eksen","Top axis",self.top_axis_var),("Üst eksen tikleri","Top-axis ticks",self.top_tick_var),("Sağ eksen","Right axis",self.right_axis_var),("Sağ eksen tikleri","Right-axis ticks",self.right_tick_var))):
            ttk.Checkbutton(opts,text=self.tr(tr,en),variable=var).grid(row=i//2,column=i%2,sticky="w",padx=6)
        ttk.Label(opts,text=self.tr("Açıklama konumu:","Legend location:")).grid(row=4,column=0,sticky="w",padx=6,pady=(5,0))
        locations=tuple(TR_LEGEND) if self.language=="tr" else tuple(TR_LEGEND.values())
        ttk.Combobox(opts,textvariable=self.legend_loc_var,values=locations,state="readonly",width=16).grid(row=4,column=1,sticky="w",pady=(5,0))
        fontbox=ttk.LabelFrame(box,text=self.tr("Yazı","Typography"),padding=7); fontbox.grid(row=len(fields)+1,column=0,columnspan=2,sticky="ew",pady=(0,8))
        ttk.Label(fontbox,text=self.tr("Yazı tipi","Font family")).grid(row=0,column=0,sticky="w")
        families=sorted(set(tkfont.families(self)))
        ttk.Combobox(fontbox,textvariable=self.plot_font_family,values=families,state="readonly",width=22).grid(row=0,column=1,padx=5)
        ttk.Label(fontbox,text=self.tr("Boyut","Size")).grid(row=1,column=0,sticky="w",pady=4)
        ttk.Spinbox(fontbox,textvariable=self.plot_font_size,from_=6,to=40,width=7).grid(row=1,column=1,sticky="w",padx=5)
        ttk.Checkbutton(fontbox,text=self.tr("Eğik yazı","Italic"),variable=self.plot_font_italic).grid(row=2,column=0,sticky="w")
        ttk.Button(fontbox,text=self.tr("Özel işaret ekle…","Insert special character…"),command=lambda:self.open_symbol_picker(win),style="Compact.TButton").grid(row=2,column=1,sticky="ew",padx=5)
        ttk.Label(fontbox,text=self.tr("Arka plan","Background")).grid(row=3,column=0,sticky="w")
        bg_entry=ttk.Entry(fontbox,textvariable=self.plot_background,width=10);bg_entry.grid(row=3,column=1,sticky="w",padx=5)
        ttk.Button(fontbox,text="…",width=3,command=lambda:self.choose_background_color(bg_entry),style="Compact.TButton").grid(row=3,column=1,sticky="e")
        def apply():
            for child in box.winfo_children():
                if isinstance(child,ttk.Entry) and hasattr(child,"source"):
                    child.source.delete(0,"end"); child.source.insert(0,child.get())
            win.destroy()
            if self.loaded_xy_series: self.draw_selected_plot()
        ttk.Button(box,text=self.tr("Uygula","Apply"),command=apply,style="Accent.TButton").grid(row=len(fields)+2,column=0,columnspan=2,sticky="ew")

    def open_symbol_picker(self,parent):
        target=simpledialog.askstring(self.tr("Hedef","Target"),self.tr("Hedef: başlık, x veya y","Target: title, x or y"),parent=parent)
        if not target:return
        normalized=target.strip().lower(); mapping={"başlık":self.line_title,"title":self.line_title,"x":self.line_xlabel,"y":self.line_ylabel}
        entry=mapping.get(normalized)
        if entry is None:messagebox.showerror(self.tr("Geçersiz hedef","Invalid target"),self.tr("Başlık, x veya y yazın.","Enter title, x or y."),parent=parent);return
        win=tk.Toplevel(parent); win.title(self.tr("Özel İşaretler","Special Characters"))
        symbols=("µ","μ","°","±","×","÷","Å","Ω","α","β","γ","δ","λ","π","σ","²","³","⁻¹","∞","≤","≥","≈","Δ","∑","√")
        def insert(symbol):entry.insert("insert",symbol);win.destroy()
        for i,symbol in enumerate(symbols):ttk.Button(win,text=symbol,width=4,command=lambda s=symbol:insert(s)).grid(row=i//7,column=i%7,padx=2,pady=2)

    def choose_background_color(self,_entry=None):
        result=colorchooser.askcolor(self.plot_background.get(),parent=self)
        if result and result[1]:self.plot_background.set(result[1])

    def apply_plot_typography(self,fig):
        family=self.plot_font_family.get() or "Arial"; size=max(6,int(self.plot_font_size.get())); style="italic" if self.plot_font_italic.get() else "normal"
        for ax in fig.axes:
            for obj in [ax.title,ax.xaxis.label,ax.yaxis.label,*ax.get_xticklabels(),*ax.get_yticklabels()]:
                obj.set_fontfamily(family); obj.set_fontsize(size); obj.set_fontstyle(style)
            legend=ax.get_legend()
            if legend:
                for text in legend.get_texts():text.set_fontfamily(family);text.set_fontsize(size);text.set_fontstyle(style)

    def toggle_dark_mode(self):
        self.dark_mode=not self.dark_mode; self.apply_theme()

    def set_language(self, language):
        self.language=language; self._build_menu()
        if hasattr(self,"sheet_window") and self.sheet_window.winfo_exists(): self.sheet_window.destroy()
        self.load_btn.configure(text=self.tr("Veri Yükle","Load Data")); self.data_btn.configure(text=self.tr("Veri Tablosu","Data Table"))
        self.graph_settings_btn.configure(text=self.tr("Grafik Ayarları","Graph Settings")); self.draw_btn.configure(text=self.tr("Grafik Oluştur","Create Graph"))
        self.preview_label_var.set(self.tr("Önizleme","Preview")); self.preview_full_btn.configure(text=self.tr("⛶ Önizlemeyi Büyüt","⛶ Expand Preview"))
        self.excel_file_label.configure(text=self.tr("Veri yüklenmedi","No data loaded") if not self.excel_path else Path(self.excel_path).name)
        self.plot_type_label.configure(text=self.tr("Grafik türü:","Graph type:")); self.line_radio.configure(text=self.tr("Çizgi","Line")); self.bar_radio.configure(text=self.tr("Sütun","Bar"))
        self.series_box.configure(text=self.tr("Seriler","Series")); self.apply_series_btn.configure(text=self.tr("Uygula","Apply")); self.reset_series_btn.configure(text=self.tr("Sıfırla","Reset"))
        self.series_tree.heading("name",text=self.tr("Seri","Series")); self.series_tree.heading("axis",text=self.tr("Y Ekseni","Y Axis"))
        self.text_tool_btn.configure(text=self.tr("Metin","Text")); self.arrow_tool_btn.configure(text=self.tr("Ok","Arrow")); self.delete_annotation_btn.configure(text=self.tr("Sil","Delete"))
        if language=="tr":
            self.series_mode_combo.configure(values=tuple(TR_MODES)); self.series_marker_combo.configure(values=tuple(TR_MARKERS)); self.series_fill_combo.configure(values=tuple(TR_FILLS)); self.series_line_style_combo.configure(values=tuple(TR_LINES)); self.series_color_combo.configure(values=tuple(TR_COLORS))
            self.series_axis_side_var.set("Sol")
            self.legend_loc_var.set(next((k for k,v in TR_LEGEND.items() if v==self.legend_loc_var.get()),self.legend_loc_var.get()))
        else:
            self.series_mode_combo.configure(values=("Line + Symbol","Line Only","Symbol Only")); self.series_marker_combo.configure(values=tuple(MARKER_OPTIONS)); self.series_fill_combo.configure(values=("Open","Filled")); self.series_line_style_combo.configure(values=tuple(LINESTYLE_OPTIONS)); self.series_color_combo.configure(values=tuple(TR_COLORS.values()))
            self.legend_loc_var.set(TR_LEGEND.get(self.legend_loc_var.get(),self.legend_loc_var.get()))
        self.series_axis_side_combo.configure(values=("Sol","Sağ") if language=="tr" else ("left","right"))
        idx=self._selected_series_index()
        if idx is not None:self._load_series_into_editor(idx)
        elif language=="tr":
            self.series_mode_var.set("Çizgi + Sembol"); self.series_marker_name_var.set("Daire (o)"); self.series_fill_var.set("Açık"); self.series_line_style_name_var.set("Düz"); self.series_color_var.set("Otomatik / Palet")
        else:
            self.series_mode_var.set("Line + Symbol"); self.series_marker_name_var.set("Circle (o)"); self.series_fill_var.set("Open"); self.series_line_style_name_var.set("Solid"); self.series_color_var.set("Auto / Palette")

    def show_about(self):
        messagebox.showinfo(self.tr("Hakkında","About"), self.tr(
            "Scientific Graph Studio (Beta)\n\nBilimsel veri dosyalarını düzenlemek, görselleştirmek ve analiz etmek için geliştirilmiştir.\nÇizgi, sütun, çoklu Y ekseni, regresyon, eğri altındaki alan, istatistik ve proje desteği içerir.\n\nAhmetcan tarafından geliştirilmiştir.",
            "Scientific Graph Studio (Beta)\n\nCreated to edit, visualize and analyze scientific CSV data.\nIncludes Line, Bar, multiple Y axes, regression, AUC, statistics and project support.\n\nby Ahmetcan"))

    def show_help(self):
        win=tk.Toplevel(self); win.title(self.tr("Yardım ve Komutlar","Help and Commands")); win.geometry("680x520")
        text=tk.Text(win,wrap="word",padx=16,pady=14); scroll=ttk.Scrollbar(win,command=text.yview); text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right",fill="y"); text.pack(fill="both",expand=True)
        content=self.tr(
"""SCIENTIFIC GRAPH STUDIO — HIZLI YARDIM

VERİ
• Veri Yükle: CSV/TXT açar; sütunlar X1–Y1 çiftleri oluşturur.
• Veri Tablosu: Çift tıkla düzenle; sürükleyerek hücre aralığı seç.
• Ctrl/Cmd+V: Excel bloğu yapıştır. Delete/Backspace: seçili hücreleri temizle.
• Başlığa çift tıkla: sütun adını değiştir.

GRAFİK
• Çizgi/Sütun: grafik türünü değiştirir ve seri panelini uyarlar.
• Grafik Ayarları: başlık, eksen, sınır, logaritmik ölçek, eksen tikleri ve açıklama kutusu.
• Metin/Ok: aracı seç, grafik üzerine tıkla. Sil: son eklenen öğeyi kaldırır.

ANALİZ
• Regresyon ve uyum raporu: eğim, kesişim, R², F ve p değeri.
• F testi, tanımlayıcı istatistik ve eğri altındaki alan.

DOSYA VE AYARLAR
• .myopj proje; .otp görünüm şablonudur.
• Ayarlar: dil, açık/koyu görünüm ve grafik seçenekleri.
• Önizlemeyi Büyüt: tam ekran; Esc ile kapatılır.""",
"""SCIENTIFIC GRAPH STUDIO — QUICK HELP

DATA
• Load Data: opens CSV/TXT; columns create X1–Y1 pairs.
• Data Table: double-click to edit; drag to select a cell range.
• Ctrl/Cmd+V: paste an Excel block. Delete/Backspace: clear selected cells.
• Double-click a heading to rename a column.

GRAPH
• Line/Bar changes graph type and adapts the series panel.
• Graph Settings controls titles, axes, limits, logarithmic scale, axis ticks and legend.
• Text/Arrow: choose a tool and click the graph. Delete removes the latest item.

ANALYSIS
• Regression/fit report: slope, intercept, R², F and p-value.
• F test, descriptive statistics and area under curve.

FILE AND SETTINGS
• .myopj is a project; .otp is an appearance template.
• Settings controls language, light/dark appearance and graph options.
• Expand Preview opens full screen; Esc closes it.""")
        text.insert("1.0",content); text.configure(state="disabled")

    def add_text_annotation(self):
        if self.current_figure is None: return
        value=simpledialog.askstring(self.tr("Metin Ekle","Add Text"),self.tr("Metin:","Text:"),parent=self)
        if not value:return
        self.annotation_mode="text"; self.annotation_text=value; self.annotation_clicks=[]
        self.preview_canvas.get_tk_widget().configure(cursor="crosshair")
        self._annotation_cid=self.preview_canvas.mpl_connect("button_press_event",self._on_annotation_click)

    def add_arrow_annotation(self):
        if self.current_figure is None:return
        self.annotation_mode="arrow"; self.annotation_clicks=[]
        self.preview_canvas.get_tk_widget().configure(cursor="crosshair")
        self._annotation_cid=self.preview_canvas.mpl_connect("button_press_event",self._on_annotation_click)
        messagebox.showinfo(self.tr("Ok Ekle","Add Arrow"),self.tr("Başlangıç ve bitiş noktalarına tıklayın.","Click the start and end points."))

    def add_shape_annotation(self,shape):
        if self.current_figure is None:return
        self.annotation_mode="shape:"+shape;self.annotation_clicks=[];self.preview_canvas.get_tk_widget().configure(cursor="crosshair")
        self._annotation_cid=self.preview_canvas.mpl_connect("button_press_event",self._on_annotation_click)

    def _on_annotation_click(self,event):
        if event.inaxes is None:return
        if self.annotation_mode=="text":
            artist=event.inaxes.text(event.xdata,event.ydata,self.annotation_text)
            self.annotation_artists.append(artist); self._finish_annotation_mode()
        elif self.annotation_mode=="arrow":
            self.annotation_clicks.append((event.xdata,event.ydata,event.inaxes))
            if len(self.annotation_clicks)==2:
                (x1,y1,ax),(x2,y2,_)=self.annotation_clicks
                artist=ax.annotate("",xy=(x2,y2),xytext=(x1,y1),arrowprops=dict(arrowstyle="->",lw=1.6,color="crimson"))
                self.annotation_artists.append(artist); self._finish_annotation_mode()
        elif self.annotation_mode and self.annotation_mode.startswith("shape:"):
            self.annotation_clicks.append((event.xdata,event.ydata,event.inaxes))
            if len(self.annotation_clicks)==2:
                from matplotlib.patches import Circle,Ellipse,Rectangle
                (x1,y1,ax),(x2,y2,_)=self.annotation_clicks;shape=self.annotation_mode.split(":",1)[1]
                if shape=="circle":artist=Circle((x1,y1),math.hypot(x2-x1,y2-y1),fill=False,color="crimson",lw=1.6)
                elif shape=="ellipse":artist=Ellipse(((x1+x2)/2,(y1+y2)/2),abs(x2-x1),abs(y2-y1),fill=False,color="crimson",lw=1.6)
                else:artist=Rectangle((min(x1,x2),min(y1,y2)),abs(x2-x1),abs(y2-y1),fill=False,color="crimson",lw=1.6)
                ax.add_patch(artist);self.annotation_artists.append(artist);self._finish_annotation_mode()
        self.preview_canvas.draw_idle()

    def _finish_annotation_mode(self):
        if hasattr(self,"_annotation_cid"): self.preview_canvas.mpl_disconnect(self._annotation_cid)
        self.annotation_mode=None; self.annotation_clicks=[]
        self.preview_canvas.get_tk_widget().configure(cursor="")

    def delete_last_annotation(self):
        if self.annotation_mode: self._finish_annotation_mode()
        if self.annotation_artists:
            artist=self.annotation_artists.pop()
            try: artist.remove()
            except ValueError: pass
            self.preview_canvas.draw_idle()

    def open_fullscreen_preview(self):
        if self.current_figure is None: messagebox.showinfo(self.tr("Önizleme","Preview"),self.tr("Önce grafik oluşturun.","Create a graph first.")); return
        win=tk.Toplevel(self); win.title("Scientific Graph Studio — "+self.tr("Önizleme","Preview")); win.geometry("1100x760");win.minsize(700,500)
        preview_figure=copy.deepcopy(self.current_figure)
        canvas=FigureCanvasTkAgg(preview_figure,master=win); canvas.draw(); canvas.get_tk_widget().pack(fill="both",expand=True)
        win.bind("<Escape>",lambda _e:win.destroy()); win.bind("<F11>",lambda _e:win.attributes("-fullscreen",not win.attributes("-fullscreen")))

    def _redraw_line_preview(self):
        if not self.loaded_xy_series: return
        fig, left = matplotlib.pyplot.subplots(figsize=(SETTINGS.figure_width, SETTINGS.figure_height), dpi=110, layout="constrained")
        apply_scientific_style(left, SETTINGS.font_size, SETTINGS.axis_linewidth, False)
        left.set_facecolor(self.plot_background.get()); fig.set_facecolor(self.plot_background.get())
        right = None
        if any(getattr(s, "y_axis_side", "left") == "right" and s.visible for s in self.loaded_xy_series):
            right = left.twinx(); apply_scientific_style(right, SETTINGS.font_size, SETTINGS.axis_linewidth, False)
            right.spines["left"].set_visible(False)
        logx, logy = self.line_logx.get(), self.line_logy.get()
        left.set_xscale("log" if logx else "linear"); left.set_yscale("log" if logy else "linear")
        if right: right.set_xscale("log" if logx else "linear"); right.set_yscale("log" if logy else "linear")
        if self.axis_break:
            axis, start, end = self.axis_break["axis"], self.axis_break["start"], self.axis_break["end"]
            gap = end - start
            visible_span = max(gap * .08, np.finfo(float).eps)
            def forward(values):
                values = np.asarray(values)
                return np.where(values <= start, values,
                    np.where(values >= end, values-gap+visible_span,
                             start+(values-start)*visible_span/gap))
            def inverse(values):
                values = np.asarray(values); edge = start + visible_span
                return np.where(values <= start, values,
                    np.where(values >= edge, values+gap-visible_span,
                             start+(values-start)*gap/visible_span))
            if axis == "x" and not logx:
                left.set_xscale("function", functions=(forward, inverse))
            elif axis == "y" and not logy:
                left.set_yscale("function", functions=(forward, inverse))
                if right: right.set_yscale("function", functions=(forward, inverse))
        if logx:
            left.xaxis.set_minor_locator(LogLocator(base=10, subs=np.arange(2,10)*.1)); left.xaxis.set_minor_formatter(NullFormatter())
        if logy:
            left.yaxis.set_minor_locator(LogLocator(base=10, subs=np.arange(2,10)*.1)); left.yaxis.set_minor_formatter(NullFormatter())
        left.set_axisbelow(True)
        left.grid(False,which="both")
        left.spines["top"].set_visible(self.top_axis_var.get()); left.spines["top"].set_linewidth(SETTINGS.axis_linewidth)
        left.spines["right"].set_visible(self.right_axis_var.get())
        left.spines["right"].set_linewidth(SETTINGS.axis_linewidth)
        left.tick_params(top=self.top_axis_var.get() and self.top_tick_var.get(), right=self.right_axis_var.get() and self.right_tick_var.get())
        if right:
            right.spines["top"].set_visible(self.top_axis_var.get()); right.spines["right"].set_visible(self.right_axis_var.get())
            right.tick_params(top=self.top_axis_var.get() and self.top_tick_var.get(),right=self.right_axis_var.get() and self.right_tick_var.get())
        palette = PALETTES[SETTINGS.palette_name]
        for j, s in enumerate(self.loaded_xy_series):
            if not s.visible: continue
            if logx and np.any(s.x <= 0) or logy and np.any(s.y <= 0):
                raise ValueError(f"{s.name}: log eksende sıfır/negatif değer var.")
            ax = right if getattr(s, "y_axis_side", "left") == "right" and right else left
            color = s.color or palette[j % len(palette)]
            marker = None if s.plot_mode == "Line Only" else s.marker
            linestyle = "none" if s.plot_mode == "Symbol Only" else (s.line_style if s.line_style != "none" else "-")
            face = color if s.marker_fill == "Filled" or marker in ("+","x","|","_",".",",",None) else "white"
            ax.plot(s.x, s.y, label=s.name, color=color, linestyle=linestyle, linewidth=s.line_width,
                    marker=marker, markersize=s.marker_size, markerfacecolor=face, markeredgecolor=color,
                    markeredgewidth=s.marker_edge_width, markevery=s.marker_every, alpha=s.alpha, zorder=3+j)
            if self.plot_type.get()=="area": ax.fill_between(s.x,s.y,0,color=color,alpha=.25)
            if j in self.fit_series:
                _series,fit,r2,_f,_p=self._calculate_fit(j); slope,intercept=fit.slope,fit.intercept
                xx = np.linspace(np.min(s.x), np.max(s.x), 200)
                ax.plot(xx, slope*xx+intercept, color="red", linewidth=1.7, label=self.tr(f"Uyum: eğim={slope:.4g}, R²={r2:.4g}",f"Fit: slope={slope:.4g}, R²={r2:.4g}"))
            if j in self.polynomial_fits:
                degree,coeff,r2=self.polynomial_fits[j];xx=np.linspace(np.min(s.x),np.max(s.x),300)
                ax.plot(xx,np.polyval(coeff,xx),color="purple",linewidth=1.8,linestyle="--",label=self.tr(f"{degree}. derece polinom, R²={r2:.4g}",f"Polynomial degree {degree}, R²={r2:.4g}"))
            if j in self.area_series:
                order=np.argsort(s.x); ax.fill_between(s.x[order], s.y[order], 0, color=color, alpha=.18, hatch="//")
                area=self.area_values.get(j,self._trapezoid(s.y[order],s.x[order]))
                self.area_values[j]=area
                ax.text(.02,.97,self.tr(f"Eğri altındaki alan ({s.name}) = {area:.6g}",f"Area under curve ({s.name}) = {area:.6g}"),transform=ax.transAxes,va="top",color=color,fontsize=max(8,SETTINGS.font_size-1),bbox=dict(facecolor="white",alpha=.8,edgecolor=color))
        style_labels(left, self.line_title.get(), self.line_xlabel.get(), self.line_ylabel.get(), SETTINGS.font_size)
        if right: right.set_ylabel(self.tr("Sağ Y", "Right Y"), fontsize=SETTINGS.font_size+1)
        apply_axis_limits(left, optional_float(self.line_ymin.get()), optional_float(self.line_ymax.get()), optional_float(self.line_ystep.get()), optional_float(self.line_xmin.get()), optional_float(self.line_xmax.get()), optional_float(self.line_xstep.get()))
        apply_comma_tick_formatter(left, not logx, not logy)
        if self.axis_break:
            ax = left; axis=self.axis_break["axis"]
            kwargs=dict(transform=ax.transAxes, color="black", clip_on=False, linewidth=1.2)
            if axis=="x":
                ax.plot((.485,.495),(-.012,.012),**kwargs); ax.plot((.505,.515),(-.012,.012),**kwargs)
            else:
                ax.plot((-.012,.012),(.485,.495),**kwargs); ax.plot((-.012,.012),(.505,.515),**kwargs)
        if self.line_legend.get():
            handles, labels = left.get_legend_handles_labels()
            if right:
                h2,l2=right.get_legend_handles_labels(); handles+=h2; labels+=l2
            left.legend(handles, labels, frameon=False, loc=self.legend_location(), ncol=int(self.legend_cols.get()),
                        prop={"family":SETTINGS.legend_font_family,"size":SETTINGS.legend_font_size})
        self.apply_plot_typography(fig)
        self._render(fig)


if __name__ == "__main__":
    matplotlib.rcParams.update({"font.family":"sans-serif", "axes.unicode_minus":False,
                                "savefig.transparent":False, "pdf.fonttype":42, "ps.fonttype":42})
    GrafikOlusturucuV2().mainloop()
