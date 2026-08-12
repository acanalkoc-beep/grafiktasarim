"""
Grafik Oluşturucu Pro v3.2.1 — Origin-style Scientific Plotting GUI (CSV)

Line plots:
    CSV layout: X1 | Y1 | X2 | Y2 | ... | X10 | Y10
    First row is treated as headers. Each X/Y pair may have a different length.
    Excel dosyanizi "CSV UTF-8 (Comma delimited)" olarak kaydedebilirsiniz.

Dependencies:
    matplotlib + numpy (CSV icin ekstra paket gerekmez)

Run:
    python grafik_olusturucu_origin_excel_v2.py
"""

from __future__ import annotations

import copy
import json
import os
import subprocess
import tempfile
import tkinter.font as tkfont
from dataclasses import asdict
from tkinter import simpledialog

import math
import csv
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk
from typing import Optional

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.ticker import AutoMinorLocator, LogLocator, MultipleLocator, NullFormatter, ScalarFormatter, FuncFormatter




# -----------------------------------------------------------------------------
# Palettes — 10 publication-oriented choices, each suitable for up to 10 curves
# -----------------------------------------------------------------------------

PALETTES = {
    "Modern Scientific": [
        "#3569B7", "#D8583A", "#2A9D78", "#7A5AA6", "#D49A2A",
        "#3F8FA3", "#B85C8A", "#6D747C", "#8B6F47", "#5C8A4C",
    ],
    "Origin Classic": [
        "#0000FF", "#FF0000", "#008000", "#FF00FF", "#00A6A6",
        "#000000", "#E69F00", "#7F3C8D", "#2E91E5", "#E15F99",
    ],
    "Okabe-Ito +": [
        "#0072B2", "#D55E00", "#009E73", "#CC79A7", "#E69F00",
        "#56B4E9", "#F0E442", "#000000", "#6A3D9A", "#8C564B",
    ],
    "Nature-like": [
        "#3B6FB6", "#C84E45", "#54A24B", "#9C6ADE", "#E2A03F",
        "#4C9DA6", "#D66BA0", "#7A7A7A", "#A46C3C", "#6B8E23",
    ],
    "Science Journal": [
        "#0C5DA5", "#FF2C00", "#00B945", "#845B97", "#FF9500",
        "#00A6D6", "#E94E77", "#474747", "#9E5A2F", "#5D8C3A",
    ],
    "Deep Muted": [
        "#4C78A8", "#E45756", "#72B7B2", "#F2CF5B", "#B279A2",
        "#FF9DA6", "#9D755D", "#BAB0AC", "#59A14F", "#EDC948",
    ],
    "Ocean & Earth": [
        "#245C8A", "#3B8EA5", "#52A675", "#88B04B", "#D9A441",
        "#C76D3B", "#9E4D64", "#725A7A", "#4E6E58", "#6A7F94",
    ],
    "Warm Contrast": [
        "#9E2A2B", "#E76F51", "#F4A261", "#E9C46A", "#2A9D8F",
        "#287271", "#264653", "#6D597A", "#B56576", "#355070",
    ],
    "High Contrast": [
        "#111111", "#D7263D", "#1B6CA8", "#2E8B57", "#F49D37",
        "#7B2CBF", "#00A6A6", "#C44536", "#5A189A", "#6C757D",
    ],
    "Monochrome": [
        "#111111", "#2A2A2A", "#444444", "#5E5E5E", "#787878",
        "#929292", "#ACACAC", "#C0C0C0", "#D0D0D0", "#E0E0E0",
    ],
}

MARKER_OPTIONS = {
    "Circle (o)": "o",
    "Square (s)": "s",
    "Triangle Up (^)": "^",
    "Triangle Down (v)": "v",
    "Triangle Left (<)": "<",
    "Triangle Right (>)": ">",
    "Diamond (D)": "D",
    "Thin Diamond (d)": "d",
    "Pentagon (p)": "p",
    "Hexagon 1 (h)": "h",
    "Hexagon 2 (H)": "H",
    "Plus (+)": "+",
    "Filled Plus (P)": "P",
    "Cross (x)": "x",
    "Filled Cross (X)": "X",
    "Star (*)": "*",
    "Point (.)": ".",
    "Pixel (,)": ",",
    "Vertical Line (|)": "|",
    "Horizontal Line (_)": "_",
    "None": None,
}

LINESTYLE_OPTIONS = {
    "Solid": "-",
    "Dashed": "--",
    "Dash-dot": "-.",
    "Dotted": ":",
    "None": "none",
}

MARKERS = [v for v in MARKER_OPTIONS.values() if v is not None]
LINESTYLES = ["-", "--", "-.", ":"]


@dataclass
class PlotSettings:
    dpi: int = 600
    font_size: int = 11
    axis_linewidth: float = 1.15
    line_width: float = 1.8
    marker_size: float = 6.0
    marker_edge_width: float = 1.15
    legend_font_family: str = "Arial"
    legend_font_size: int = 10
    bar_width: float = 0.68
    error_linewidth: float = 1.15
    capsize: float = 4.0
    figure_width: float = 7.2
    figure_height: float = 4.8
    palette_name: str = "Modern Scientific"


@dataclass
class XYSeries:
    name: str
    x: np.ndarray
    y: np.ndarray
    x_header: str = ""
    y_header: str = ""

    # Per-series visual controls (editable in GUI)
    plot_mode: str = "Line + Symbol"   # Line + Symbol | Line Only | Symbol Only
    marker: Optional[str] = "o"
    marker_fill: str = "Open"          # Open | Filled
    line_style: str = "-"
    line_width: float = 1.8
    marker_size: float = 6.0
    marker_edge_width: float = 1.15
    marker_every: int = 1
    alpha: float = 1.0
    visible: bool = True
    color: Optional[str] = None          # None = use selected global palette
    y_axis_side: str = "left"            # left | right (v2 multi-axis)


SETTINGS = PlotSettings()


# -----------------------------------------------------------------------------
# Generic parsing helpers (bar chart side)
# -----------------------------------------------------------------------------

def split_tokens(text: str) -> list[str]:
    clean = text.replace("\t", ",").replace(";", ",")
    return [part.strip() for part in clean.split(",") if part.strip()]


def parse_float_list(text: str) -> list[float]:
    vals = split_tokens(text)
    if not vals:
        return []
    try:
        return [float(v.replace(",", ".")) for v in vals]
    except ValueError as exc:
        raise ValueError("Sayısal veriler geçerli sayılar olmalıdır.") from exc


def parse_series_block(text: str) -> list[tuple[str, list[float]]]:
    rows = [r.strip() for r in text.splitlines() if r.strip()]
    if not rows:
        return []
    out = []
    for idx, row in enumerate(rows, start=1):
        if ":" in row:
            name, values = row.split(":", 1)
            name = name.strip() or f"Seri {idx}"
        else:
            name, values = f"Seri {idx}", row
        parsed = parse_float_list(values)
        if not parsed:
            raise ValueError(f"'{name}' serisinde veri bulunamadı.")
        out.append((name, parsed))
    return out


def parse_error_block(text: str, series: list[tuple[str, list[float]]]) -> dict[str, list[float]]:
    if not text.strip():
        return {}
    rows = [r.strip() for r in text.splitlines() if r.strip()]
    errors = {}
    named = any(":" in row for row in rows)
    if named:
        for row in rows:
            if ":" not in row:
                raise ValueError("Error bar satırlarının tamamı 'Seri Adı: değerler' biçiminde olmalı.")
            name, values = row.split(":", 1)
            errors[name.strip()] = parse_float_list(values)
    else:
        if len(rows) not in (1, len(series)):
            raise ValueError("Error bar için ya tek satır ya da her seri için bir satır girin.")
        if len(rows) == 1 and len(series) == 1:
            errors[series[0][0]] = parse_float_list(rows[0])
        else:
            for (name, _), row in zip(series, rows):
                errors[name] = parse_float_list(row)
    return errors


def optional_float(text: str) -> Optional[float]:
    text = text.strip()
    return None if not text else float(text.replace(",", "."))


def _to_float(value) -> Optional[float]:
    """Hem 3.14 hem 3,14 biçimini okur; yaygın binlik ayraçlarını da tolere eder."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        return v if math.isfinite(v) else None

    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        v = float(text)
        return v if math.isfinite(v) else None
    except (ValueError, TypeError):
        return None


# -----------------------------------------------------------------------------
# CSV reader for X1,Y1, X2,Y2 ... X10,Y10 — no openpyxl required
# -----------------------------------------------------------------------------

def _read_csv_rows(path: str) -> list[list[str]]:
    """Destek: ; , TAB ayıracı ve 3.14 / 3,14 ondalık biçimleri."""
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(16384)
        f.seek(0)

        lines = [ln for ln in sample.splitlines()[:10] if ln.strip()]
        semicolons = sum(ln.count(";") for ln in lines)
        tabs = sum(ln.count("\t") for ln in lines)

        if semicolons > 0 and semicolons >= tabs:
            delimiter = ";"
        elif tabs > 0:
            delimiter = "\t"
        else:
            try:
                delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
            except csv.Error:
                delimiter = ","

        reader = csv.reader(f, delimiter=delimiter)
        return [row for row in reader if any(str(v).strip() for v in row)]


def load_csv_xy_pairs(
    path: str,
    max_series: int = 10,
    first_row_header: bool = True,
    sort_by_x: bool = False,
) -> list[XYSeries]:
    rows = _read_csv_rows(path)
    if not rows:
        raise ValueError("CSV dosyası boş.")

    max_col = min(max(len(r) for r in rows), max_series * 2)
    if max_col < 2:
        raise ValueError("Dosyada en az iki sütun (X1 ve Y1) bulunmalıdır.")

    header = rows[0] if first_row_header else []
    data_rows = rows[1:] if first_row_header else rows
    output = []

    for pair_idx, x_col in enumerate(range(0, max_col, 2), start=1):
        y_col = x_col + 1
        if y_col >= max_col or pair_idx > max_series:
            break

        xh = str(header[x_col]).strip() if first_row_header and x_col < len(header) and str(header[x_col]).strip() else f"X{pair_idx}"
        yh = str(header[y_col]).strip() if first_row_header and y_col < len(header) and str(header[y_col]).strip() else f"Y{pair_idx}"

        xs, ys = [], []
        for row in data_rows:
            xv = _to_float(row[x_col] if x_col < len(row) else None)
            yv = _to_float(row[y_col] if y_col < len(row) else None)
            if xv is None or yv is None:
                continue
            xs.append(xv)
            ys.append(yv)

        if not xs:
            continue

        x_arr = np.asarray(xs, dtype=float)
        y_arr = np.asarray(ys, dtype=float)

        if sort_by_x:
            order = np.argsort(x_arr, kind="stable")
            x_arr = x_arr[order]
            y_arr = y_arr[order]

        output.append(XYSeries(name=f"Series {pair_idx}", x=x_arr, y=y_arr, x_header=xh, y_header=yh, marker=MARKERS[(pair_idx - 1) % len(MARKERS)]))

    if not output:
        raise ValueError("Geçerli XY çifti bulunamadı. Düzen: X1 | Y1 | X2 | Y2 | ... | X10 | Y10")
    return output


# -----------------------------------------------------------------------------
# Publication-style plot helpers
# -----------------------------------------------------------------------------

def apply_scientific_style(ax, font_size: int, axis_linewidth: float, grid: bool) -> None:
    ax.set_facecolor("white")
    ax.figure.set_facecolor("white")
    for side in ("left", "right", "top", "bottom"):
        ax.spines[side].set_visible(True)
        ax.spines[side].set_linewidth(axis_linewidth)
        ax.spines[side].set_color("#1A1A1A")

    ax.tick_params(
        axis="both", which="major", direction="in", length=5.5,
        width=axis_linewidth, top=True, right=True, pad=6,
        labelsize=font_size, colors="#1A1A1A",
    )
    ax.tick_params(
        axis="both", which="minor", direction="in", length=3.0,
        width=max(0.7, axis_linewidth * 0.8), top=True, right=True,
        colors="#1A1A1A",
    )
    ax.xaxis.set_minor_locator(AutoMinorLocator(2))
    ax.yaxis.set_minor_locator(AutoMinorLocator(2))
    if grid:
        ax.grid(axis="both", which="major", linewidth=0.55, alpha=0.18, zorder=0)
    else:
        ax.grid(False)


def apply_scalar_formatter(ax) -> None:
    formatter_x = ScalarFormatter(useMathText=True)
    formatter_x.set_powerlimits((-3, 4))
    formatter_y = ScalarFormatter(useMathText=True)
    formatter_y.set_powerlimits((-3, 4))
    ax.xaxis.set_major_formatter(formatter_x)
    ax.yaxis.set_major_formatter(formatter_y)


def _comma_number(x: float, pos=None) -> str:
    """Grafikte 3.14 yerine 3,14 gösterir."""
    if abs(x) < 1e-14:
        x = 0.0
    if x != 0 and (abs(x) >= 1e5 or abs(x) < 1e-4):
        return f"{x:.3e}".replace(".", ",")
    return f"{x:g}".replace(".", ",")


def apply_comma_tick_formatter(ax, x_axis: bool = True, y_axis: bool = True) -> None:
    formatter = FuncFormatter(_comma_number)
    if x_axis:
        ax.xaxis.set_major_formatter(formatter)
    if y_axis:
        ax.yaxis.set_major_formatter(formatter)


def style_labels(ax, title: str, xlabel: str, ylabel: str, font_size: int) -> None:
    ax.set_xlabel(xlabel, fontsize=font_size + 1, labelpad=8)
    ax.set_ylabel(ylabel, fontsize=font_size + 1, labelpad=9)
    if title.strip():
        ax.set_title(title.strip(), fontsize=font_size + 2, pad=12, weight="semibold")


def apply_axis_limits(ax, y_min, y_max, y_step, x_min=None, x_max=None, x_step=None) -> None:
    if y_min is not None or y_max is not None:
        lo, hi = ax.get_ylim()
        ax.set_ylim(y_min if y_min is not None else lo, y_max if y_max is not None else hi)
    if y_step is not None:
        if y_step <= 0:
            raise ValueError("Y adımı sıfırdan büyük olmalıdır.")
        ax.yaxis.set_major_locator(MultipleLocator(y_step))
    if x_min is not None or x_max is not None:
        lo, hi = ax.get_xlim()
        ax.set_xlim(x_min if x_min is not None else lo, x_max if x_max is not None else hi)
    if x_step is not None:
        if x_step <= 0:
            raise ValueError("X adımı sıfırdan büyük olmalıdır.")
        ax.xaxis.set_major_locator(MultipleLocator(x_step))


def add_legend(ax, enabled: bool, font_size: int, location: str, ncol: int = 1) -> None:
    if not enabled:
        return
    handles, labels = ax.get_legend_handles_labels()
    if not handles:
        return
    ax.legend(
        frameon=False,
        prop={"family": SETTINGS.legend_font_family, "size": SETTINGS.legend_font_size},
        loc=location,
        ncol=max(1, ncol),
        handlelength=2.2,
        columnspacing=1.2,
        borderaxespad=0.7,
    )


def create_bar_figure(
    categories: list[str], series: list[tuple[str, list[float]]], errors: dict[str, list[float]],
    title: str, xlabel: str, ylabel: str, y_min: Optional[float], y_max: Optional[float],
    y_step: Optional[float], grid: bool, legend: bool, value_labels: bool,
    different_colors: bool, palette_name: str, value_label_mode: str = "Above",
) -> matplotlib.figure.Figure:
    if not categories or not series:
        raise ValueError("Kategori ve en az bir Y serisi gereklidir.")
    n = len(categories)
    for name, vals in series:
        if len(vals) != n:
            raise ValueError(f"'{name}' serisi {len(vals)} değer içeriyor; kategori sayısı {n}.")
        if name in errors and len(errors[name]) != n:
            raise ValueError(f"'{name}' error bar sayısı veri sayısıyla eşleşmiyor.")

    palette = PALETTES[palette_name]
    fig, ax = plt.subplots(figsize=(SETTINGS.figure_width, SETTINGS.figure_height), dpi=110, layout="constrained")
    apply_scientific_style(ax, SETTINGS.font_size, SETTINGS.axis_linewidth, grid)
    x = np.arange(n, dtype=float)
    count = len(series)
    total_group_width = min(0.82, max(0.48, SETTINGS.bar_width))
    width = total_group_width / count

    for j, (name, vals) in enumerate(series):
        xpos = x + (j - (count - 1) / 2) * width
        colors = [palette[i % len(palette)] for i in range(n)] if different_colors and count == 1 else palette[j % len(palette)]
        bars = ax.bar(xpos, vals, width=width * 0.91, color=colors, edgecolor="#202020", linewidth=0.75, label=name, zorder=3)
        if name in errors:
            ax.errorbar(xpos, vals, yerr=np.asarray(errors[name], dtype=float), fmt="none", ecolor="#202020",
                        elinewidth=SETTINGS.error_linewidth, capsize=SETTINGS.capsize,
                        capthick=SETTINGS.error_linewidth, zorder=5)
        if value_labels:
            labels = [f"{v:g}".replace(".", ",") for v in vals]
            fs = max(8, SETTINGS.font_size - 2)

            if value_label_mode == "Inside white box":
                # Place labels inside the bars with a white background for maximum readability.
                for rect, label, value in zip(bars, labels, vals):
                    h = rect.get_height()
                    y = h * 0.82 if h >= 0 else h * 0.82
                    ax.text(
                        rect.get_x() + rect.get_width() / 2,
                        y,
                        label,
                        ha="center",
                        va="center",
                        fontsize=fs,
                        color="#202020",
                        bbox=dict(
                            boxstyle="round,pad=0.22",
                            facecolor="white",
                            edgecolor="none",
                            alpha=0.96,
                        ),
                        zorder=8,
                    )
            elif value_label_mode == "Inside plain":
                ax.bar_label(
                    bars,
                    labels=labels,
                    label_type="center",
                    padding=0,
                    fontsize=fs,
                    color="white",
                    fontweight="bold",
                )
            else:
                ax.bar_label(
                    bars,
                    labels=labels,
                    padding=3,
                    fontsize=fs,
                    color="#202020",
                )

    ax.set_xticks(x)
    ax.set_xticklabels(categories)
    ax.margins(x=0.035)
    style_labels(ax, title, xlabel, ylabel, SETTINGS.font_size)
    apply_axis_limits(ax, y_min, y_max, y_step)
    apply_comma_tick_formatter(ax, x_axis=False, y_axis=True)
    if legend and count > 1:
        add_legend(ax, True, SETTINGS.font_size, "best")
    return fig


def create_line_figure_from_xy(
    xy_series: list[XYSeries],
    title: str,
    xlabel: str,
    ylabel: str,
    y_min: Optional[float],
    y_max: Optional[float],
    y_step: Optional[float],
    x_min: Optional[float],
    x_max: Optional[float],
    x_step: Optional[float],
    grid: bool,
    legend: bool,
    palette_name: str,
    log_x: bool = False,
    log_y: bool = False,
    legend_location: str = "best",
    legend_columns: int = 1,
) -> matplotlib.figure.Figure:
    if not xy_series:
        raise ValueError("Çizilecek CSV serisi bulunamadı.")

    palette = PALETTES[palette_name]
    fig, ax = plt.subplots(
        figsize=(SETTINGS.figure_width, SETTINGS.figure_height),
        dpi=110,
        layout="constrained",
    )
    apply_scientific_style(ax, SETTINGS.font_size, SETTINGS.axis_linewidth, grid)

    if log_x:
        ax.set_xscale("log")
        ax.xaxis.set_minor_locator(LogLocator(base=10.0, subs=np.arange(2, 10) * 0.1, numticks=100))
        ax.xaxis.set_minor_formatter(NullFormatter())
    if log_y:
        ax.set_yscale("log")
        ax.yaxis.set_minor_locator(LogLocator(base=10.0, subs=np.arange(2, 10) * 0.1, numticks=100))
        ax.yaxis.set_minor_formatter(NullFormatter())
    if grid and (log_x or log_y):
        ax.grid(True, which="major", linewidth=0.65, alpha=0.28, zorder=0)
        ax.grid(True, which="minor", linewidth=0.42, linestyle=":", alpha=0.24, zorder=0)

    plotted = 0
    for j, s in enumerate(xy_series[:10]):
        if not s.visible:
            continue

        if log_x and np.any(s.x <= 0):
            raise ValueError(f"'{s.name}' serisinde log-X için sıfır veya negatif X değeri var.")
        if log_y and np.any(s.y <= 0):
            raise ValueError(f"'{s.name}' serisinde log-Y için sıfır veya negatif Y değeri var.")

        color = s.color if s.color else palette[j % len(palette)]

        if s.plot_mode == "Line Only":
            marker = None
            linestyle = s.line_style if s.line_style != "none" else "-"
        elif s.plot_mode == "Symbol Only":
            marker = s.marker
            linestyle = "none"
        else:  # Line + Symbol
            marker = s.marker
            linestyle = s.line_style if s.line_style != "none" else "-"

        # "+" / "x" markers do not have a meaningful face fill.
        if marker in ("+", "x", "|", "_", ".", ",", None):
            marker_face = color
        else:
            marker_face = color if s.marker_fill == "Filled" else "white"

        ax.plot(
            s.x,
            s.y,
            label=s.name,
            color=color,
            linestyle=linestyle,
            linewidth=max(0.1, float(s.line_width)),
            marker=marker,
            markersize=max(0.5, float(s.marker_size)),
            markerfacecolor=marker_face,
            markeredgecolor=color,
            markeredgewidth=max(0.1, float(s.marker_edge_width)),
            markevery=max(1, int(s.marker_every)) if marker is not None else None,
            alpha=min(1.0, max(0.05, float(s.alpha))),
            zorder=3 + j,
        )
        plotted += 1

    if plotted == 0:
        raise ValueError("Görünür seri yok. En az bir seriyi görünür yapın.")

    # Lineer eksenlerde 3,14 göster; log formatter'ını koru.
    apply_comma_tick_formatter(ax, x_axis=not log_x, y_axis=not log_y)

    if x_min is None and x_max is None:
        ax.margins(x=0.025)

    style_labels(ax, title, xlabel, ylabel, SETTINGS.font_size)
    apply_axis_limits(ax, y_min, y_max, y_step, x_min, x_max, x_step)

    if legend:
        add_legend(ax, True, SETTINGS.font_size, legend_location, legend_columns)

    return fig


# -----------------------------------------------------------------------------
# GUI
# -----------------------------------------------------------------------------

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

class PlotCreatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Grafik Oluşturucu Pro v2 — CSV / Origin Style")
        self.geometry("1280x820")
        self.minsize(1100, 700)

        self.current_figure: Optional[matplotlib.figure.Figure] = None
        self.preview_canvas: Optional[FigureCanvasTkAgg] = None
        self.toolbar: Optional[NavigationToolbar2Tk] = None
        self.excel_path: Optional[str] = None
        self.loaded_xy_series: list[XYSeries] = []
        self._series_editor_loading = False
        self._series_applying = False

        self._setup_ttk()
        self._build_ui()
        self.after(80, self._center_window)

    def _setup_ttk(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TButton", padding=(10, 7))
        style.configure("Accent.TButton", padding=(9, 5), font=("Arial", 9, "bold"))
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"))

    def _center_window(self) -> None:
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self) -> None:
        top = ttk.Frame(self, padding=(12, 10))
        top.pack(fill="x")
        ttk.Label(top, text="Scientific Graph Studio v2", font=("Arial", 15, "bold")).pack(side="left")
        ttk.Button(top, text="⚙ Grafik Ayarları", command=self.open_settings).pack(side="right")
        ttk.Button(top, text="💾 Dışa Aktar", command=self.export_current).pack(side="right", padx=(0, 8))

        body = ttk.Panedwindow(self, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        control_outer = ttk.Frame(body)
        preview_outer = ttk.Frame(body)
        body.add(control_outer, weight=0)
        body.add(preview_outer, weight=1)

        control_canvas = tk.Canvas(control_outer, width=470, highlightthickness=0)
        scroll = ttk.Scrollbar(control_outer, orient="vertical", command=control_canvas.yview)
        self.controls = ttk.Frame(control_canvas, padding=(4, 4, 10, 8))
        self.controls.bind("<Configure>", lambda e: control_canvas.configure(scrollregion=control_canvas.bbox("all")))
        control_canvas.create_window((0, 0), window=self.controls, anchor="nw")
        control_canvas.configure(yscrollcommand=scroll.set)
        control_canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.tabs = ttk.Notebook(self.controls)
        self.tabs.pack(fill="both", expand=True)
        self.bar_tab = ttk.Frame(self.tabs, padding=10)
        self.line_tab = ttk.Frame(self.tabs, padding=10)
        self.tabs.add(self.bar_tab, text="📊 Sütun")
        self.tabs.add(self.line_tab, text="📈 CSV Çizgi")
        self._build_bar_tab()
        self._build_line_tab()

        ttk.Label(preview_outer, text="Önizleme", font=("Arial", 11, "bold")).pack(anchor="w", pady=(2, 6))
        self.preview_frame = ttk.Frame(preview_outer)
        self.preview_frame.pack(fill="both", expand=True)
        self._show_placeholder()

    @staticmethod
    def _labeled_entry(parent, label: str, default: str = "", width: int = 44):
        ttk.Label(parent, text=label).pack(anchor="w", pady=(7, 2))
        e = ttk.Entry(parent, width=width)
        e.insert(0, default)
        e.pack(fill="x")
        return e

    @staticmethod
    def _text_box(parent, label: str, height: int = 4, default: str = ""):
        ttk.Label(parent, text=label).pack(anchor="w", pady=(7, 2))
        t = tk.Text(parent, height=height, wrap="none", undo=True, font=("Menlo", 10))
        t.insert("1.0", default)
        t.pack(fill="x")
        return t

    @staticmethod
    def _three_entries(parent, labels: tuple[str, str, str], defaults=("", "", "")):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(7, 0))
        entries = []
        for i, (lab, default) in enumerate(zip(labels, defaults)):
            sub = ttk.Frame(frame)
            sub.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 5, 0))
            frame.columnconfigure(i, weight=1)
            ttk.Label(sub, text=lab).pack(anchor="w")
            e = ttk.Entry(sub, width=10)
            e.insert(0, default)
            e.pack(fill="x")
            entries.append(e)
        return entries

    def _build_bar_tab(self) -> None:
        p = self.bar_tab
        self.bar_x = self._labeled_entry(p, "X kategorileri (virgül ile)", "A, B, C, D")
        self.bar_y = self._text_box(p, "Y serileri — her satır: Seri Adı: değerler", 4, "Series 1: 24, 38, 31, 46")
        self.bar_err = self._text_box(p, "Y error bar (opsiyonel; aynı seri adıyla)", 3, "Series 1: 2.1, 3.0, 2.5, 3.4")
        self.bar_title = self._labeled_entry(p, "Başlık (opsiyonel)", "")
        self.bar_xlabel = self._labeled_entry(p, "X ekseni etiketi", "Condition")
        self.bar_ylabel = self._labeled_entry(p, "Y ekseni etiketi", "Response")
        self.bar_ymin, self.bar_ymax, self.bar_ystep = self._three_entries(
            p, ("Y min", "Y max", "Y adım"), ("0", "", "")
        )

        opt = ttk.LabelFrame(p, text="Görünüm", padding=8)
        opt.pack(fill="x", pady=10)

        self.bar_grid = tk.BooleanVar(value=False)
        self.bar_legend = tk.BooleanVar(value=True)
        self.bar_labels = tk.BooleanVar(value=False)
        self.bar_multicolor = tk.BooleanVar(value=False)

        ttk.Checkbutton(opt, text="Grid", variable=self.bar_grid).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Legend (çoklu seri)", variable=self.bar_legend).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Değerleri göster", variable=self.bar_labels).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Tek seride her sütunu farklı renklendir", variable=self.bar_multicolor).grid(row=3, column=0, sticky="w")

        ttk.Label(opt, text="Değer konumu:").grid(row=2, column=1, sticky="e", padx=(16, 5))
        self.bar_label_mode = tk.StringVar(value="Above")
        ttk.Combobox(
            opt,
            textvariable=self.bar_label_mode,
            values=["Above", "Inside white box", "Inside plain"],
            state="readonly",
            width=18,
        ).grid(row=2, column=2, sticky="w")

        ttk.Button(
            p,
            text="Grafiği Oluştur / Güncelle",
            style="Accent.TButton",
            command=self.draw_bar,
        ).pack(fill="x", pady=(4, 6))

    def _build_line_tab(self) -> None:
        p = self.line_tab

        csv_box = ttk.LabelFrame(p, text="CSV Veri Kaynağı", padding=5)
        csv_box.pack(fill="x", pady=(0, 5))

        csv_top = ttk.Frame(csv_box)
        csv_top.pack(fill="x")
        ttk.Button(csv_top, text="📂 CSV Seç", command=self.choose_csv, width=12).pack(side="left")
        self.excel_file_label = ttk.Label(csv_top, text="Dosya seçilmedi", foreground="#666666")
        self.excel_file_label.pack(side="left", padx=(7, 0), fill="x", expand=True)

        csv_opts = ttk.Frame(csv_box)
        csv_opts.pack(fill="x", pady=(3, 0))
        self.first_row_header = tk.BooleanVar(value=True)
        self.sort_by_x = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            csv_opts,
            text="İlk satır başlık",
            variable=self.first_row_header,
            command=self.reload_excel_preview,
        ).pack(side="left")
        ttk.Checkbutton(
            csv_opts,
            text="X'e göre sırala",
            variable=self.sort_by_x,
            command=self.reload_excel_preview,
        ).pack(side="left", padx=(10, 0))
        ttk.Label(
            csv_opts,
            text="X1|Y1 ... X10|Y10 • 3.14 / 3,14",
            foreground="#666666",
        ).pack(side="right")

        summary_box = ttk.LabelFrame(p, text="Yüklenen Seriler — seçip aşağıdan düzenleyin (maks. 10)", padding=7)
        summary_box.pack(fill="x", pady=(0, 8))

        self.series_tree = ttk.Treeview(
            summary_box,
            columns=("idx", "name", "n", "mode", "marker", "line", "color"),
            show="headings",
            height=6,
            selectmode="browse",
        )
        headings = {
            "idx": "#",
            "name": "Seri adı",
            "n": "N",
            "mode": "Mod",
            "marker": "Symbol",
            "line": "Line",
            "color": "Color",
        }
        widths = {"idx": 30, "name": 128, "n": 42, "mode": 82, "marker": 62, "line": 58, "color": 70}
        for col in ("idx", "name", "n", "mode", "marker", "line", "color"):
            self.series_tree.heading(col, text=headings[col])
            self.series_tree.column(col, width=widths[col], anchor="center" if col != "name" else "w")
        self.series_tree.pack(fill="x")
        self.series_tree.bind("<<TreeviewSelect>>", self._on_series_select)

        editor = ttk.LabelFrame(p, text="Seçili Seri Ayarları", padding=8)
        editor.pack(fill="x", pady=(0, 9))

        # Row 0: name + visible
        ttk.Label(editor, text="Seri adı").grid(row=0, column=0, sticky="w")
        self.series_name_var = tk.StringVar(value="Series 1")
        self.series_name_entry = ttk.Entry(editor, textvariable=self.series_name_var, width=22)
        self.series_name_entry.grid(row=0, column=1, sticky="ew", padx=(5, 10))
        self.series_visible_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(editor, text="Görünür", variable=self.series_visible_var).grid(row=0, column=2, sticky="w")

        # Row 1: plot mode
        ttk.Label(editor, text="Çizim modu").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.series_mode_var = tk.StringVar(value="Line + Symbol")
        self.series_mode_combo = ttk.Combobox(
            editor,
            textvariable=self.series_mode_var,
            values=["Line + Symbol", "Line Only", "Symbol Only"],
            state="readonly",
            width=18,
        )
        self.series_mode_combo.grid(row=1, column=1, sticky="w", padx=(5, 10), pady=(5, 0))

        ttk.Label(editor, text="Symbol").grid(row=1, column=2, sticky="e", pady=(5, 0))
        self.series_marker_name_var = tk.StringVar(value="Circle (o)")
        self.series_marker_combo = ttk.Combobox(
            editor,
            textvariable=self.series_marker_name_var,
            values=list(MARKER_OPTIONS.keys()),
            state="readonly",
            width=19,
        )
        self.series_marker_combo.grid(row=1, column=3, sticky="w", padx=(5, 0), pady=(5, 0))

        # Row 2: marker fill / size / every
        ttk.Label(editor, text="Symbol dolgu").grid(row=2, column=0, sticky="w", pady=(5, 0))
        self.series_fill_var = tk.StringVar(value="Open")
        self.series_fill_combo = ttk.Combobox(
            editor,
            textvariable=self.series_fill_var,
            values=["Open", "Filled"],
            state="readonly",
            width=10,
        )
        self.series_fill_combo.grid(row=2, column=1, sticky="w", padx=(5, 10), pady=(5, 0))

        ttk.Label(editor, text="Boyut").grid(row=2, column=2, sticky="e", pady=(5, 0))
        self.series_marker_size_var = tk.StringVar(value=str(SETTINGS.marker_size))
        self.series_marker_size_spin = ttk.Spinbox(
            editor, from_=0.5, to=30, increment=0.5,
            textvariable=self.series_marker_size_var, width=7
        )
        self.series_marker_size_spin.grid(row=2, column=3, sticky="w", padx=(5, 0), pady=(5, 0))

        ttk.Label(editor, text="Symbol kalınlığı").grid(row=8, column=0, sticky="w", pady=(5, 0))
        self.series_marker_edge_width_var = tk.StringVar(value=str(SETTINGS.marker_edge_width))
        self.series_marker_edge_width_spin = ttk.Spinbox(
            editor, from_=0.1, to=8.0, increment=0.1,
            textvariable=self.series_marker_edge_width_var, width=8
        )
        self.series_marker_edge_width_spin.grid(row=8, column=1, sticky="w", padx=(5, 10), pady=(5, 0))

        ttk.Label(editor, text="Symbol aralığı").grid(row=3, column=0, sticky="w", pady=(5, 0))
        self.series_marker_every_var = tk.StringVar(value="1")
        self.series_marker_every_spin = ttk.Spinbox(
            editor, from_=1, to=10000,
            textvariable=self.series_marker_every_var, width=8
        )
        self.series_marker_every_spin.grid(row=3, column=1, sticky="w", padx=(5, 10), pady=(5, 0))

        # Row 3/4: line style
        ttk.Label(editor, text="Çizgi stili").grid(row=3, column=2, sticky="e", pady=(5, 0))
        self.series_line_style_name_var = tk.StringVar(value="Solid")
        self.series_line_style_combo = ttk.Combobox(
            editor,
            textvariable=self.series_line_style_name_var,
            values=list(LINESTYLE_OPTIONS.keys()),
            state="readonly",
            width=12,
        )
        self.series_line_style_combo.grid(row=3, column=3, sticky="w", padx=(5, 0), pady=(5, 0))

        ttk.Label(editor, text="Çizgi kalınlığı").grid(row=4, column=0, sticky="w", pady=(5, 0))
        self.series_line_width_var = tk.StringVar(value=str(SETTINGS.line_width))
        self.series_line_width_spin = ttk.Spinbox(
            editor, from_=0.1, to=8.0, increment=0.1,
            textvariable=self.series_line_width_var, width=8
        )
        self.series_line_width_spin.grid(row=4, column=1, sticky="w", padx=(5, 10), pady=(5, 0))

        ttk.Label(editor, text="Opacity").grid(row=4, column=2, sticky="e", pady=(5, 0))
        self.series_alpha_var = tk.StringVar(value="1.0")
        self.series_alpha_spin = ttk.Spinbox(
            editor, from_=0.05, to=1.0, increment=0.05,
            textvariable=self.series_alpha_var, width=8
        )
        self.series_alpha_spin.grid(row=4, column=3, sticky="w", padx=(5, 0), pady=(5, 0))

        ttk.Label(editor, text="Seri rengi").grid(row=5, column=0, sticky="w", pady=(5, 0))
        color_row = ttk.Frame(editor)
        color_row.grid(row=5, column=1, columnspan=3, sticky="ew", padx=(5, 0), pady=(5, 0))
        self.series_color_var = tk.StringVar(value="Auto / Palette")
        self.series_color_combo = ttk.Combobox(
            color_row,
            textvariable=self.series_color_var,
            values=[
                "Auto / Palette",
                "Black", "Red", "Blue", "Green", "Orange", "Purple",
                "Cyan", "Magenta", "Gray", "Brown",
            ],
            state="readonly",
            width=16,
        )
        self.series_color_combo.pack(side="left")
        self.series_color_preview = tk.Label(color_row, text="   ", relief="solid", bd=1, bg="#FFFFFF")
        self.series_color_preview.pack(side="left", padx=(6, 4))
        ttk.Button(
            color_row,
            text="Özel Renk…",
            width=11,
            command=self.choose_series_custom_color,
        ).pack(side="left")

        # Seçim değişikliklerini anında uygula.
        for widget in (
            self.series_mode_combo,
            self.series_marker_combo,
            self.series_fill_combo,
            self.series_line_style_combo,
            self.series_color_combo,
        ):
            widget.bind("<<ComboboxSelected>>", self._apply_series_on_event)

        for widget in (
            self.series_name_entry,
            self.series_marker_size_spin,
            self.series_marker_edge_width_spin,
            self.series_marker_every_spin,
            self.series_line_width_spin,
            self.series_alpha_spin,
        ):
            widget.bind("<Return>", self._apply_series_on_event)
            widget.bind("<FocusOut>", self._apply_series_on_event)

        self.series_visible_var.trace_add(
            "write",
            lambda *_: self._apply_series_on_event()
            if (not self._series_editor_loading
                and not self._series_applying
                and self._selected_series_index() is not None)
            else None
        )

        editor.columnconfigure(1, weight=1)
        editor.columnconfigure(3, weight=1)

        btnrow = ttk.Frame(editor)
        btnrow.grid(row=6, column=0, columnspan=4, sticky="ew", pady=(8, 0))
        ttk.Button(btnrow, text="Seçili Seriye Uygula", command=self.apply_selected_series_settings).pack(side="left", fill="x", expand=True)
        ttk.Button(btnrow, text="Tüm Serilere Varsayılan Stil", command=self.reset_all_series_styles).pack(side="left", fill="x", expand=True, padx=(6, 0))

        self.line_title = self._labeled_entry(p, "Başlık (opsiyonel)", "")
        self.line_xlabel = self._labeled_entry(p, "X ekseni etiketi", "X")
        self.line_ylabel = self._labeled_entry(p, "Y ekseni etiketi", "Y")
        self.line_ymin, self.line_ymax, self.line_ystep = self._three_entries(
            p, ("Y min", "Y max", "Y adım"), ("", "", "")
        )
        self.line_xmin, self.line_xmax, self.line_xstep = self._three_entries(
            p, ("X min", "X max", "X adım"), ("", "", "")
        )

        opt = ttk.LabelFrame(p, text="Genel Çizgi Grafiği Ayarları", padding=6)
        opt.pack(fill="x", pady=6)
        ttk.Label(
            opt,
            text="Palette, yalnız rengi 'Auto / Palette' olan serilere uygulanır.",
            foreground="#666666",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 3))
        self.line_grid = tk.BooleanVar(value=False)
        self.line_legend = tk.BooleanVar(value=True)
        self.line_logx = tk.BooleanVar(value=False)
        self.line_logy = tk.BooleanVar(value=False)

        ttk.Checkbutton(opt, text="Grid", variable=self.line_grid).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Legend", variable=self.line_legend).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(opt, text="Log X", variable=self.line_logx).grid(row=2, column=1, sticky="w", padx=(18, 0))
        ttk.Checkbutton(opt, text="Log Y", variable=self.line_logy).grid(row=1, column=1, sticky="w", padx=(18, 0))

        ttk.Label(opt, text="Legend konumu:").grid(row=4, column=0, sticky="w", pady=(6, 0))
        self.legend_loc_var = tk.StringVar(value="best")
        locs = [
            "best", "upper right", "upper left", "lower right", "lower left",
            "center right", "center left", "upper center", "lower center", "center",
        ]
        ttk.Combobox(
            opt,
            textvariable=self.legend_loc_var,
            values=locs,
            state="readonly",
            width=15,
        ).grid(row=4, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

        ttk.Label(opt, text="Legend sütunu:").grid(row=3, column=0, sticky="w", pady=(6, 0))
        self.legend_cols = ttk.Spinbox(opt, from_=1, to=5, width=8)
        self.legend_cols.set("1")
        self.legend_cols.grid(row=3, column=1, sticky="w", padx=(6, 0), pady=(6, 0))

        ttk.Button(
            p,
            text="CSV Verisinden Grafiği Oluştur",
            style="Accent.TButton",
            command=self.draw_line,
        ).pack(fill="x", pady=(4, 6))

    def _apply_series_on_event(self, event=None) -> None:
        """Seçili seri editöründeki değişiklikleri güvenli biçimde uygular."""
        if self._series_editor_loading or self._series_applying:
            return
        if self._selected_series_index() is not None:
            self.apply_selected_series_settings()

    def choose_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="CSV dosyasi sec",
            filetypes=[("CSV / Text", "*.csv *.txt"), ("Tum dosyalar", "*.*")],
        )
        if not path:
            return
        try:
            self.excel_path = path  # mevcut state adini geriye uyumluluk icin koruyoruz
            self.excel_file_label.configure(text=Path(path).name, foreground="#1A1A1A")
            self.reload_excel_preview()
        except Exception as exc:
            self.excel_path = None
            messagebox.showerror("CSV acilamadi", str(exc))

    def reload_excel_preview(self) -> None:
        if not self.excel_path:
            return
        try:
            old_styles = {}
            for i, s in enumerate(self.loaded_xy_series):
                old_styles[i] = {
                    "name": s.name,
                    "plot_mode": s.plot_mode,
                    "marker": s.marker,
                    "marker_fill": s.marker_fill,
                    "line_style": s.line_style,
                    "line_width": s.line_width,
                    "marker_size": s.marker_size,
                    "marker_edge_width": s.marker_edge_width,
                    "marker_every": s.marker_every,
                    "alpha": s.alpha,
                    "visible": s.visible,
                    "color": s.color,
                }

            new_series = load_csv_xy_pairs(
                self.excel_path,
                max_series=10,
                first_row_header=self.first_row_header.get(),
                sort_by_x=self.sort_by_x.get(),
            )

            # Preserve program-side names/styles when reloading the same CSV.
            for i, s in enumerate(new_series):
                if i in old_styles:
                    for key, value in old_styles[i].items():
                        setattr(s, key, value)
                else:
                    s.name = f"Series {i + 1}"
                    s.plot_mode = "Line + Symbol"
                    s.marker = MARKERS[i % len(MARKERS)]
                    s.marker_fill = "Open"
                    s.line_style = LINESTYLES[i % len(LINESTYLES)]
                    s.line_width = SETTINGS.line_width
                    s.marker_size = SETTINGS.marker_size
                    s.marker_edge_width = SETTINGS.marker_edge_width
                    s.marker_every = 1
                    s.alpha = 1.0
                    s.visible = True
                    s.color = None

            self.loaded_xy_series = new_series
            self._refresh_series_tree()

            if self.loaded_xy_series:
                first_item = self.series_tree.get_children()[0]
                self.series_tree.selection_set(first_item)
                self.series_tree.focus(first_item)
                self._load_series_into_editor(0)

        except Exception as exc:
            self.loaded_xy_series = []
            for item in self.series_tree.get_children():
                self.series_tree.delete(item)
            messagebox.showerror("CSV verisi okunamadı", str(exc))

    def _marker_display_name(self, marker) -> str:
        for name, value in MARKER_OPTIONS.items():
            if value == marker:
                return name
        return "None"

    def _linestyle_display_name(self, linestyle: str) -> str:
        for name, value in LINESTYLE_OPTIONS.items():
            if value == linestyle:
                return name
        return "Solid"

    def _refresh_series_tree(self) -> None:
        for item in self.series_tree.get_children():
            self.series_tree.delete(item)

        for idx, s in enumerate(self.loaded_xy_series):
            marker_name = self._marker_display_name(s.marker)
            marker_short = marker_name.split(" (")[0]
            line_name = self._linestyle_display_name(s.line_style)
            if not s.visible:
                mode_text = "Hidden"
            else:
                mode_text = s.plot_mode

            self.series_tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(
                    idx + 1,
                    s.name,
                    len(s.x),
                    mode_text,
                    marker_short,
                    line_name,
                    (s.color if s.color else "Palette"),
                ),
            )

    def _selected_series_index(self) -> Optional[int]:
        selection = self.series_tree.selection()
        if not selection:
            return None
        try:
            return int(selection[0])
        except (ValueError, TypeError):
            return None

    def _on_series_select(self, event=None) -> None:
        idx = self._selected_series_index()
        if idx is not None:
            self._load_series_into_editor(idx)

    def _named_color_to_hex(self, name: str) -> Optional[str]:
        mapping = {
            "Black": "#000000",
            "Red": "#D62728",
            "Blue": "#1F77B4",
            "Green": "#2CA02C",
            "Orange": "#FF7F0E",
            "Purple": "#9467BD",
            "Cyan": "#17BECF",
            "Magenta": "#E377C2",
            "Gray": "#7F7F7F",
            "Brown": "#8C564B",
        }
        return mapping.get(name)

    def _hex_to_named_color(self, value: Optional[str]) -> str:
        if not value:
            return "Auto / Palette"
        for name in ("Black", "Red", "Blue", "Green", "Orange", "Purple", "Cyan", "Magenta", "Gray", "Brown"):
            if self._named_color_to_hex(name).lower() == str(value).lower():
                return name
        return str(value)

    def _update_series_color_preview(self, value: Optional[str]) -> None:
        try:
            if value:
                self.series_color_preview.configure(bg=value)
            else:
                self.series_color_preview.configure(bg="#FFFFFF")
        except Exception:
            self.series_color_preview.configure(bg="#FFFFFF")

    def choose_series_custom_color(self) -> None:
        idx = self._selected_series_index()
        if idx is None:
            messagebox.showwarning("Seri seçilmedi", "Önce yüklenen seriler listesinden bir seri seçin.")
            return
        initial = self.loaded_xy_series[idx].color or "#1F77B4"
        result = colorchooser.askcolor(color=initial, title="Seri rengini seç")
        if not result or not result[1]:
            return
        hex_color = result[1].upper()
        self.series_color_var.set(hex_color)
        self._update_series_color_preview(hex_color)
        self.apply_selected_series_settings()

    def _load_series_into_editor(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.loaded_xy_series):
            return

        self._series_editor_loading = True
        try:
            s = self.loaded_xy_series[idx]
            self.series_name_var.set(s.name)
            self.series_visible_var.set(s.visible)
            self.series_mode_var.set(s.plot_mode)
            self.series_marker_name_var.set(self._marker_display_name(s.marker))
            self.series_fill_var.set(s.marker_fill)
            self.series_marker_size_var.set(f"{s.marker_size:g}")
            self.series_marker_edge_width_var.set(f"{s.marker_edge_width:g}")
            self.series_marker_every_var.set(str(s.marker_every))
            self.series_line_style_name_var.set(self._linestyle_display_name(s.line_style))
            self.series_line_width_var.set(f"{s.line_width:g}")
            self.series_alpha_var.set(f"{s.alpha:g}")
            self.series_color_var.set(self._hex_to_named_color(s.color))
            self._update_series_color_preview(s.color)
        finally:
            self._series_editor_loading = False

    def apply_selected_series_settings(self) -> None:
        if self._series_editor_loading or self._series_applying:
            return

        idx = self._selected_series_index()
        if idx is None:
            messagebox.showwarning("Seri seçilmedi", "Önce yüklenen seriler listesinden bir seri seçin.")
            return

        self._series_applying = True
        try:
            s = self.loaded_xy_series[idx]
            name = self.series_name_var.get().strip() or f"Series {idx + 1}"
            marker = MARKER_OPTIONS.get(self.series_marker_name_var.get(), "o")
            line_style = LINESTYLE_OPTIONS.get(self.series_line_style_name_var.get(), "-")
            marker_size = float(self.series_marker_size_var.get().replace(",", "."))
            marker_edge_width = float(self.series_marker_edge_width_var.get().replace(",", "."))
            marker_every = int(float(self.series_marker_every_var.get().replace(",", ".")))
            line_width = float(self.series_line_width_var.get().replace(",", "."))
            alpha = float(self.series_alpha_var.get().replace(",", "."))

            color_choice = self.series_color_var.get().strip()
            if color_choice == "Auto / Palette":
                custom_color = None
            elif color_choice.startswith("#") and len(color_choice) in (4, 7):
                custom_color = color_choice
            else:
                custom_color = self._named_color_to_hex(color_choice)
                if custom_color is None:
                    raise ValueError("Geçersiz seri rengi.")

            if marker_size <= 0:
                raise ValueError("Symbol boyutu sıfırdan büyük olmalıdır.")
            if marker_edge_width <= 0:
                raise ValueError("Symbol kalınlığı sıfırdan büyük olmalıdır.")
            if marker_every < 1:
                raise ValueError("Symbol aralığı en az 1 olmalıdır.")
            if line_width <= 0:
                raise ValueError("Çizgi kalınlığı sıfırdan büyük olmalıdır.")
            if not (0.05 <= alpha <= 1.0):
                raise ValueError("Opacity 0,05 ile 1,00 arasında olmalıdır.")

            s.name = name
            s.visible = self.series_visible_var.get()
            s.plot_mode = self.series_mode_var.get()
            s.marker = marker
            s.marker_fill = self.series_fill_var.get()
            s.marker_size = marker_size
            s.marker_edge_width = marker_edge_width
            s.marker_every = marker_every
            s.line_style = line_style
            s.line_width = line_width
            s.alpha = alpha
            s.color = custom_color

            self._update_series_color_preview(s.color)
            self._refresh_series_tree()
            self.series_tree.selection_set(str(idx))
            self.series_tree.focus(str(idx))

        except Exception as exc:
            messagebox.showerror("Seri ayarları uygulanamadı", str(exc))
            return
        finally:
            self._series_applying = False

        # Redraw only after all callbacks are unlocked.
        if self.excel_path and self.loaded_xy_series:
            try:
                self._redraw_line_preview()
            except Exception as exc:
                messagebox.showerror("Önizleme güncellenemedi", str(exc))

    def reset_all_series_styles(self) -> None:
        if not self.loaded_xy_series:
            return

        for i, s in enumerate(self.loaded_xy_series):
            s.name = f"Series {i + 1}"
            s.visible = True
            s.plot_mode = "Line + Symbol"
            s.marker = MARKERS[i % len(MARKERS)]
            s.marker_fill = "Open"
            s.line_style = LINESTYLES[i % len(LINESTYLES)]
            s.line_width = SETTINGS.line_width
            s.marker_size = SETTINGS.marker_size
            s.marker_edge_width = SETTINGS.marker_edge_width
            s.marker_every = 1
            s.alpha = 1.0
            s.color = None

        self._refresh_series_tree()
        first_item = self.series_tree.get_children()[0]
        self.series_tree.selection_set(first_item)
        self.series_tree.focus(first_item)
        self._load_series_into_editor(0)

        if self.excel_path and self.loaded_xy_series:
            self._redraw_line_preview()

    def _show_placeholder(self) -> None:
        for child in self.preview_frame.winfo_children():
            child.destroy()
        frame = ttk.Frame(self.preview_frame, padding=25)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Grafik önizlemesi burada görünecek.", font=("Arial", 12)).place(relx=0.5, rely=0.48, anchor="center")
        ttk.Label(frame, text="Çizgi grafiği için CSV seçin; sütun grafiği için soldan veri girin.").place(relx=0.5, rely=0.53, anchor="center")

    def _render(self, fig) -> None:
        if self.current_figure is not None:
            plt.close(self.current_figure)
        self.current_figure = fig
        for child in self.preview_frame.winfo_children():
            child.destroy()
        canvas = FigureCanvasTkAgg(fig, master=self.preview_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar = NavigationToolbar2Tk(canvas, self.preview_frame, pack_toolbar=False)
        toolbar.update()
        toolbar.pack(fill="x", pady=(4, 0))
        self.preview_canvas = canvas
        self.toolbar = toolbar

    def draw_bar(self) -> None:
        try:
            categories = split_tokens(self.bar_x.get())
            series = parse_series_block(self.bar_y.get("1.0", "end"))
            errors = parse_error_block(self.bar_err.get("1.0", "end"), series)
            fig = create_bar_figure(
                categories, series, errors, self.bar_title.get(), self.bar_xlabel.get(), self.bar_ylabel.get(),
                optional_float(self.bar_ymin.get()), optional_float(self.bar_ymax.get()), optional_float(self.bar_ystep.get()),
                self.bar_grid.get(), self.bar_legend.get(), self.bar_labels.get(), self.bar_multicolor.get(), SETTINGS.palette_name,
                self.bar_label_mode.get(),
            )
            self._render(fig)
        except Exception as exc:
            messagebox.showerror("Grafik oluşturulamadı", str(exc))

    def _redraw_line_preview(self) -> None:
        """Mevcut seri ayarlarıyla çizgi grafiğini yeniden oluşturur."""
        if not self.excel_path or not self.loaded_xy_series:
            return

        legend_columns = int(self.legend_cols.get())
        fig = create_line_figure_from_xy(
            xy_series=self.loaded_xy_series,
            title=self.line_title.get(),
            xlabel=self.line_xlabel.get(),
            ylabel=self.line_ylabel.get(),
            y_min=optional_float(self.line_ymin.get()),
            y_max=optional_float(self.line_ymax.get()),
            y_step=optional_float(self.line_ystep.get()),
            x_min=optional_float(self.line_xmin.get()),
            x_max=optional_float(self.line_xmax.get()),
            x_step=optional_float(self.line_xstep.get()),
            grid=self.line_grid.get(),
            legend=self.line_legend.get(),
            palette_name=SETTINGS.palette_name,
            log_x=self.line_logx.get(),
            log_y=self.line_logy.get(),
            legend_location=self.legend_loc_var.get(),
            legend_columns=legend_columns,
        )
        self._render(fig)

    def draw_line(self) -> None:
        try:
            if not self.excel_path:
                raise ValueError("Önce bir CSV dosyası seçin.")
            if not self.loaded_xy_series:
                raise ValueError("CSV içinde geçerli XY serisi bulunamadı.")

            self._redraw_line_preview()

        except Exception as exc:
            messagebox.showerror("Grafik oluşturulamadı", str(exc))

    def open_settings(self) -> None:
        win = tk.Toplevel(self)
        win.title("Grafik Ayarları")
        win.resizable(False, False)
        box = ttk.Frame(win, padding=14)
        box.pack(fill="both", expand=True)

        entries = {}
        fields = [
            ("dpi", "Export DPI", str(SETTINGS.dpi)),
            ("font_size", "Font size", str(SETTINGS.font_size)),
            ("axis_linewidth", "Axis linewidth", str(SETTINGS.axis_linewidth)),
            ("line_width", "Line width", str(SETTINGS.line_width)),
            ("marker_size", "Marker size", str(SETTINGS.marker_size)),
            ("marker_edge_width", "Symbol edge width", str(SETTINGS.marker_edge_width)),
            ("legend_font_family", "Legend font family", SETTINGS.legend_font_family),
            ("legend_font_size", "Legend font size", str(SETTINGS.legend_font_size)),
            ("bar_width", "Bar group width", str(SETTINGS.bar_width)),
            ("error_linewidth", "Error-bar linewidth", str(SETTINGS.error_linewidth)),
            ("capsize", "Error-bar cap size", str(SETTINGS.capsize)),
            ("figure_width", "Figure width (inch)", str(SETTINGS.figure_width)),
            ("figure_height", "Figure height (inch)", str(SETTINGS.figure_height)),
        ]
        for row, (key, label, default) in enumerate(fields):
            ttk.Label(box, text=label).grid(row=row, column=0, sticky="w", pady=3, padx=(0, 10))
            e = ttk.Entry(box, width=14)
            e.insert(0, default)
            e.grid(row=row, column=1, sticky="ew", pady=3)
            entries[key] = e

        ttk.Label(box, text="Palette (10 seçenek)").grid(row=len(fields), column=0, sticky="w", pady=(8, 3))
        palette_var = tk.StringVar(value=SETTINGS.palette_name)
        palette = ttk.Combobox(box, textvariable=palette_var, values=list(PALETTES), state="readonly", width=22)
        palette.grid(row=len(fields), column=1, sticky="ew", pady=(8, 3))

        # Palette swatches preview
        swatch_frame = ttk.Frame(box)
        swatch_frame.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="ew", pady=(5, 4))
        swatches = []
        for i in range(10):
            c = tk.Canvas(swatch_frame, width=26, height=18, highlightthickness=1, highlightbackground="#AAAAAA")
            c.pack(side="left", padx=2)
            swatches.append(c)

        def update_swatches(*_):
            cols = PALETTES[palette_var.get()]
            for c, color in zip(swatches, cols):
                c.delete("all")
                c.create_rectangle(0, 0, 30, 22, fill=color, outline="")

        palette_var.trace_add("write", update_swatches)
        update_swatches()

        def save_settings():
            try:
                SETTINGS.dpi = int(entries["dpi"].get())
                SETTINGS.font_size = int(entries["font_size"].get())
                SETTINGS.axis_linewidth = float(entries["axis_linewidth"].get())
                SETTINGS.line_width = float(entries["line_width"].get())
                SETTINGS.marker_size = float(entries["marker_size"].get())
                SETTINGS.marker_edge_width = float(entries["marker_edge_width"].get())
                SETTINGS.legend_font_family = entries["legend_font_family"].get().strip() or "Arial"
                SETTINGS.legend_font_size = int(entries["legend_font_size"].get())
                SETTINGS.bar_width = float(entries["bar_width"].get())
                SETTINGS.error_linewidth = float(entries["error_linewidth"].get())
                SETTINGS.capsize = float(entries["capsize"].get())
                SETTINGS.figure_width = float(entries["figure_width"].get())
                SETTINGS.figure_height = float(entries["figure_height"].get())
                SETTINGS.palette_name = palette_var.get()
                if SETTINGS.dpi <= 0 or SETTINGS.font_size <= 0 or SETTINGS.figure_width <= 0 or SETTINGS.figure_height <= 0:
                    raise ValueError("DPI, font ve figür boyutları pozitif olmalıdır.")
                win.destroy()
                if self.tabs.index(self.tabs.select()) == 0:
                    self.draw_bar()
                elif self.excel_path:
                    self.draw_line()
            except Exception as exc:
                messagebox.showerror("Geçersiz ayar", str(exc), parent=win)

        ttk.Button(box, text="Uygula", command=save_settings, style="Accent.TButton").grid(
            row=len(fields) + 2, column=0, columnspan=2, sticky="ew", pady=(12, 0)
        )
        win.transient(self)
        win.grab_set()

    def export_current(self) -> None:
        if self.current_figure is None:
            messagebox.showinfo("Dışa aktar", "Önce bir grafik oluşturun.")
            return
        path = filedialog.asksaveasfilename(
            title="Grafiği dışa aktar",
            defaultextension=".png",
            filetypes=[
                ("PNG - yüksek çözünürlük", "*.png"),
                ("PDF - vektörel", "*.pdf"),
                ("SVG - vektörel", "*.svg"),
                ("TIFF - yayın", "*.tif *.tiff"),
                ("Tüm dosyalar", "*.*"),
            ],
        )
        if not path:
            return
        try:
            suffix = Path(path).suffix.lower()
            kwargs = {"bbox_inches": "tight", "pad_inches": 0.05, "facecolor": "white"}
            if suffix in {".png", ".tif", ".tiff", ".jpg", ".jpeg"}:
                kwargs["dpi"] = SETTINGS.dpi
            self.current_figure.savefig(path, **kwargs)
            messagebox.showinfo("Tamam", f"Grafik kaydedildi:\n{path}")
        except Exception as exc:
            messagebox.showerror("Kaydetme hatası", str(exc))


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
    matplotlib.rcParams.update({
        "font.family": "sans-serif",
        "axes.unicode_minus": False,
        "savefig.transparent": False,
        "pdf.fonttype": 42,
        "ps.fonttype": 42
    })
    GrafikOlusturucuV2().mainloop()
