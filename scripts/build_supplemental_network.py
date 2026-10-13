#!/usr/bin/env python3
"""Recreate the image-backed network reliability box plots from worksheet data.

Each modelled result is retained. Plotly computes linear-interpolation quartiles
and Tukey whiskers from those results, rather than digitising the source images.
"""

from __future__ import annotations

from collections import Counter
import json
import math
from pathlib import Path
import statistics
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "data/sources/2026-wem-esoo-data-register.xlsx"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
YEARS = [f"{year}-{str(year + 1)[2:]}" for year in range(2026, 2036)]
THRESHOLD = 0.0002  # Source values are percentage points, not fractions.
SOURCE_URL = "https://www.aemo.com.au/-/media/files/electricity/wem/planning_and_forecasting/esoo/2026/2026-wem-esoo-data-register.xlsx"


def worksheet_rows(archive, part, first_row):
    strings = [
        "".join(t.text or "" for t in item.findall(".//s:t", NS))
        for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))
    ]
    rows = []
    for row in ET.fromstring(archive.read(part)).findall("s:sheetData/s:row", NS):
        row_number = int(row.attrib["r"])
        if row_number < first_row:
            continue
        values = {}
        for cell in row:
            value = cell.find("s:v", NS)
            if value is None:
                continue
            column = "".join(c for c in cell.attrib["r"] if c.isalpha())
            values[column] = (
                strings[int(value.text)]
                if cell.get("t") == "s"
                else float(value.text)
            )
        if values.get("B") in YEARS:
            rows.append((row_number, values))
    return rows


def numeric_percent(value):
    """The exceptional text cell carries %, while numeric cells use %-point units."""
    if isinstance(value, str):
        assert value.endswith("%"), f"Unexpected non-numeric USE value: {value!r}"
        value = float(value[:-1])
    assert isinstance(value, (float, int)) and math.isfinite(value)
    return value


def axis(title="", **extra):
    return {
        "title": {"text": title, "standoff": 14},
        "automargin": True,
        "gridcolor": "#dfe7eb",
        "zerolinecolor": "#b8c7cf",
        **extra,
    }


def base_layout(height=600):
    return {
        "height": height,
        "autosize": True,
        "margin": {"l": 78, "r": 32, "t": 50, "b": 105},
        "paper_bgcolor": "#ffffff",
        "plot_bgcolor": "#ffffff",
        "font": {"family": "Inter, Arial, sans-serif", "color": "#173543", "size": 12},
        "legend": {"orientation": "h", "x": 0, "y": 1.12},
        "hovermode": "closest",
        "boxmode": "group",
        "boxgap": 0.4,
        "boxgroupgap": 0.18,
    }


def box_trace(rows, column, name, color, fill, points=False, xaxis="x", yaxis="y"):
    return {
        "type": "box",
        "name": name,
        "x": [values["B"] for _, values in rows],
        "y": [numeric_percent(values[column]) for _, values in rows],
        "customdata": [row_number for row_number, _ in rows],
        "xaxis": xaxis,
        "yaxis": yaxis,
        "quartilemethod": "linear",
        "boxmean": True,
        # Plotly draws min/max whiskers when boxpoints=False. Retain the Tukey
        # whiskers while hiding outlier marks for Figure 42, as its image does.
        "boxpoints": points if points else "outliers",
        "jitter": 0,
        "pointpos": 0,
        "line": {"color": color, "width": 1.5},
        "fillcolor": fill,
        "marker": {"color": color, "size": 3, "opacity": 0.55 if points else 0},
        "hoveron": "boxes+points" if points else "boxes",
        "hovertemplate": "%{x}<br>%{y:.6f}%<extra>" + name + "</extra>",
        "legendgroup": name,
        "offsetgroup": name,
        "meta": {"sourceColumn": column, "unit": "percentage points"},
    }


def threshold_line(xref="paper", yref="y", **extra):
    return {
        "type": "line",
        "xref": xref,
        "yref": yref,
        "x0": 0,
        "x1": 1,
        "y0": THRESHOLD,
        "y1": THRESHOLD,
        "line": {"color": "#ad2353", "width": 1.7, "dash": "dash"},
        **extra,
    }


def summary(rows, columns):
    by_year = {}
    for year in YEARS:
        selected = [values for _, values in rows if values["B"] == year]
        by_year[year] = {}
        for column, name in columns.items():
            values = [numeric_percent(item[column]) for item in selected]
            ordered = sorted(values)
            def plotly_linear_quantile(fraction):
                rank = len(ordered) * fraction - 0.5
                low = max(0, math.floor(rank))
                high = min(len(ordered) - 1, math.ceil(rank))
                return ordered[low] + (ordered[high] - ordered[low]) * (rank - low)
            q1, median, q3 = [plotly_linear_quantile(fraction) for fraction in (0.25, 0.5, 0.75)]
            iqr = q3 - q1
            by_year[year][name] = {
                "count": len(values),
                "minimum": min(values),
                "q1": q1,
                "median": median,
                "mean": statistics.mean(values),
                "q3": q3,
                "maximum": max(values),
                "lower_whisker": min(v for v in values if v >= q1 - 1.5 * iqr),
                "upper_whisker": max(v for v in values if v <= q3 + 1.5 * iqr),
            }
    return by_year


def build_f42(rows, sheet):
    data = [
        box_trace(rows, "E", "System USE", "#44b9a9", "rgba(68,185,169,0.38)"),
        box_trace(rows, "D", "Sub-regional USE", "#7d87bd", "rgba(125,135,189,0.4)"),
    ]
    early_rows = [row for row in rows if row[1]["B"] in YEARS[:5]]
    for column, name, color, fill in [
        ("E", "System USE", "#44b9a9", "rgba(68,185,169,0.38)"),
        ("D", "Sub-regional USE", "#7d87bd", "rgba(125,135,189,0.4)"),
    ]:
        trace = box_trace(early_rows, column, name, color, fill, xaxis="x2", yaxis="y2")
        trace["showlegend"] = False
        data.append(trace)
    layout = base_layout(650)
    layout.update({
        "xaxis": axis("Capacity year", type="category", categoryorder="array", categoryarray=YEARS, tickangle=-35, showgrid=False),
        "yaxis": axis("Unserved energy (% of annual consumption)", range=[-0.025, 2], tickformat=".3f", ticksuffix="%", dtick=0.25),
        "xaxis2": axis("", domain=[0.065, 0.46], anchor="y2", type="category", categoryorder="array", categoryarray=YEARS[:5], tickangle=-30, showgrid=False, tickfont={"size": 10}),
        "yaxis2": axis("", domain=[0.34, 0.79], anchor="x2", range=[-0.0002, 0.04], tickformat=".3f", ticksuffix="%", dtick=0.01, tickfont={"size": 10}, zeroline=False),
        "shapes": [
            threshold_line(),
            {"type": "rect", "xref": "x2 domain", "yref": "y2 domain", "x0": 0, "x1": 1, "y0": 0, "y1": 1, "fillcolor": "#ffffff", "line": {"color": "#aab8c0", "width": 1}, "layer": "below"},
            threshold_line(xref="x2 domain", yref="y2"),
        ],
        "annotations": [{"xref": "paper", "yref": "y", "x": 1, "y": THRESHOLD, "text": "0.0002% reliability standard", "showarrow": False, "xanchor": "right", "yanchor": "bottom", "yshift": 6, "font": {"color": "#ad2353", "size": 11}}],
    })
    return {
        "id": "supplement-ch6-f42",
        "sheet": sheet["name"],
        "title": sheet["title"].strip(),
        "section": "Chapter 6",
        "kind": "interactive",
        "image_file": "data/analysis/images/image18.png",
        "data": data,
        "layout": layout,
        "source_file": WORKBOOK.name,
        "source_series_count": 2,
        "sources": [
            {"series_name": name, "x": f"'{sheet['name']}'!$B$24:$B$17242", "y": f"'{sheet['name']}'!${column}$24:${column}$17242"}
            for column, name in [("E", "System USE"), ("D", "Sub-regional USE")]
        ],
        "notes": [
            "Recreated from all 17,219 modelled results per series in the worksheet. Values already use percentage points; 1.0 means 1% of annual consumption.",
            "Boxes show the 25th–75th percentiles, with the median and dashed mean; whiskers extend to the most extreme result within 1.5 times the interquartile range. Quartiles use linear interpolation. Outlier points are hidden as in the source figure, but all results are retained in the chart data.",
            "The inset repeats 2026-27 to 2030-31 on a finer vertical scale. The dashed reference line is the 0.0002% reliability standard.",
        ],
    }


def build_f43(rows, sheet):
    regions = [
        ("D", "North", "#4a9f91", "rgba(74,159,145,0.3)"),
        ("E", "Metro North", "#72a168", "rgba(114,161,104,0.35)"),
        ("F", "South", "#669db0", "rgba(102,157,176,0.32)"),
        ("G", "Metro South", "#e78f7b", "rgba(231,143,123,0.4)"),
        ("H", "East", "#9380aa", "rgba(147,128,170,0.4)"),
    ]
    domains = [([0, 0.455], [0.72, 0.96]), ([0.545, 1], [0.72, 0.96]), ([0, 0.455], [0.38, 0.62]), ([0.545, 1], [0.38, 0.62]), ([0, 0.455], [0.04, 0.28])]
    data = []
    layout = base_layout(900)
    layout["boxmode"] = "overlay"
    layout["margin"] = {"l": 80, "r": 30, "t": 25, "b": 70}
    layout["annotations"] = []
    layout["shapes"] = []
    for i, ((column, name, color, fill), (xdomain, ydomain)) in enumerate(zip(regions, domains), start=1):
        suffix = "" if i == 1 else str(i)
        xref, yref = f"x{suffix}", f"y{suffix}"
        trace = box_trace(rows, column, name, color, fill, points="outliers", xaxis=xref, yaxis=yref)
        trace["showlegend"] = False
        data.append(trace)
        layout[f"xaxis{suffix}"] = axis("", domain=xdomain, anchor=yref, type="category", categoryorder="array", categoryarray=YEARS, tickangle=-38, showgrid=False, tickfont={"size": 10})
        layout[f"yaxis{suffix}"] = axis("", domain=ydomain, anchor=xref, range=[-0.002, 0.31], dtick=0.1, tickformat=".3f", ticksuffix="%", hoverformat=".6f", zeroline=False)
        layout["annotations"].append({"xref": "paper", "yref": "paper", "x": sum(xdomain) / 2, "y": ydomain[1], "yshift": 12, "text": name, "showarrow": False, "font": {"size": 15}, "xanchor": "center", "yanchor": "bottom"})
        layout["shapes"].append(threshold_line(xref=f"{xref} domain", yref=yref))
    layout["annotations"].extend([
        {"xref": "paper", "yref": "paper", "x": 0, "y": 0.5, "xshift": -64, "text": "Unserved energy (% of annual consumption)", "textangle": -90, "showarrow": False, "xanchor": "center", "font": {"size": 12}},
        {"xref": "paper", "yref": "paper", "x": 0.77, "y": 0.16, "text": "Dashed reference: 0.0002%<br>Each panel uses the same vertical scale", "showarrow": False, "font": {"size": 11, "color": "#566c78"}},
    ])
    return {
        "id": "supplement-ch6-f43",
        "sheet": sheet["name"],
        "title": sheet["title"].strip(),
        "section": "Chapter 6",
        "kind": "interactive",
        "image_file": "data/analysis/images/image19.png",
        "data": data,
        "layout": layout,
        "source_file": WORKBOOK.name,
        "source_series_count": 5,
        "sources": [
            {"series_name": name, "x": f"'{sheet['name']}'!$B$34:$B$17252", "y": f"'{sheet['name']}'!${column}$34:${column}$17252"}
            for column, name, _, _ in regions
        ],
        "notes": [
            "Recreated from all 17,219 modelled results for each of the five sub-regions. All panels use the same vertical scale; zoom to inspect small values. Numeric values already use percentage points.",
            "Boxes show the 25th–75th percentiles, with the median and dashed mean. Whiskers extend to the most extreme result within 1.5 times the interquartile range; points beyond the whiskers are shown. Quartiles use linear interpolation.",
            "The worksheet header ‘Metor South’ is displayed as ‘Metro South’, matching the source figure. Cell G1119 is stored as the text ‘1.28664568327889E-05%’; its numeric percentage-point value is retained, and verified against the corresponding sub-regional total in Figure 42.",
        ],
    }


def main():
    inventory = json.loads((ROOT / "data/analysis/inventory.json").read_text())
    sheets = {sheet["name"]: sheet for sheet in inventory["sheets"]}
    with zipfile.ZipFile(WORKBOOK) as archive:
        f42_rows = worksheet_rows(archive, sheets["Ch 6_F.42"]["part"], 24)
        f43_rows = worksheet_rows(archive, sheets["Ch 6_F.43"]["part"], 34)
    assert len(f42_rows) == len(f43_rows) == 17219
    largest_error = 0
    for (f42_row, f42), (f43_row, f43) in zip(f42_rows, f43_rows):
        assert f43_row == f42_row + 10 and f42["B"] == f43["B"]
        total = sum(numeric_percent(f43[col]) for col in "DEFGH")
        error = abs(total - numeric_percent(f42["D"]))
        largest_error = max(largest_error, error)
        assert math.isclose(total, f42["D"], abs_tol=1e-12, rel_tol=1e-10)
    specs = [build_f42(f42_rows, sheets["Ch 6_F.42"]), build_f43(f43_rows, sheets["Ch 6_F.43"])]
    output = ROOT / "src/charts/supplemental_network.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(specs, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n")
    review = {
        "source": str(WORKBOOK.relative_to(ROOT)),
        "figures": [42, 43],
        "rows_per_figure": len(f42_rows),
        "samples_per_year": dict(Counter(values["B"] for _, values in f42_rows)),
        "unit": "percentage points; values are not multiplied by 100",
        "method": "Full worksheet samples; Plotly linear quartiles use zero-based rank n*p - 0.5 with linear interpolation (Langford method 10); Tukey 1.5-IQR whiskers; median and arithmetic mean. F42 hides outlier points as the source image does; F43 shows them.",
        "normalizations": [{"cell": "'Ch 6_F.43'!G1119", "source": "1.28664568327889E-05%", "numeric_percentage_points": 1.28664568327889e-05, "evidence": "Sum of the five regional values matches 'Ch 6_F.42'!D1109."}],
        "regional_sum_check": {"rows_checked": len(f42_rows), "largest_absolute_difference": largest_error, "pass": True},
        "figure42_summary": summary(f42_rows, {"D": "Sub-regional USE", "E": "System USE"}),
        "figure43_summary": summary(f43_rows, {"D": "North", "E": "Metro North", "F": "South", "G": "Metro South", "H": "East"}),
    }
    (ROOT / "data/analysis/supplemental-network-review.json").write_text(json.dumps(review, ensure_ascii=False, indent=2, allow_nan=False) + "\n")
    print(f"Wrote {len(specs)} network figures to {output.relative_to(ROOT)} ({output.stat().st_size:,} bytes)")
    print(f"Verified all {len(f42_rows):,} regional sums; maximum floating-point difference {largest_error:.3g}.")


if __name__ == "__main__":
    main()
