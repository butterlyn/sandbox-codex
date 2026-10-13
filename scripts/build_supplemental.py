#!/usr/bin/env python3
"""Rebuild image-backed register figures from their underlying worksheet cells.

These are separate from native Excel chart objects. All numerical points are
retained; maps, diagrams and duplicate overview images are labelled references.
Run from any working directory: python scripts/build_supplemental.py
"""

from __future__ import annotations

import datetime as dt
import json
import math
from pathlib import Path
import re
import statistics
import struct
import xml.etree.ElementTree as ET
import zipfile


ROOT = Path(__file__).resolve().parents[1]
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
XLSX = ROOT / "data/sources/2026-wem-esoo-data-register.xlsx"
INVENTORY = json.loads((ROOT / "data/analysis/inventory.json").read_text())
SHEETS = {s["name"]: s for s in INVENTORY["sheets"]}
WORKBOOK_URL = "https://www.aemo.com.au/-/media/files/electricity/wem/planning_and_forecasting/esoo/2026/2026-wem-esoo-data-register.xlsx?rev=26835913281149c6bb2c55eb7556d8de&sc_lang=en"
REPORT_URL = "https://www.aemo.com.au/-/media/files/electricity/wem/planning_and_forecasting/esoo/2026/2026-wem-electricity-statement-of-opportunities.pdf?rev=3aa69153c92f4437ac7032402a261f9c&sc_lang=en"


def column_number(name: str) -> int:
    n = 0
    for letter in name:
        n = n * 26 + ord(letter) - 64
    return n


class Workbook:
    def __init__(self):
        self.zip = zipfile.ZipFile(XLSX)
        self.strings = [
            "".join(node.text or "" for node in si.iter(f"{{{NS['s']}}}t"))
            for si in ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        ]

    def cells(self, sheet: str) -> dict[str, str | float]:
        """Use OOXML cached values, including all rows of the availability curve."""
        result = {}
        with self.zip.open(SHEETS[sheet]["part"]) as stream:
            for _, node in ET.iterparse(stream, events=("end",)):
                if node.tag == f"{{{NS['s']}}}c":
                    value = node.find("s:v", NS)
                    if value is not None and value.text is not None:
                        if node.get("t") == "s":
                            result[node.get("r")] = self.strings[int(value.text)]
                        elif node.get("t") != "e":
                            result[node.get("r")] = float(value.text)
                    node.clear()
                elif node.tag == f"{{{NS['s']}}}row":
                    node.clear()
        return result


def section(sheet):
    return "Appendices" if sheet.startswith("A") else "Chapter " + re.search(r"Ch (\d+)", sheet)[1]


def source(sheet, cell_range, role):
    return {"sheet": sheet, "range": cell_range, "role": role}


def base(sheet, suffix=None):
    stem = re.sub(r"[^a-z0-9]+", "-", sheet.lower()).strip("-")
    # Keep the chapter identifier readable and consistent with other workers.
    stem = stem.replace("ch-", "ch").replace("-f-", "-f")
    images = [i["file"] for i in INVENTORY["images"] if i["sheet"] == sheet and not i["file"].endswith("/image1.png")]
    return {
        "id": "supplement-" + stem + ("-" + suffix if suffix else ""),
        "sheet": sheet,
        "title": SHEETS[sheet]["title"].strip(),
        "section": section(sheet),
        "kind": "interactive",
        "source_format": "worksheet data behind image",
        "data": [],
        "layout": {},
        "notes": [],
        "sources": [],
        "image_file": images[0],
        "publication": "AEMO, 2026 WEM Electricity Statement of Opportunities, data register (version 2)",
        "source_url": WORKBOOK_URL,
    }


def values(cells, col, first, last):
    result = [cells[f"{col}{i}"] for i in range(first, last + 1)]
    assert all(isinstance(v, (float, int)) and math.isfinite(v) for v in result)
    return result


def build_capacity(workbook):
    sheet = "Ch 3_F.21"
    cells = workbook.cells(sheet)
    spec = base(sheet)
    years = ["Committed capacity", "Anticipated capacity", "Proposed capacity"]
    other_cols = ["H", "I", "J", "K", "L", "M"]
    groups = [("ESR", ["D"], "#9f50a0"), ("Wind", ["E"], "#a7d77e"),
              ("Gas", ["F"], "#39bdaa"), ("Solar", ["G"], "#ffbd24"),
              ("Other", other_cols, "#6f6879")]
    for name, cols, color in groups:
        spec["data"].append({
            "type": "bar", "name": name, "orientation": "h", "y": years,
            "x": [sum(cells[f"{col}{row}"] for col in cols) for row in range(24, 27)],
            "marker": {"color": color},
            "hovertemplate": "%{y}<br>" + name + ": %{x:.4%}<extra></extra>",
        })
    for row in range(24, 27):
        assert abs(sum(cells[f"{col}{row}"] for col in "DEFGHIJKLM") - 1) < 1e-12
    spec["layout"] = {
        "barmode": "stack", "height": 440,
        "xaxis": {"title": {"text": "Share of new available capacity in 2035–36"}, "tickformat": ".0%", "range": [0, 1]},
        "yaxis": {"autorange": "reversed", "automargin": True},
        "legend": {"orientation": "h", "y": -0.2},
        "margin": {"l": 165, "r": 20, "t": 25, "b": 100},
    }
    spec["notes"] = [
        "The workbook infographic is recreated as a 100% stacked bar chart using its unrounded worksheet percentages.",
        "Other combines waste, coal, demand-side participation, distillate, hydro and ESR/DSP, matching the infographic's grouping. Its exact committed share is 8.426%, rather than the image's rounded 9% label.",
        "Committed, anticipated and proposed are project development categories; these proportions do not describe the whole installed fleet.",
    ]
    spec["sources"] = [source(sheet, "B23:N26", "Technology shares and total by project development category")]
    return spec, {"numeric_source_cells": 33, "plotted_points": 15, "transformation": "Exact grouped sums; fraction values retained for percentage formatting."}


def build_unserved_boxplot(workbook):
    sheet = "Ch 5_F.33"
    cells = workbook.cells(sheet)
    spec = base(sheet)
    years = [cells[f"{col}24"] for col in "CDEFGHIJKL"]
    summary = []
    for i, col in enumerate("CDEFGHIJKL"):
        y = values(cells, col, 25, 1624)
        trace = {"type": "box", "name": years[i], "y": y, "boxpoints": "all", "jitter": 0.35,
                 "pointpos": 0, "boxmean": True, "quartilemethod": "linear", "showlegend": False,
                 "marker": {"color": "rgba(118,118,118,0.18)", "size": 3},
                 "line": {"color": "#743480", "width": 1.3}, "fillcolor": "rgba(137,78,151,0.40)",
                 "hovertemplate": years[i] + "<br>Unserved energy: %{y:.7f}%<extra></extra>"}
        spec["data"].append(trace)
        if i < 5:
            # Separate zoom panel avoids the original inset hiding plotted data.
            zoom = dict(trace)
            zoom.update({"xaxis": "x2", "yaxis": "y2", "legendgroup": years[i]})
            spec["data"].append(zoom)
        summary.append({"capacity_year": years[i], "samples": len(y), "min": min(y), "max": max(y), "mean": statistics.mean(y)})
    spec["layout"] = {
        "height": 700, "showlegend": False,
        "xaxis": {"title": {"text": "Capacity year"}, "type": "category", "categoryorder": "array", "categoryarray": years,
                  "anchor": "y", "tickangle": -28},
        "yaxis": {"title": {"text": "Unserved energy (% of annual consumption)"}, "domain": [0, 0.60], "range": [-0.035, 1.5]},
        "xaxis2": {"type": "category", "categoryorder": "array", "categoryarray": years[:5], "domain": [0, 0.60], "anchor": "y2", "tickangle": -20},
        "yaxis2": {"title": {"text": "Early years (%)"}, "domain": [0.78, 1], "range": [0, 0.002], "tickformat": ".4f", "anchor": "x2"},
        "shapes": [
            {"type": "line", "xref": "x domain", "x0": 0, "x1": 1, "yref": "y", "y0": 0.0002, "y1": 0.0002, "line": {"color": "#a61f44", "width": 2, "dash": "dash"}},
            {"type": "line", "xref": "x2 domain", "x0": 0, "x1": 1, "yref": "y2", "y0": 0.0002, "y1": 0.0002, "line": {"color": "#a61f44", "width": 2, "dash": "dash"}},
        ],
        "annotations": [{"xref": "paper", "yref": "paper", "x": 0.98, "y": 0.97, "xanchor": "right", "yanchor": "top", "showarrow": False,
                         "align": "left", "text": "1,600 simulations per capacity year<br>Red dashed line: 0.0002% reliability threshold<br>Boxes: middle 50%; dashed centre: mean<br>Whiskers: observations within 1.5 × IQR", "font": {"size": 12}}],
        "margin": {"l": 85, "r": 25, "t": 30, "b": 90},
    }
    spec["notes"] = [
        "All 16,000 simulation results are retained: 16 weather reference years × 100 Monte Carlo forced-outage iterations for each of 10 capacity years.",
        "The original inset is shown as a separate early-years zoom panel. Both panels use the same source results; the red line is the 0.0002% annual-consumption reliability threshold.",
        "Values are percentage points, as shown by the figure title and y-axis. The workbook's explanatory note says '(MWh)' but its supplied values and chart are percentages.",
        "Boxes show the interquartile range using linear quantiles; whiskers extend to observations within 1.5 × IQR. Every observation is shown and the dashed box centre indicates its mean.",
    ]
    spec["sources"] = [source(sheet, "C24:L1624", "Capacity-year labels and all 1,600 sorted results for each year"), source(sheet, "B19", "Box-and-whisker methodology note")]
    return spec, {"numeric_source_cells": 16000, "plotted_points": 24000, "unique_observations": 16000, "summary": summary,
                  "transformation": "Unscaled percentages, linear quartiles; early five years repeated only in the zoom panel."}


COLORS = {
    "Coal": "#363b37", "Wind": "#a5d779", "Solar": "#ffb914", "Distillate": "#eb6654", "Gas": "#3bbcab",
    "Waste": "#dce780", "ESR (discharge)": "#9e50a1", "Demand-side participation": "#6268b0",
    "Unserved energy": "#a21b42", "Distributed PV": "#f9df85", "Headroom": "#f2d9f1", "Outages": "#cecece",
}


def excel_datetime(date, time):
    # Column B already advances the calendar date. Column C is elapsed days,
    # formatted as a clock time: use its fractional part to avoid double-counting.
    return (dt.datetime(1899, 12, 30) + dt.timedelta(seconds=round((date + time % 1) * 86400))).isoformat()


def build_dispatch(workbook, sheet):
    cells = workbook.cells(sheet)
    spec = base(sheet)
    first, last = 28, 124
    x = [excel_datetime(cells[f"B{i}"], cells[f"C{i}"]) for i in range(first, last + 1)]
    assert all((dt.datetime.fromisoformat(b) - dt.datetime.fromisoformat(a)).total_seconds() == 1800 for a, b in zip(x, x[1:]))
    for col in "DEFGHIJKLMNO":
        name = cells[f"{col}27"]
        if name == "Unserved Energy":
            name = "Unserved energy"
        y = values(cells, col, first, last)
        # Coal is identically zero in Figure 35 and omitted in that image's legend.
        if name == "Coal" and not any(y):
            continue
        spec["data"].append({
            "type": "scatter", "mode": "lines", "name": name, "x": x, "y": y, "stackgroup": "supply",
            "line": {"color": COLORS[name], "width": 0.6}, "fillcolor": COLORS[name],
            "hovertemplate": "%{x|%d %b %Y %H:%M}<br>" + name + ": %{y:,.2f} MW<extra></extra>",
        })
    spec["data"].append({
        "type": "scatter", "mode": "lines", "name": "ESR (charge)", "x": x,
        "y": [-v for v in values(cells, "R", first, last)], "fill": "tozeroy", "fillcolor": "rgba(157,77,160,0.53)",
        "line": {"color": "#a457a6", "width": 1},
        "hovertemplate": "%{x|%d %b %Y %H:%M}<br>ESR charging (plotted below zero): %{y:,.2f} MW<extra></extra>",
    })
    for col, name, dash in [("P", "Unscheduled operational demand", "solid"), ("Q", "Operational demand", "dashdot")]:
        spec["data"].append({
            "type": "scatter", "mode": "lines", "name": name, "x": x, "y": values(cells, col, first, last),
            "line": {"color": "#101216", "width": 2.5, "dash": dash},
            "hovertemplate": "%{x|%d %b %Y %H:%M}<br>" + name + ": %{y:,.2f} MW<extra></extra>",
        })
    spec["data"].append({
        "type": "scatter", "mode": "lines", "name": "ESR state of charge (%)", "x": x, "y": values(cells, "S", first, last),
        "yaxis": "y2", "line": {"color": "#a455a7", "width": 2.3, "dash": "dash"},
        "hovertemplate": "%{x|%d %b %Y %H:%M}<br>ESR state of charge: %{y:.2f}%<extra></extra>",
    })
    spec["layout"] = {
        "height": 670, "hovermode": "x unified",
        "xaxis": {"type": "date", "title": {"text": "Modelled local date and time"}, "tickformat": "%d %b<br>%H:%M", "dtick": 21600000, "range": [x[0], x[-1]]},
        "yaxis": {"title": {"text": "Generation, ESR and demand (MW)"}, "range": [-2000, 8000], "tickformat": ",.0f"},
        "yaxis2": {"title": {"text": "ESR state of charge (%)"}, "range": [-25, 100], "tickvals": [0, 25, 50, 75, 100],
                   "overlaying": "y", "side": "right", "showgrid": False, "zeroline": False},
        "legend": {"orientation": "h", "y": -0.22, "font": {"size": 11}},
        "margin": {"l": 75, "r": 75, "t": 25, "b": 195},
    }
    spec["notes"] = [
        "All 97 half-hour samples are retained, including the endpoint of the two-day period. Dates and times are decoded directly from the worksheet's Excel serial values.",
        "Generation, distributed PV, available headroom and outages are stacked above zero as in the original. Charging is shown below zero; its source values are positive and are sign-inverted only for display.",
        "ESR state of charge uses the right-hand percentage axis, aligned at zero with the MW axis. Demand uses the left-hand MW axis.",
    ]
    if sheet == "Ch 5_F.34":
        spec["notes"].append(str(cells["B26"]).removeprefix("Note: "))
        spec["notes"].append("Report context: the first day has generation headroom and batteries charge ahead of the event. On the second evening headroom is exhausted, DSP is unavailable on Sunday, and battery injection is insufficient to avoid unserved energy.")
    else:
        spec["notes"].append("Report context: low winter temperatures coincide with low wind. Batteries charge and discharge on the first day, but low overnight renewables and limited thermal headroom prevent sufficient recharge before the second day's morning and evening peaks.")
    spec["sources"] = [source(sheet, "B27:S124", "Half-hour dispatch, generation, demand, storage charging and state of charge"),
                       {"url": REPORT_URL, "role": "Report section 5.4.1, interpretation of the sample event"}]
    return spec, {"numeric_source_cells": 1746, "unique_time_samples": 97, "plotted_points": sum(len(t["y"]) for t in spec["data"]),
                  "start": x[0], "end": x[-1], "transformation": "Decode B+fractional_part(C) Excel dates; negate column R for below-zero charging; all other values unchanged."}


def build_daily_distribution(workbook):
    sheet = "Ch 5_F.36"
    cells = workbook.cells(sheet)
    spec = base(sheet)
    time = [round(cells[f"B{r}"] * 24, 8) for r in range(27, 75)]
    assert time == [i / 2 for i in range(48)]
    seasons = ["Shoulder", "Summer", "Winter"]
    colors = ["#ffba12", "#a82046", "#36bfae"]
    domains = [([0, 0.45], [0.59, 1]), ([0.55, 1], [0.59, 1]), ([0, 0.45], [0, 0.41]), ([0.55, 1], [0, 0.41])]
    annotations = []
    for panel, first_col in enumerate(["C", "F", "I", "L"]):
        year = cells[f"{first_col}25"]
        suffix = str(panel + 1) if panel else ""
        xref, yref = "x" + suffix, "y" + suffix
        for j, season in enumerate(seasons):
            col = chr(ord(first_col) + j)
            assert cells[f"{col}26"] == season
            spec["data"].append({
                "type": "scatter", "mode": "lines", "name": season, "legendgroup": season, "showlegend": panel == 0,
                "x": time, "y": values(cells, col, 27, 74), "xaxis": xref, "yaxis": yref,
                "line": {"color": colors[j], "width": 2.4},
                "customdata": [f"{int(h):02d}:{'30' if h % 1 else '00'}" for h in time],
                "hovertemplate": year + " · " + season + "<br>%{customdata}<br>Average USE: %{y:,.5f} MWh<extra></extra>",
            })
        xd, yd = domains[panel]
        spec["layout"]["xaxis" + suffix] = {"domain": xd, "anchor": yref, "tickvals": [0, 4, 8, 12, 16, 20],
                                                  "ticktext": ["00:00", "04:00", "08:00", "12:00", "16:00", "20:00"],
                                                  "range": [0, 23.5], "title": {"text": "Time"} if panel > 1 else {}}
        spec["layout"]["yaxis" + suffix] = {"domain": yd, "anchor": xref, "range": [0, 150], "dtick": 40,
                                                  "title": {"text": "Average USE (MWh)"} if panel % 2 == 0 else {}}
        annotations.append({"xref": "paper", "yref": "paper", "x": sum(xd) / 2, "y": yd[1] + 0.035,
                            "text": year, "showarrow": False, "font": {"size": 14}, "xanchor": "center"})
    spec["layout"].update({"height": 660, "annotations": annotations, "legend": {"orientation": "h", "y": -0.15, "groupclick": "togglegroup"},
                           "margin": {"l": 75, "r": 25, "t": 35, "b": 100}})
    spec["notes"] = ["All 48 half-hour trading intervals for each season and capacity year are retained. All four panels share the same MWh scale; legend controls toggle a season across all panels.",
                     "Shoulder, summer and winter are the workbook's season categories. Values are the published average unserved energy (MWh) for each interval."]
    spec["sources"] = [source(sheet, "B25:N74", "Capacity years, season labels and half-hour average unserved-energy profiles")]
    return spec, {"numeric_source_cells": 624, "plotted_points": 576, "series_count": 12, "transformation": "Trading interval fractions converted exactly to half-hour clock times; energy values unchanged."}


def build_availability(workbook):
    sheet = "Ch 5_F.40"
    cells = workbook.cells(sheet)
    spec = base(sheet)
    first, last = 27, 280346
    summary = []
    for demand_col, margin_col, color in [("B", "C", "#ac2046"), ("D", "E", "#753b86")]:
        year = cells[f"{demand_col}24"]
        demand = values(cells, demand_col, first, last)
        allowance = values(cells, margin_col, first, last)
        y = [d + m for d, m in zip(demand, allowance)]
        assert len(y) == 280320
        assert all(a >= b for a, b in zip(y, y[1:]))
        assert len(set(allowance)) == 1
        spec["data"].append({
            "type": "scattergl", "mode": "lines", "name": year, "x0": 1, "dx": 1, "y": y,
            "line": {"color": color, "width": 2}, "fill": "tozeroy", "fillcolor": "rgba(139,79,116,0.20)",
            "hovertemplate": year + "<br>Descending trading-interval rank: %{x:,}<br>Minimum capacity requirement: %{y:,.3f} MW<br>Reserve and other allowances: " + f"{allowance[0]:,.0f}" + " MW<extra></extra>",
        })
        summary.append({"capacity_year": year, "samples": len(y), "allowance_MW": allowance[0], "minimum_MW": min(y), "maximum_MW": max(y)})
    spec["layout"] = {
        "height": 510, "hovermode": "closest",
        "xaxis": {"title": {"text": "Trading intervals ranked by descending demand"}, "range": [1, 280320], "tickformat": ",.0f"},
        "yaxis": {"title": {"text": "Minimum capacity requirement (MW)"}, "rangemode": "tozero", "tickformat": ",.0f"},
        "legend": {"orientation": "h", "y": -0.2}, "margin": {"l": 75, "r": 25, "t": 25, "b": 100},
    }
    spec["notes"] = [
        "Every one of the 280,320 ranked half-hour source observations is retained for each capacity year (560,640 plotted points in total). No sampling or curve simplification is applied.",
        "Each curve is the worksheet demand plus reserve margin, Intermittent Load allowance, Regulation Raise and ESR Duration Requirement uplift: 1,243 MW in 2027–28 and 1,386 MW in 2028–29.",
        "The horizontal axis is the rank across the published reference-year observations, not a date or consecutive operational chronology. WebGL renders the full dataset; hover or zoom for the exact local value.",
    ]
    spec["sources"] = [source(sheet, "B24:E26", "Capacity-year labels, allowance definition and MW units"),
                       source(sheet, "B27:C280346", "2027–28 demand plus reserve and other allowances"),
                       source(sheet, "D27:E280346", "2028–29 demand plus reserve and other allowances"),
                       {"url": REPORT_URL, "role": "Report section 5.9, Availability Curves definition"}]
    return spec, {"numeric_source_cells": 1121280, "plotted_points": 560640, "summary": summary,
                  "transformation": "Exact row-wise sums B+C and D+E; rank starts at 1, all 280,320 samples per trace; constant allowance columns not duplicated in plot arrays."}


REFERENCE_INFO = {
    "Ch 1_F.3": ("diagram", "The Planning Criterion's three limbs and their relationship to Peak and Flexible Reserve Capacity Targets."),
    "Ch 1_F.4": ("diagram", "How the WEM ESOO connects government planning, AEMO forecasting, Western Power network planning and investor decisions."),
    "Ch 1_F.5": ("diagram", "Qualitative comparison of Slower Growth, Step Change and Accelerated Transition scenario assumptions; no numerical radar-axis values are supplied."),
    "Ch 1_F.6": ("diagram", "Relationship between generation, auxiliary loads, network losses and delivered/underlying consumption and demand."),
    "Ch 4_F.30": ("diagram", "Relationship between Mid Peak ESROI, ESR Duration Requirement and the contiguous Peak ESROD period."),
    "Ch 6_F.41": ("map", "Map identifying the SWIS network sub-regions."),
    "Ch 6_F.44": ("map", "Transmission network and facilities in the East sub-region."),
    "Ch 6_F.45": ("map", "Transmission network and facilities in the Kwinana area."),
    "Ch 6_F.46": ("map", "Transmission network and facilities in the North sub-region."),
    "Ch 7_F.48": ("map", "Geographic capacity graphic showing the technology mix and declared sent-out capacity of proposed projects across North, East and South SWIS regions."),
    "A1_F.12a": ("duplicate overview", "Overview panel also supplied as a separate native Excel chart on the following worksheet."),
    "A1_F.19": ("duplicate overview", "Overview illustrating equal weighting of project-based and economic-based data-centre forecasts. All three component plots have separate interactive recreations."),
}


def build_references(workbook):
    specs, reviews = [], []
    panel_map = {
        "image24.png": ("temperature-response", "Household electricity consumption response to temperature, 2020–21 to 2024–25 (kWh)", "A1_F.12b"),
        "image26.png": ("base-load", "Average annual household base load consumption, 2020–21 to 2035–36 (kWh)", "A1_F.12c"),
        "image28.png": ("cooling-response", "Average annual household cooling response consumption, 2020–21 to 2035–36 (kWh/CDD)", "A1_F.12d"),
        "image30.png": ("heating-response", "Average annual household heating response consumption, 2020–21 to 2035–36 (kWh/HDD)", "A1_F.12e"),
    }
    for sheet, (reference_type, description) in REFERENCE_INFO.items():
        cells = workbook.cells(sheet)
        numeric = {k: v for k, v in cells.items() if isinstance(v, (int, float))}
        assert not numeric, (sheet, numeric)
        image_records = [i for i in INVENTORY["images"] if i["sheet"] == sheet and not i["file"].endswith("/image1.png")]
        for record in image_records:
            image_name = Path(record["file"]).name
            panel = panel_map.get(image_name)
            spec = base(sheet, panel[0] if panel else None)
            spec.update({"kind": "reference", "source_format": "image only", "reference_type": reference_type,
                         "image_file": record["file"], "description": description})
            if panel:
                spec["title"] = "Figure 12 overview — " + panel[1]
                spec["related_sheets"] = [panel[2]]
            if sheet == "A1_F.19":
                spec["related_sheets"] = ["A1_F.19b", "A1_F.19c", "A1_F.19d"]
            with (ROOT / record["file"]).open("rb") as f:
                f.read(16)
                width, height = struct.unpack(">II", f.read(8))
            spec["image_width"], spec["image_height"] = width, height
            spec["notes"] = ["Original AEMO figure retained as a zoomable reference. This worksheet contains the image and text labels only; it supplies no numerical series or geographic coordinates to recreate interactively."]
            if panel:
                spec["notes"] = ["Original overview panel. The same plot is recreated interactively from its numerical worksheet on " + panel[2] + "."]
            if sheet == "A1_F.19":
                spec["notes"] = ["Original methodology overview retained to show the equal-weight blending relationship. The project-based, economic-based and final forecast plots are recreated interactively from A1_F.19b, A1_F.19c and A1_F.19d."]
            for address, value in cells.items():
                if isinstance(value, str) and value.strip().lower().startswith(("note:", "notes:", "source:")):
                    spec["notes"].append(value.strip())
            spec["sources"] = [{"sheet": sheet, "range": "embedded image", "role": reference_type, "image_part": record["part"]}]
            specs.append(spec)
            reviews.append({"id": spec["id"], "sheet": sheet, "numeric_source_cells": 0, "kind": "reference", "reference_type": reference_type,
                            "image_file": record["file"], "dimensions": [width, height], "related_sheets": spec.get("related_sheets", [])})
    return specs, reviews


def main():
    workbook = Workbook()
    specs, review = [], []
    builders = [build_capacity, build_unserved_boxplot,
                lambda w: build_dispatch(w, "Ch 5_F.34"), lambda w: build_dispatch(w, "Ch 5_F.35"),
                build_daily_distribution, build_availability]
    for builder in builders:
        spec, audit = builder(workbook)
        specs.append(spec)
        review.append({"id": spec["id"], "sheet": spec["sheet"], "kind": spec["kind"], "series_count": len(spec["data"]), **audit})
        print(f"Built {spec['sheet']}: {len(spec['data'])} traces, {audit['plotted_points']:,} plotted points")
    references, reference_review = build_references(workbook)
    specs.extend(references)
    review.extend(reference_review)
    sheet_order = {s["name"]: i for i, s in enumerate(INVENTORY["sheets"])}
    specs.sort(key=lambda s: (sheet_order[s["sheet"]], s["id"]))
    assert len({s["id"] for s in specs}) == len(specs)
    output = ROOT / "src/charts/supplemental.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(specs, separators=(",", ":"), ensure_ascii=False, allow_nan=False) + "\n")
    review_path = ROOT / "data/analysis/supplemental-review.json"
    review_path.write_text(json.dumps({
        "coverage": {"interactive_figures": 6, "reference_images": len(references), "reference_sheets": len(REFERENCE_INFO),
                     "delegated_numeric_sheets": ["Ch 6_F.42", "Ch 6_F.43", "A4_F.25", "A4_F.26", "A4_F.27", "A4_F.28"]},
        "source_workbook": str(XLSX.relative_to(ROOT)), "items": review,
    }, indent=2, ensure_ascii=False, allow_nan=False) + "\n")
    print(f"Wrote {len(specs)} specs ({output.stat().st_size:,} bytes) and {review_path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
