"""Read Excel's saved OOXML chart caches and reproduce them as Plotly specs.

The source register is never recalculated.  ``load_chart`` exposes untouched
numeric caches, formulas, axes and drawing styles; ``build_chart`` translates
those into browser-ready traces.  Missing points stay None.  Explicit Excel
display-unit scaling and stacked-line accumulation are recorded in trace.meta.
"""
from __future__ import annotations

import colorsys
import copy
import datetime as dt
import json
import math
import re
import zipfile
from functools import lru_cache
from pathlib import Path

from lxml import etree as ET
from openpyxl.utils.cell import range_boundaries, get_column_letter

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "data/sources/2026-wem-esoo-data-register.xlsx"
NS = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
PALETTE = ["#003B5C", "#40C1AC", "#E56A54", "#F2A900", "#8A1B61", "#009CDE", "#A4DBE8", "#373A36"]
UNITS = {"hundreds": 100, "thousands": 1000, "tenThousands": 10000,
         "hundredThousands": 100000, "millions": 1e6, "tenMillions": 1e7,
         "hundredMillions": 1e8, "billions": 1e9, "trillions": 1e12}


def val(node, path, default=None):
    found = node.find(path, NS) if node is not None else None
    return found.get("val", default) if found is not None else default


def text_at(node, path, default=None):
    found = node.find(path, NS) if node is not None else None
    return found.text if found is not None else default


def number(value):
    if value is None or str(value).strip() in ("", "#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NUM!", "#NAME?"):
        return None
    try:
        n = float(value)
        return n if math.isfinite(n) else None
    except (TypeError, ValueError):
        return None


@lru_cache(maxsize=1)
def inventory():
    return json.loads((ROOT / "data/analysis/inventory.json").read_text())


@lru_cache(maxsize=1)
def archive():
    return zipfile.ZipFile(WORKBOOK)


@lru_cache(maxsize=1)
def shared_strings():
    if "xl/sharedStrings.xml" not in archive().namelist():
        return []
    root = ET.fromstring(archive().read("xl/sharedStrings.xml"))
    return ["".join(n.itertext()) for n in root.findall("s:si", NS)]


@lru_cache(maxsize=1)
def sheet_parts():
    root = ET.fromstring(archive().read("xl/workbook.xml"))
    rels = ET.fromstring(archive().read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    result = {}
    for sheet in root.findall("s:sheets/s:sheet", NS):
        target = targets[sheet.get("{%s}id" % NS["r"])]
        result[sheet.get("name")] = target.lstrip("/") if target.startswith("/") else "xl/" + target
    return result


@lru_cache(maxsize=128)
def sheet_cells(sheet):
    """Raw cached cell values, avoiding date coercion and formula evaluation."""
    if sheet not in sheet_parts():
        return {}
    root = ET.fromstring(archive().read(sheet_parts()[sheet]))
    result = {}
    for cell in root.findall("s:sheetData/s:row/s:c", NS):
        raw = text_at(cell, "s:v")
        typ = cell.get("t")
        if typ == "s" and raw is not None:
            value = shared_strings()[int(raw)]
        elif typ == "inlineStr":
            value = "".join(cell.xpath("s:is//s:t/text()", namespaces=NS))
        elif typ in ("str", "e"):
            value = raw
        else:
            value = number(raw)
        result[cell.get("r")] = value
    return result


@lru_cache(maxsize=128)
def hidden_cells(sheet):
    root = ET.fromstring(archive().read(sheet_parts()[sheet]))
    cols = set()
    for col in root.findall("s:cols/s:col", NS):
        if col.get("hidden") == "1":
            cols.update(range(int(col.get("min")), int(col.get("max")) + 1))
    rows = {int(row.get("r")) for row in root.findall("s:sheetData/s:row", NS) if row.get("hidden") == "1"}
    return cols, rows


def formula_values(formula, visible_only=False):
    """Resolve a simple local sheet range to raw saved values, or None."""
    if not formula or "#REF!" in formula or "[" in formula:
        return None
    match = re.fullmatch(r"(?:'((?:[^']|'')+)'|([^!]+))!([$A-Z0-9:]+)", formula)
    if not match:
        return None
    sheet = (match.group(1) or match.group(2)).replace("''", "'")
    if sheet not in sheet_parts():
        return None
    try:
        a, b, c, d = range_boundaries(match.group(3))
    except ValueError:
        return None
    if not all((a, b, c, d)):
        return None
    cells = sheet_cells(sheet)
    hidden_cols, hidden_rows = hidden_cells(sheet) if visible_only else (set(), set())
    return [cells.get(f"{get_column_letter(col)}{row}") for row in range(b, d + 1) for col in range(a, c + 1)
            if row not in hidden_rows and col not in hidden_cols]


@lru_cache(maxsize=1)
def theme_colors():
    root = ET.fromstring(archive().read("xl/theme/theme1.xml"))
    scheme = root.find("a:themeElements/a:clrScheme", NS)
    colors = {}
    for entry in scheme:
        color = entry[0]
        colors[ET.QName(entry).localname] = color.get("lastClr", color.get("val", "000000"))
    colors.update({"tx1": colors.get("dk1", "000000"), "tx2": colors.get("dk2", "000000"),
                   "bg1": colors.get("lt1", "FFFFFF"), "bg2": colors.get("lt2", "FFFFFF")})
    return colors


def color_of(node, default=None):
    """Resolve OOXML theme, luminance, tint, shade and transparency."""
    if node is None:
        return default
    if ET.QName(node).localname in ("solidFill", "fgClr", "bgClr"):
        node = next(iter(node), None)
    if node is None:
        return default
    tag = ET.QName(node).localname
    if tag == "schemeClr":
        hexcolor = theme_colors().get(node.get("val"), "777777")
    elif tag in ("srgbClr", "sysClr"):
        hexcolor = node.get("lastClr", node.get("val", "777777"))
    else:
        return default
    try:
        rgb = [int(hexcolor[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    except ValueError:
        return default
    h, l, s = colorsys.rgb_to_hls(*rgb)
    alpha = 1
    for transform in node:
        op = ET.QName(transform).localname
        v = float(transform.get("val", "100000")) / 100000
        if op == "lumMod": l *= v
        elif op == "lumOff": l += v
        elif op == "alpha": alpha = v
        elif op == "tint": l = l * v + (1 - v)
        elif op == "shade": l *= v
        elif op == "satMod": s *= v
    rgb = [round(v * 255) for v in colorsys.hls_to_rgb(h, min(1, max(0, l)), min(1, max(0, s)))]
    if alpha < 1:
        return f"rgba({rgb[0]},{rgb[1]},{rgb[2]},{alpha:g})"
    return "#" + "".join(f"{v:02X}" for v in rgb)


def style_of(element, fallback, chart_type):
    shape = element.find("c:spPr", NS)
    line = shape.find("a:ln", NS) if shape is not None else None
    fill = color_of(shape.find("a:solidFill", NS), fallback) if shape is not None else fallback
    no_fill = shape is not None and shape.find("a:noFill", NS) is not None
    line_color = color_of(line.find("a:solidFill", NS), fallback) if line is not None else fallback
    line_none = line is not None and line.find("a:noFill", NS) is not None
    dash = val(line, "a:prstDash", "solid")
    dash = {"sysDash": "dash", "sysDot": "dot", "lgDash": "longdash", "dashDot": "dashdot",
            "lgDashDot": "longdashdot", "lgDashDotDot": "longdashdot", "sysDashDot": "dashdot",
            "sysDashDotDot": "dashdot"}.get(dash, dash)
    if dash not in ("solid", "dot", "dash", "longdash", "dashdot", "longdashdot"):
        dash = "solid"
    marker = element.find("c:marker", NS)
    symbol = val(marker, "c:symbol", "none")
    marker_shape = marker.find("c:spPr", NS) if marker is not None else None
    marker_fill = color_of(marker_shape.find("a:solidFill", NS), line_color) if marker_shape is not None else line_color
    marker_open = marker_shape is not None and marker_shape.find("a:noFill", NS) is not None
    marker_map = {"circle": "circle", "dash": "line-ew", "diamond": "diamond", "dot": "circle",
                  "plus": "cross", "square": "square", "star": "star", "triangle": "triangle-up",
                  "x": "x", "auto": "circle", "none": "circle"}
    psymbol = marker_map.get(symbol, "circle")
    if marker_open and psymbol in ("circle", "diamond", "square", "triangle-up"):
        psymbol += "-open"
    width = float(line.get("w", "28575")) / 12700 if line is not None else 2.25
    pattern = shape.find("a:pattFill", NS) if shape is not None else None
    pattern_style = None
    if pattern is not None:
        fill = color_of(pattern.find("a:bgClr", NS), fill)
        pname = pattern.get("prst", "")
        pshape = "|" if "Vert" in pname else "-" if "Horz" in pname else "x" if pname == "diagCross" else "/" if "UpDiag" in pname else "\\" if "DnDiag" in pname else "+" if pname == "cross" else "."
        pattern_style = {"shape": pshape,
                         "fgcolor": color_of(pattern.find("a:fgClr", NS), line_color), "solidity": 0.2}
    return {"fill": "rgba(0,0,0,0)" if no_fill else fill,
            "color": line_color if chart_type in ("lineChart", "scatterChart") else fill,
            "line": {"color": line_color, "width": 0 if line_none else width, "dash": dash},
            "line_none": line_none, "no_fill": no_fill,
            "marker": {"color": marker_fill, "symbol": psymbol, "size": float(val(marker, "c:size", 5)) * 1.33},
            "has_marker": symbol != "none", "pattern": pattern_style}


def cache_values(parent):
    """Return values, formula, format and inner-to-outer category levels."""
    if parent is None:
        return {"values": [], "formula": None, "format": None, "levels": []}
    formula = text_at(parent, ".//c:f")
    cache = next((n for n in parent.iter() if ET.QName(n).localname in
                  ("numCache", "strCache", "numLit", "strLit", "multiLvlStrCache")), None)
    if cache is None:
        direct = parent.find("c:v", NS)
        values = [direct.text] if direct is not None else (formula_values(formula) or [])
        return {"values": values, "formula": formula, "format": None, "levels": []}
    numeric = ET.QName(cache).localname in ("numCache", "numLit")
    levels = cache.findall("c:lvl", NS)
    count = int(val(cache, "c:ptCount", 0))
    def points(holder, fill_forward=False):
        pts = holder.findall("c:pt", NS)
        length = max([count] + [int(p.get("idx")) + 1 for p in pts])
        result = [None] * length
        for p in pts:
            raw = text_at(p, "c:v")
            result[int(p.get("idx"))] = number(raw) if numeric else raw
        if fill_forward:
            current = None
            for i, v in enumerate(result):
                if v is not None: current = v
                elif current is not None: result[i] = current
        return result
    level_values = [points(level, fill_forward=i > 0) for i, level in enumerate(levels)]
    if levels:
        values = [" · ".join(str(level[i]) for level in reversed(level_values) if level[i] is not None)
                  for i in range(count)]
    else:
        values = points(cache)
    return {"values": values, "formula": formula, "format": text_at(cache, "c:formatCode"),
            "levels": level_values}


def rich_text(node):
    if node is None: return ""
    paragraphs = node.findall(".//a:p", NS)
    if paragraphs:
        return " ".join("".join(p.xpath(".//a:t/text()", namespaces=NS)) for p in paragraphs).strip()
    return "".join(node.xpath(".//c:v/text()", namespaces=NS)).strip()


def parse_axis(node):
    unit = val(node, "c:dispUnits/c:builtInUnit")
    return {"id": val(node, "c:axId"), "type": ET.QName(node).localname,
            "position": val(node, "c:axPos"), "title": rich_text(node.find("c:title", NS)),
            "min": number(val(node, "c:scaling/c:min")), "max": number(val(node, "c:scaling/c:max")),
            "reverse": val(node, "c:scaling/c:orientation") == "maxMin",
            "log_base": number(val(node, "c:scaling/c:logBase")),
            "visible": val(node, "c:delete", "0") != "1",
            "labels": val(node, "c:tickLblPos") != "none", "cross_axis": val(node, "c:crossAx"),
            "major_unit": number(val(node, "c:majorUnit")),
            "format": node.find("c:numFmt", NS).get("formatCode") if node.find("c:numFmt", NS) is not None else None,
            "display_unit": unit, "scale": number(val(node, "c:dispUnits/c:custUnit")) or UNITS.get(unit, 1),
            "grid": node.find("c:majorGridlines", NS) is not None,
            "label_rotation": number(node.find("c:txPr/a:bodyPr", NS).get("rot")) if node.find("c:txPr/a:bodyPr", NS) is not None else None}


def same_value(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-10, abs_tol=1e-8)
    if a in (None, "") and b in (None, "", "#N/A", "#REF!", "#VALUE!"):
        return True
    return a == b


def load_chart(chart_number):
    """Parse a chart. Series x/y are raw, padded only during build_chart()."""
    entry = next(e for e in inventory()["charts"] if e["chart_number"] == int(chart_number))
    result = copy.deepcopy(entry)
    root = ET.fromstring(archive().read(entry["xml_part"]))
    plot = root.find("c:chart/c:plotArea", NS)
    result["series"], result["groups"], result["axes"], result["notes"] = [], [], {}, []
    result["blank_mode"] = val(root, "c:chart/c:dispBlanksAs", "gap")
    result["plot_visible_only"] = val(root, "c:chart/c:plotVisOnly", "1") == "1"
    result["embedded_title"] = rich_text(root.find("c:chart/c:title", NS))
    hidden_legend = {int(val(n, "c:idx")) for n in root.findall("c:chart/c:legend/c:legendEntry", NS)
                     if val(n, "c:delete") == "1"}
    for node in plot:
        tag = ET.QName(node).localname
        if tag.endswith("Ax"):
            axis = parse_axis(node)
            result["axes"][axis["id"]] = axis
        elif tag.endswith("Chart"):
            group_index = len(result["groups"])
            group = {"type": tag, "index": group_index, "grouping": val(node, "c:grouping", "standard"),
                     "bar_direction": val(node, "c:barDir", "col"), "axis_ids": [n.get("val") for n in node.findall("c:axId", NS)],
                     "gap_width": number(val(node, "c:gapWidth", 150)), "overlap": number(val(node, "c:overlap", 0)),
                     "scatter_style": val(node, "c:scatterStyle", "lineMarker"), "series": []}
            result["groups"].append(group)
            for element in node.findall("c:ser", NS):
                index = int(val(element, "c:idx", len(result["series"])))
                order = int(val(element, "c:order", index))
                name_data = cache_values(element.find("c:tx", NS))
                name = str(next((v for v in name_data["values"] if v is not None), f"Series {index + 1}")).strip()
                x_node = element.find("c:xVal", NS) if tag == "scatterChart" else element.find("c:cat", NS)
                y_node = element.find("c:yVal", NS) if tag == "scatterChart" else element.find("c:val", NS)
                x_data, y_data = cache_values(x_node), cache_values(y_node)
                style = style_of(element, PALETTE[index % len(PALETTE)], tag)
                point_colors = {}
                for dp in element.findall("c:dPt", NS):
                    point_colors[int(val(dp, "c:idx", 0))] = style_of(dp, style["fill"], tag)["fill"]
                s = {"index": index, "order": order, "name": name, "group": group_index,
                     "chart_type": tag, "grouping": group["grouping"], "axis_ids": group["axis_ids"],
                     "x": x_data["values"], "y": y_data["values"], "x_levels": x_data["levels"],
                     "x_formula": x_data["formula"], "y_formula": y_data["formula"], "name_formula": name_data["formula"],
                     "x_format": x_data["format"], "y_format": y_data["format"], "color": style["color"],
                     "style": style, "point_colors": point_colors, "showlegend": order not in hidden_legend,
                     "smooth": val(element, "c:smooth", "0") == "1"}
                dl = element.find("c:dLbls", NS)
                if dl is None: dl = node.find("c:dLbls", NS)
                s["show_values"] = val(dl, "c:showVal", "0") == "1"
                s["labels"] = {int(val(d, "c:idx", 0)): rich_text(d.find("c:tx", NS))
                               for d in (dl.findall("c:dLbl", NS) if dl is not None else []) if d.find("c:tx", NS) is not None}
                raw_values = formula_values(s["y_formula"], visible_only=result["plot_visible_only"])
                mismatch = []
                if raw_values is not None:
                    for i in range(max(len(raw_values), len(s["y"]))):
                        a = s["y"][i] if i < len(s["y"]) else None
                        b = raw_values[i] if i < len(raw_values) else None
                        if not same_value(a, b): mismatch.append(i)
                source_kind = "worksheet" if raw_values is not None else "literal" if not s["y_formula"] else "broken-reference" if "#REF!" in s["y_formula"] else "unresolved-reference"
                s["validation"] = {"source_resolved": raw_values is not None, "source_kind": source_kind, "visible_cells_only": result["plot_visible_only"], "mismatch_count": len(mismatch),
                                   "mismatch_indices": mismatch[:20], "cache_points": len(s["y"])}
                if mismatch:
                    result["notes"].append(f"Saved chart cache for ‘{name}’ differs from the saved worksheet at {len(mismatch)} point(s); the published chart cache is retained.")
                if s["y_formula"] and "#REF!" in s["y_formula"]:
                    result["notes"].append(f"‘{name}’ has an invalid Excel value reference ({s['y_formula']}); only its saved chart cache is available.")
                result["series"].append(s)
                group["series"].append(index)
    # Excel reuses sibling categories when a series omits c:cat. A broken
    # category reference may likewise be repaired from an unambiguous peer.
    candidates = [s for s in result["series"] if s["x"] and not (s["x_formula"] and "#REF!" in s["x_formula"])]
    for s in result["series"]:
        bad_ref = bool(s["x_formula"] and "#REF!" in s["x_formula"])
        if not s["x"] or bad_ref:
            peers = [p for p in candidates if len(p["x"]) >= len(s["y"]) and p["axis_ids"] == s["axis_ids"]]
            if peers:
                peer = min(peers, key=lambda p: len(p["x"]))
                s["x"], s["x_levels"] = copy.deepcopy(peer["x"]), copy.deepcopy(peer["x_levels"])
                s["x_format"] = peer["x_format"]
                s["category_inherited_from"] = peer["index"]
                s["resolved_x_formula"] = peer["x_formula"]
                if bad_ref:
                    result["notes"].append(f"Invalid category reference {s['x_formula']} repaired using sibling series categories {peer['x_formula']}; values are unchanged.")
            else:
                s["x"] = list(range(1, len(s["y"]) + 1))
                s["generated_categories"] = True
    return result


def axis_layout(axis):
    a = {"title": {"text": axis["title"], "standoff": 14}, "automargin": True,
         "showgrid": axis["grid"], "gridcolor": "#E5EAF0", "zerolinecolor": "#BAC5CE",
         "showline": False, "showticklabels": axis["visible"] and axis["labels"],
         "visible": axis["visible"], "fixedrange": False}
    scale = axis["scale"]
    if axis["type"] == "catAx": a["type"] = "category"
    elif axis["type"] == "dateAx": a["type"] = "date"
    else: a["type"] = "linear"
    if axis["log_base"]: a["type"] = "log"
    if axis["min"] is not None and axis["max"] is not None:
        a["range"] = [axis["min"] / scale, axis["max"] / scale]
        if axis["reverse"]: a["range"].reverse()
    elif axis["min"] is not None or axis["max"] is not None:
        # Plotly 3.1 can discard a partially specified range during coercion.
        # minallowed/maxallowed preserve Excel's fixed endpoint while allowing
        # the opposite endpoint to fit the data (verified against browser output).
        if axis["reverse"]: a["autorange"] = "reversed"
        bound = "minallowed" if axis["min"] is not None else "maxallowed"
        a["autorangeoptions"] = {bound: (axis["min"] if axis["min"] is not None else axis["max"]) / scale}
    elif axis["reverse"]: a["autorange"] = "reversed"
    if axis["major_unit"] and a["type"] == "linear": a["dtick"] = axis["major_unit"] / scale
    fmt = axis["format"] or ""
    if "%" in fmt: a["tickformat"] = ".0%" if "." not in fmt else ".1%"
    elif "#,##0" in fmt: a["tickformat"] = ",.0f" if "." not in fmt else ",.1f"
    rot = axis["label_rotation"]
    if rot is not None and -5400000 <= rot <= 5400000: a["tickangle"] = rot / 60000
    return a


def excel_time(value):
    if not isinstance(value, (int, float)): return value
    minutes = round((value % 1) * 24 * 60) % (24 * 60)
    return f"{minutes // 60:02}:{minutes % 60:02}"


def excel_date(value):
    if not isinstance(value, (int, float)): return value
    return (dt.datetime(1899, 12, 30) + dt.timedelta(days=value)).isoformat()


def build_chart(chart_number):
    chart = load_chart(chart_number)
    notes = chart["notes"].copy()
    axes = chart["axes"]
    axis_map, layout = {}, {"hovermode": "x unified", "dragmode": "zoom",
        "paper_bgcolor": "#FFFFFF", "plot_bgcolor": "#FFFFFF",
        "font": {"family": "Arial, sans-serif", "size": 12, "color": "#283C49"},
        "margin": {"l": 75, "r": 45, "t": 24, "b": 100},
        "legend": {"orientation": "h", "y": -0.23, "x": 0, "font": {"size": 11}, "traceorder": "normal"}}
    counts = {"x": 0, "y": 0}
    for aid, axis in axes.items():
        dim = "x" if axis["position"] in ("b", "t") else "y"
        counts[dim] += 1
        suffix = "" if counts[dim] == 1 else str(counts[dim])
        axis_map[aid] = dim + suffix
        a = axis_layout(axis)
        a["side"] = {"l": "left", "r": "right", "b": "bottom", "t": "top"}.get(axis["position"], "bottom")
        if counts[dim] > 1:
            a["overlaying"] = dim
            if dim == "x": a["showgrid"] = False
        layout[f"{dim}axis{suffix}"] = a
    for aid, axis in axes.items():
        key = axis_map[aid]
        layout[key[0] + "axis" + key[1:]]["anchor"] = axis_map.get(axis["cross_axis"], "y" if key[0] == "x" else "x")
    data, sources = [], []
    area_groups_seen = set()
    line_stacks = {}
    for series in chart["series"]:
        group = chart["groups"][series["group"]]
        xtype = series["chart_type"]
        xaid = next((a for a in series["axis_ids"] if axis_map.get(a, "x").startswith("x")), None)
        yaid = next((a for a in series["axis_ids"] if axis_map.get(a, "y").startswith("y")), None)
        xaxis, yaxis = axes.get(xaid, {}), axes.get(yaid, {})
        xref, yref = axis_map.get(xaid, "x"), axis_map.get(yaid, "y")
        scale = yaxis.get("scale", 1)
        xs, raw_ys = list(series["x"]), list(series["y"])
        count = max(len(xs), len(raw_ys))
        xs += [None] * (count - len(xs))
        raw_ys += [None] * (count - len(raw_ys))
        ys = [v / scale if isinstance(v, (int, float)) else None for v in raw_ys]
        metadata = {"sourceSeriesIndex": series["index"], "sourceSeriesOrder": series["order"],
                    "excelChartType": xtype, "sourceGroup": series["group"], "sourceValueCount": len(series["y"]),
                    "yScaleFactor": scale}
        # Stacked lines in Excel store increments; their displayed line is a
        # cumulative total. Keep cache values as customdata for audit/export.
        if xtype == "lineChart" and series["grouping"] in ("stacked", "percentStacked"):
            cumulative = line_stacks.setdefault(series["group"], [0] * count)
            if len(cumulative) < count: cumulative += [0] * (count - len(cumulative))
            for i, y in enumerate(ys):
                if y is not None: cumulative[i] += y
            ys = [cumulative[i] if y is not None else None for i, y in enumerate(ys)]
            metadata["stackedLine"] = True
        if xaxis.get("type") == "dateAx":
            xs = [excel_date(v) for v in xs]
        elif series["x_format"] and re.search(r"h+:m", series["x_format"], flags=re.I):
            xs = [excel_time(v) for v in xs]
        elif xaxis.get("type") == "catAx":
            xs = [str(int(v)) if isinstance(v, (int, float)) and int(v) == v else str(v) if v is not None else None for v in xs]
        style = series["style"]
        trace = {"name": series["name"], "x": xs, "y": ys, "xaxis": xref, "yaxis": yref,
                 "showlegend": series["showlegend"], "legendrank": series["order"], "meta": metadata,
                 "customdata": raw_ys, "hovertemplate": "%{x}<br>%{fullData.name}: %{y:,.3f}<extra></extra>"}
        if xtype == "barChart":
            trace["type"] = "bar"
            colors = [series["point_colors"].get(i, style["fill"]) for i in range(count)] if series["point_colors"] else style["fill"]
            trace["marker"] = {"color": colors, "line": {"width": style["line"]["width"], "color": style["line"]["color"]}}
            if style["pattern"]: trace["marker"]["pattern"] = style["pattern"]
            if group["bar_direction"] == "bar":
                trace["orientation"], trace["x"], trace["y"] = "h", ys, xs
            if series["grouping"] in ("stacked", "percentStacked"):
                layout["barmode"] = "relative"
                if series["grouping"] == "percentStacked": layout["barnorm"] = "percent"
            elif "barmode" not in layout: layout["barmode"] = "group"
            layout["bargap"] = min(0.8, group["gap_width"] / (100 + group["gap_width"]))
        else:
            trace["type"] = "scatter"
            marker_enabled = style["has_marker"]
            line_enabled = not style["line_none"]
            if xtype == "scatterChart" and group["scatter_style"] == "marker": line_enabled = False
            trace["mode"] = "lines+markers" if line_enabled and marker_enabled else "lines" if line_enabled else "markers" if marker_enabled else "none"
            trace["line"] = copy.deepcopy(style["line"])
            if series["smooth"]: trace["line"]["shape"] = "spline"
            trace["marker"] = style["marker"]
            trace["connectgaps"] = chart["blank_mode"] == "span"
            if xtype == "areaChart":
                trace["mode"] = "lines"
                trace["fillcolor"] = style["fill"]
                if style["pattern"]:
                    trace["fillpattern"] = {**style["pattern"], "bgcolor": style["fill"], "fillmode": "replace"}
                trace["fill"] = "tozeroy"
                if series["grouping"] in ("stacked", "percentStacked"):
                    trace["stackgroup"] = f"area-{series['group']}"
                    trace["fill"] = "tonexty" if series["group"] in area_groups_seen else "tozeroy"
                    area_groups_seen.add(series["group"])
                    trace["stackgaps"] = "infer zero"
                    if series["grouping"] == "percentStacked": trace["groupnorm"] = "percent"
                # Excel filled-area borders are often explicitly disabled.
                if style["line_none"]: trace["line"]["width"] = 0
        if series["show_values"] or series["labels"]:
            trace["text"] = [series["labels"].get(i, f"{v:,.0f}" if isinstance(v, (int, float)) and series["show_values"] else "") for i, v in enumerate(ys)]
            if trace["type"] == "bar": trace["textposition"] = "auto"
            else:
                trace["mode"] = (trace["mode"] + "+text") if trace["mode"] != "none" else "text"
                trace["textposition"] = "top center"
        data.append(trace)
        sources.append({"series_index": series["index"], "series_name": series["name"],
                        "name": series["name_formula"], "x": series.get("resolved_x_formula", series["x_formula"]),
                        "original_x": series["x_formula"], "y": series["y_formula"], "validation": series["validation"]})
    # Category axes cannot distinguish repeated labels, and Plotly does not
    # calculate stackgroup filled areas on category axes at all. Use positions
    # for either case, retaining source labels in the hover and tick text.
    for xref in {t.get("xaxis", "x") for t in data}:
        key = xref[0] + "axis" + xref[1:]
        if layout.get(key, {}).get("type") != "category": continue
        axis_traces = [t for t in data if t.get("xaxis", "x") == xref and t.get("orientation") != "h"]
        longest = max(axis_traces, key=lambda t: len(t["x"]), default=None)
        if longest is None: continue
        labels = longest["x"]
        nonempty = [x for x in labels if x is not None]
        duplicated = len(set(nonempty)) != len(nonempty)
        stacked_area = any(t.get("stackgroup") for t in axis_traces)
        if duplicated or stacked_area:
            category_positions = {label: i for i, label in enumerate(labels) if label is not None}
            if not duplicated:
                for trace in axis_traces:
                    for label in trace["x"]:
                        if label is not None and label not in category_positions:
                            category_positions[label] = len(labels)
                            labels.append(label)
            for trace in axis_traces:
                original = trace["x"]
                trace["x"] = list(range(len(original))) if duplicated else [category_positions.get(label, i) for i, label in enumerate(original)]
                trace["customdata"] = [[label, raw] for label, raw in zip(original, trace["customdata"])]
                trace["hovertemplate"] = "%{customdata[0]}<br>%{fullData.name}: %{y:,.3f}<extra></extra>"
                trace["meta"]["positionalCategories"] = True
            step = max(1, math.ceil(len(labels) / 18))
            layout[key].update({"type": "linear", "tickmode": "array", "tickvals": list(range(0, len(labels), step)),
                                "ticktext": labels[::step]})
    if any(s["x_levels"] for s in chart["series"]):
        notes.append("Multi-level Excel categories are combined into unique labels; all category levels remain in the hover labels.")
    if any(t["meta"].get("stackedLine") for t in data):
        notes.append("Stacked line series are displayed as cumulative totals, as in Excel; source increments are retained for verification.")
    if any(t["meta"]["yScaleFactor"] != 1 for t in data):
        notes.append("Values and axes use the Excel chart’s explicit display-unit scaling; original cached values are retained.")
    title = chart["title"].strip()
    same_sheet = [e for e in inventory()["charts"] if e["sheet"] == chart["sheet"]]
    subtitle = ""
    if len(same_sheet) > 1:
        names = [s["name"] for s in chart["series"]]
        subtitle = " / ".join(names[:2]) + (f" (+{len(names)-2} series)" if len(names) > 2 else "")
    return {"id": chart["id"], "chart_number": chart["chart_number"], "sheet": chart["sheet"],
            "title": title, "subtitle": subtitle, "section": chart["section"], "kind": "interactive",
            "data": data, "layout": layout, "notes": list(dict.fromkeys(notes)), "sources": sources,
            "excel_types": chart["types"], "source_file": WORKBOOK.name,
            "source_series_count": chart["series_count"], "source_xml": chart["xml_part"]}


def write_group(numbers, output):
    specs = [build_chart(n) for n in numbers]
    target = ROOT / output
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(specs, ensure_ascii=False, indent=2, allow_nan=False) + "\n")
    return specs
