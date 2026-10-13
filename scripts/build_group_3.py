#!/usr/bin/env python3
"""Build and verify Excel chart objects 29–41, including chart-specific semantics."""

from __future__ import annotations

import json
import math
from pathlib import Path

from chart_common import build_chart, excel_date, excel_time, formula_values, load_chart


ROOT = Path(__file__).resolve().parents[1]
NUMBERS = range(29, 42)


def same(left, right):
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return math.isclose(left, right, rel_tol=1e-10, abs_tol=1e-8)
    if left is None and right is None:
        return True
    if isinstance(left, str) and isinstance(right, str):
        return left.strip() == right.strip()
    return left == right


def audit_chart(raw, spec):
    """Check chart cache values/categories against their saved worksheet ranges."""
    checks = []
    rendered = {trace["meta"]["sourceSeriesIndex"]: trace for trace in spec["data"]}
    for series in raw["series"]:
        trace = rendered[series["index"]]
        scale = trace["meta"]["yScaleFactor"]
        source_y = list(series["y"]) + [None] * (len(trace["y"]) - len(series["y"]))
        for i, value in enumerate(source_y):
            assert same(trace["y"][i], value / scale if value is not None else None), (raw["chart_number"], series["index"], i)
        assert len(trace["x"]) == len(trace["y"])
        for direction in ("x", "y"):
            formula = series.get(f"resolved_{direction}_formula", series[f"{direction}_formula"])
            original_formula = series[f"{direction}_formula"]
            saved = formula_values(formula)
            cache = series[direction]
            if saved is None:
                # Figure 47's waterfall baseline and RCT are literal chart data;
                # the same saved values are independently present in the sheet.
                literal_range = {
                    (34, 4, "y"): "'Ch 7_F.47'!$G$26:$G$30",
                    (34, 3, "y"): "'Ch 7_F.47'!$C$26:$C$30",
                }.get((raw["chart_number"], series["index"], direction))
                if literal_range:
                    saved = formula_values(literal_range)
                    formula = literal_range
            if saved is not None:
                length = max(len(cache), len(saved))
                mismatches = [i for i in range(length) if not same(cache[i] if i < len(cache) else None, saved[i] if i < len(saved) else None)]
                assert not mismatches, (raw["chart_number"], series["index"], direction, mismatches[:10])
                checks.append({
                    "series_index": series["index"], "direction": direction,
                    "formula": formula, "original_formula": original_formula,
                    "points_checked": length, "mismatch_count": 0,
                    "normalization": "Whitespace-only category differences are ignored." if direction == "x" and raw["chart_number"] == 33 else None,
                })
        assert not trace.get("connectgaps", False), "Source empty cells must remain gaps."
    return {
        "chart_number": raw["chart_number"], "sheet": raw["sheet"],
        "title": spec["title"], "series_count": len(raw["series"]),
        "numeric_points": sum(sum(v is not None for v in trace["y"]) for trace in spec["data"]),
        "checks": checks, "notes": spec["notes"],
        "rendered_values_match_chart_cache": True,
    }


def apply_adjustments(number, raw, spec):
    data, layout = spec["data"], spec["layout"]
    by_index = {series["index"]: series for series in raw["series"]}
    # Excel category spaces are typographical, not separate positions.
    for trace in data:
        trace["x"] = [x.strip() if isinstance(x, str) else x for x in trace["x"]]

    if number in (29, 31):
        categories = [excel_time(x) for x in raw["series"][0]["x"]]
        for trace in data:
            # Figure 31's reference line inherits numeric categories but no
            # number format from Excel; use the bar series' clock formatting.
            trace["x"] = categories.copy()
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
        layout["xaxis"].update({"type": "category", "categoryorder": "array", "categoryarray": categories, "tickmode": "array", "tickvals": categories[::2], "tickangle": -40})
        if number == 29:
            layout["yaxis"]["title"]["text"] = "Demand and ESR dispatch (MW)"
            spec["notes"].append("Trading intervals run from 08:00 to 07:30 the following day. The dashed Peak ESROD Demand line shares the same 48 intervals and MW axis as the stacked columns.")

    if number == 30:
        for trace in data:
            trace["x"] = list(range(17569))
            trace["hovertemplate"] = "Trading interval %{x:,.0f}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
            trace["line"]["simplify"] = True
        layout["xaxis"].update({"type": "linear", "range": [0, 17568], "dtick": 2000, "tickformat": ",.0f"})
        spec["notes"].append("All 17,569 saved interval positions are retained. The 2020-21, 2021-22 and 2024-25 profiles contain 17,521 values and 48 trailing blanks; the 2023-24 profile and dispatch-threshold series contain all 17,569 values.")

    if number in (32, 33, 35, 37, 41):
        layout["xaxis"]["title"]["text"] = "Capacity year"

    if number == 34:
        # The transparent 'Total' series is an offset, essential to the
        # waterfall. The separate Excel line axis must have the same MW scale.
        categories = data[0]["x"].copy()
        for trace in data:
            trace["x"] = categories.copy()
            trace["xaxis"] = "x"
            trace["yaxis"] = "y"
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
            if trace["meta"]["sourceSeriesIndex"] == 4:
                trace["showlegend"] = False
                trace["hoverinfo"] = "skip"
                trace.pop("hovertemplate", None)
                trace["meta"]["waterfallBaseline"] = True
            elif trace["meta"]["sourceSeriesIndex"] in (0, 1):
                trace["text"] = ["", "", "", "", f"{trace['y'][4]:,.0f}"]
                trace["textposition"] = "inside"
                trace["insidetextfont"] = {"color": "#ffffff"}
            elif trace["meta"]["sourceSeriesIndex"] == 2:
                # The anticipated increment belongs to Class 2, so a label
                # attached to the zero-height Class 3 bar would not be visible.
                trace["text"][3] = ""
        layout.pop("xaxis2", None)
        layout.pop("yaxis2", None)
        layout["yaxis"].update({"range": [0, 8000], "dtick": 1000})
        layout["xaxis"].update({"categoryorder": "array", "categoryarray": data[0]["x"], "tickangle": 0})
        layout["barmode"] = "stack"
        anticipated_top = sum(trace["y"][3] or 0 for trace in data if trace["type"] == "bar")
        layout["annotations"] = [
            {"x": "Anticipated", "y": anticipated_top, "text": "622", "showarrow": False, "yshift": 14, "font": {"size": 11}},
            {"xref": "paper", "x": 0, "y": 6323, "text": "RCT 6,323", "showarrow": False, "yshift": 13, "xanchor": "left", "font": {"size": 11}},
        ]
        spec["notes"].append("The transparent baseline series reproduces Excel’s waterfall offsets for committed capacity, RLM uplift and anticipated capacity. The RCT is 6,323 MW on the same MW scale; the final column is total forecast available capacity.")

    if number == 35:
        categories = [value.strip() for value in formula_values("'Ch 7_F.49'!$B$28:$B$39")]
        for trace in data:
            trace["x"] = categories.copy()
            trace["xaxis"] = "x"
            if trace["meta"]["sourceSeriesIndex"] == 5:
                trace["yaxis"] = "y2"
                trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
            else:
                trace["yaxis"] = "y"
                trace["zorder"] = 10
                trace["hovertemplate"] = "%{x}<br>%{fullData.name}: $%{y:,.2f} per MW per year<extra></extra>"
        layout.pop("xaxis2", None)
        layout["xaxis"].update({"categoryorder": "array", "categoryarray": categories, "range": [-0.5, 11.5]})
        # Confirmed against report Figure 49: the right-hand zero is halfway up
        # the figure, while the left-hand currency axis begins at zero.
        layout["yaxis"].update({"range": [0, 500000], "autorange": False, "dtick": 50000, "tickprefix": "$", "tickformat": ",.0f", "anchor": "x"})
        layout["yaxis2"].update({"range": [-1000, 1000], "autorange": False, "dtick": 200, "anchor": "x", "overlaying": "y", "side": "right", "tickformat": ",.0f", "zeroline": True})
        layout["margin"]["r"] = 90
        spec["notes"].append("Price lines use the left axis ($ per MW per year); excess/deficit columns use the right axis (MW). Axis ranges match report Figure 49. The 2028-29 capacity and Peak RC Price cells are blank; only Benchmark RC Price is published for that year.")

    if number == 36:
        # Excel declares dateAx despite supplying literal financial-year strings.
        # Treat those labels as categories, avoiding invalid browser date parsing.
        categories = raw["series"][0]["x"]
        layout["xaxis"].update({"type": "category", "categoryorder": "array", "categoryarray": categories, "tickmode": "array", "tickvals": categories[::2], "tickangle": -40, "title": {"text": "Financial year", "standoff": 14}})
        layout["yaxis"].update({"tickformat": ".1%", "hoverformat": ".3%"})
        for trace in data:
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:.3%}<extra></extra>"
        spec["notes"].append("Growth values are stored as fractions and formatted as percentages. Financial-year labels are retained exactly; they are not parsed as calendar dates.")

    if number in (38, 39, 40):
        dates = [excel_date(value)[:10] for value in formula_values(f"'{spec['sheet']}'!$C$24:$C$227")]
        for trace in data:
            series = by_index[trace["meta"]["sourceSeriesIndex"]]
            trace["customdata"] = [[year, value, date] for year, value, date in zip(series["x"], trace["y"], dates)]
            trace["hovertemplate"] = "%{customdata[0]} · %{customdata[2]}<br>%{fullData.name}: %{y:,.2f}" + (" MW degraded" if number in (39, 40) else " connections") + "<extra></extra>"
        layout["xaxis"]["title"]["text"] = "Capacity year (monthly observations)"
        spec["sources"].append({"series_name": "Monthly dates for hover labels", "x": f"'{spec['sheet']}'!$C$24:$C$227", "y": None})
        spec["notes"].append("Each capacity-year label spans 12 monthly observations. Every monthly point is preserved at a unique position, with its worksheet date shown on hover.")
        if number in (39, 40):
            spec["notes"].append("The 2026 Actual series ends in June 2025, and the 2026 scenario series begin in April 2026. The intervening blank source cells are retained; no values are interpolated.")

    if number == 41:
        for trace in data:
            assert trace["meta"]["yScaleFactor"] == 1000
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.3f} TWh<extra></extra>"
        spec["notes"].append("Source cells are in GWh. Values are divided by 1,000 to reproduce the Excel chart’s TWh display unit; original GWh values remain in the saved chart data.")

    return spec


def main():
    specs = []
    reviews = []
    for number in NUMBERS:
        raw = load_chart(number)
        spec = apply_adjustments(number, raw, build_chart(number))
        reviews.append(audit_chart(raw, spec))
        specs.append(spec)
    target = ROOT / "src/charts/group_3.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(specs, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n")
    review = {
        "charts": reviews,
        "chart_count": len(specs),
        "cache_values_verified": sum(chart["numeric_points"] for chart in reviews),
        "cache_and_worksheet_positions_checked": sum(check["points_checked"] for chart in reviews for check in chart["checks"]),
        "source_issues": [
            "Chart 32 has a broken category formula for its surplus/shortfall series; the sibling series’ saved year labels are used, without changing any y values.",
            "Chart 33 cached categories contain whitespace absent from the worksheet; trimming is display-only.",
            "Chart 36 declares an Excel date axis but supplies financial-year text categories; the recreation displays those text labels.",
        ],
        "source_file": "data/sources/2026-wem-esoo-data-register.xlsx",
    }
    (ROOT / "data/analysis/group_3-review.json").write_text(json.dumps(review, ensure_ascii=False, indent=2, allow_nan=False) + "\n")
    print(f"Wrote {len(specs)} charts to {target.relative_to(ROOT)} ({target.stat().st_size:,} bytes)")
    print(f"Verified {review['cache_values_verified']:,} numeric values and {review['cache_and_worksheet_positions_checked']:,} source positions.")


if __name__ == "__main__":
    main()
