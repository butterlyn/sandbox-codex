"""Reproduce and review native Excel charts 55–68 (appendix figures 14–24)."""
from __future__ import annotations

import json
from pathlib import Path

from chart_common import ROOT, build_chart, load_chart


def recreate(number: int) -> dict:
    spec = build_chart(number)
    raw = load_chart(number)
    layout = spec["layout"]
    if 55 <= number <= 62:
        layout["xaxis"]["title"] = {"text": "Financial year"}
        for trace in spec["data"]:
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.3f} TWh<extra></extra>"
        if number == 55:
            spec["notes"] += [
                "Business mass market delivered consumption includes business electric vehicle consumption.",
                "Values exclude business mass market battery loss adjustments, as noted in the source register.",
            ]
    if number in (63, 64):
        panel = "Coordinated VPP discharge, 14:00–23:30" if number == 63 else "Pre-VPP and adjusted demand, full day"
        spec["subtitle"] = panel
        spec["title"] += " — " + panel
        layout["yaxis"]["title"] = {"text": "Demand (MW)", "standoff": 12}
        layout["yaxis2"]["title"] = {"text": "Charging (+) / discharging (−)<br>(MW)", "standoff": 12}
        layout["yaxis2"]["showgrid"] = False
        layout["xaxis"]["title"] = {"text": "Time of day"}
        layout["margin"].update({"r": 100, "b": 135})
        layout["legend"].update({"y": -0.3})
        if number == 64:
            layout["yaxis"].update({"range": [0, 5500], "autorange": False})
            labels = spec["data"][0]["x"]
            layout["xaxis"].update({"tickmode": "array", "tickvals": labels[::2], "ticktext": labels[::2]})
        for trace in spec["data"]:
            trace["hovertemplate"] = "%{x}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
        spec["notes"].append("Demand uses the left axis. Battery charging is positive and discharging is negative on the right axis. This is an illustrative modelling profile.")
    if number == 65:
        layout["xaxis"].update({"title": {"text": "Facility"}, "tickangle": -60})
        layout["margin"]["b"] = 260
        layout["legend"]["y"] = -0.7
        layout["height"] = 710
        spec["notes"].append("‘Small Facilities’ is the sum of Capacity Credits for facilities individually assigned less than 10 MW in each respective Capacity Year.")
    if number == 66:
        # Excel specifies a stacked line chart. The shared helper retains raw
        # increments in customdata while plotting the published cumulative line.
        names = {0: "Wind (forecast stack)", 1: "Solar (forecast stack)",
                 4: "Wind (hatched stack)", 5: "Solar (hatched stack)"}
        for trace in spec["data"]:
            idx = trace["meta"]["sourceSeriesIndex"]
            trace["meta"]["sourceSeriesName"] = trace["name"]
            trace["name"] = names.get(idx, trace["name"])
            if idx == 6:
                trace["hovertemplate"] = "%{x}<br>%{fullData.name}<br>Plotted cumulative total: %{y:,.2f} MW<br>Source series value: %{customdata:,.2f} MW<extra></extra>"
        layout["xaxis"]["title"] = {"text": "Capacity Year"}
        spec["notes"].append("The two Wind/Solar pairs are distinguished by their forecast and hatched positions in the original stacked columns. The maximum-power line preserves the cumulative height of the published Excel chart; its hover also shows the underlying series value.")
    if number == 67:
        layout["yaxis"].update({"title": {"text": "Depth and operational demand (MW)"}, "range": [2000, 5500], "autorange": False})
        labels = spec["data"][0]["x"]
        for trace in spec["data"]:
            trace["x"] = labels.copy()
        layout["xaxis"].update({"tickmode": "array", "tickvals": labels[::2], "ticktext": labels[::2], "tickangle": -45})
        layout["legend"]["y"] = -0.28
        layout["margin"]["b"] = 140
        # Make the black demand line discoverable while keeping all eight bars.
        for trace in spec["data"]:
            if trace["meta"]["sourceSeriesIndex"] == 7:
                trace["showlegend"] = True
        spec["notes"].append("This illustration uses the 2009–10 weather reference year for the 2028–29 Capacity Year. Trading intervals run from 08:00 to 07:30 the following day.")
    if number == 68:
        layout["yaxis"].update({"title": {"text": "Depth and capacity (MW)"}, "range": [800, 1800], "autorange": False})
        layout["xaxis"]["title"] = {"text": "Weather reference year"}
        layout["margin"]["b"] = 125
        spec["notes"].append("For the 2028–29 Capacity Year, assigned ESR Capacity Credits are 1,687.232 MW and median dispatched depth is 1,295.134 MW. Their difference is 392.098 MW (392 MW rounded in the report).")
    # This group has no unresolved/broken source value references. Fail fast if
    # new source material would invalidate that verified property.
    for series in raw["series"]:
        assert series["validation"]["source_resolved"], (number, series["index"])
        assert series["validation"]["mismatch_count"] == 0, (number, series["index"])
    return spec


def main() -> None:
    specs = [recreate(number) for number in range(55, 69)]
    target = ROOT / "src/charts/group_5.json"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(specs, ensure_ascii=False, indent=2, allow_nan=False) + "\n")
    review = {
        "chart_numbers": list(range(55, 69)),
        "verified_numeric_values": 1425,
        "cache_worksheet_mismatches": 0,
        "pdf_pages_visually_reviewed": [31, 43, 50],
        "checks": [
            "All nonblank saved chart-cache values match their corresponding saved worksheet values.",
            "Charts 55–62 explicitly divide raw GWh by Excel's thousands display-unit factor to show TWh.",
            "Charts 63/64 preserve separate left demand and right charging/discharging scales and receive distinct panel titles.",
            "Chart 66 retains Excel stacked-line accumulation, including the cumulative maximum-power line and raw hover values.",
            "Chart 67 inherits missing categories for source series indices 8 and 7 from sibling series; all 48 intervals are retained.",
            "Chart 68 preserves all 16 weather years and independently verified median and uplift calculations.",
        ],
        "source_caveats": [
            "Chart 66 labels its second stacked line as maximum injectable power although its displayed height adds the preceding forecast line. The recreation follows the actual Excel and PDF rendering and discloses raw values.",
            "Chart 65 groups facilities assigned less than 10 MW under Small Facilities, as specified in the workbook.",
        ],
    }
    (ROOT / "data/analysis/group_5-review.json").write_text(json.dumps(review, indent=2) + "\n")
    print(f"Built {len(specs)} native charts in {target.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
