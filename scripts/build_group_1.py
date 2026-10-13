"""Executive summary and demand charts (native Excel objects 1–14)."""
from __future__ import annotations

import json

from chart_common import ROOT, build_chart, load_chart


def scenario_capacity(spec):
    """Explain stacked scenario increments and remove broken Excel placeholders."""
    source = load_chart(2)
    by_index = {s["index"]: s for s in source["series"]}
    broken = [t for t in spec["data"] if t["name"] == "#REF!"]
    spec["excluded_series"] = [{"sourceSeriesIndex": t["meta"]["sourceSeriesIndex"],
                                 "name": t["name"], "cached_y": t["customdata"],
                                 "reason": "Invalid value/name reference; cached single value 1 is not a forecast series."}
                                for t in broken]
    spec["data"] = [t for t in spec["data"] if t["name"] != "#REF!"]
    technology_colors = {s["name"]: s["style"]["fill"] for s in source["series"]
                         if s["chart_type"] == "barChart" and s["order"] < 9}
    seen = set()
    for trace in spec["data"]:
        series = by_index[trace["meta"]["sourceSeriesIndex"]]
        if trace["type"] != "bar":
            trace["showlegend"] = True
            continue
        order = series["order"]
        stage = "Existing (Low)" if order < 10 else "Committed increment (Expected)" if order < 21 else "Anticipated increment (High)"
        technology = trace["name"]
        trace["name"] = technology
        trace["legendgroup"] = technology
        trace["showlegend"] = technology not in seen
        trace["legendrank"] = series["order"] if technology not in seen else 100 + series["order"]
        seen.add(technology)
        trace["marker"]["color"] = technology_colors[technology]
        if order >= 10:
            trace["marker"]["color"] = "#FFFFFF"
            trace["marker"]["pattern"] = {"shape": "." if order < 21 else "|",
                "fgcolor": technology_colors[technology], "bgcolor": "#FFFFFF", "solidity": 0.45}
        trace["hovertemplate"] = f"%{{x}}<br>{technology} · {stage}: %{{y:,.3f}} MW<extra></extra>"
        trace["meta"]["scenarioComponent"] = stage
    spec["notes"] = [n for n in spec["notes"] if "‘#REF!’" not in n]
    spec["notes"].append("The workbook contains three broken #REF! placeholder series with a single cached value of 1. They are excluded from the plotted forecast and retained in provenance.")
    spec["notes"].append("Bars stack existing capacity, committed increments and anticipated increments; solid, dotted and striped fills follow the Excel chart. Each technology legend entry toggles its components together.")
    spec["layout"]["legend"]["groupclick"] = "togglegroup"
    return spec


def battery_profiles(spec):
    """Keep every half-hour and make the day/time labels readable."""
    number = spec["chart_number"]
    labels = list(spec["data"][0]["x"])
    for trace in spec["data"]:
        original = list(trace["x"])
        raw = trace["customdata"]
        trace["x"] = list(range(len(original)))
        trace["customdata"] = [[label, value] for label, value in zip(original, raw)]
        trace["hovertemplate"] = "%{customdata[0]}<br>%{fullData.name}: %{y:,.2f} MW<extra></extra>"
        trace["meta"]["positionalCategories"] = True
    xaxis = spec["layout"]["xaxis"]
    xaxis.update({"type": "linear", "tickmode": "array", "tickangle": 0,
                  "title": {"text": "Modelled half-hour interval"}})
    if number in (11, 12):
        ticks = list(range(0, 144, 12))
        xaxis.update({"tickvals": ticks, "ticktext": [labels[i].replace(" · ", "<br>") for i in ticks],
                      "range": [-0.5, 143.5]})
        spec["layout"]["shapes"] = [{"type": "line", "xref": "x", "yref": "paper", "x0": n - 0.5, "x1": n - 0.5,
                                      "y0": 0, "y1": 1, "line": {"color": "#CBD5DF", "dash": "dot", "width": 1}}
                                     for n in (48, 96)]
    else:
        xaxis.update({"tickvals": list(range(len(labels))), "ticktext": [label.split(" · ")[-1] for label in labels],
                      "range": [-0.5, len(labels) - 0.5]})
    subtitles = {11: "Battery charging and discharging · three modelled days",
                 12: "Operational demand before and after distributed batteries · three modelled days",
                 13: "VPP charging and discharging · Day 3, 15:30–21:00",
                 14: "Operational demand before and after the VPP · Day 3, 15:30–21:00"}
    spec["subtitle"] = subtitles[number]
    spec["related_chart_ids"] = [f"chart-{(number + 1 if number % 2 else number - 1):02}"]
    spec["layout"]["margin"]["b"] = 120
    return spec


def create_specs():
    specs = [build_chart(n) for n in range(1, 15)]
    for spec in specs:
        if spec["chart_number"] == 2:
            scenario_capacity(spec)
        elif spec["chart_number"] in (11, 12, 13, 14):
            battery_profiles(spec)
        # Validate raw source caches independently of intentional rendering
        # transformations (unit scaling and cumulative stacked lines).
        source = load_chart(spec["chart_number"])
        mismatches = [s["name"] for s in source["series"] if s["validation"]["mismatch_count"]]
        if mismatches:
            raise ValueError(f"Unexpected cache/source mismatch: {spec['id']} {mismatches}")
        spec["verification"] = {"raw_series_count": len(source["series"]),
                                "resolved_series_count": sum(s["validation"]["source_resolved"] for s in source["series"]),
                                "cache_source_mismatches": 0,
                                "preserved_null_values": sum(v is None for s in source["series"] for v in s["y"])}
    return specs


def main():
    specs = create_specs()
    out = ROOT / "src/charts/group_1.json"
    out.write_text(json.dumps(specs, ensure_ascii=False, indent=2, allow_nan=False) + "\n")
    print(f"Wrote {len(specs)} chart specifications to {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
