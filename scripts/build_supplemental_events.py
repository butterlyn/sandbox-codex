#!/usr/bin/env python3
"""Recreate image-backed Appendix A4 event charts from their worksheet rows.

Source values and classifications take precedence where the embedded image and
worksheet disagree. Bubble area is proportional to supplied unserved energy.
"""

from __future__ import annotations

import collections
import datetime as dt
import json
import math
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "data/sources/2026-wem-esoo-data-register.xlsx"
CONDITIONS = {
    "Outages >= Reserve Margin MW": ("High outages", "#DF171B"),
    "Renewables <= 15% availability": ("Low renewables", "#1B67E9"),
    "Battery SoC < 50%": ("Low battery SoC", "#FFD321"),
    "Outages >= Reserve Margin MW & Battery SoC < 50%": ("High outages + low battery SoC", "#F47C16"),
    "Outages >= Reserve Margin MW & Renewables <= 15% availability": ("High outages + low renewables", "#7F4086"),
    "Renewables <= 15% availability & Battery SoC < 50%": ("Low battery SoC + low renewables", "#8C9C80"),
    "All three drivers": ("All three drivers", "#703600"),
    "Other": ("Other", "#D8D8D8"),
}


def normalized_cause(value):
    return value.replace("<50%", "< 50%")


def build_sheet(workbook, figure):
    name = f"A4_F.{figure}"
    sheet = workbook[name]
    last_row = {25: 89, 26: 3509}[figure]
    records = []
    for row_number, row in enumerate(
        sheet.iter_rows(min_row=24, max_row=last_row, min_col=2, max_col=5, values_only=True), 24
    ):
        date, demand, energy, cause = row
        assert isinstance(date, dt.datetime), (name, row_number, "date", date)
        assert isinstance(demand, (int, float)) and math.isfinite(demand)
        assert isinstance(energy, (int, float)) and math.isfinite(energy) and energy > 0
        category = normalized_cause(cause)
        assert category in CONDITIONS, (name, row_number, cause)
        records.append({
            "row": row_number, "date": date.strftime("%Y-%m-%d"),
            "demand": demand, "energy": energy, "cause": cause, "category": category,
        })

    # The raw MWh values remain in marker.size; sizeref controls only their
    # drawing scale. A minimum 3px radius keeps small events discoverable.
    size_reference = 2 * max(record["energy"] for record in records) / 52**2
    traces = []
    sources = []
    for category, (label, color) in CONDITIONS.items():
        selected = [record for record in records if record["category"] == category]
        if not selected:
            continue
        traces.append({
            "name": label, "type": "scatter", "mode": "markers",
            "x": [record["date"] for record in selected],
            "y": [record["demand"] for record in selected],
            "marker": {
                "size": [record["energy"] for record in selected],
                "sizemode": "area", "sizeref": size_reference, "sizemin": 3,
                "color": color, "opacity": 0.76,
                "line": {"color": "#30383B", "width": 0.7},
            },
            "customdata": [[record["energy"], record["cause"], record["row"]] for record in selected],
            "hovertemplate": (
                "%{x|%d %b %Y}<br>Operational demand: %{y:,.2f} MW"
                "<br>Unserved energy: %{customdata[0]:,.3f} MWh"
                "<br>%{customdata[1]}<extra>%{fullData.name}</extra>"
            ),
            "meta": {
                "sourceSheet": name, "sourceRows": [record["row"] for record in selected],
                "sourceCategory": category, "sourceValueCount": len(selected),
                "sizeUnit": "MWh", "sizeEncoding": "area", "yScaleFactor": 1,
            },
        })
        sources.append({
            "sheet": name, "range": f"B24:E{last_row}", "role": label,
            "filter": f"Cause = {category}",
        })

    layout = {
        "font": {"family": "Arial, sans-serif", "color": "#223345", "size": 12},
        "paper_bgcolor": "rgba(0,0,0,0)", "plot_bgcolor": "#FFFFFF",
        "height": 590,
        "margin": {"l": 82, "r": 30, "t": 20, "b": 140},
        "hovermode": "closest", "dragmode": "zoom",
        "legend": {
            "orientation": "h", "y": -0.24, "x": 0, "xanchor": "left", "yanchor": "top",
            "font": {"size": 11}, "itemsizing": "constant",
        },
        "xaxis": {
            "type": "date", "title": {"text": "Event date"}, "automargin": True,
            "range": ["2026-10-01", "2031-10-01"] if figure == 25 else ["2031-10-01", "2035-10-15"],
            "tickformat": "%b %Y", "dtick": "M6", "tickangle": -35,
            "showgrid": False, "fixedrange": False,
        },
        "yaxis": {
            "title": {"text": "Operational demand (MW)"}, "automargin": True,
            "range": [2800, 5800] if figure == 25 else [2000, 6400],
            "tickformat": ",.0f", "dtick": 500, "showgrid": True,
            "gridcolor": "#E5EAF0", "zeroline": False, "fixedrange": False,
        },
    }
    notes = [
        f"All {len(records):,} worksheet records are shown. Bubble area represents unserved energy (MWh); "
        "the smallest bubbles have a minimum visible size. Hover for exact values.",
        "High outages: outages ≥ reserve margin. Low renewables: availability ≤ 15%. "
        "Low battery SoC: state of charge < 50%.",
    ]
    caveats = []
    if figure == 25:
        caveats.append(
            "The workbook title and embedded image end in 2029–30, but the worksheet data continues to "
            "11 August 2031. The interactive chart includes those later records."
        )
        caveats.append(
            "Colors follow the worksheet's Cause values, which differ from the embedded image for some events."
        )
    else:
        caveats.append(
            "The source Cause column contains five categories; the embedded image legend contains eight. "
            "Colors here follow the supplied worksheet classifications."
        )
    notes.extend(caveats)
    spec = {
        "id": f"supplement-a4-f{figure}", "sheet": name,
        "title": sheet["A1"].value.strip(), "section": "Appendices",
        "kind": "interactive", "types": ["bubble"],
        "data": traces, "layout": layout, "notes": notes, "sources": sources,
        "image_file": f"data/analysis/images/image{33 if figure == 25 else 34}.png",
    }

    # Trace-level round trip protects dates, duplicate events, sizes and causes
    # from accidental aggregation or mismatched arrays during reconstruction.
    round_trip = {}
    for trace in traces:
        count = len(trace["x"])
        assert count == len(trace["y"]) == len(trace["marker"]["size"]) == len(trace["customdata"])
        for i, row_number in enumerate(trace["meta"]["sourceRows"]):
            assert row_number not in round_trip
            round_trip[row_number] = (
                trace["x"][i], trace["y"][i], trace["marker"]["size"][i], trace["customdata"][i][1]
            )
    assert len(round_trip) == len(records) == last_row - 23
    for record in records:
        assert round_trip[record["row"]] == (
            record["date"], record["demand"], record["energy"], record["cause"]
        )
    review = {
        "id": spec["id"], "sheet": name, "source_range": f"B24:E{last_row}",
        "source_records": len(records), "plotted_records": sum(len(trace["x"]) for trace in traces),
        "date_range": [min(record["date"] for record in records), max(record["date"] for record in records)],
        "demand_range_MW": [min(record["demand"] for record in records), max(record["demand"] for record in records)],
        "unserved_energy_range_MWh": [min(record["energy"] for record in records), max(record["energy"] for record in records)],
        "category_counts": dict(collections.Counter(record["category"] for record in records)),
        "marker_size_encoding": "MWh proportional to area, minimum visible size for tiny values",
        "checks": [
            "all source rows retained exactly once, including duplicate dates",
            "every plotted date, demand, MWh size and Cause round-trips to its source row",
            "all supplied categories mapped explicitly without inferred reclassification",
        ],
        "source_caveats": caveats, "status": "passed",
    }
    return spec, review


def main():
    workbook = load_workbook(WORKBOOK, read_only=True, data_only=True)
    specs = []
    reviews = []
    for figure in (25, 26):
        spec, review = build_sheet(workbook, figure)
        specs.append(spec)
        reviews.append(review)
    workbook.close()
    target = ROOT / "src/charts/supplemental_events.json"
    target.write_text(json.dumps(specs, ensure_ascii=False, separators=(",", ":"), allow_nan=False) + "\n")
    (ROOT / "data/analysis/supplemental-events-review.json").write_text(
        json.dumps({"charts": reviews, "status": "passed"}, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    )
    print(f"Built and verified 2 supplemental charts with {sum(item['source_records'] for item in reviews):,} observations")


if __name__ == "__main__":
    main()
