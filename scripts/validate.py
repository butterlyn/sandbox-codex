"""Verify source integrity, complete workbook coverage and native chart values."""
from collections import Counter
import hashlib
import json
import math
from pathlib import Path
from zipfile import ZipFile
from chart_common import load_chart

ROOT = Path(__file__).resolve().parents[1]


def validate():
    sources = json.loads((ROOT / "data/sources/sources.json").read_text())
    for source in sources:
        content = (ROOT / "data/sources" / source["filename"]).read_bytes()
        assert len(content) == source["bytes"], source["filename"]
        assert hashlib.sha256(content).hexdigest() == source["sha256"], source["filename"]
    inventory = json.loads((ROOT / "data/analysis/inventory.json").read_text())
    with ZipFile(ROOT / "data/sources/2026-wem-esoo-data-register.xlsx") as archive:
        chart_parts = {n for n in archive.namelist() if n.startswith("xl/charts/chart") and n.endswith(".xml")}
    assert chart_parts == {chart["xml_part"] for chart in inventory["charts"]}
    figures = []
    for path in sorted((ROOT / "src/charts").glob("*.json")):
        figures.extend(json.loads(path.read_text(), parse_constant=lambda value: (_ for _ in ()).throw(ValueError(value))))
    ids = [f["id"] for f in figures]
    assert len(ids) == len(set(ids)), "Duplicate figure identifiers"
    native = [f for f in figures if f.get("chart_number") is not None]
    assert sorted(f["chart_number"] for f in native) == list(range(1, 69)), "Missing/duplicate native chart object"
    expected_sheets = {s["name"] for s in inventory["sheets"] if s["title"].startswith("Figure ")}
    actual_sheets = {f["sheet"] for f in figures}
    assert expected_sheets == actual_sheets, {"missing": sorted(expected_sheets - actual_sheets), "extra": sorted(actual_sheets - expected_sheets)}
    for f in figures:
        assert f["kind"] in {"interactive", "reference", "unavailable"}, f["id"]
        if f["kind"] == "interactive":
            assert f["data"], f["id"]
            for trace in f["data"]:
                horizontal = trace.get("orientation") == "h"
                values = trace.get("x" if horizontal else "y", trace.get("values", []))
                assert values, (f["id"], trace.get("name"))
                assert all(v is None or isinstance(v, (float, int)) and math.isfinite(v) for v in values), (f["id"], trace.get("name"))
                x = trace.get("y" if horizontal else "x", trace.get("labels"))
                if x is not None:
                    if x and isinstance(x[0], list):
                        assert all(len(level) == len(values) for level in x), f["id"]
                    else:
                        assert len(x) == len(values), (f["id"], trace.get("name"), len(x), len(values))
        elif f["kind"] == "unavailable":
            assert not f.get("data"), "Placeholders must not be plotted as forecast data"
            assert f.get("notes"), f["id"]
        if f.get("image_file"):
            assert (ROOT / f["image_file"]).is_file(), f["id"]
    audit = []
    for number in range(1, 69):
        raw = load_chart(number)
        for series in raw["series"]:
            check = series["validation"]
            assert check["mismatch_count"] == 0, (number, series["index"], check)
            if not check["source_resolved"]:
                assert not series["y_formula"] or "#REF!" in series["y_formula"], (number, series["y_formula"])
            audit.append({"chart": number, "series_index": series["index"], "formula": series["y_formula"],
                          "numeric_points": sum(isinstance(v, (int, float)) for v in series["y"]), **check})
    kinds = Counter(f["kind"] for f in figures)
    report = {"status": "passed", "source_files_verified": len(sources), "figure_worksheets_covered": len(expected_sheets),
              "native_chart_objects": len(native), "figures": len(figures), "kinds": dict(kinds),
              "native_series": len(audit), "native_cache_positions": sum(row["cache_points"] for row in audit),
              "numeric_cache_points": sum(row["numeric_points"] for row in audit),
              "source_resolved_series": sum(row["source_resolved"] for row in audit),
              "numeric_values_verified": sum(row["numeric_points"] for row in audit if row["source_resolved"]),
              "native_cache_source_mismatches": sum(row["mismatch_count"] for row in audit), "series_audit": audit}
    (ROOT / "data/analysis/validation.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({k: v for k, v in report.items() if k != "series_audit"}, indent=2))


if __name__ == "__main__":
    validate()
