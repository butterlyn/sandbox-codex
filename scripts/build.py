"""Assemble every independently generated figure into one fully offline HTML file."""
import base64
import gzip
import json
import mimetypes
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]


def build():
    inventory = json.loads((ROOT / "data/analysis/inventory.json").read_text())
    order = {sheet["name"]: i for i, sheet in enumerate(inventory["sheets"])}
    figures = []
    groups = sorted((ROOT / "src/charts").glob("*.json"))
    if not groups:
        raise ValueError("No chart specs. Run the chart group generators first.")
    for path in groups:
        group = json.loads(path.read_text())
        if not isinstance(group, list):
            raise ValueError(f"{path} must contain a figure array")
        for figure in group:
            figure.setdefault("kind", "interactive")
            figure.setdefault("notes", [])
            figure.setdefault("sources", [])
            if figure.get("image_file"):
                image_path = ROOT / figure["image_file"]
                mime = mimetypes.guess_type(image_path)[0] or "image/png"
                figure["image_data"] = "data:" + mime + ";base64," + base64.b64encode(image_path.read_bytes()).decode()
            figures.append(figure)
    figures.sort(key=lambda f: (order.get(f["sheet"], 999), f.get("chart_number", 999), f["id"]))
    ids = [f["id"] for f in figures]
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate figure IDs")
    native = {f["chart_number"] for f in figures if f.get("chart_number") is not None}
    missing = set(range(1, inventory["chart_count"] + 1)) - native
    if missing:
        raise ValueError(f"Missing native chart objects: {sorted(missing)}")
    figure_sheets = {s["name"] for s in inventory["sheets"] if s["title"].startswith("Figure ")}
    uncovered = figure_sheets - {figure["sheet"] for figure in figures}
    if uncovered:
        raise ValueError(f"Missing figure worksheets: {sorted(uncovered)}")
    payload = {"title": "2026 WEM ESOO interactive data register", "figures": figures,
               "sources": json.loads((ROOT / "data/sources/sources.json").read_text())}
    serialized = json.dumps(payload, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode()
    compressed = gzip.compress(serialized, compresslevel=9, mtime=0)
    html = (ROOT / "src/index.html").read_text()
    assets = {"/*__APP_CSS__*/": (ROOT / "src/styles.css").read_text(),
              "/*__PLOTLY__*/": (ROOT / "vendor/plotly-3.1.0.min.js").read_text(),
              "/*__APP_JS__*/": (ROOT / "src/app.js").read_text(),
              "__CHART_PAYLOAD__": base64.b64encode(compressed).decode()}
    for marker, content in assets.items():
        if marker in ("/*__PLOTLY__*/", "/*__APP_JS__*/"):
            content = re.sub(r"</script", r"<\\/script", content, flags=re.I)
        html = html.replace(marker, content)
    # Arial in minimal Linux browser images lacks the diagonal-arrow glyph.
    html = html.replace("↗", "→")
    (ROOT / "index.html").write_text(html)
    summary = {"figures": len(figures), "native_chart_objects": len(native),
               "interactive": sum(f["kind"] == "interactive" for f in figures),
               "reference": sum(f["kind"] == "reference" for f in figures),
               "unavailable": sum(f["kind"] == "unavailable" for f in figures),
               "payload_bytes": len(serialized), "gzip_bytes": len(compressed), "html_bytes": len(html.encode()),
               "groups": [p.name for p in groups]}
    (ROOT / "data/analysis/build-summary.json").write_text(json.dumps(summary, indent=2) + "\n")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    build()
