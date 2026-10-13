"""Inventory every worksheet, native chart object and embedded figure in the XLSX."""
import json
from pathlib import Path
import posixpath
from zipfile import ZipFile
from lxml import etree as E

ROOT = Path(__file__).resolve().parents[1]
NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def extract():
    with ZipFile(ROOT / "data/sources/2026-wem-esoo-data-register.xlsx") as archive:
        members = set(archive.namelist())

        def relations(part):
            path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            if path not in members:
                return {}
            return {n.get("Id"): posixpath.normpath(posixpath.join(posixpath.dirname(part), n.get("Target")))
                    for n in E.fromstring(archive.read(path))}

        strings = ["".join(n.xpath(".//s:t/text()", namespaces=NS))
                   for n in E.fromstring(archive.read("xl/sharedStrings.xml"))]
        workbook_relations = relations("xl/workbook.xml")
        charts, images, sheets = [], [], []
        for sheet in E.fromstring(archive.read("xl/workbook.xml")).findall("s:sheets/s:sheet", NS):
            name = sheet.get("name")
            part = workbook_relations[sheet.get("{" + NS["r"] + "}id")]
            element = E.fromstring(archive.read(part))
            cells = {}
            for row in element.findall("s:sheetData/s:row", NS):
                if int(row.get("r")) > 15:
                    break
                for cell in row:
                    value = cell.findtext("s:v", namespaces=NS)
                    if value is not None:
                        cells[cell.get("r")] = strings[int(value)] if cell.get("t") == "s" else value
            title = cells.get("A1", name)
            counts = {"charts": 0, "images": 0}
            drawings = []
            sheet_relations = relations(part)
            for drawing in element.findall("s:drawing", NS):
                drawing_part = sheet_relations[drawing.get("{" + NS["r"] + "}id")]
                drawing_relations = relations(drawing_part)
                drawing_element = E.fromstring(archive.read(drawing_part))
                drawings.append(drawing_part)
                for chart in drawing_element.findall(".//c:chart", NS):
                    chart_part = drawing_relations[chart.get("{" + NS["r"] + "}id")]
                    number = int(chart_part.split("chart")[-1].split(".")[0])
                    chart_element = E.fromstring(archive.read(chart_part))
                    path = ROOT / f"data/analysis/chart-xml/chart{number:02d}.xml"
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(archive.read(chart_part))
                    section = "Appendices" if name.startswith("A") else ("Executive summary" if name.startswith("ES") else "Chapter " + name.split("_")[0].replace("Ch ", ""))
                    charts.append({"id": f"chart-{number:02d}", "chart_number": number, "sheet": name,
                                   "sheet_part": part, "title": title, "xml_part": chart_part,
                                   "xml_file": str(path.relative_to(ROOT)), "section": section,
                                   "types": [E.QName(n).localname for n in chart_element.findall(".//c:plotArea/*", NS) if E.QName(n).localname.endswith("Chart")],
                                   "series_count": len(chart_element.findall(".//c:ser", NS))})
                    counts["charts"] += 1
                for blip in drawing_element.findall(".//a:blip", NS):
                    relationship = blip.get("{" + NS["r"] + "}embed")
                    if relationship and relationship in drawing_relations:
                        image_part = drawing_relations[relationship]
                        destination = ROOT / "data/analysis/images" / posixpath.basename(image_part)
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        destination.write_bytes(archive.read(image_part))
                        images.append({"sheet": name, "title": title, "part": image_part, "file": str(destination.relative_to(ROOT))})
                        counts["images"] += 1
            sheets.append({"name": name, "title": title, "part": part, "top_cells": cells, "drawings": drawings, **counts})
        charts.sort(key=lambda item: item["chart_number"])
        result = {"chart_count": len(charts), "charts": charts, "sheets": sheets, "images": images}
        (ROOT / "data/analysis/inventory.json").write_text(json.dumps(result, indent=2) + "\n")
        print(f"Inventoried {len(sheets)} worksheets, {len(charts)} native charts and {len(images)} image relationships.")


if __name__ == "__main__":
    extract()
