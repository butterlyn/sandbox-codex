#!/usr/bin/env python3
"""Exercise the integrated explorer in Chromium, with external requests blocked.

The managed Chromium policy blocks file:// navigation. A temporary loopback-only
HTTP server serves the same self-contained index.html; external dependencies are
still denied and recorded. No publication, source links, or remote services are
invoked by this audit.
"""

from __future__ import annotations

import argparse
import contextlib
import csv
import datetime as dt
import functools
import hashlib
import http.server
import io
import json
import math
from pathlib import Path
import threading
import time
import traceback
from urllib.parse import urlsplit

from playwright.sync_api import sync_playwright


ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS = ROOT / "artifacts"


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *_args):
        pass


@contextlib.contextmanager
def local_server():
    handler = functools.partial(QuietHandler, directory=str(ROOT))
    server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/index.html"
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=3)


class Audit:
    def __init__(self, expected_count):
        self.result = {
            "started_at": dt.datetime.now(dt.timezone.utc).isoformat(),
            "index_sha256": hashlib.sha256((ROOT / "index.html").read_bytes()).hexdigest(),
            "expected_figure_count": expected_count,
            "browser": "Chromium /usr/bin/chromium via Python Playwright",
            "transport": "temporary loopback HTTP server",
            "verification_limit": (
                "Direct file:// navigation is blocked by the managed browser's administrative policy "
                "(ERR_BLOCKED_BY_ADMINISTRATOR). The identical single HTML file was tested over "
                "loopback HTTP with all external requests denied."
            ),
            "external_requests": [], "local_requests": [],
            "page_errors": [], "console_errors": [],
            "figures": [], "checks": [], "screenshots": [],
        }
        self.current_figure = "initialization"

    def check(self, name, callback):
        start = time.monotonic()
        try:
            details = callback()
            self.result["checks"].append({"name": name, "status": "passed", "seconds": round(time.monotonic() - start, 2), "details": details})
            print(f"PASS {name}", flush=True)
            return details
        except Exception as error:
            self.result["checks"].append({"name": name, "status": "failed", "seconds": round(time.monotonic() - start, 2), "error": str(error), "traceback": traceback.format_exc(limit=3)})
            print(f"FAIL {name}: {error}", flush=True)
            return None

    def screenshot(self, page, filename, *, full_page=False):
        # Capture the settled layout, including the mobile drawer transition.
        page.wait_for_function("innerWidth > 800 || document.querySelector('#sidebar').classList.contains('open') || document.querySelector('#sidebar').getBoundingClientRect().right <= 0.1")
        if page.locator("#toast").is_visible():
            page.locator("#toast").wait_for(state="hidden", timeout=7000)
        for close in page.locator(".plotly-notifier .notifier-close").all():
            close.click(force=True)
        target = ARTIFACTS / filename
        page.screenshot(path=str(target), full_page=full_page)
        self.result["screenshots"].append(str(target.relative_to(ROOT)))

    def select(self, page, figure_id):
        self.current_figure = figure_id
        returned = page.evaluate("""async id => await Promise.race([
            window.WEM_APP.selectFigure(id, {updateHash:false}),
            new Promise((_,reject)=>setTimeout(()=>reject(new Error('Figure rendering exceeded 90 seconds')),90000))
        ])""", figure_id)
        assert returned == figure_id, (figure_id, returned)
        page.wait_for_function("id => window.WEM_APP.lastRendered === id", arg=figure_id)

    def attach(self, page, context, url):
        allowed_origin = urlsplit(url).netloc

        def route_request(route):
            request_url = route.request.url
            parsed = urlsplit(request_url)
            if parsed.scheme in ("http", "https", "ws", "wss") and parsed.netloc != allowed_origin:
                self.result["external_requests"].append({"url": request_url, "figure": self.current_figure, "action": "aborted"})
                route.abort("blockedbyclient")
            else:
                if parsed.scheme in ("http", "https"):
                    self.result["local_requests"].append(request_url)
                route.continue_()

        context.route("**/*", route_request)
        page.on("pageerror", lambda error: self.result["page_errors"].append({"figure": self.current_figure, "error": str(error)}))
        page.on("console", lambda message: self.result["console_errors"].append({"figure": self.current_figure, "text": message.text}) if message.type == "error" else None)

    def run(self, page, context, url):
        self.attach(page, context, url)
        page.goto(url, wait_until="load", timeout=90000)
        page.wait_for_function("typeof window.WEM_READY !== 'undefined'", timeout=90000)
        assert page.evaluate("async () => await window.WEM_READY") is True
        figures = page.evaluate("window.WEM_APP.state.figures.map(f=>({id:f.id,kind:f.kind,sheet:f.sheet,title:f.title}))")
        self.result["figure_count"] = len(figures)
        self.result["kind_counts"] = {kind: sum(f["kind"] == kind for f in figures) for kind in ("interactive", "reference", "unavailable")}
        self.check("complete inventory loaded", lambda: self.check_inventory(page, figures))
        self.screenshot(page, "qa-desktop-first-chart.png", full_page=True)

        screenshot_ids = {
            "chart-35": "qa-multiple-axes.png",
            "supplement-ch6-f42": "qa-box-plot.png",
            "supplement-ch5-f40": "qa-availability.png",
            "chart-21": "qa-source-issue.png",
        }
        for figure in figures:
            start = time.monotonic()
            entry = {"id": figure["id"], "kind": figure["kind"], "sheet": figure["sheet"]}
            try:
                self.select(page, figure["id"])
                entry.update(self.inspect_figure(page, figure))
                entry["status"] = "passed"
                if figure["id"] in screenshot_ids:
                    page.locator(".figure-card").scroll_into_view_if_needed()
                    self.screenshot(page, screenshot_ids[figure["id"]])
            except Exception as error:
                entry.update({"status": "failed", "error": str(error)})
                print(f"FAIL figure {figure['id']}: {error}", flush=True)
            entry["seconds"] = round(time.monotonic() - start, 2)
            self.result["figures"].append(entry)
            if len(self.result["figures"]) % 10 == 0:
                print(f"Rendered {len(self.result['figures'])}/{len(figures)} figures", flush=True)

        self.check("hover shows underlying observation", lambda: self.check_hover(page))
        self.check("legend swatches match plotted marker colors", lambda: self.check_legend_colors(page))
        self.check("legend toggle, isolate and show-all", lambda: self.check_legend(page))
        self.check("drag zoom and reset restore original view", lambda: self.check_zoom_reset(page))
        self.check("search and type filters", lambda: self.check_filters(page))
        self.check("figure navigation and permalink", lambda: self.check_navigation(page))
        self.check("data table paginates all observations", lambda: self.check_table(page))
        self.check("CSV exports exact unsampled figure data", lambda: self.check_csv(page))
        self.check("PNG export downloads a real image", lambda: self.check_png(page))
        self.check("reference image zoom and fit", lambda: self.check_reference(page))
        self.check("source issue disables unavailable actions", lambda: self.check_unavailable(page))
        self.check("mobile layout and navigation", lambda: self.check_mobile(page))
        self.check("no external runtime dependencies requested", self.check_external)
        self.check("no browser JavaScript errors", self.check_errors)

    def check_inventory(self, page, figures):
        assert len(figures) == self.result["expected_figure_count"], (len(figures), self.result["expected_figure_count"])
        assert len({figure["id"] for figure in figures}) == len(figures)
        assert page.locator("#figure-list .figure-link").count() == len(figures)
        assert self.result["kind_counts"] == {"interactive": 80, "reference": 15, "unavailable": 2}, self.result["kind_counts"]
        return self.result["kind_counts"]

    def inspect_figure(self, page, figure):
        assert page.locator("#figure-title").inner_text().strip()
        if figure["kind"] == "reference":
            page.wait_for_function("(() => {const image=document.getElementById('reference-image');return image && image.complete && image.naturalWidth>0;})()")
            return page.locator("#reference-image").evaluate("image => ({image_width:image.naturalWidth,image_height:image.naturalHeight})")
        if figure["kind"] == "unavailable":
            assert page.locator("#unavailable").is_visible()
            assert "broken" in page.locator("#unavailable").inner_text().lower()
            return {"source_issue_displayed": True}
        details = page.locator("#plot").evaluate(r"""plot => {
            const numeric = value => typeof value==='number' && Number.isFinite(value);
            const flatten = value => Array.isArray(value) ? value.flat(Infinity) : ArrayBuffer.isView(value) ? Array.from(value) : [];
            const valid = (plot.data||[]).reduce((sum,trace)=>sum+['x','y','values','z','median','q1','q3'].reduce((n,key)=>n+flatten(trace[key]).filter(numeric).length,0),0);
            const calcValid=(plot.calcdata||[]).flat().filter(point=>numeric(point.x)||numeric(point.y)||numeric(point.med)||numeric(point.s)||numeric(point.v)).length;
            const geometry=plot.querySelectorAll('.scatterlayer path.js-line,.scatterlayer path.point,.barlayer path,.boxlayer path,.pielayer path.surface,.heatmaplayer image,.contourlayer path').length;
            const fullTraces=(plot._fullData||[]).filter(trace=>trace.visible!==false&&trace.visible!=='legendonly');
            const webgl=fullTraces.some(trace=>/gl$/.test(trace.type)) && plot.querySelectorAll('canvas').length>0;
            return {input_traces:(plot.data||[]).length,full_traces:fullTraces.length,numeric_values:valid,calculated_points:calcValid,svg_count:plot.querySelectorAll('svg.main-svg').length,geometry_count:geometry,webgl,
                axes:Object.fromEntries(Object.entries(plot._fullLayout||{}).filter(([key])=>/^[xy]axis\d*$/.test(key)).map(([key,axis])=>[key,{type:axis.type,range:axis.range}]))};
        }""")
        assert details["input_traces"] > 0 and details["numeric_values"] > 0, details
        assert details["full_traces"] > 0 and details["svg_count"] > 0, details
        assert details["geometry_count"] > 0 or details["webgl"], details
        assert details["calculated_points"] > 0 or details["webgl"], details
        if figure["id"] == "chart-01":
            assert math.isclose(details["axes"]["yaxis"]["range"][1], 7500, abs_tol=1e-7), details["axes"]
            details["fixed_axis_bound_verified"] = "y maximum 7,500 MW"
        if figure["id"] == "chart-38":
            assert math.isclose(details["axes"]["yaxis"]["range"][0], 1000000, abs_tol=1e-7), details["axes"]
            details["fixed_axis_bound_verified"] = "y minimum 1,000,000"
        if figure["id"] == "chart-35":
            categories = page.locator("#plot").evaluate("plot=>({categories:plot._fullLayout.xaxis._categories,traces:plot._fullData.map(trace=>trace.x)})")
            expected = [f"{year}-{str(year + 1)[-2:]}" for year in range(2017, 2029)]
            assert categories["categories"] == expected, categories
            assert all(values == expected for values in categories["traces"]), categories
            details["shared_categories_verified"] = expected
        return details

    def check_hover(self, page):
        self.select(page, "chart-50")
        point = page.locator("#plot .scatterlayer path.point").first
        point.hover(force=True)
        page.wait_for_function("document.querySelector('#plot .hoverlayer')?.textContent.includes('kWh')")
        text = page.locator("#plot .hoverlayer").text_content()
        assert "°C" in text and "kWh" in text, text
        page.mouse.move(5, 5)
        return {"tooltip": text[:300]}

    def check_legend(self, page):
        self.select(page, "chart-01")
        button = page.locator("#legend [data-trace]").first
        index = int(button.get_attribute("data-trace"))
        button.click()
        page.wait_for_function("i => document.querySelector('#plot').data[i].visible==='legendonly'", arg=index)
        assert button.get_attribute("aria-pressed") == "false"
        button.click()
        page.wait_for_function("i => document.querySelector('#plot').data[i].visible===true", arg=index)
        button.click(modifiers=["Shift"])
        page.wait_for_function("document.querySelector('#plot').data.some(t=>t.visible==='legendonly')")
        page.locator("#show-all-series").click()
        page.wait_for_function("!document.querySelector('#plot').data.some(t=>t.showlegend!==false && (t.visible===false||t.visible==='legendonly'))")
        # Figure 2 has repeated scenario components grouped under shared legend
        # labels. Show all must restore the hidden legendgroup companions too.
        self.select(page, "chart-02")
        grouped = page.locator("#plot").evaluate("""plot => {
            const first=plot.data.findIndex(trace=>trace.showlegend!==false && trace.legendgroup && plot.data.filter(t=>t.legendgroup===trace.legendgroup).length>1);
            return {first,group:first<0?null:plot.data[first].legendgroup,indices:first<0?[]:plot.data.map((trace,index)=>trace.legendgroup===plot.data[first].legendgroup?index:-1).filter(index=>index>=0)};
        }""")
        assert grouped["first"] >= 0 and len(grouped["indices"]) > 1, grouped
        group_button = page.locator(f'#legend [data-trace="{grouped["first"]}"]')
        group_button.click()
        page.wait_for_function("indices=>indices.every(i=>document.querySelector('#plot').data[i].visible==='legendonly')", arg=grouped["indices"])
        group_button.click(modifiers=["Shift"])
        page.wait_for_function("indices=>indices.every(i=>document.querySelector('#plot').data[i].visible===true)", arg=grouped["indices"])
        page.locator("#show-all-series").click()
        page.wait_for_function("!document.querySelector('#plot').data.some(trace=>(trace.showlegend!==false || trace.legendgroup) && (trace.visible===false||trace.visible==='legendonly'))")
        return {"toggled_trace": index, "isolate_and_restore": True, "scenario_group": grouped}

    def check_legend_colors(self, page):
        self.select(page, "chart-50")
        colors = page.locator("#plot").evaluate("""plot => {
            const rgb = value => (value.match(/[\\d.]+/g)||[]).slice(0,3).map(Number);
            const traces=[...plot.querySelectorAll('.scatterlayer .trace')];
            return plot.data.map((trace,index)=>{
                if (trace.mode!=='markers') return null;
                const point=traces[index]?.querySelector('path.point');
                const swatch=document.querySelector('#legend [data-trace="'+index+'"] .legend-swatch');
                if (!point || !swatch) return {name:trace.name,missing:true};
                const pointColor=getComputedStyle(point).fill, legendColor=getComputedStyle(swatch).backgroundColor;
                return {name:trace.name,pointColor,legendColor,match:JSON.stringify(rgb(pointColor))===JSON.stringify(rgb(legendColor))};
            }).filter(Boolean);
        }""")
        assert len(colors) == 5 and all(color.get("match") for color in colors), colors
        return colors

    def check_zoom_reset(self, page):
        self.select(page, "chart-50")
        before = page.locator("#plot").evaluate("plot=>plot._fullLayout.xaxis.range")
        box = page.locator("#plot .nsewdrag").first.bounding_box()
        assert box and box["width"] > 100
        page.mouse.move(box["x"] + box["width"] * .25, box["y"] + box["height"] * .2)
        page.mouse.down()
        page.mouse.move(box["x"] + box["width"] * .7, box["y"] + box["height"] * .7, steps=12)
        page.mouse.up()
        page.wait_for_function("before => JSON.stringify(document.querySelector('#plot')._fullLayout.xaxis.range)!==JSON.stringify(before)", arg=before)
        zoomed = page.locator("#plot").evaluate("plot=>plot._fullLayout.xaxis.range")
        page.locator("#reset").click()
        page.wait_for_function("before => JSON.stringify(document.querySelector('#plot')._fullLayout.xaxis.range)===JSON.stringify(before)", arg=before)
        return {"original_range": before, "zoomed_range": zoomed}

    def check_filters(self, page):
        page.locator("#search").fill("temperature")
        count = page.locator("#figure-list .figure-link").count()
        assert 0 < count < self.result["figure_count"], count
        assert "temperature" in page.locator("#figure-list").inner_text().lower()
        page.locator("#kind").select_option("unavailable")
        assert page.locator("#figure-list .figure-link").count() == 0
        assert "No figures match" in page.locator("#figure-list").inner_text()
        page.locator("#search").fill("")
        assert page.locator("#figure-list .figure-link").count() == 2
        page.locator("#clear-filters").click()
        assert page.locator("#figure-list .figure-link").count() == self.result["figure_count"]
        page.locator("#section").select_option(label="Executive summary")
        assert page.locator("#figure-list .figure-link").count() == 2
        page.locator("#clear-filters").click()
        return {"temperature_results": count, "source_issue_results": 2, "executive_summary_results": 2}

    def check_navigation(self, page):
        page.locator('#figure-list [data-figure="chart-01"]').click()
        page.wait_for_function("window.WEM_APP.lastRendered==='chart-01'")
        assert page.url.endswith("#chart-01"), page.url
        page.locator("#next-figure").click()
        page.wait_for_function("window.WEM_APP.lastRendered==='chart-02'")
        page.locator("#previous-figure").click()
        page.wait_for_function("window.WEM_APP.lastRendered==='chart-01'")
        assert page.locator("#figure-permalink").get_attribute("href") == "#chart-01"
        return {"next": "chart-02", "previous": "chart-01", "permalink": "#chart-01"}

    def check_table(self, page):
        self.select(page, "chart-50")
        page.locator("#tab-data").click()
        page.locator("#data-series").select_option(label="2020-21 data")
        rows = page.locator("#data-table tbody tr")
        assert rows.count() == 100
        assert rows.first.locator("td").first.inner_text() == "1"
        first_y = float(rows.first.locator("td").nth(2).inner_text())
        assert math.isclose(first_y, 13.020458570757599, rel_tol=1e-12)
        page.locator("#data-next").click()
        assert rows.first.locator("td").first.inner_text() == "101"
        page.locator("#data-prev").click()
        assert rows.first.locator("td").first.inner_text() == "1"
        for _ in range(3):
            page.locator("#data-next").click()
        assert rows.count() == 65
        assert rows.first.locator("td").first.inner_text() == "301"
        assert rows.last.locator("td").first.inner_text() == "365"
        assert page.locator("#data-next").is_disabled()
        page.locator("#tab-chart").click()
        return {"series": "2020-21 data", "total_observations": 365, "last_page_rows": 65, "first_y": first_y}

    def check_csv(self, page):
        self.select(page, "chart-50")
        with page.expect_download(timeout=30000) as event:
            page.locator("#export-csv").click()
        download = event.value
        target = ARTIFACTS / "qa-chart-50.csv"
        download.save_as(str(target))
        rows = list(csv.DictReader(io.StringIO(target.read_text(encoding="utf-8-sig"))))
        assert len(rows) == 2026, len(rows)
        assert len({row["series"] for row in rows}) == 10
        data = [row for row in rows if row["series"] == "2020-21 data"]
        model = [row for row in rows if row["series"] == "2020-21 model"]
        assert len(data) == 365 and len(model) == 40
        assert math.isclose(float(data[0]["x"]), 18.3979166666667, rel_tol=1e-12)
        assert math.isclose(float(data[0]["y"]), 13.020458570757599, rel_tol=1e-12)
        assert float(model[0]["x"]) == 1 and float(model[-1]["x"]) == 40
        return {"filename": download.suggested_filename, "data_rows": len(rows), "series": 10, "artifact": str(target.relative_to(ROOT))}

    def check_png(self, page):
        self.select(page, "chart-01")
        with page.expect_download(timeout=90000) as event:
            page.locator("#export-image").click()
        download = event.value
        target = ARTIFACTS / "qa-chart-01-export.png"
        download.save_as(str(target))
        content = target.read_bytes()
        assert content.startswith(b"\x89PNG\r\n\x1a\n") and len(content) > 20000
        width = int.from_bytes(content[16:20], "big")
        height = int.from_bytes(content[20:24], "big")
        assert width >= 1500 and height >= 850, (width, height)
        page.wait_for_function("document.querySelector('#plot').layout.showlegend === false")
        return {"filename": download.suggested_filename, "bytes": len(content), "width": width, "height": height, "artifact": str(target.relative_to(ROOT))}

    def check_reference(self, page):
        self.select(page, "supplement-ch1-f3")
        page.wait_for_function("document.getElementById('reference-image').complete && document.getElementById('reference-image').naturalWidth>0")
        before = page.locator("#reference-image").evaluate("image=>image.getBoundingClientRect().width")
        page.locator("#image-plus").click()
        after = page.locator("#reference-image").evaluate("image=>image.getBoundingClientRect().width")
        assert after > before * 1.2, (before, after)
        assert page.locator("#image-scale").inner_text() == "125%"
        page.locator("#image-minus").click()
        assert page.locator("#image-scale").inner_text() == "Fit to width"
        page.locator("#image-plus").click()
        page.locator("#reset").click()
        assert page.locator("#image-scale").inner_text() == "Fit to width"
        return {"fit_width": before, "zoomed_width": after}

    def check_unavailable(self, page):
        self.select(page, "chart-21")
        for selector in ("#export-csv", "#export-image", "#reset", "#tab-data"):
            assert page.locator(selector).is_disabled(), selector
        page.locator("#tab-source").click()
        assert "broken" in page.locator("#panel-source").inner_text().lower()
        page.locator("#tab-chart").click()
        return {"chart": "chart-21", "disabled_actions": ["CSV", "PNG", "Reset", "Data"]}

    def check_mobile(self, page):
        page.set_viewport_size({"width": 390, "height": 844})
        self.select(page, "chart-01")
        page.evaluate("window.scrollTo(0,0)")
        page.wait_for_function("document.querySelector('#plot').getBoundingClientRect().width < 390")
        assert page.locator("#menu-button").is_visible()
        overflow = page.evaluate("({viewport:innerWidth,document:document.documentElement.scrollWidth,body:document.body.scrollWidth})")
        assert overflow["document"] <= overflow["viewport"] + 1, overflow
        self.screenshot(page, "qa-mobile-first-chart.png", full_page=True)
        page.locator("#menu-button").click()
        assert page.locator("#menu-button").get_attribute("aria-expanded") == "true"
        page.locator("#search").fill("temperature")
        page.locator('#figure-list [data-figure="chart-50"]').click()
        page.wait_for_function("window.WEM_APP.lastRendered==='chart-50'")
        assert page.locator("#menu-button").get_attribute("aria-expanded") == "false"
        assert page.locator("#scrim").is_hidden()
        overflow = page.evaluate("({viewport:innerWidth,document:document.documentElement.scrollWidth})")
        assert overflow["document"] <= overflow["viewport"] + 1, overflow
        self.screenshot(page, "qa-mobile-temperature.png", full_page=True)
        return {"viewport": [390, 844], "document_width": overflow["document"], "menu_navigation": "chart-50"}

    def check_external(self):
        assert not self.result["external_requests"], self.result["external_requests"]
        return {"external_requests": 0, "local_request_count": len(self.result["local_requests"])}

    def check_errors(self):
        assert not self.result["page_errors"], self.result["page_errors"]
        assert not self.result["console_errors"], self.result["console_errors"]
        return {"page_errors": 0, "console_errors": 0}

    def save(self):
        self.result["index_unchanged_during_test"] = self.result["index_sha256"] == hashlib.sha256((ROOT / "index.html").read_bytes()).hexdigest()
        if not self.result["index_unchanged_during_test"]:
            self.result["checks"].append({"name": "built HTML unchanged during audit", "status": "failed", "error": "index.html was rebuilt during this run; rerun against the final artifact"})
        failed_checks = [check for check in self.result["checks"] if check["status"] == "failed"]
        failed_figures = [figure for figure in self.result["figures"] if figure["status"] == "failed"]
        self.result["finished_at"] = dt.datetime.now(dt.timezone.utc).isoformat()
        self.result["failed_checks"] = len(failed_checks)
        self.result["failed_figures"] = len(failed_figures)
        self.result["status"] = "passed" if not (failed_checks or failed_figures or self.result.get("fatal_error")) else "failed"
        target = ROOT / "data/analysis/browser-check.json"
        target.write_text(json.dumps(self.result, ensure_ascii=False, indent=2) + "\n")
        print(json.dumps({"status": self.result["status"], "figures": len(self.result["figures"]), "failed_figures": len(failed_figures), "failed_checks": len(failed_checks), "index_unchanged_during_test": self.result["index_unchanged_during_test"]}), flush=True)
        return self.result["status"] == "passed"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--expected-count", type=int, default=97)
    args = parser.parse_args()
    ARTIFACTS.mkdir(exist_ok=True)
    audit = Audit(args.expected_count)
    try:
        with local_server() as url, sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=True, executable_path="/usr/bin/chromium", args=["--no-sandbox"])
            context = browser.new_context(viewport={"width": 1440, "height": 1000}, device_scale_factor=1, accept_downloads=True)
            page = context.new_page()
            page.set_default_timeout(20000)
            audit.run(page, context, url)
            browser.close()
    except Exception as error:
        audit.result["fatal_error"] = str(error)
        audit.result["fatal_traceback"] = traceback.format_exc()
        print(traceback.format_exc(), flush=True)
    success = audit.save()
    raise SystemExit(0 if success else 1)


if __name__ == "__main__":
    main()
