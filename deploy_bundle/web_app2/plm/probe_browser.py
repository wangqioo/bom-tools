"""Probe PLM browser automation step by step.

This script is intentionally diagnostic. It does not download attachments by
default; it opens the browser, runs small PLM actions, and saves screenshots plus
JSON snapshots for every stage.
"""

from __future__ import annotations

import argparse
import getpass
import json
import re
import time
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

try:
    from .automation import (
        PLM_SEARCH_URL,
        START_URL,
        _click_first_search_result,
        _click_text_any_page,
        _click_detail_content,
        _dump_page_debug,
        _find_search_input,
        _search_result_visible,
        _type_top_search,
        _wait_for_plm_home,
        click_opening_page,
        login_if_present,
        wait_for_eip_ready,
    )
except ImportError:
    from automation import (
        PLM_SEARCH_URL,
        START_URL,
        _click_first_search_result,
        _click_text_any_page,
        _click_detail_content,
        _dump_page_debug,
        _find_search_input,
        _search_result_visible,
        _type_top_search,
        _wait_for_plm_home,
        click_opening_page,
        login_if_present,
        wait_for_eip_ready,
    )


def output_dir() -> Path:
    path = Path("web_app2") / "outputs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def debug_dir() -> Path:
    path = output_dir() / "plm_browser_probe"
    path.mkdir(parents=True, exist_ok=True)
    return path


def stage_log(stage: str, message: str) -> None:
    print(f"[{stage}] {message}", flush=True)


def dump_stage(page, stage: str) -> None:
    out = output_dir()
    _dump_page_debug(page, out, f"probe_{stage}")
    summary = {
        "stage": stage,
        "url": page.url,
        "title": "",
        "pages": [],
        "frames": [],
    }
    try:
        summary["title"] = page.title()
    except PlaywrightError:
        pass
    for idx, opened in enumerate(page.context.pages):
        try:
            summary["pages"].append({"index": idx, "url": opened.url, "title": opened.title()})
        except PlaywrightError:
            summary["pages"].append({"index": idx, "url": opened.url})
    for idx, frame in enumerate(page.frames):
        try:
            summary["frames"].append(
                {
                    "index": idx,
                    "url": frame.url,
                    "text_sample": frame.evaluate(
                        "() => document.body ? document.body.innerText.slice(0, 1200) : ''"
                    ),
                }
            )
        except PlaywrightError as exc:
            summary["frames"].append({"index": idx, "url": frame.url, "error": str(exc)})
    stamp = time.strftime("%Y%m%d_%H%M%S")
    (debug_dir() / f"{stamp}_{stage}_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def find_pdf_links(page) -> list[dict[str, str]]:
    links: list[dict[str, str]] = []
    for frame in page.frames:
        try:
            frame_links = frame.evaluate(
                r"""() => Array.from(document.querySelectorAll('a'))
                    .map((a) => ({
                        href: a.href || '',
                        text: (a.innerText || a.textContent || '').trim(),
                        title: a.getAttribute('title') || ''
                    }))
                    .filter((item) => /\.pdf(\?|$)/i.test(item.href) || /\.pdf$/i.test(item.text) || /\.pdf$/i.test(item.title))"""
            )
            if frame_links:
                links.extend(frame_links)
        except PlaywrightError:
            continue
    return links

def find_attachment_controls(page) -> list[dict[str, object]]:
    controls: list[dict[str, object]] = []
    script = r"""() => {
        const visible = (el) => {
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
        };
        const clean = (text) => (text || '').replace(/\s+/g, ' ').trim();
        return Array.from(document.querySelectorAll('*')).map((el, index) => {
            const rect = el.getBoundingClientRect();
            const text = clean(el.innerText || el.textContent || el.value || '');
            const ownText = clean(Array.from(el.childNodes).filter((n) => n.nodeType === Node.TEXT_NODE).map((n) => n.textContent).join(' '));
            const cls = typeof el.className === 'string' ? el.className : '';
            const href = el.href || el.getAttribute('href') || '';
            const src = el.src || el.getAttribute('src') || '';
            const title = el.getAttribute('title') || '';
            const aria = el.getAttribute('aria-label') || '';
            const role = el.getAttribute('role') || '';
            const type = el.getAttribute('type') || '';
            const tag = el.tagName;
            const combined = [text, ownText, cls, href, src, title, aria, role, type, tag].join(' ');
            const relevant = /\.pdf|checkbox|checker|x-grid|download|wtcore|netmarkets|table|grid/i.test(combined) ||
                ['INPUT', 'IMG', 'A', 'BUTTON'].includes(tag);
            if (!relevant) return null;
            return {
                index, tag, id: el.id || '', cls, role, type, text: text.slice(0, 300), ownText: ownText.slice(0, 300),
                title, aria, href, src,
                checked: typeof el.checked === 'boolean' ? el.checked : null,
                visible: visible(el),
                rect: { left: Math.round(rect.left), top: Math.round(rect.top), width: Math.round(rect.width), height: Math.round(rect.height) }
            };
        }).filter(Boolean);
    }"""
    for frame_index, frame in enumerate(page.frames):
        try:
            items = frame.evaluate(script) or []
            for item in items:
                item["frame_index"] = frame_index
                item["frame_url"] = frame.url
                controls.append(item)
        except PlaywrightError as exc:
            controls.append({"frame_index": frame_index, "frame_url": frame.url, "error": str(exc)})
    return controls

def probe(username: str, password: str, hqpn: str, *, direct_search: bool, keep_open_ms: int) -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            stage_log("open_eip", START_URL)
            page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
            dump_stage(page, "01_eip_loaded")

            stage_log("login", "checking SSO and waiting for EIP")
            wait_for_eip_ready(page, username, password)
            dump_stage(page, "02_eip_ready")

            stage_log("open_plm", "clicking PLM entry")
            plm_page = click_opening_page(
                page,
                page.locator("a").filter(has_text=re.compile(r"^PLM$")),
                timeout=30000,
            )
            login_if_present(plm_page, username, password, timeout=1000)
            plm_page = _wait_for_plm_home(context, plm_page, username, password)
            dump_stage(plm_page, "03_plm_home")

            if direct_search:
                stage_log("search_page", f"direct goto {PLM_SEARCH_URL}")
                search_page = plm_page
                search_page.goto(PLM_SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
                search_page.wait_for_timeout(15000)
            else:
                stage_log("search_page", "opening through feature map")
                _click_text_any_page(context, "\u529f\u80fd\u5730\u56fe")
                search_page = _click_text_any_page(context, "\u641c\u7d22")
                search_page.wait_for_timeout(3000)
            dump_stage(search_page, "04_search_page")

            stage_log("search_input", "detecting search input")
            locator = _find_search_input(search_page)
            if locator is None:
                raise RuntimeError("search input not found")
            try:
                box = locator.bounding_box(timeout=3000)
            except PlaywrightError:
                box = None
            stage_log("search_input", f"found, box={box}")

            stage_log("search", f"searching {hqpn}")
            _type_top_search(search_page, hqpn)
            search_page.wait_for_timeout(3000)
            dump_stage(search_page, "05_after_search")
            if not _search_result_visible(search_page, hqpn):
                raise RuntimeError(f"search result text not visible: {hqpn}")

            stage_log("result", "opening first matching result")
            detail_page = _click_first_search_result(context, search_page, hqpn)
            detail_page.wait_for_timeout(3000)
            dump_stage(detail_page, "06_detail_page")

            stage_log("content", "opening content tab")
            _click_detail_content(detail_page)
            detail_page.wait_for_timeout(3000)
            dump_stage(detail_page, "07_content_tab")

            controls = find_attachment_controls(detail_page)
            controls_file = debug_dir() / f"{time.strftime('%Y%m%d_%H%M%S')}_attachment_controls.json"
            controls_file.write_text(json.dumps(controls, ensure_ascii=False, indent=2), encoding="utf-8")
            stage_log("controls", f"found {len(controls)} relevant controls; saved {controls_file}")
            for item in controls[:30]:
                stage_log("controls", json.dumps(item, ensure_ascii=False))

            pdfs = find_pdf_links(detail_page)
            stage_log("pdf", f"found {len(pdfs)} PDF links")
            for item in pdfs[:20]:
                stage_log("pdf", json.dumps(item, ensure_ascii=False))

            stage_log("done", f"keeping browser open for {keep_open_ms} ms")
            detail_page.wait_for_timeout(keep_open_ms)
        except Exception as exc:
            stage_log("failed", str(exc))
            try:
                active = [p for p in context.pages if not p.is_closed()][-1]
                dump_stage(active, "99_failed")
                active.wait_for_timeout(keep_open_ms)
            except Exception:
                pass
            raise
        finally:
            context.close()
            browser.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Step-by-step PLM browser automation probe.")
    parser.add_argument("--username", required=True, help="EIP/PLM username")
    parser.add_argument("--password", help="EIP/PLM password; prompted if omitted")
    parser.add_argument("--hqpn", required=True, help="HQ material number")
    parser.add_argument(
        "--direct-search",
        action="store_true",
        help="Go directly to the PLM search hash URL instead of clicking feature map/search.",
    )
    parser.add_argument("--keep-open-ms", type=int, default=30000, help="Keep browser open after success/failure.")
    args = parser.parse_args()

    password = args.password or getpass.getpass("Password: ")
    probe(args.username, password, args.hqpn, direct_search=args.direct_search, keep_open_ms=args.keep_open_ms)


if __name__ == "__main__":
    main()





