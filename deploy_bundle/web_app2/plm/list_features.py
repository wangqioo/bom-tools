"""List visible PLM feature-map entries.

Run from the repository root:
    python web_app2/plm/list_features.py --username YOUR_ID
"""

from __future__ import annotations

import argparse
import getpass
import json
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

try:
    from .automation import (
        START_URL,
        click_opening_page,
        login_if_present,
        wait_for_eip_ready,
    )
except ImportError:
    from automation import (
        START_URL,
        click_opening_page,
        login_if_present,
        wait_for_eip_ready,
    )


def _clean_text(value: str) -> str:
    return " ".join((value or "").split())


def _extract_visible_entries(page) -> list[dict[str, str]]:
    raw_entries = page.evaluate(
        """() => Array.from(document.querySelectorAll(
            "a,button,[role='button'],[role='link'],.el-tree-node__label,.ant-tree-title,.x-tree-node-text,li,td"
        )).map((el) => ({
            text: (el.innerText || el.textContent || "").trim(),
            tag: el.tagName,
            role: el.getAttribute("role") || "",
            title: el.getAttribute("title") || "",
            href: el.getAttribute("href") || "",
            visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
        }))"""
    )

    seen: set[str] = set()
    entries: list[dict[str, str]] = []
    for item in raw_entries:
        if not item.get("visible"):
            continue
        text = _clean_text(item.get("title") or item.get("text") or "")
        if not text or len(text) < 2 or len(text) > 80:
            continue
        if text in seen:
            continue
        seen.add(text)
        entries.append(
            {
                "text": text,
                "tag": item.get("tag") or "",
                "role": item.get("role") or "",
                "href": item.get("href") or "",
            }
        )
    return entries


def _page_label(page) -> str:
    try:
        return f"title={page.title()!r} url={page.url}"
    except PlaywrightError:
        return f"url={page.url}"


def _find_text_page(context, patterns: list[str]):
    candidates = [page for page in context.pages if not page.is_closed()]
    for pattern in patterns:
        for candidate in reversed(candidates):
            locator = candidate.get_by_text(pattern, exact=False).first
            try:
                locator.wait_for(state="visible", timeout=3000)
                return candidate, locator, pattern
            except (PlaywrightTimeoutError, PlaywrightError):
                continue
    return None, None, ""


def _click_feature_map(context, fallback_page):
    patterns = ["功能地图", "功能菜单", "应用菜单", "应用中心", "菜单", "地图"]
    page, locator, pattern = _find_text_page(context, patterns)
    if not page or not locator:
        return fallback_page, False, ""
    try:
        with page.expect_popup(timeout=5000) as popup_info:
            locator.click(timeout=10000)
        opened = popup_info.value
        opened.wait_for_load_state("domcontentloaded", timeout=30000)
        return opened, True, pattern
    except PlaywrightTimeoutError:
        locator.click(timeout=10000)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=30000)
        except PlaywrightError:
            pass
        return page, True, pattern
    except PlaywrightError:
        open_pages = [candidate for candidate in context.pages if not candidate.is_closed()]
        return (open_pages[-1] if open_pages else page), True, pattern


def _write_diagnostics(context, output: Path | None) -> None:
    if not output:
        return
    diag_dir = output.parent / f"{output.stem}_diagnostics"
    diag_dir.mkdir(parents=True, exist_ok=True)
    pages = [page for page in context.pages if not page.is_closed()]
    summary = []
    for index, page in enumerate(pages, 1):
        label = _page_label(page)
        summary.append(label)
        try:
            entries = _extract_visible_entries(page)
            (diag_dir / f"page_{index}_visible_text.json").write_text(
                json.dumps(entries, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except PlaywrightError as exc:
            (diag_dir / f"page_{index}_error.txt").write_text(str(exc), encoding="utf-8")
        try:
            page.screenshot(path=str(diag_dir / f"page_{index}.png"), full_page=True)
        except PlaywrightError:
            pass
    (diag_dir / "pages.txt").write_text("\n".join(summary), encoding="utf-8")
    print(f"已保存诊断信息：{diag_dir}")


def list_plm_features(username: str, password: str, *, output: Path | None, headless: bool) -> list[dict[str, str]]:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            print("打开 EIP...")
            page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
            wait_for_eip_ready(page, username, password)

            print("进入 PLM...")
            plm_page = click_opening_page(
                page,
                page.locator("a").filter(has_text=r"PLM"),
                timeout=30000,
            )
            login_if_present(plm_page, username, password, timeout=500)
            print("当前已打开页面：")
            for opened_page in context.pages:
                if not opened_page.is_closed():
                    print(f"  - {_page_label(opened_page)}")

            print("打开功能地图...")
            feature_page, clicked, matched_text = _click_feature_map(context, plm_page)
            if clicked:
                print(f"已点击：{matched_text}")
            else:
                print("未找到可见的功能地图入口，改为导出当前页面可见文本。")
                _write_diagnostics(context, output)
            try:
                feature_page.wait_for_load_state("domcontentloaded", timeout=30000)
            except PlaywrightError:
                pass
            feature_page.wait_for_timeout(3000)

            entries = _extract_visible_entries(feature_page)
        finally:
            context.close()
            browser.close()

    if output:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"已保存：{output}")

    return entries


def main() -> int:
    parser = argparse.ArgumentParser(description="List visible PLM feature-map entries.")
    parser.add_argument("--username", "-u", help="EIP/PLM username.")
    parser.add_argument("--password", "-p", help="EIP/PLM password. Omit to prompt securely.")
    parser.add_argument("--headless", action="store_true", help="Run Chromium in headless mode.")
    parser.add_argument("--output", "-o", type=Path, help="JSON output path.")
    args = parser.parse_args()

    username = args.username or input("账号：").strip()
    password = args.password or getpass.getpass("密码：")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output or Path("web_app2") / "outputs" / f"plm_feature_map_{timestamp}.json"

    entries = list_plm_features(username, password, output=output, headless=args.headless)
    print(f"共发现 {len(entries)} 个可见文本项：")
    for entry in entries:
        print(f"- {entry['text']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
