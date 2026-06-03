"""Try the PLM feature-map search flow for one HQ material number.

Run from the repository root:
    python web_app2/plm/try_hq_search.py --username YOUR_ID --hqpn HQ111120B1009
"""

from __future__ import annotations

import argparse
import getpass
import json
import re
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


def _page_label(page) -> str:
    try:
        return f"title={page.title()!r} url={page.url}"
    except PlaywrightError:
        return f"url={page.url}"


def _debug_dir() -> Path:
    path = Path("web_app2") / "outputs" / "plm_hq_search_debug"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _output_dir() -> Path:
    path = Path("web_app2") / "outputs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _mark_and_shot(page, locator, label: str) -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    try:
        locator.evaluate(
            """(el, label) => {
                el.scrollIntoView({block: "center", inline: "center"});
                el.style.outline = "4px solid #ff2d55";
                el.style.boxShadow = "0 0 0 4px rgba(255,45,85,.25)";
                el.setAttribute("data-plm-debug-target", label);
            }""",
            label,
        )
        page.screenshot(path=str(_debug_dir() / f"{stamp}_{label}.png"), full_page=True)
        print(f"已保存目标截图：{_debug_dir() / f'{stamp}_{label}.png'}")
    except PlaywrightError:
        pass


def _dump_search_page_controls(page, label: str) -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    base = _debug_dir() / f"{stamp}_{label}"
    try:
        page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
        controls = []
        for frame_index, frame in enumerate(page.frames):
            try:
                frame_controls = frame.evaluate(
                    """(frameIndex) => Array.from(document.querySelectorAll(
                        "input, textarea, [contenteditable='true']"
                    )).map((el, index) => {
                        const rect = el.getBoundingClientRect();
                        const style = window.getComputedStyle(el);
                        return {
                            frameIndex,
                            frameUrl: location.href,
                            index,
                            tag: el.tagName,
                            type: el.getAttribute("type") || "",
                            name: el.getAttribute("name") || "",
                            id: el.getAttribute("id") || "",
                            placeholder: el.getAttribute("placeholder") || "",
                            title: el.getAttribute("title") || "",
                            ariaLabel: el.getAttribute("aria-label") || "",
                            value: el.value || el.textContent || "",
                            visible: rect.width > 0 && rect.height > 0 &&
                                style.visibility !== "hidden" && style.display !== "none",
                            rect: {
                                left: Math.round(rect.left),
                                top: Math.round(rect.top),
                                right: Math.round(rect.right),
                                bottom: Math.round(rect.bottom),
                                width: Math.round(rect.width),
                                height: Math.round(rect.height),
                            },
                        };
                    })""",
                    frame_index,
                )
                controls.extend(frame_controls)
            except Exception:
                continue
        base.with_suffix(".json").write_text(json.dumps(controls, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"已保存搜索页控件诊断：{base.with_suffix('.png')} / {base.with_suffix('.json')}")
    except Exception as exc:
        print(f"保存搜索页诊断失败：{exc}")


def _dump_search_results(page, label: str) -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    base = _debug_dir() / f"{stamp}_{label}"
    try:
        page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
    except PlaywrightError:
        pass

    payload = {"frames": []}
    for frame_index, frame in enumerate(page.frames):
        try:
            data = frame.evaluate(
                """(frameIndex) => {
                    const clean = (text) => (text || "").replace(/\\s+/g, " ").trim();
                    const rows = Array.from(document.querySelectorAll("tr")).slice(0, 80).map((tr) =>
                        Array.from(tr.querySelectorAll("th,td")).map((cell) => clean(cell.innerText || cell.textContent))
                            .filter(Boolean)
                    ).filter((row) => row.length);
                    const links = Array.from(document.querySelectorAll("a")).slice(0, 120).map((a) => ({
                        text: clean(a.innerText || a.textContent),
                        href: a.href || a.getAttribute("href") || "",
                    })).filter((item) => item.text || item.href);
                    const visibleText = clean(document.body ? document.body.innerText : "").slice(0, 6000);
                    return {
                        frameIndex,
                        frameUrl: location.href,
                        title: document.title || "",
                        rows,
                        links,
                        visibleText,
                    };
                }""",
                frame_index,
            )
            payload["frames"].append(data)
        except Exception as exc:
            payload["frames"].append({"frameIndex": frame_index, "error": str(exc)})

    base.with_suffix(".json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保存搜索结果诊断：{base.with_suffix('.png')} / {base.with_suffix('.json')}")


def _click_first_result_number(context, page, hqpn: str):
    before = set(context.pages)
    try:
        page.screenshot(path=str(_debug_dir() / "before_first_result_coordinate_click.png"), full_page=True)
    except PlaywrightError:
        pass

    try:
        link = page.locator(f"a:has-text('{hqpn}')").first
        link.wait_for(state="visible", timeout=3000)
        _mark_and_shot(page, link, "first_result_number")
        with page.expect_popup(timeout=5000) as popup_info:
            link.click(timeout=10000)
        opened = popup_info.value
        opened.wait_for_load_state("domcontentloaded", timeout=30000)
    except (PlaywrightTimeoutError, PlaywrightError):
        print("未能通过链接定位点击第一行编号，改用表格坐标点击。")
        page.mouse.click(265, 343)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=30000)
        except PlaywrightError:
            pass
        page.wait_for_timeout(3000)
        new_pages = [candidate for candidate in context.pages if candidate not in before and not candidate.is_closed()]
        opened = new_pages[-1] if new_pages else page

    opened.wait_for_timeout(3000)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    shot = _debug_dir() / f"{stamp}_after_first_result_click.png"
    try:
        opened.screenshot(path=str(shot), full_page=True)
        print(f"已保存打开结果页截图：{shot}")
    except PlaywrightError:
        pass
    print(f"打开结果页：{_page_label(opened)}")
    return opened


def _click_detail_tab(page, text: str):
    candidates = [
        page.get_by_role("button", name=text).first,
        page.get_by_role("tab", name=text).first,
        page.locator(f"a:has-text('{text}')").first,
        page.get_by_text(text, exact=True).first,
    ]
    last_error = None
    for locator in candidates:
        try:
            locator.wait_for(state="visible", timeout=5000)
            _mark_and_shot(page, locator, f"click_detail_{text}")
            locator.click(timeout=10000)
            page.wait_for_timeout(3000)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            shot = _debug_dir() / f"{stamp}_after_click_detail_{text}.png"
            try:
                page.screenshot(path=str(shot), full_page=True)
                print(f"已保存点击{text}后截图：{shot}")
            except PlaywrightError:
                pass
            return
        except PlaywrightError as exc:
            last_error = exc
    raise RuntimeError(f"未找到详情页按钮/页签：{text}，最后错误：{last_error}")


def _dump_content_page(page, label: str) -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    base = _debug_dir() / f"{stamp}_{label}"
    try:
        page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
    except PlaywrightError:
        pass

    payload = {"frames": []}
    for frame_index, frame in enumerate(page.frames):
        try:
            data = frame.evaluate(
                """(frameIndex) => {
                    const clean = (text) => (text || "").replace(/\\s+/g, " ").trim();
                    const checkboxes = Array.from(document.querySelectorAll("input[type='checkbox']")).map((el, index) => {
                        const rect = el.getBoundingClientRect();
                        const style = window.getComputedStyle(el);
                        return {
                            index,
                            checked: !!el.checked,
                            disabled: !!el.disabled,
                            visible: rect.width > 0 && rect.height > 0 &&
                                style.visibility !== "hidden" && style.display !== "none",
                            rect: {
                                left: Math.round(rect.left),
                                top: Math.round(rect.top),
                                right: Math.round(rect.right),
                                bottom: Math.round(rect.bottom),
                                width: Math.round(rect.width),
                                height: Math.round(rect.height),
                            },
                        };
                    });
                    const links = Array.from(document.querySelectorAll("a, button")).slice(0, 160).map((el) => ({
                        text: clean(el.innerText || el.textContent),
                        title: el.getAttribute("title") || "",
                        href: el.href || el.getAttribute("href") || "",
                    })).filter((item) => item.text || item.title || item.href);
                    return {
                        frameIndex,
                        frameUrl: location.href,
                        title: document.title || "",
                        visibleText: clean(document.body ? document.body.innerText : "").slice(0, 8000),
                        checkboxes,
                        links,
                    };
                }""",
                frame_index,
            )
            payload["frames"].append(data)
        except Exception as exc:
            payload["frames"].append({"frameIndex": frame_index, "error": str(exc)})
    base.with_suffix(".json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已保存内容页诊断：{base.with_suffix('.png')} / {base.with_suffix('.json')}")


def _select_attachments_and_download(page, hqpn: str) -> None:
    page.wait_for_timeout(3000)
    _dump_content_page(page, "content_before_attachment_download")

    checked = False
    print("点击附件表头总复选框：x=59, y=276")
    try:
        page.mouse.click(59, 276)
        page.wait_for_timeout(1000)
        checked = True
    except PlaywrightError:
        pass

    print(f"附件总复选框勾选结果：{checked}")
    try:
        page.screenshot(path=str(_debug_dir() / "content_after_check_attachments.png"), full_page=True)
    except PlaywrightError:
        pass

    if not checked:
        raise RuntimeError("未能勾选附件总复选框，停止下载")

    page.wait_for_timeout(1000)

    download_candidates = ["下载选定的文件", "下载选定文件", "下载所选文件", "下载选定项"]
    last_error = None
    for text in download_candidates:
        for frame in page.frames:
            locator = frame.get_by_text(text, exact=True).first
            try:
                locator.wait_for(state="visible", timeout=1500)
                _mark_and_shot(page, locator, f"click_{text}")
                try:
                    with page.expect_download(timeout=30000) as download_info:
                        locator.click(timeout=10000)
                    download = download_info.value
                    suggested = download.suggested_filename or f"{hqpn}_attachments.zip"
                    target = _output_dir() / suggested
                    download.save_as(str(target))
                    print(f"已下载附件：{target}")
                    return
                except PlaywrightTimeoutError:
                    locator.click(timeout=10000)
                    page.wait_for_timeout(5000)
                    print(f"已点击：{text}，但未捕获到浏览器下载事件")
                    return
            except PlaywrightError as exc:
                last_error = exc
                continue

    print(f"未通过文本找到下载按钮，改用附件区下载图标坐标：x=91, y=236；最后错误：{last_error}")
    try:
        with page.expect_download(timeout=30000) as download_info:
            page.mouse.click(91, 236)
        download = download_info.value
        suggested = download.suggested_filename or f"{hqpn}_attachments.zip"
        target = _output_dir() / suggested
        download.save_as(str(target))
        print(f"已下载附件：{target}")
        return
    except PlaywrightTimeoutError:
        context = page.context
        before = set(context.pages)
        page.mouse.click(91, 236)
        page.wait_for_timeout(5000)
        new_pages = [p for p in context.pages if p not in before and not p.is_closed()]
        preview_page = new_pages[-1] if new_pages else page
        try:
            preview_page.screenshot(path=str(_debug_dir() / "single_attachment_preview.png"), full_page=True)
        except PlaywrightError:
            pass
        for target in (
            preview_page.get_by_role("button", name="下载").first,
            preview_page.get_by_role("link", name="下载").first,
            preview_page.locator("a:has-text('下载')").first,
            preview_page.get_by_text("下载", exact=True).first,
        ):
            try:
                target.wait_for(state="visible", timeout=5000)
                with preview_page.expect_download(timeout=30000) as download_info:
                    target.click(timeout=10000)
                download = download_info.value
                suggested = download.suggested_filename or f"{hqpn}_attachment"
                target_path = _output_dir() / suggested
                download.save_as(str(target_path))
                print(f"已从预览页下载附件：{target_path}")
                return
            except (PlaywrightTimeoutError, PlaywrightError):
                continue
        raise RuntimeError("单附件进入预览页后未找到下载按钮")
    except PlaywrightError as exc:
        raise RuntimeError(f"未找到“下载选定的文件”按钮/链接：{exc}") from exc


def _visible_text_locator(context, text: str, timeout: int = 3000):
    for page in reversed([p for p in context.pages if not p.is_closed()]):
        locator = page.get_by_text(text, exact=True).first
        try:
            locator.wait_for(state="visible", timeout=timeout)
            return page, locator
        except (PlaywrightTimeoutError, PlaywrightError):
            continue
    return None, None


def _wait_for_plm_ready(context, plm_page, username: str, password: str):
    for _ in range(24):
        pages = [p for p in context.pages if not p.is_closed()]
        for candidate in reversed(pages):
            login_if_present(candidate, username, password, timeout=500)
            for text in ("功能地图", "搜索"):
                locator = candidate.get_by_text(text, exact=True).first
                try:
                    locator.wait_for(state="visible", timeout=1500)
                    return candidate
                except (PlaywrightTimeoutError, PlaywrightError):
                    pass
        try:
            if "Loading" not in (plm_page.title() or ""):
                plm_page.wait_for_load_state("domcontentloaded", timeout=3000)
        except PlaywrightError:
            pass
        plm_page.wait_for_timeout(3000)
    return plm_page


def _click_text_maybe_popup(context, text: str, timeout: int = 10000):
    page, locator = _visible_text_locator(context, text)
    if not page or not locator:
        pages = "\n".join(_page_label(p) for p in context.pages if not p.is_closed())
        raise RuntimeError(f"未找到可见入口：{text}\n当前页面：\n{pages}")
    _mark_and_shot(page, locator, f"click_{text}")

    before = set(context.pages)
    try:
        with page.expect_popup(timeout=5000) as popup_info:
            locator.click(timeout=timeout)
        opened = popup_info.value
        opened.wait_for_load_state("domcontentloaded", timeout=30000)
        return opened
    except PlaywrightTimeoutError:
        locator.click(timeout=timeout)
        try:
            page.wait_for_load_state("domcontentloaded", timeout=30000)
        except PlaywrightError:
            pass
        new_pages = [candidate for candidate in context.pages if candidate not in before and not candidate.is_closed()]
        return new_pages[-1] if new_pages else page


def _fill_first_visible_search_input(page, hqpn: str) -> None:
    frame_candidates = []
    for frame in page.frames:
        try:
            input_index = frame.evaluate(
                """() => {
                    const inputs = Array.from(document.querySelectorAll("input"));
                    const scored = inputs.map((el, index) => {
                        const rect = el.getBoundingClientRect();
                        const style = window.getComputedStyle(el);
                        const type = (el.getAttribute("type") || "text").toLowerCase();
                        const visible = rect.width > 40 && rect.height > 10 &&
                            style.visibility !== "hidden" && style.display !== "none";
                        const inTopBar = rect.top >= 0 && rect.top < 80;
                        const onRight = rect.left > window.innerWidth * 0.55;
                        const textLike = ["", "text", "search"].includes(type);
                        return {
                            index,
                            ok: visible && inTopBar && onRight && textLike,
                            right: rect.right,
                            width: rect.width,
                        };
                    }).filter(item => item.ok);
                    scored.sort((a, b) => b.right - a.right || b.width - a.width);
                    return scored.length ? scored[0].index : -1;
                }"""
            )
            if input_index >= 0:
                frame_candidates.append(frame.locator("input").nth(input_index))
        except PlaywrightError:
            continue

    for frame in page.frames:
        for selector in [
            "input[name*='keyword' i]:visible",
            "input[id*='keyword' i]:visible",
            "input[name*='search' i]:visible",
            "input[id*='search' i]:visible",
            "input[type='text']:visible",
            "input:not([type]):visible",
        ]:
            frame_candidates.append(frame.locator(selector).first)

    try:
        handle = page.main_frame.evaluate_handle(
            """() => {
                const controls = Array.from(document.querySelectorAll(
                    "input, textarea, [contenteditable='true']"
                )).filter((el) => {
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return rect.width > 0 && rect.height > 0 &&
                        style.visibility !== "hidden" &&
                        style.display !== "none" &&
                        rect.top < Math.max(260, window.innerHeight * 0.35);
                });
                controls.sort((a, b) => {
                    const ar = a.getBoundingClientRect();
                    const br = b.getBoundingClientRect();
                    return br.right - ar.right || ar.top - br.top;
                });
                return controls[0] || null;
            }"""
        )
        right_top_input = handle.as_element()
        if right_top_input:
            frame_candidates.append(right_top_input)
    except PlaywrightError:
        pass
    frame_candidates.extend(
        [
            page.locator("input:visible").last,
            page.get_by_role("textbox").last,
            page.locator("textarea:visible").last,
            page.locator("[contenteditable='true']:visible").last,
            page.get_by_role("textbox").first,
            page.locator("input:visible").first,
            page.locator("textarea:visible").first,
            page.locator("[contenteditable='true']:visible").first,
        ]
    )
    last_error = None
    for locator in frame_candidates:
        try:
            locator.wait_for(state="visible", timeout=5000)
            _mark_and_shot(page, locator, "hqpn_input")
            box = locator.bounding_box(timeout=5000)
            if box:
                print(
                    "将点击输入框中心："
                    f"x={round(box['x'] + box['width'] / 2)}, "
                    f"y={round(box['y'] + box['height'] / 2)}, "
                    f"w={round(box['width'])}, h={round(box['height'])}"
                )
                page.mouse.click(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
            else:
                locator.click(timeout=5000)
            locator.evaluate(
                """(el, value) => {
                    el.focus();
                    el.value = value;
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                }""",
                hqpn,
            )
            page.wait_for_timeout(300)
            try:
                active_value = page.evaluate(
                    """() => {
                        const el = document.activeElement;
                        return {
                            tag: el ? el.tagName : "",
                            type: el ? (el.getAttribute("type") || "") : "",
                            id: el ? (el.getAttribute("id") || "") : "",
                            name: el ? (el.getAttribute("name") || "") : "",
                            value: el ? (el.value || el.textContent || "") : "",
                        };
                    }"""
                )
                print(f"当前焦点元素：{active_value}")
            except PlaywrightError:
                pass
            try:
                locator.press("Enter", timeout=3000)
                page.wait_for_timeout(1000)
            except PlaywrightError:
                pass
            page.mouse.click(1172, 20)
            try:
                page.screenshot(path=str(_debug_dir() / "after_hqpn_type.png"), full_page=True)
            except PlaywrightError:
                pass
            return
        except PlaywrightError as exc:
            last_error = exc

    raise RuntimeError(f"未找到可输入料号的输入框：{last_error}")


def run(username: str, password: str, hqpn: str) -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        try:
            print("打开 EIP...")
            page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
            wait_for_eip_ready(page, username, password)

            print("进入 PLM...")
            plm_page = click_opening_page(
                page,
                page.locator("a").filter(has_text=re.compile(r"^PLM$")),
                timeout=30000,
            )
            login_if_present(plm_page, username, password, timeout=500)
            print("等待 PLM 页面加载完成...")
            plm_page = _wait_for_plm_ready(context, plm_page, username, password)

            print("当前页面：")
            for opened_page in context.pages:
                if not opened_page.is_closed():
                    print(f"  - {_page_label(opened_page)}")

            print("打开功能地图...")
            feature_page = _click_text_maybe_popup(context, "功能地图")
            print(f"功能地图页面：{_page_label(feature_page)}")

            print("点击搜索...")
            search_page = _click_text_maybe_popup(context, "搜索")
            print(f"搜索页面：{_page_label(search_page)}")
            search_page.wait_for_timeout(1500)
            _dump_search_page_controls(search_page, "after_search_click")

            print(f"输入料号并回车：{hqpn}")
            _fill_first_visible_search_input(search_page, hqpn)
            search_page.wait_for_timeout(5000)
            _dump_search_results(search_page, "after_hqpn_search")

            print("点击搜索结果第一行编号...")
            detail_page = _click_first_result_number(context, search_page, hqpn)

            print("点击详情页：内容")
            _click_detail_tab(detail_page, "内容")

            print("勾选内容页附件并点击下载选定的文件...")
            _select_attachments_and_download(detail_page, hqpn)

            print("已完成输入。浏览器保持打开 10 秒，方便你观察结果...")
            detail_page.wait_for_timeout(10000)
        except Exception:
            print("执行失败，浏览器保持打开 20 秒，方便观察当前页面...")
            try:
                page.wait_for_timeout(20000)
            except PlaywrightError:
                pass
            raise
        finally:
            context.close()
            browser.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Try PLM search for one HQ material number.")
    parser.add_argument("--username", "-u", help="EIP/PLM username.")
    parser.add_argument("--password", "-p", help="EIP/PLM password. Omit to prompt securely.")
    parser.add_argument("--hqpn", default="HQ111120B1009", help="HQ material number to search.")
    args = parser.parse_args()

    username = args.username or input("账号：").strip()
    password = args.password or getpass.getpass("密码：")
    run(username, password, args.hqpn.strip())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
