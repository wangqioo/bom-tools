import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from zipfile import ZIP_DEFLATED, ZipFile

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, Playwright, TimeoutError as PlaywrightTimeoutError


START_URL = "https://eip.evex-tech.com/"
PLM_SEARCH_URL = "http://plm.evex-tech.com/Windchill/app/#ptc1/ext/huaqin/homePage/searchFunction"


@dataclass(frozen=True)
class PlmFeature:
    key: str
    label: str
    entry_name: str
    query_button: str = "查询"
    export_button: str = "结果导出"


FEATURES: dict[str, PlmFeature] = {
    "spec_reverse_material": PlmFeature(
        key="spec_reverse_material",
        label="规格型号反查物料",
        entry_name="规格型号反查物料",
    ),
}


LogFn = Callable[[str], None]


def _noop_log(_: str) -> None:
    return


@dataclass(frozen=True)
class ClickTarget:
    frame_index: int
    probe_id: str
    score: int
    reasons: list[str] = field(default_factory=list)
    tag: str = ""
    text: str = ""
    title: str = ""
    href: str = ""
    frame_url: str = ""
    rect: dict | None = None


def _clean_text(value: str | None) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _target_json(targets: list[ClickTarget], limit: int = 20) -> list[dict]:
    payload = []
    for target in targets[:limit]:
        payload.append({
            "frame_index": target.frame_index,
            "probe_id": target.probe_id,
            "score": target.score,
            "reasons": target.reasons,
            "tag": target.tag,
            "text": target.text[:300],
            "title": target.title,
            "href": target.href,
            "frame_url": target.frame_url,
            "rect": target.rect,
        })
    return payload


def require_feature(feature_key: str) -> PlmFeature:
    try:
        return FEATURES[feature_key]
    except KeyError as exc:
        valid = ", ".join(sorted(FEATURES))
        raise ValueError(f"Unknown feature: {feature_key}. Valid features: {valid}") from exc


def fill_login(page: Page, username: str, password: str) -> None:
    username_box = page.locator("#tbLoginName, input[placeholder='请输入用户名']").first
    username_box.wait_for(state="visible", timeout=60000)
    username_box.click(timeout=10000)
    username_box.fill(username, timeout=10000)
    username_box.press("Tab", timeout=10000)
    page.keyboard.type(password, delay=30)
    page.keyboard.press("Enter")
    try:
        page.wait_for_load_state("domcontentloaded", timeout=60000)
    except PlaywrightError:
        pass


def login_if_present(page: Page, username: str, password: str, timeout: int = 15000) -> bool:
    try:
        page.locator("#tbLoginName, input[placeholder='请输入用户名']").first.wait_for(
            state="visible", timeout=timeout
        )
        fill_login(page, username, password)
        return True
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def click_opening_page(page: Page, locator, timeout: int = 15000) -> Page:
    context = page.context
    pages_before = set(context.pages)

    try:
        with page.expect_popup(timeout=timeout) as popup_info:
            locator.click(timeout=timeout)
        new_page = popup_info.value
        new_page.wait_for_load_state("domcontentloaded", timeout=60000)
        return new_page
    except PlaywrightTimeoutError:
        try:
            page.wait_for_load_state("domcontentloaded", timeout=60000)
        except PlaywrightError:
            pass
        return page
    except PlaywrightError:
        pages_after = [candidate for candidate in context.pages if candidate not in pages_before]
        if pages_after:
            pages_after[-1].wait_for_load_state("domcontentloaded", timeout=60000)
            return pages_after[-1]
        open_pages = [candidate for candidate in context.pages if not candidate.is_closed()]
        if open_pages:
            return open_pages[-1]
        raise


def by_button_or_text(page: Page, text: str):
    return page.get_by_role("button", name=text).or_(page.get_by_text(text, exact=True)).first


def click_by_text(page: Page, text: str, timeout: int = 30000) -> None:
    locator = by_button_or_text(page, text)
    locator.wait_for(state="visible", timeout=timeout)
    locator.click(timeout=timeout)


def click_when_ready(page: Page, text: str, quick_timeout: int = 2000, fallback_timeout: int = 30000) -> None:
    locator = by_button_or_text(page, text)
    try:
        locator.click(timeout=quick_timeout)
        return
    except PlaywrightError:
        pass
    locator.wait_for(state="visible", timeout=fallback_timeout)
    locator.click(timeout=fallback_timeout)


def _scan_click_targets(
    page: Page,
    *,
    query: str,
    exact: bool = False,
    row_text: str = "",
    href_exclude: list[str] | None = None,
    prefer_text: list[str] | None = None,
    tag_boost: list[str] | None = None,
    limit: int = 40,
) -> list[ClickTarget]:
    query_clean = _clean_text(query)
    row_clean = _clean_text(row_text)
    href_exclude = [item.lower() for item in (href_exclude or [])]
    prefer_text = [_clean_text(item).lower() for item in (prefer_text or []) if _clean_text(item)]
    tag_boost = [item.upper() for item in (tag_boost or [])]
    stamp = f"plm_probe_{int(time.time() * 1000)}"
    targets: list[ClickTarget] = []

    script = r"""({ stamp, frameIndex }) => {
        const visible = (el) => {
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
                style.visibility !== 'hidden' && style.opacity !== '0';
        };
        const clean = (text) => (text || '').replace(/\s+/g, ' ').trim();
        const clickableSelector = [
            'a', 'button', 'input', 'textarea', 'select', '[role]', '[onclick]',
            '.x-btn', '.x-grid-row', '.x-grid3-row', '.x-tab-strip-text', 'span', 'div', 'td'
        ].join(',');
        return Array.from(document.querySelectorAll(clickableSelector)).map((el, index) => {
            const rect = el.getBoundingClientRect();
            const probeId = `${stamp}_${frameIndex}_${index}`;
            try { el.setAttribute('data-plm-probe-id', probeId); } catch (_) {}
            const row = el.closest('tr,.x-grid-row,.x-grid3-row,[role=row]');
            return {
                probe_id: probeId,
                tag: el.tagName || '',
                text: clean(el.innerText || el.textContent || el.value || ''),
                own_text: clean(Array.from(el.childNodes || []).filter((n) => n.nodeType === Node.TEXT_NODE).map((n) => n.textContent).join(' ')),
                id: el.id || '',
                name: el.getAttribute('name') || '',
                cls: typeof el.className === 'string' ? el.className : '',
                role: el.getAttribute('role') || '',
                type: el.getAttribute('type') || '',
                title: el.getAttribute('title') || '',
                aria: el.getAttribute('aria-label') || '',
                href: el.href || el.getAttribute('href') || '',
                disabled: Boolean(el.disabled || el.getAttribute('aria-disabled') === 'true'),
                visible: visible(el),
                row_text: clean(row ? (row.innerText || row.textContent || '') : ''),
                rect: { left: Math.round(rect.left), top: Math.round(rect.top), width: Math.round(rect.width), height: Math.round(rect.height) },
            };
        }).filter((item) => item.visible && !item.disabled && item.rect.width > 0 && item.rect.height > 0);
    }"""

    for frame_index, frame in enumerate(page.frames):
        try:
            raw_items = frame.evaluate(script, {"stamp": stamp, "frameIndex": frame_index}) or []
        except PlaywrightError:
            continue
        for item in raw_items:
            fields = [
                item.get("text", ""), item.get("own_text", ""), item.get("title", ""),
                item.get("aria", ""), item.get("id", ""), item.get("name", ""),
                item.get("href", ""), item.get("row_text", ""), item.get("cls", ""),
            ]
            combined = _clean_text(" ".join(str(value or "") for value in fields))
            combined_l = combined.lower()
            query_l = query_clean.lower()
            row_l = item.get("row_text", "").lower()
            href_l = item.get("href", "").lower()
            text_l = _clean_text(item.get("text", "")).lower()
            title_l = _clean_text(item.get("title", "")).lower()
            tag = str(item.get("tag", "")).upper()
            score = 0
            reasons: list[str] = []

            if query_clean:
                if exact and (text_l == query_l or title_l == query_l):
                    score += 100
                    reasons.append("exact text/title")
                elif not exact and query_l in combined_l:
                    score += 55
                    reasons.append("contains query")
                elif exact and query_l in combined_l:
                    score += 30
                    reasons.append("contains exact query")
                else:
                    continue

            if row_clean:
                if row_clean.lower() in row_l:
                    score += 35
                    reasons.append("row contains target")
                else:
                    continue

            if href_exclude and any(pattern in href_l for pattern in href_exclude):
                score -= 80
                reasons.append("excluded href pattern")
            if prefer_text and any(value in combined_l for value in prefer_text):
                score += 15
                reasons.append("preferred context")
            if tag in tag_boost:
                score += 10
                reasons.append("preferred tag")
            if tag in {"A", "BUTTON", "INPUT"}:
                score += 8
                reasons.append("native clickable")
            if item.get("role") in {"button", "link", "tab", "row"}:
                score += 6
                reasons.append("interactive role")
            rect = item.get("rect") or {}
            if rect.get("width", 0) > 600 or rect.get("height", 0) > 120:
                score -= 12
                reasons.append("large container")

            if score <= 0:
                continue
            targets.append(ClickTarget(
                frame_index=frame_index,
                probe_id=item.get("probe_id", ""),
                score=score,
                reasons=reasons,
                tag=tag,
                text=_clean_text(item.get("text", "")),
                title=_clean_text(item.get("title", "")),
                href=item.get("href", ""),
                frame_url=frame.url,
                rect=rect,
            ))

    targets.sort(key=lambda target: target.score, reverse=True)
    return targets[:limit]


def _write_target_debug(page: Page, output_dir: Path, label: str, targets: list[ClickTarget]) -> None:
    debug_dir = output_dir / "plm_hq_search_debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    try:
        (debug_dir / f"{stamp}_{label}_targets.json").write_text(
            json.dumps(_target_json(targets, limit=50), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


def _click_target(context, page: Page, target: ClickTarget, timeout: int = 10000) -> Page:
    if target.frame_index >= len(page.frames):
        raise RuntimeError(f"Target frame no longer exists: {target.frame_index}")
    frame = page.frames[target.frame_index]
    locator = frame.locator(f"[data-plm-probe-id='{target.probe_id}']").first
    before = set(context.pages)

    try:
        locator.scroll_into_view_if_needed(timeout=timeout)
    except PlaywrightError:
        pass

    try:
        locator.click(timeout=timeout)
    except PlaywrightError:
        frame.evaluate(
            """(probeId) => {
                const el = document.querySelector(`[data-plm-probe-id='${probeId}']`);
                if (!el) throw new Error(`target not found: ${probeId}`);
                el.scrollIntoView({ block: 'center', inline: 'center' });
                el.click();
            }""",
            target.probe_id,
        )

    page.wait_for_timeout(800)
    new_pages = [candidate for candidate in context.pages if candidate not in before and not candidate.is_closed()]
    opened = new_pages[-1] if new_pages else page
    try:
        opened.wait_for_load_state("domcontentloaded", timeout=30000)
    except PlaywrightError:
        pass
    return opened

def wait_for_eip_ready(page: Page, username: str, password: str) -> None:
    plm_link = page.locator("a").filter(has_text=re.compile(r"^PLM$"))
    for _ in range(8):
        login_if_present(page, username, password, timeout=5000)
        try:
            plm_link.wait_for(state="visible", timeout=15000)
            return
        except PlaywrightTimeoutError:
            if "sso.huaqin.com/login" not in page.url and "callback" not in page.url:
                page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
            else:
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=30000)
                except PlaywrightError:
                    pass
    plm_link.wait_for(state="visible", timeout=30000)


def wait_for_query_result(page: Page, export_button: str) -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=30000)
    except PlaywrightError:
        pass

    result_markers = [
        page.get_by_text(export_button, exact=True),
        page.locator("table").first,
        page.locator(".el-table, .ant-table, .vxe-table").first,
    ]
    for marker in result_markers:
        try:
            marker.wait_for(state="visible", timeout=15000)
            return
        except PlaywrightTimeoutError:
            continue
    page.wait_for_timeout(3000)


def export_result(page: Page, upload_file: Path, export_button_text: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    export_button = by_button_or_text(page, export_button_text)
    export_button.wait_for(state="visible", timeout=30000)

    with page.expect_download(timeout=120000) as download_info:
        export_button.click(timeout=30000)
    download = download_info.value

    suggested_name = download.suggested_filename or f"{upload_file.stem}_结果导出.xlsx"
    output_path = output_dir / suggested_name
    download.save_as(str(output_path))
    return output_path


def run_plm_feature(
    playwright: Playwright,
    *,
    username: str,
    password: str,
    feature: PlmFeature,
    upload_file: Path,
    output_dir: Path,
    headless: bool = False,
    log: LogFn = _noop_log,
) -> Path:
    upload_file = upload_file.expanduser().resolve()
    if not upload_file.exists():
        raise FileNotFoundError(f"Upload file not found: {upload_file}")
    if not upload_file.is_file():
        raise ValueError(f"Upload path is not a file: {upload_file}")

    log("启动浏览器")
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    try:
        log("打开 EIP")
        page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
        wait_for_eip_ready(page, username, password)

        log("进入 PLM")
        plm_page = click_opening_page(
            page,
            page.locator("a").filter(has_text=re.compile(r"^PLM$")),
            timeout=30000,
        )
        login_if_present(plm_page, username, password, timeout=500)

        log("打开功能地图")
        click_when_ready(plm_page, "功能地图", quick_timeout=2000, fallback_timeout=30000)

        log(f"打开功能：{feature.entry_name}")
        feature_link = plm_page.get_by_role("link", name=feature.entry_name).or_(
            plm_page.get_by_text(feature.entry_name, exact=True)
        ).first
        target_page = click_opening_page(plm_page, feature_link, timeout=30000)

        log(f"上传文件：{upload_file.name}")
        file_input = target_page.locator("input[type='file']").first
        file_input.set_input_files(str(upload_file))

        log(f"点击{feature.query_button}")
        click_by_text(target_page, feature.query_button, timeout=30000)

        log("等待结果")
        wait_for_query_result(target_page, feature.export_button)

        log(f"点击{feature.export_button}")
        output_path = export_result(target_page, upload_file, feature.export_button, output_dir)
        log(f"导出完成：{output_path}")
        return output_path
    finally:
        context.close()
        browser.close()


def _wait_for_plm_home(context, page: Page, username: str, password: str) -> Page:
    for _ in range(24):
        for candidate in reversed([p for p in context.pages if not p.is_closed()]):
            login_if_present(candidate, username, password, timeout=500)
            for text in ("功能地图", "搜索"):
                try:
                    candidate.get_by_text(text, exact=True).first.wait_for(state="visible", timeout=1500)
                    return candidate
                except (PlaywrightTimeoutError, PlaywrightError):
                    pass
        page.wait_for_timeout(3000)
    return page


def _wait_for_extjs(page: Page, timeout_ms: int = 15000) -> None:
    """Windchill ExtJS pages often render after domcontentloaded."""
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for frame in page.frames:
            try:
                ready = frame.evaluate(
                    """() => {
                        const bodyText = document.body ? document.body.innerText : '';
                        return Boolean(window.Ext || document.querySelector('#keywordkeywordField_SearchTextBox') ||
                            bodyText.includes('搜索') || bodyText.includes('内容'));
                    }"""
                )
                if ready:
                    return
            except PlaywrightError:
                continue
        page.wait_for_timeout(500)


def _open_plm_search_page(context, page: Page, username: str, password: str, log: LogFn = _noop_log) -> Page:
    candidates = [p for p in context.pages if not p.is_closed()]
    search_page = candidates[-1] if candidates else page
    login_if_present(search_page, username, password, timeout=500)
    log("直接进入 PLM 搜索页")
    search_page.goto(PLM_SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
    login_if_present(search_page, username, password, timeout=3000)
    _wait_for_extjs(search_page, timeout_ms=18000)
    return search_page
def _click_text_any_page(context, text: str, timeout: int = 10000) -> Page:
    for page in reversed([p for p in context.pages if not p.is_closed()]):
        targets = _scan_click_targets(
            page,
            query=text,
            exact=True,
            tag_boost=["A", "BUTTON", "SPAN", "DIV"],
            limit=8,
        )
        if targets:
            last_error = None
            for target in targets:
                try:
                    return _click_target(context, page, target, timeout=timeout)
                except (PlaywrightError, RuntimeError) as exc:
                    last_error = exc
            if last_error:
                continue

        locator = page.get_by_text(text, exact=True).first
        try:
            locator.wait_for(state="visible", timeout=3000)
        except (PlaywrightTimeoutError, PlaywrightError):
            continue
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
            new_pages = [p for p in context.pages if p not in before and not p.is_closed()]
            return new_pages[-1] if new_pages else page
    raise RuntimeError(f"Visible entry not found: {text}")

def _find_search_input(page: Page):
    selectors = [
        "#keywordkeywordField_SearchTextBox",
        "input[id*='keyword' i]",
        "input[name*='keyword' i]",
        "input[id*='search' i]",
        "input[name*='search' i]",
    ]
    for _ in range(30):
        for frame in page.frames:
            for selector in selectors:
                locator = frame.locator(selector).first
                try:
                    locator.wait_for(state="visible", timeout=500)
                    return locator
                except (PlaywrightTimeoutError, PlaywrightError):
                    continue
        page.wait_for_timeout(500)
    return None


def _search_result_visible(page: Page, hqpn: str) -> bool:
    for frame in page.frames:
        try:
            found = frame.evaluate(
                """(partNumber) => Boolean(document.body && (document.body.innerText || '').includes(partNumber))""",
                hqpn,
            )
            if found:
                return True
        except PlaywrightError:
            continue
    return False

def _type_top_search(page: Page, hqpn: str) -> None:
    locator = _find_search_input(page)
    if locator is None:
        raise RuntimeError("未找到 PLM 搜索输入框：页面内不存在 keyword/search 输入框")

    last_error = None
    for attempt in range(3):
        try:
            locator.click(timeout=5000)
            locator.fill(hqpn, timeout=5000)
            page.wait_for_timeout(300)
            locator.press("Enter", timeout=5000)
            page.wait_for_timeout(3000)
            if _search_result_visible(page, hqpn):
                return
        except PlaywrightError as exc:
            last_error = exc

        try:
            locator.evaluate(
                """(el, value) => {
                    el.focus();
                    el.value = value;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                }""",
                hqpn,
            )
            page.keyboard.press("Enter")
            page.wait_for_timeout(3000)
            if _search_result_visible(page, hqpn):
                return
        except PlaywrightError as exc:
            last_error = exc

        try:
            clicked = False
            for frame in page.frames:
                clicked = frame.evaluate(
                    """() => {
                        const visible = (el) => {
                            const rect = el.getBoundingClientRect();
                            const style = window.getComputedStyle(el);
                            return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
                        };
                        const controls = Array.from(document.querySelectorAll('button,a,input[type=button],span,div'));
                        const button = controls.find((el) => visible(el) && /^(搜索|查询)$/.test((el.innerText || el.value || '').trim()));
                        if (!button) return false;
                        button.click();
                        return true;
                    }"""
                )
                if clicked:
                    break
            if clicked:
                page.wait_for_timeout(4000)
                if _search_result_visible(page, hqpn):
                    return
        except PlaywrightError as exc:
            last_error = exc

    raise RuntimeError(f"PLM 搜索已输入但未出现料号结果：{hqpn}，最后错误：{last_error}")
def _click_first_search_result(context, page: Page, hqpn: str, output_dir: Path | None = None) -> Page:
    last_error = None
    for _ in range(20):
        targets = _scan_click_targets(
            page,
            query=hqpn,
            row_text=hqpn,
            href_exclude=["edrview.jsp", "getPartStructureED"],
            prefer_text=["Design", hqpn],
            tag_boost=["A"],
            limit=12,
        )
        if output_dir:
            _write_target_debug(page, output_dir, "search_result", targets)
        for target in targets:
            if any(pattern in target.href.lower() for pattern in ("edrview.jsp", "getpartstructureed")):
                continue
            try:
                opened = _click_target(context, page, target, timeout=10000)
                try:
                    opened.wait_for_load_state("domcontentloaded", timeout=30000)
                except PlaywrightError:
                    pass
                if "edrview.jsp" in opened.url or "getPartStructureED" in opened.url:
                    last_error = RuntimeError("Opened structure preview instead of material detail page")
                    continue
                return opened
            except (PlaywrightError, RuntimeError) as exc:
                last_error = exc

        for frame in page.frames:
            try:
                clicked = frame.evaluate(
                    r"""(partNumber) => {
                        const visible = (el) => {
                            const rect = el.getBoundingClientRect();
                            const style = window.getComputedStyle(el);
                            return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
                                style.visibility !== 'hidden';
                        };
                        const rows = Array.from(document.querySelectorAll('tr, .x-grid-row, .x-grid3-row, [role=row]'));
                        const scoredRows = rows
                            .map((row) => ({ row, text: (row.innerText || row.textContent || '').replace(/\s+/g, ' ') }))
                            .filter((item) => item.text.includes(partNumber))
                            .sort((a, b) => {
                                const ad = /Design/i.test(a.text) ? 0 : 1;
                                const bd = /Design/i.test(b.text) ? 0 : 1;
                                return ad - bd;
                            });
                        for (const item of scoredRows) {
                            const links = Array.from(item.row.querySelectorAll('a'));
                            const link = links.find((a) => visible(a) && !/(edrview\.jsp|getPartStructureED)/i.test(a.href || '') &&
                                ((a.innerText || a.textContent || '').includes(partNumber) || (a.href || '').includes(partNumber)));
                            if (link) {
                                link.scrollIntoView({ block: 'center', inline: 'center' });
                                link.click();
                                return true;
                            }
                        }
                        return false;
                    }""",
                    hqpn,
                )
                if not clicked:
                    continue
                page.wait_for_timeout(1500)
                new_pages = [p for p in context.pages if not p.is_closed()]
                opened = new_pages[-1] if new_pages else page
                try:
                    opened.wait_for_load_state("domcontentloaded", timeout=30000)
                except PlaywrightError:
                    pass
                if "edrview.jsp" in opened.url or "getPartStructureED" in opened.url:
                    raise RuntimeError("Opened structure preview instead of material detail page")
                return opened
            except (PlaywrightError, RuntimeError) as exc:
                last_error = exc
        page.wait_for_timeout(1000)
    if output_dir:
        _dump_page_debug(page, output_dir, "search_result_failed")
    raise RuntimeError(f"Clickable material result not found: {hqpn}; last error: {last_error}")

def _click_detail_content(page: Page, output_dir: Path | None = None) -> None:
    last_error = None
    for _ in range(10):
        targets = _scan_click_targets(
            page,
            query="内容",
            exact=True,
            tag_boost=["A", "BUTTON", "SPAN", "DIV"],
            limit=10,
        )
        if output_dir:
            _write_target_debug(page, output_dir, "content_tab", targets)
        for target in targets:
            try:
                _click_target(page.context, page, target, timeout=10000)
                page.wait_for_timeout(1200)
                return
            except (PlaywrightError, RuntimeError) as exc:
                last_error = exc

        for frame in page.frames:
            try:
                clicked = frame.evaluate(
                    """() => {
                        const visible = (el) => {
                            const rect = el.getBoundingClientRect();
                            const style = window.getComputedStyle(el);
                            return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
                                style.visibility !== 'hidden';
                        };
                        const candidates = Array.from(document.querySelectorAll('a,button,span,div,[role=tab]'))
                            .filter((el) => visible(el) && (el.innerText || el.textContent || '').trim() === '内容');
                        if (!candidates.length) return false;
                        candidates[0].scrollIntoView({ block: 'center', inline: 'center' });
                        candidates[0].click();
                        return true;
                    }"""
                )
                if clicked:
                    page.wait_for_timeout(1200)
                    return
            except PlaywrightError as exc:
                last_error = exc
        page.wait_for_timeout(500)
    if output_dir:
        _dump_page_debug(page, output_dir, "content_tab_failed")
    raise RuntimeError(f"Detail content tab not found: {last_error}")

def _dump_page_debug(page: Page, output_dir: Path, label: str) -> None:
    debug_dir = output_dir / "plm_hq_search_debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    try:
        page.screenshot(path=str(debug_dir / f"{stamp}_{label}.png"), full_page=True)
    except PlaywrightError:
        pass
    payload = []
    for index, frame in enumerate(page.frames):
        try:
            payload.append({
                "index": index,
                "url": frame.url,
                "text": frame.evaluate("() => document.body ? document.body.innerText.slice(0, 12000) : ''"),
                "inputs": frame.evaluate("""() => Array.from(document.querySelectorAll('input,textarea,[contenteditable=true]')).map((el, i) => {
                    const rect = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    return {
                        index: i,
                        id: el.id || '',
                        name: el.getAttribute('name') || '',
                        type: el.getAttribute('type') || '',
                        value: el.value || el.textContent || '',
                        visible: rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden',
                        rect: { left: Math.round(rect.left), top: Math.round(rect.top), width: Math.round(rect.width), height: Math.round(rect.height) }
                    };
                })"""),
            })
        except PlaywrightError as exc:
            payload.append({"index": index, "url": frame.url, "error": str(exc)})
    try:
        import json as _json
        (debug_dir / f"{stamp}_{label}.json").write_text(_json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass

def _download_selected_attachments(page: Page, hqpn: str, output_dir: Path, log: LogFn = _noop_log) -> Path:
    start = time.monotonic()

    def step(message: str) -> None:
        line = f"{message}（{time.monotonic() - start:.1f}s）"
        log(line)
        try:
            with (output_dir / "plm_hq_attachment_timing.log").open("a", encoding="utf-8") as fp:
                fp.write(line + "\n")
        except OSError:
            pass

    def zip_paths(files: list[Path]) -> Path:
        zip_path = output_dir / f"{hqpn}_附件.zip"
        with ZipFile(zip_path, "w", ZIP_DEFLATED) as zf:
            for file_path in files:
                zf.write(file_path, arcname=file_path.name)
        for file_path in files:
            try:
                file_path.unlink()
            except OSError:
                pass
        return zip_path

    def safe_name(name: str, index: int) -> str:
        cleaned = re.sub(r'[\\/:*?"<>|]+', '_', (name or '').strip())
        if not cleaned:
            cleaned = f"{hqpn}_附件_{index}.pdf"
        lowered = cleaned.lower()
        if lowered.endswith('.pdf.crdownload'):
            cleaned = cleaned[:-11]
        elif not lowered.endswith(('.pdf', '.zip')):
            cleaned += '.pdf'
        return cleaned

    def find_pdf_links() -> list[dict[str, str]]:
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
                        .filter((item) => /\.pdf(\.crdownload)?(\?|$)/i.test(item.href) || /\.pdf(\.crdownload)?$/i.test(item.text) || /\.pdf(\.crdownload)?$/i.test(item.title))"""
                )
                if frame_links:
                    links.extend(frame_links)
            except PlaywrightError:
                continue
        deduped: list[dict[str, str]] = []
        seen = set()
        for item in links:
            key = item.get('href') or item.get('text') or item.get('title')
            if not key or key in seen:
                continue
            seen.add(key)
            deduped.append(item)
        return deduped

    def direct_download(url: str, name: str, index: int) -> Path | None:
        if not url.lower().startswith(("http://", "https://")):
            return None
        try:
            response = page.context.request.get(url, timeout=60000)
            if not response.ok:
                return None
            body = response.body()
            if not body:
                return None
            raw_path = output_dir / safe_name(name, index)
            raw_path.write_bytes(body)
            return raw_path
        except Exception:
            return None

    def click_pdf_link(link_text: str, href: str) -> Path | None:
        context = page.context
        before = set(context.pages)
        for frame in page.frames:
            try:
                clicked = frame.evaluate(
                    """({ text, href }) => {
                        const visible = (el) => {
                            const rect = el.getBoundingClientRect();
                            const style = window.getComputedStyle(el);
                            return rect.width > 0 && rect.height > 0 && style.display !== 'none' &&
                                style.visibility !== 'hidden';
                        };
                        const links = Array.from(document.querySelectorAll('a'));
                        const link = links.find((a) => visible(a) && ((href && a.href === href) ||
                            (text && ((a.innerText || a.textContent || '').trim() === text || (a.getAttribute('title') || '') === text))));
                        if (!link) return false;
                        link.scrollIntoView({ block: 'center', inline: 'center' });
                        link.click();
                        return true;
                    }""",
                    {"text": link_text, "href": href},
                )
                if not clicked:
                    continue
                page.wait_for_timeout(1200)
                new_pages = [p for p in context.pages if p not in before and not p.is_closed()]
                preview_page = new_pages[-1] if new_pages else page
                return download_preview_resource(preview_page)
            except PlaywrightError:
                continue
        return None

    def download_preview_resource(preview_page: Page) -> Path | None:
        for _ in range(10):
            urls = [preview_page.url]
            try:
                urls.extend(
                    preview_page.evaluate(
                        """() => Array.from(document.querySelectorAll('embed,iframe,object,a'))
                            .map((el) => el.src || el.data || el.href || '')
                            .filter(Boolean)"""
                    )
                )
            except PlaywrightError:
                pass
            for index, url in enumerate(dict.fromkeys(urls), start=1):
                if not url.lower().startswith(("http://", "https://")):
                    continue
                downloaded = direct_download(url, f"{hqpn}_附件_{index}.pdf", index)
                if downloaded:
                    return downloaded
            preview_page.wait_for_timeout(500)
        return None
    def download_checked_pdf_rows() -> list[Path]:
        label = "???????"

        def save_pdf_viewer_resource(pages_before) -> list[Path]:
            page.wait_for_timeout(2500)
            candidates = [candidate for candidate in page.context.pages if not candidate.is_closed()]
            new_pages = [candidate for candidate in candidates if candidate not in pages_before]
            for candidate in list(reversed(new_pages)) + list(reversed(candidates)):
                url = candidate.url or ""
                if "application/pdf" not in url and ".pdf" not in url.lower() and "doDirectDownload" not in url:
                    continue
                name = f"{hqpn}_attachment.pdf"
                try:
                    from urllib.parse import parse_qs, unquote, urlparse
                    query = parse_qs(urlparse(url).query)
                    if query.get("ofn"):
                        name = unquote(query["ofn"][0])
                except Exception:
                    pass
                output_path = direct_download(url, name, 1)
                if output_path:
                    step(f"Downloaded PDF viewer resource: {output_path.name}")
                    return [output_path]
            return []

        for frame in page.frames:
            try:
                checker = frame.locator('.x-grid3-hd-checker').first
                if checker.count() == 0:
                    continue
                checker.click(timeout=5000)
                page.wait_for_timeout(800)

                pages_before = set(page.context.pages)
                try:
                    with page.expect_download(timeout=10000) as download_info:
                        frame.get_by_text(label, exact=True).click(timeout=10000, force=True)
                    download = download_info.value
                    suggested = download.suggested_filename or f"{hqpn}_attachments.zip"
                    output_path = output_dir / safe_name(suggested, 1)
                    download.save_as(str(output_path))
                    step(f"Downloaded selected attachments: {output_path.name}")
                    return [output_path]
                except PlaywrightTimeoutError:
                    viewer_files = save_pdf_viewer_resource(pages_before)
                    if viewer_files:
                        return viewer_files
            except PlaywrightTimeoutError:
                continue
            except PlaywrightError:
                continue
        return []
    output_dir.mkdir(parents=True, exist_ok=True)
    page.wait_for_timeout(2500)
    links = find_pdf_links()
    step(f"识别 PDF 附件：{len(links)} 个")

    downloaded_files: list[Path] = []
    for index, link in enumerate(links, start=1):
        name = link.get('text') or link.get('title') or f"{hqpn}_附件_{index}.pdf"
        href = link.get('href') or ''
        downloaded = direct_download(href, name, index)
        if not downloaded:
            downloaded = click_pdf_link(name, href)
        if downloaded:
            downloaded_files.append(downloaded)
            step(f"已下载 PDF：{downloaded.name}")

    if downloaded_files:
        return zip_paths(downloaded_files)

    checked_files = download_checked_pdf_rows()
    if checked_files:
        return zip_paths(checked_files)


    pdf_targets = _scan_click_targets(
        page,
        query="pdf",
        exact=False,
        prefer_text=["附件", "下载", hqpn],
        tag_boost=["A", "BUTTON", "SPAN", "DIV"],
        limit=30,
    )
    attachment_targets = _scan_click_targets(
        page,
        query="附件",
        exact=False,
        prefer_text=["pdf", "下载", hqpn],
        tag_boost=["A", "BUTTON", "SPAN", "DIV"],
        limit=30,
    )
    _write_target_debug(page, output_dir, "attachment_pdf", pdf_targets)
    _write_target_debug(page, output_dir, "attachment_controls", attachment_targets)
    debug_dir = output_dir / "plm_hq_search_debug"
    debug_dir.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(debug_dir / "attachment_page.png"), full_page=True)
    except PlaywrightError:
        pass
    raise RuntimeError("Downloadable PDF attachment links were not detected")

def run_hq_attachment_download(
    playwright: Playwright,
    *,
    username: str,
    password: str,
    hqpn: str,
    output_dir: Path,
    headless: bool = False,
    log: LogFn = _noop_log,
) -> Path:
    hqpn = (hqpn or "").strip()
    if not hqpn:
        raise ValueError("HQ 料号不能为空")

    log("启动浏览器")
    browser = playwright.chromium.launch(headless=headless)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()
    try:
        log("打开 EIP")
        page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)
        wait_for_eip_ready(page, username, password)

        log("进入 PLM")
        plm_page = click_opening_page(
            page,
            page.locator("a").filter(has_text=re.compile(r"^PLM$")),
            timeout=30000,
        )
        login_if_present(plm_page, username, password, timeout=500)
        plm_page = _wait_for_plm_home(context, plm_page, username, password)

        log("Open PLM search page")
        search_page = _open_plm_search_page(context, plm_page, username, password, log=log)

        log(f"搜索料号：{hqpn}")
        _type_top_search(search_page, hqpn)
        try:
            search_page.locator(f"a:has-text('{hqpn}')").first.wait_for(state="visible", timeout=8000)
        except PlaywrightError:
            search_page.wait_for_timeout(2000)

        log("打开第一条搜索结果")
        detail_page = _click_first_search_result(context, search_page, hqpn, output_dir=output_dir)

        log("进入内容页")
        _click_detail_content(detail_page, output_dir=output_dir)

        log("勾选附件并下载")
        output_path = _download_selected_attachments(detail_page, hqpn, output_dir, log=log)
        log(f"下载完成：{output_path}")
        return output_path
    finally:
        context.close()
        browser.close()
