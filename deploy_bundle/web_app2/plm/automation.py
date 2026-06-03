import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, Playwright, TimeoutError as PlaywrightTimeoutError


START_URL = "https://eip.evex-tech.com/"


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


def _click_text_any_page(context, text: str, timeout: int = 10000) -> Page:
    for page in reversed([p for p in context.pages if not p.is_closed()]):
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
    raise RuntimeError(f"未找到可见入口：{text}")


def _type_top_search(page: Page, hqpn: str) -> None:
    # Windchill top search box sits in the fixed header. The classic page exposes
    # awkward frames, so use the verified viewport coordinate.
    page.mouse.click(1082, 20)
    page.wait_for_timeout(300)
    page.keyboard.press("Control+A")
    page.keyboard.press("Backspace")
    page.keyboard.type(hqpn, delay=120)
    page.wait_for_timeout(300)
    page.keyboard.press("Enter")


def _click_first_search_result(context, page: Page, hqpn: str) -> Page:
    before = set(context.pages)
    try:
        link = page.locator(f"a:has-text('{hqpn}')").first
        link.wait_for(state="visible", timeout=3000)
        with page.expect_popup(timeout=5000) as popup_info:
            link.click(timeout=10000)
        opened = popup_info.value
        opened.wait_for_load_state("domcontentloaded", timeout=30000)
        return opened
    except (PlaywrightTimeoutError, PlaywrightError):
        page.mouse.click(265, 343)
        page.wait_for_timeout(3000)
        new_pages = [p for p in context.pages if p not in before and not p.is_closed()]
        return new_pages[-1] if new_pages else page


def _click_detail_content(page: Page) -> None:
    for locator in (
        page.get_by_role("button", name="内容").first,
        page.get_by_role("tab", name="内容").first,
        page.locator("a:has-text('内容')").first,
        page.get_by_text("内容", exact=True).first,
    ):
        try:
            locator.wait_for(state="visible", timeout=5000)
            locator.click(timeout=10000)
            page.wait_for_timeout(3000)
            return
        except PlaywrightError:
            continue
    raise RuntimeError("未找到详情页“内容”页签")


def _download_selected_attachments(page: Page, hqpn: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    page.wait_for_timeout(3000)
    page.mouse.click(59, 276)
    page.wait_for_timeout(1000)

    try:
        with page.expect_download(timeout=30000) as download_info:
            page.mouse.click(91, 236)
        download = download_info.value
        suggested = download.suggested_filename or f"{hqpn}_attachments.zip"
        output_path = output_dir / suggested
        download.save_as(str(output_path))
        return output_path
    except PlaywrightTimeoutError:
        page.mouse.click(91, 236)
        raise RuntimeError("已点击下载选定文件，但未捕获到下载事件")


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

        log("打开功能地图")
        _click_text_any_page(context, "功能地图")

        log("打开搜索")
        search_page = _click_text_any_page(context, "搜索")

        log(f"搜索料号：{hqpn}")
        _type_top_search(search_page, hqpn)
        search_page.wait_for_timeout(5000)

        log("打开第一条搜索结果")
        detail_page = _click_first_search_result(context, search_page, hqpn)

        log("进入内容页")
        _click_detail_content(detail_page)

        log("勾选附件并下载")
        output_path = _download_selected_attachments(detail_page, hqpn, output_dir)
        log(f"下载完成：{output_path}")
        return output_path
    finally:
        context.close()
        browser.close()
