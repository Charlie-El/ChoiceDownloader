import argparse
import logging
import os
import re
import shutil
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
from pywinauto.mouse import move, scroll


APP_DIR = Path(__file__).resolve().parent
DEFAULT_EXE = ""
DEFAULT_BASE_DIR = ""
DEFAULT_LOG_DIR = ""
SAFE_WINDOW_KEYWORD = "Choice"
GUI_WINDOW_EXCLUDE_HINTS = ("公告下载助手", "announcement_workbench", "visual studio code")
MAX_BATCH_DOWNLOAD_COUNT = 5
DEFAULT_MAX_BATCH_DOWNLOAD_COUNT = 100

COORDINATE_DEFAULTS = {
    "left_nav_scroll": (250, 980),
    "company_announcement": (150, 1442),
    "all_announcements": (900, 360),
    "financial_report": (900, 470),
    "batch_download": (2300, 370),
    "popup_browse": (1532, 780),
    "popup_range_checkbox": (1070, 940),
    "popup_range_input": (1340, 930),
    "popup_download": (1530, 990),
    "return_home": (110, 118),
}

CALIBRATION_TARGET_CHOICES = tuple(COORDINATE_DEFAULTS.keys())

LEFT_NAV_SCROLL_POINT = COORDINATE_DEFAULTS["left_nav_scroll"]
LEFT_NAV_COMPANY_ANNOUNCEMENT_POINT = COORDINATE_DEFAULTS["company_announcement"]
ALL_ANNOUNCEMENTS_POINT = COORDINATE_DEFAULTS["all_announcements"]
FINANCIAL_REPORT_POINT = COORDINATE_DEFAULTS["financial_report"]
ANNOUNCEMENT_BATCH_BUTTON_POINT = COORDINATE_DEFAULTS["batch_download"]
POPUP_BATCH_BROWSE_BUTTON_POINT = COORDINATE_DEFAULTS["popup_browse"]
POPUP_BATCH_RANGE_CHECKBOX_POINT = COORDINATE_DEFAULTS["popup_range_checkbox"]
POPUP_BATCH_RANGE_INPUT_POINT = COORDINATE_DEFAULTS["popup_range_input"]
POPUP_BATCH_DOWNLOAD_BUTTON_POINT = COORDINATE_DEFAULTS["popup_download"]
RETURN_HOME_POINT = COORDINATE_DEFAULTS["return_home"]


class StopRequestedError(RuntimeError):
    pass


class CompanyNotFoundError(RuntimeError):
    pass


class LatestFilterError(RuntimeError):
    pass


def setup_logging(log_dir: Path, extra_handlers: list[logging.Handler] | None = None) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"choice_run_{datetime.now():%Y%m%d_%H%M%S}.log"
    handlers: list[logging.Handler] = [
        logging.FileHandler(log_path, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
    if extra_handlers:
        handlers.extend(extra_handlers)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True,
    )
    logging.info("logging started: %s", log_path)
    return log_path


def fail_safe(message: str) -> int:
    logging.error(message)
    return 1


def check_stop(stop_event=None):
    if stop_event is not None and stop_event.is_set():
        logging.info("stop requested")
        raise StopRequestedError("Run stopped by user.")


def controlled_sleep(seconds: float, stop_event=None, interval: float = 0.1):
    if seconds <= 0:
        check_stop(stop_event)
        return
    end_time = time.time() + seconds
    while time.time() < end_time:
        check_stop(stop_event)
        time.sleep(min(interval, end_time - time.time()))
    check_stop(stop_event)


def parse_coordinate_overrides(raw_items: list[str] | None) -> dict[str, tuple[int, int]]:
    overrides: dict[str, tuple[int, int]] = {}
    for raw_item in raw_items or []:
        if "=" not in raw_item:
            raise ValueError(f"Coordinate override must be name=x,y: {raw_item}")
        name, raw_point = raw_item.split("=", 1)
        name = name.strip()
        if name not in COORDINATE_DEFAULTS:
            raise ValueError(f"Unknown coordinate target: {name}")
        parts = [part.strip() for part in raw_point.split(",")]
        if len(parts) != 2:
            raise ValueError(f"Coordinate override must be name=x,y: {raw_item}")
        try:
            x = int(parts[0])
            y = int(parts[1])
        except ValueError as exc:
            raise ValueError(f"Coordinate values must be integers: {raw_item}") from exc
        overrides[name] = (x, y)
    return overrides


def get_coordinate(name: str, coordinate_overrides: dict[str, tuple[int, int]] | None = None) -> tuple[int, int]:
    if coordinate_overrides and name in coordinate_overrides:
        return coordinate_overrides[name]
    return COORDINATE_DEFAULTS[name]


def ensure_safe_report_name(report_name: str, whitelist: set[str]):
    if report_name not in whitelist:
        raise ValueError(
            f"Report name '{report_name}' is not in the whitelist: {sorted(whitelist)}"
        )


def ensure_download_count(count: int, max_count: int = MAX_BATCH_DOWNLOAD_COUNT):
    if count < 1:
        raise ValueError("Batch download count must be at least 1.")
    if count > max_count:
        raise ValueError(
            f"Batch download count {count} exceeds safe limit {max_count}."
        )


def rect_tuple(rect):
    return (rect.left, rect.top, rect.right, rect.bottom, rect.width(), rect.height())


def wait_for_dialog(keyword: str, timeout: int = 10, stop_event=None):
    deadline = time.time() + timeout
    while time.time() < deadline:
        check_stop(stop_event)
        for backend in ("win32", "uia"):
            for win in Desktop(backend=backend).windows():
                try:
                    title = win.window_text()
                except Exception:
                    continue
                if keyword in title:
                    logging.info("found dialog via %s: %s", backend, title)
                    return win
        controlled_sleep(0.5, stop_event)
    raise RuntimeError(f"Dialog with keyword '{keyword}' not found.")


def wait_for_dialog_to_close(keyword: str, timeout: int = 10, stop_event=None) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        check_stop(stop_event)
        found = False
        for backend in ("win32", "uia"):
            for win in Desktop(backend=backend).windows():
                try:
                    if keyword in win.window_text():
                        found = True
                        break
                except Exception:
                    continue
            if found:
                break
        if not found:
            logging.info("dialog closed: %s", keyword)
            return True
        controlled_sleep(0.3, stop_event)
    return False


def find_choice_window(timeout: int = 30, stop_event=None):
    for _ in range(timeout):
        check_stop(stop_event)
        matches = []
        for win in Desktop(backend="uia").windows():
            try:
                title = win.window_text()
                if any(exclude.lower() in title.lower() for exclude in GUI_WINDOW_EXCLUDE_HINTS):
                    continue
                if SAFE_WINDOW_KEYWORD in title:
                    matches.append(win)
            except Exception:
                continue
        if matches:
            return matches[0]
        controlled_sleep(1, stop_event)
    return None


def launch_or_connect(exe_path: Path, stop_event=None):
    window = find_choice_window(timeout=3, stop_event=stop_event)
    if window:
        logging.info("attached existing window: %s", window.window_text())
        return window

    if not exe_path.exists():
        raise FileNotFoundError(f"Choice executable not found: {exe_path}")

    logging.info("starting executable: %s", exe_path)
    if exe_path.suffix.lower() == ".lnk":
        os.startfile(str(exe_path))
    else:
        Application(backend="uia").start(str(exe_path))
    window = find_choice_window(timeout=30, stop_event=stop_event)
    if not window:
        raise RuntimeError("Choice window did not appear within 30 seconds.")
    logging.info("started and connected window: %s", window.window_text())
    return window


def find_bottom_search_edit(window):
    edits = [ctrl for ctrl in window.descendants() if ctrl.element_info.control_type == "Edit"]
    candidates = []
    window_rect = window.rectangle()
    min_top = max(window_rect.bottom - 80, 0)

    for edit in edits:
        try:
            rect = edit.rectangle()
        except Exception:
            continue
        if (
            rect.width() >= 180
            and rect.height() >= 20
            and rect.top >= min_top
            and rect.right >= window_rect.right - 700
        ):
            candidates.append((rect.top, -rect.width(), edit))

    if not candidates:
        raise RuntimeError("No likely bottom search input was found.")

    candidates.sort()
    return candidates[0][2]


def input_bottom_search(window, text: str, stop_event=None):
    search_edit = find_bottom_search_edit(window)
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.3, stop_event)
    search_edit.click_input()
    controlled_sleep(0.5, stop_event)
    send_keys("^a{BACKSPACE}")
    send_keys(text, with_spaces=True)
    logging.info("typed text into bottom search: %s", text)


def submit_search_with_enter(window, stop_event=None):
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.3, stop_event)
    send_keys("{ENTER}")
    logging.info("submitted bottom search with Enter")


def send_f9(window, stop_event=None):
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.5, stop_event)
    send_keys("{F9}")
    logging.info("sent F9 to active Choice window")


def click_window_relative_point(window, point: tuple[int, int], step_name: str, stop_event=None):
    rect = window.rectangle()
    x = rect.left + point[0]
    y = rect.top + point[1]
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.3, stop_event)
    window.click_input(coords=point)
    logging.info("clicked %s at relative=%s absolute=(%s, %s)", step_name, point, x, y)


def move_mouse_to_window_relative_point(window, point: tuple[int, int], step_name: str, stop_event=None):
    rect = window.rectangle()
    x = rect.left + point[0]
    y = rect.top + point[1]
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.3, stop_event)
    move(coords=(x, y))
    logging.info("moved mouse to %s at relative=%s absolute=(%s, %s)", step_name, point, x, y)


def scroll_window_relative_point(window, point: tuple[int, int], wheel_dist: int, step_name: str, stop_event=None):
    rect = window.rectangle()
    x = rect.left + point[0]
    y = rect.top + point[1]
    check_stop(stop_event)
    window.set_focus()
    controlled_sleep(0.3, stop_event)
    scroll(coords=(x, y), wheel_dist=wheel_dist)
    logging.info("scrolled %s at relative=%s absolute=(%s, %s) wheel=%s", step_name, point, x, y, wheel_dist)


def scroll_left_nav_to_bottom(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    for index in range(4):
        scroll_window_relative_point(
            window,
            get_coordinate("left_nav_scroll", coordinate_overrides),
            -14,
            f"left_nav_down_{index + 1}",
            stop_event=stop_event,
        )
        controlled_sleep(0.4, stop_event)


def open_company_announcement(
    window,
    report_name: str,
    enter_wait_seconds: float,
    post_f9_wait_seconds: float,
    stop_event=None,
    coordinate_overrides: dict[str, tuple[int, int]] | None = None,
):
    input_bottom_search(window, report_name, stop_event=stop_event)
    controlled_sleep(1.0, stop_event)
    logging.info("waited 1.0 second after typing company name before pressing Enter")
    submit_search_with_enter(window, stop_event=stop_event)
    controlled_sleep(enter_wait_seconds, stop_event)
    send_f9(window, stop_event=stop_event)
    controlled_sleep(post_f9_wait_seconds, stop_event)
    scroll_left_nav_to_bottom(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    controlled_sleep(0.8, stop_event)
    click_window_relative_point(
        window,
        get_coordinate("company_announcement", coordinate_overrides),
        "company_announcement",
        stop_event=stop_event,
    )
    controlled_sleep(3.2, stop_event)
    logging.info("waited 3.2 seconds before all announcements filter")


def click_all_announcements(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    click_window_relative_point(
        window,
        get_coordinate("all_announcements", coordinate_overrides),
        "all_announcements",
        stop_event=stop_event,
    )
    controlled_sleep(1.0, stop_event)


def click_financial_report(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    click_window_relative_point(
        window,
        get_coordinate("financial_report", coordinate_overrides),
        "financial_report",
        stop_event=stop_event,
    )
    controlled_sleep(1.2, stop_event)


def open_financial_reports(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    click_all_announcements(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    click_financial_report(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)


def move_to_calibration_target(
    window,
    target_name: str,
    report_name: str,
    enter_wait_seconds: float,
    post_f9_wait_seconds: float,
    stop_event=None,
    coordinate_overrides: dict[str, tuple[int, int]] | None = None,
):
    if target_name not in COORDINATE_DEFAULTS:
        raise ValueError(f"Unknown calibration target: {target_name}")

    if target_name == "return_home":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("return_home", coordinate_overrides),
            "calibrate_return_home",
            stop_event=stop_event,
        )
        return

    input_bottom_search(window, report_name, stop_event=stop_event)
    controlled_sleep(1.0, stop_event)
    logging.info("waited 1.0 second after typing company name before pressing Enter")
    submit_search_with_enter(window, stop_event=stop_event)
    controlled_sleep(enter_wait_seconds, stop_event)
    send_f9(window, stop_event=stop_event)
    controlled_sleep(post_f9_wait_seconds, stop_event)

    if target_name == "left_nav_scroll":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("left_nav_scroll", coordinate_overrides),
            "calibrate_left_nav_scroll",
            stop_event=stop_event,
        )
        return

    scroll_left_nav_to_bottom(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    controlled_sleep(0.8, stop_event)
    if target_name == "company_announcement":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("company_announcement", coordinate_overrides),
            "calibrate_company_announcement",
            stop_event=stop_event,
        )
        return

    click_window_relative_point(
        window,
        get_coordinate("company_announcement", coordinate_overrides),
        "company_announcement",
        stop_event=stop_event,
    )
    controlled_sleep(3.2, stop_event)
    logging.info("waited 3.2 seconds before all announcements filter")

    if target_name == "all_announcements":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("all_announcements", coordinate_overrides),
            "calibrate_all_announcements",
            stop_event=stop_event,
        )
        return

    click_all_announcements(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    if target_name == "financial_report":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("financial_report", coordinate_overrides),
            "calibrate_financial_report",
            stop_event=stop_event,
        )
        return

    click_financial_report(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    if target_name == "batch_download":
        move_mouse_to_window_relative_point(
            window,
            get_coordinate("batch_download", coordinate_overrides),
            "calibrate_batch_download",
            stop_event=stop_event,
        )
        return

    popup_targets = {"popup_browse", "popup_range_checkbox", "popup_range_input", "popup_download"}
    if target_name in popup_targets:
        click_window_relative_point(
            window,
            get_coordinate("batch_download", coordinate_overrides),
            "batch_download_button",
            stop_event=stop_event,
        )
        controlled_sleep(1.2, stop_event)
        move_mouse_to_window_relative_point(
            window,
            get_coordinate(target_name, coordinate_overrides),
            f"calibrate_{target_name}",
            stop_event=stop_event,
        )
        return

    raise ValueError(f"Unsupported calibration target: {target_name}")


def select_folder_with_dialog(path: Path):
    time.sleep(1.0)
    focus_deadline = time.time() + 5
    while True:
        folder_dialog = wait_for_dialog("浏览文件夹", timeout=10)
        try:
            folder_dialog.set_focus()
            break
        except Exception:
            if time.time() >= focus_deadline:
                raise
            time.sleep(0.3)
    time.sleep(0.3)
    send_keys("%d")
    time.sleep(0.4)
    send_keys("^a{BACKSPACE}")
    time.sleep(0.2)
    send_keys(str(path), with_spaces=True)
    time.sleep(0.3)
    send_keys("{ENTER}")
    time.sleep(0.4)
    send_keys("{ENTER}")
    if not wait_for_dialog_to_close("浏览文件夹", timeout=3):
        send_keys("{ENTER}")
        time.sleep(0.4)
        send_keys("{ENTER}")
        if not wait_for_dialog_to_close("浏览文件夹", timeout=4):
            raise RuntimeError("Folder dialog did not close after pressing Enter.")
    logging.info("selected folder in browser dialog: %s", path)


def select_folder_with_dialog_retry(path: Path, stop_event=None):
    controlled_sleep(1.0, stop_event)
    wait_for_dialog("浏览文件夹", timeout=10, stop_event=stop_event)
    controlled_sleep(0.3, stop_event)
    for attempt in range(3):
        check_stop(stop_event)
        send_keys("%d")
        controlled_sleep(0.5, stop_event)
        send_keys("^a{BACKSPACE}")
        controlled_sleep(0.2, stop_event)
        send_keys(str(path), with_spaces=True)
        controlled_sleep(0.3, stop_event)
        send_keys("{ENTER}")
        controlled_sleep(0.4, stop_event)
        send_keys("{ENTER}")
        controlled_sleep(0.4, stop_event)
        if wait_for_dialog_to_close("浏览文件夹", timeout=4, stop_event=stop_event):
            logging.info("folder dialog closed on attempt %s", attempt + 1)
            logging.info("selected folder in browser dialog: %s", path)
            return
        controlled_sleep(0.4, stop_event)
    raise RuntimeError("Folder dialog did not close after pressing Enter.")


def configure_batch_download_dialog(
    window,
    download_root: Path,
    batch_count: int,
    skip_folder_dialog: bool = False,
    stop_event=None,
    coordinate_overrides: dict[str, tuple[int, int]] | None = None,
):
    controlled_sleep(1.2, stop_event)
    if skip_folder_dialog:
        logging.info("skipped folder dialog handling by option")
    else:
        click_window_relative_point(
            window,
            get_coordinate("popup_browse", coordinate_overrides),
            "batch_browse_button",
            stop_event=stop_event,
        )
        try:
            select_folder_with_dialog_retry(download_root, stop_event=stop_event)
        except RuntimeError as exc:
            message = str(exc)
            if "浏览文件夹" in message or "Dialog with keyword" in message:
                raise CompanyNotFoundError("没有找到公司：未找到“浏览文件夹”弹窗，可能没有进入正确的公司公告下载页面。") from exc
            raise
    click_window_relative_point(
        window,
        get_coordinate("popup_range_checkbox", coordinate_overrides),
        "batch_range_checkbox",
        stop_event=stop_event,
    )
    controlled_sleep(0.2, stop_event)
    click_window_relative_point(
        window,
        get_coordinate("popup_range_input", coordinate_overrides),
        "batch_range_input",
        stop_event=stop_event,
    )
    controlled_sleep(0.2, stop_event)
    send_keys("{RIGHT}{RIGHT}{RIGHT}")
    controlled_sleep(0.1, stop_event)
    send_keys("{BACKSPACE}{BACKSPACE}{BACKSPACE}")
    controlled_sleep(0.1, stop_event)
    send_keys(str(batch_count))
    logging.info("set batch range count: %s", batch_count)
    click_window_relative_point(
        window,
        get_coordinate("popup_download", coordinate_overrides),
        "batch_download_confirm",
        stop_event=stop_event,
    )
    logging.info("submitted batch download dialog")


def click_batch_download(
    window,
    download_root: Path,
    batch_count: int,
    skip_folder_dialog: bool = False,
    stop_event=None,
    coordinate_overrides: dict[str, tuple[int, int]] | None = None,
):
    click_window_relative_point(
        window,
        get_coordinate("batch_download", coordinate_overrides),
        "batch_download_button",
        stop_event=stop_event,
    )
    configure_batch_download_dialog(
        window,
        download_root,
        batch_count,
        skip_folder_dialog=skip_folder_dialog,
        stop_event=stop_event,
        coordinate_overrides=coordinate_overrides,
    )


def wait_for_batch_files(
    download_root: Path,
    before_paths: set[Path],
    expected_count: int,
    stable_seconds: float = 2.0,
    timeout: int = 60,
    stop_event=None,
) -> list[Path]:
    deadline = time.time() + timeout
    last_signature = None
    stable_since = None
    max_seen_count = 0

    while time.time() < deadline:
        check_stop(stop_event)
        current_paths = {path for path in download_root.rglob("*") if path.is_file()}
        new_paths = sorted(current_paths - before_paths, key=lambda p: p.stat().st_mtime, reverse=True)
        filtered_paths = [path for path in new_paths if "最新" not in path.parts]
        if len(filtered_paths) > max_seen_count:
            max_seen_count = len(filtered_paths)
            logging.info("downloaded file count updated: %s/%s", max_seen_count, expected_count)

        if filtered_paths:
            signature = tuple(
                (str(path), path.stat().st_size, int(path.stat().st_mtime))
                for path in sorted(filtered_paths, key=lambda p: str(p))
            )
            if signature == last_signature:
                if stable_since is None:
                    stable_since = time.time()
                elif time.time() - stable_since >= stable_seconds:
                    logging.info("detected downloaded files: %s", [str(p) for p in filtered_paths])
                    return filtered_paths
            else:
                last_signature = signature
                stable_since = None

        controlled_sleep(1.0, stop_event)

    if max_seen_count > 0:
        raise RuntimeError(
            f"Downloaded files did not stabilize in time: {download_root}. "
            f"Observed {max_seen_count} file(s), expected up to {expected_count}."
        )
    raise RuntimeError(f"No downloaded files appeared in time: {download_root}")


def extract_announcement_date(path: Path):
    match = re.search(r"(20\d{2}-\d{2}-\d{2})", path.name)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%Y-%m-%d")
    except ValueError:
        return None


def parse_filename_keywords(raw_keywords: str | None) -> list[str]:
    if not raw_keywords:
        return []
    return [keyword.strip() for keyword in re.split(r"[,，]", raw_keywords) if keyword.strip()]


def normalize_filename_match_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).upper()
    return "".join(char for char in text if char.isalnum())


def normalize_keyword_match_mode(raw_mode: str | None) -> str:
    mode = (raw_mode or "or").strip().lower()
    if mode not in {"and", "or"}:
        raise ValueError(f"Unsupported keyword match mode: {raw_mode}")
    return mode


def filename_matches_keywords(path: Path, keywords: list[str], match_mode: str = "or") -> bool:
    if not keywords:
        return True
    upper_name = path.name.upper()
    normalized_name = normalize_filename_match_text(path.name)
    normalized_mode = normalize_keyword_match_mode(match_mode)
    checks = [
        keyword.upper() in upper_name or normalize_filename_match_text(keyword) in normalized_name
        for keyword in keywords
    ]
    return all(checks) if normalized_mode == "and" else any(checks)


def get_download_folder(downloaded_paths: list[Path]) -> Path:
    parents = {path.parent for path in downloaded_paths}
    if len(parents) != 1:
        raise RuntimeError(f"Downloaded files are not in a single folder: {sorted(str(p) for p in parents)}")
    return next(iter(parents))


def copy_latest_file_to_latest_folder(
    downloaded_paths: list[Path],
    filename_keywords: list[str] | None = None,
    keyword_match_mode: str = "or",
    latest_only: bool = True,
) -> list[Path]:
    download_folder = get_download_folder(downloaded_paths)
    latest_dir = download_folder / "最新"
    latest_dir.mkdir(exist_ok=True)
    matched_paths = [
        path for path in downloaded_paths if filename_matches_keywords(path, filename_keywords or [], keyword_match_mode)
    ]
    if not matched_paths:
        if filename_keywords:
            raise LatestFilterError(
                f"没筛选成功：下载文件中没有匹配关键词 {filename_keywords}，匹配方式={normalize_keyword_match_mode(keyword_match_mode)}。"
            )
        raise LatestFilterError("没筛选成功：没有可复制到“最新”文件夹的下载文件。")

    for existing_path in latest_dir.iterdir():
        if existing_path.is_file():
            existing_path.unlink()

    sorted_paths = sorted(
        matched_paths,
        key=lambda p: (
            extract_announcement_date(p) or datetime.min,
            p.stat().st_mtime,
        ),
        reverse=True,
    )
    selected_paths = sorted_paths[:1] if latest_only else sorted_paths
    copied_targets: list[Path] = []
    for source_path in selected_paths:
        target = latest_dir / source_path.name
        shutil.copy2(source_path, target)
        copied_targets.append(target)

    logging.info(
        "copied %s filtered file(s) to latest dir with keywords=%s latest_only=%s: %s",
        len(copied_targets),
        filename_keywords or [],
        latest_only,
        [str(path) for path in copied_targets],
    )
    return copied_targets


def get_popup_batch_target_point(
    target_name: str,
    coordinate_overrides: dict[str, tuple[int, int]] | None = None,
) -> tuple[int, int]:
    mapping = {
        "browse": "popup_browse",
        "range_checkbox": "popup_range_checkbox",
        "range_input": "popup_range_input",
        "download": "popup_download",
    }
    if target_name not in mapping:
        raise ValueError(f"unknown popup target: {target_name}")
    return get_coordinate(mapping[target_name], coordinate_overrides)


def return_to_homepage(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    click_window_relative_point(
        window,
        get_coordinate("return_home", coordinate_overrides),
        "return_home",
        stop_event=stop_event,
    )
    controlled_sleep(1.0, stop_event)


def try_return_to_homepage(window, stop_event=None, coordinate_overrides: dict[str, tuple[int, int]] | None = None):
    try:
        return_to_homepage(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    except Exception as exc:
        logging.warning("failed to return homepage after error: %s", exc)


def remove_evidence_dir(workdir: Path):
    evidence_dir = workdir / "evidence"
    if evidence_dir.exists():
        shutil.rmtree(evidence_dir)
        logging.info("removed evidence dir: %s", evidence_dir)
    else:
        logging.info("evidence dir already absent: %s", evidence_dir)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Choice announcement downloader.")
    parser.add_argument("--exe", default=DEFAULT_EXE, help="Full path to Choice.exe or a Choice shortcut.")
    parser.add_argument("--report-name", default="", help="Whitelisted company name")
    parser.add_argument(
        "--allowed-report",
        action="append",
        dest="allowed_reports",
        help="Whitelist entry for allowed company names. Can be repeated.",
    )
    parser.add_argument(
        "--base-dir",
        default=str(DEFAULT_BASE_DIR),
        help="Base output directory used directly for downloads.",
    )
    parser.add_argument(
        "--log-dir",
        default=str(DEFAULT_LOG_DIR),
        help="Directory for step-by-step logs.",
    )
    parser.add_argument(
        "--batch-count",
        type=int,
        default=0,
        help="How many recent announcements to batch-download.",
    )
    parser.add_argument(
        "--max-batch-count",
        type=int,
        default=0,
        help="Maximum allowed batch download count for this run.",
    )
    parser.add_argument(
        "--enter-wait-seconds",
        type=float,
        default=4.0,
        help="How long to wait after Enter before F9.",
    )
    parser.add_argument(
        "--post-f9-wait-seconds",
        type=float,
        default=5.0,
        help="How long to wait after F9 before clicking company announcement.",
    )
    parser.add_argument(
        "--skip-navigation",
        action="store_true",
        help="Assume the current page is already the company announcement page.",
    )
    parser.add_argument(
        "--pause-before-batch",
        action="store_true",
        help="Pause after reaching the company announcement page, before clicking batch download.",
    )
    parser.add_argument(
        "--show-batch-target",
        action="store_true",
        help="Move mouse to the batch download target position and pause before clicking.",
    )
    parser.add_argument(
        "--pause-after-batch-click",
        action="store_true",
        help="Pause immediately after clicking batch download.",
    )
    parser.add_argument(
        "--show-popup-batch-target",
        choices=["browse", "range_checkbox", "range_input", "download"],
        help="After clicking batch download, move the mouse to a popup target and pause.",
    )
    parser.add_argument(
        "--skip-folder-dialog",
        action="store_true",
        help="Skip the browse-folder step in the batch download popup.",
    )
    parser.add_argument(
        "--show-return-home-target",
        action="store_true",
        help="After download completes, move the mouse to the return-home target and pause.",
    )
    parser.add_argument(
        "--skip-return-home",
        action="store_true",
        help="Skip clicking back to the homepage after download completes.",
    )
    parser.add_argument(
        "--pause-before-return-home",
        action="store_true",
        help="Pause after download completes, before clicking back to the homepage.",
    )
    parser.add_argument(
        "--navigation-only",
        action="store_true",
        help="Stop after opening the company announcement page.",
    )
    parser.add_argument(
        "--filename-keywords",
        default="",
        help="Comma-separated filename keywords used to keep matching downloaded files. Leave empty to keep all filenames.",
    )
    parser.add_argument(
        "--keyword-match-mode",
        choices=["and", "or"],
        default="or",
        help="How filename keywords are matched: 'or' keeps files containing any keyword, 'and' requires all keywords.",
    )
    parser.add_argument(
        "--keep-all-matches",
        action="store_true",
        help="Keep all keyword-matching files instead of only the newest dated file.",
    )
    parser.add_argument(
        "--coordinate",
        action="append",
        dest="coordinates",
        help="Override a relative click target as name=x,y. Can be repeated.",
    )
    parser.add_argument(
        "--calibrate-target",
        choices=CALIBRATION_TARGET_CHOICES,
        help="Move the mouse to the selected target and stop before performing business actions.",
    )
    return parser


def parse_args(argv: list[str] | None = None):
    return build_parser().parse_args(argv)


def run_with_args(args, extra_log_handlers: list[logging.Handler] | None = None, stop_event=None):
    is_calibration_run = bool(getattr(args, "calibrate_target", None))
    if not str(args.exe or "").strip():
        return fail_safe("Choice executable path is required.")
    if not str(args.report_name or "").strip():
        return fail_safe("Company name is required.")
    if not is_calibration_run and not str(args.base_dir or "").strip():
        return fail_safe("Download directory is required.")
    if not str(args.log_dir or "").strip():
        return fail_safe("Log directory is required.")
    whitelist = set(args.allowed_reports or [args.report_name])
    base_dir = Path(args.base_dir) if str(args.base_dir or "").strip() else Path.cwd()
    log_dir = Path(args.log_dir)
    setup_logging(log_dir, extra_handlers=extra_log_handlers)

    try:
        coordinate_overrides = parse_coordinate_overrides(getattr(args, "coordinates", None))
        ensure_safe_report_name(args.report_name, whitelist)
        if not is_calibration_run:
            ensure_download_count(args.batch_count, args.max_batch_count)
            base_dir.mkdir(parents=True, exist_ok=True)
        remove_evidence_dir(Path.cwd())
    except Exception as exc:
        return fail_safe(str(exc))

    try:
        window = launch_or_connect(Path(args.exe), stop_event=stop_event)
        logging.info("connected window rect: %s", rect_tuple(window.rectangle()))
    except Exception as exc:
        return fail_safe(str(exc))

    try:
        calibrate_target = getattr(args, "calibrate_target", None)
        if calibrate_target:
            move_to_calibration_target(
                window,
                calibrate_target,
                args.report_name,
                enter_wait_seconds=args.enter_wait_seconds,
                post_f9_wait_seconds=args.post_f9_wait_seconds,
                stop_event=stop_event,
                coordinate_overrides=coordinate_overrides,
            )
            logging.info("calibration target reached: %s", calibrate_target)
            return 0

        if not args.skip_navigation:
            open_company_announcement(
                window,
                args.report_name,
                enter_wait_seconds=args.enter_wait_seconds,
                post_f9_wait_seconds=args.post_f9_wait_seconds,
                stop_event=stop_event,
                coordinate_overrides=coordinate_overrides,
            )
        if args.navigation_only:
            logging.info("navigation-only mode finished")
            return 0
        if args.show_batch_target:
            move_mouse_to_window_relative_point(
                window,
                get_coordinate("batch_download", coordinate_overrides),
                "batch_download_target",
                stop_event=stop_event,
            )
            input("Mouse moved to batch download target. Adjust the coordinate if needed, then press Enter to continue...")
        if args.pause_before_batch:
            logging.info("paused before batch download click; inspect the UI and press Enter to continue")
            input("Paused before batch download. Press Enter to continue...")

        open_financial_reports(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)

        before_paths = {path for path in base_dir.rglob("*") if path.is_file()}

        if args.pause_after_batch_click or args.show_popup_batch_target:
            click_window_relative_point(
                window,
                get_coordinate("batch_download", coordinate_overrides),
                "batch_download_button",
                stop_event=stop_event,
            )
            if args.show_popup_batch_target:
                target_point = get_popup_batch_target_point(args.show_popup_batch_target, coordinate_overrides)
                move_mouse_to_window_relative_point(
                    window,
                    target_point,
                    f"popup_batch_{args.show_popup_batch_target}_target",
                    stop_event=stop_event,
                )
                input(
                    f"Mouse moved to popup target '{args.show_popup_batch_target}'. "
                    "Adjust the coordinate if needed, then press Enter to continue..."
                )
            if args.pause_after_batch_click:
                input("Paused after batch download click. Inspect the result, then press Enter to continue...")
            configure_batch_download_dialog(
                window,
                base_dir,
                args.batch_count,
                args.skip_folder_dialog,
                stop_event=stop_event,
                coordinate_overrides=coordinate_overrides,
            )
        else:
            click_batch_download(
                window,
                base_dir,
                args.batch_count,
                args.skip_folder_dialog,
                stop_event=stop_event,
                coordinate_overrides=coordinate_overrides,
            )

        downloaded_paths = wait_for_batch_files(base_dir, before_paths, args.batch_count, stop_event=stop_event)
        copy_latest_file_to_latest_folder(
            downloaded_paths,
            filename_keywords=parse_filename_keywords(args.filename_keywords),
            keyword_match_mode=args.keyword_match_mode,
            latest_only=not args.keep_all_matches,
        )

        if args.show_return_home_target:
            move_mouse_to_window_relative_point(
                window,
                get_coordinate("return_home", coordinate_overrides),
                "return_home_target",
                stop_event=stop_event,
            )
            input("Mouse moved to return-home target. Adjust the coordinate if needed, then press Enter to continue...")
        if args.pause_before_return_home and not args.skip_return_home:
            input("Paused before returning home. Press Enter to continue...")
        if not args.skip_return_home:
            return_to_homepage(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
    except StopRequestedError as exc:
        return fail_safe(str(exc))
    except Exception as exc:
        try:
            if "window" in locals() and window is not None and not getattr(args, "skip_return_home", False):
                try_return_to_homepage(window, stop_event=stop_event, coordinate_overrides=coordinate_overrides)
        except StopRequestedError:
            return fail_safe("Run stopped by user.")
        return fail_safe(f"automation failed: {exc}")

    logging.info("download flow finished")
    return 0


def main(argv: list[str] | None = None):
    args = parse_args(argv)
    return run_with_args(args)


if __name__ == "__main__":
    raise SystemExit(main())
