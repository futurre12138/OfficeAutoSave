import argparse
import datetime as dt
import time
from typing import Iterable, Optional, Tuple

import pythoncom
import win32com.client


WORD_PROGIDS = (
    "Word.Application",  # Microsoft Word
    "KWPS.Application",  # WPS 文字（常见）
)
EXCEL_PROGIDS = (
    "Excel.Application",  # Microsoft Excel
    "Ket.Application",  # WPS 表格（常见）
)
PPT_PROGIDS = (
    "PowerPoint.Application",  # Microsoft PowerPoint
    "WPP.Application",  # WPS 演示（常见）
)


def now_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(message: str) -> None:
    print(f"[{now_str()}] {message}", flush=True)


def safe_name(obj, fallback: str = "<unknown>") -> str:
    for attr in ("Name", "FullName"):
        try:
            v = getattr(obj, attr, None)
            if v:
                return str(v)
        except Exception:
            continue
    return fallback


def vendor_from_progid(progid: str) -> str:
    if progid.lower() in {"kwps.application", "ket.application", "wpp.application"}:
        return "WPS"
    return "Office"


def get_active_app(progids: Iterable[str]) -> Tuple[Optional[object], Optional[str]]:
    for progid in progids:
        try:
            app = win32com.client.GetActiveObject(progid)
            return app, progid
        except Exception:
            continue
    return None, None


def iter_word_docs(app) -> Iterable:
    try:
        docs = getattr(app, "Documents", None)
        if docs is None:
            return []
        return list(docs)
    except Exception:
        return []


def iter_excel_books(app) -> Iterable:
    try:
        wbs = getattr(app, "Workbooks", None)
        if wbs is None:
            return []
        return list(wbs)
    except Exception:
        return []


def iter_ppt_presentations(app) -> Iterable:
    try:
        pres = getattr(app, "Presentations", None)
        if pres is None:
            return []
        return list(pres)
    except Exception:
        return []


def try_save(obj) -> bool:
    try:
        if bool(getattr(obj, "ReadOnly", False)):
            return False
        if bool(getattr(obj, "Saved", False)):
            return False
        obj.Save()
        return True
    except Exception:
        return False


def save_word_documents() -> int:
    saved = 0
    app, progid = get_active_app(WORD_PROGIDS)
    if app is None or progid is None:
        return 0

    vendor = vendor_from_progid(progid)
    for doc in iter_word_docs(app):
        try:
            if try_save(doc):
                saved += 1
                log(f"{vendor} 文字 已保存: {safe_name(doc)}")
        except Exception as exc:
            log(f"{vendor} 文字 保存失败 {safe_name(doc)}: {exc}")
    return saved


def save_excel_workbooks() -> int:
    saved = 0
    app, progid = get_active_app(EXCEL_PROGIDS)
    if app is None or progid is None:
        return 0

    vendor = vendor_from_progid(progid)
    for book in iter_excel_books(app):
        try:
            if try_save(book):
                saved += 1
                log(f"{vendor} 表格 已保存: {safe_name(book)}")
        except Exception as exc:
            log(f"{vendor} 表格 保存失败 {safe_name(book)}: {exc}")
    return saved


def save_powerpoint_presentations() -> int:
    saved = 0
    app, progid = get_active_app(PPT_PROGIDS)
    if app is None or progid is None:
        return 0

    vendor = vendor_from_progid(progid)
    for presentation in iter_ppt_presentations(app):
        try:
            if try_save(presentation):
                saved += 1
                log(f"{vendor} 演示 已保存: {safe_name(presentation)}")
        except Exception as exc:
            log(f"{vendor} 演示 保存失败 {safe_name(presentation)}: {exc}")
    return saved


def save_all() -> int:
    total = 0
    total += save_word_documents()
    total += save_excel_workbooks()
    total += save_powerpoint_presentations()
    return total


def run(interval_seconds: int) -> None:
    log(f"Office/WPS 自动保存已启动，间隔 {interval_seconds} 秒。按 Ctrl+C 停止。")
    log("提示：只会连接“已在运行中”的 Office/WPS，不会自动启动应用。")
    while True:
        pythoncom.CoInitialize()
        try:
            saved_count = save_all()
            if saved_count == 0:
                log("本轮无需保存。")
            else:
                log(f"本轮完成，已保存 {saved_count} 个文件。")
        finally:
            pythoncom.CoUninitialize()
        time.sleep(interval_seconds)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Office/WPS 自动保存脚本（文字/表格/演示）"
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=30,
        help="自动保存间隔秒数（默认 30）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
#    if args.interval < 5:
#        raise ValueError("间隔不能小于 5 秒，建议 10 秒以上。")
    run(args.interval)
