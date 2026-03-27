import argparse
import datetime as dt
import time
from typing import Iterable

import pythoncom
import win32com.client


def now_str() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(message: str) -> None:
    print(f"[{now_str()}] {message}", flush=True)


def safe_name(obj, fallback: str = "<unknown>") -> str:
    try:
        return str(obj.Name)
    except Exception:
        return fallback


def iter_word_docs(app) -> Iterable:
    try:
        return list(app.Documents)
    except Exception:
        return []


def iter_excel_books(app) -> Iterable:
    try:
        return list(app.Workbooks)
    except Exception:
        return []


def iter_ppt_presentations(app) -> Iterable:
    try:
        return list(app.Presentations)
    except Exception:
        return []


def save_word_documents() -> int:
    saved = 0
    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:
        return 0

    for doc in iter_word_docs(word):
        try:
            if bool(doc.ReadOnly):
                continue
            if bool(doc.Saved):
                continue
            doc.Save()
            saved += 1
            log(f"Word 已保存: {safe_name(doc)}")
        except Exception as exc:
            log(f"Word 保存失败 {safe_name(doc)}: {exc}")
    return saved


def save_excel_workbooks() -> int:
    saved = 0
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return 0

    for book in iter_excel_books(excel):
        try:
            if bool(book.ReadOnly):
                continue
            if bool(book.Saved):
                continue
            book.Save()
            saved += 1
            log(f"Excel 已保存: {safe_name(book)}")
        except Exception as exc:
            log(f"Excel 保存失败 {safe_name(book)}: {exc}")
    return saved


def save_powerpoint_presentations() -> int:
    saved = 0
    try:
        powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        return 0

    for presentation in iter_ppt_presentations(powerpoint):
        try:
            if bool(presentation.Saved):
                continue
            presentation.Save()
            saved += 1
            log(f"PowerPoint 已保存: {safe_name(presentation)}")
        except Exception as exc:
            log(f"PowerPoint 保存失败 {safe_name(presentation)}: {exc}")
    return saved


def save_all() -> int:
    total = 0
    total += save_word_documents()
    total += save_excel_workbooks()
    total += save_powerpoint_presentations()
    return total


def run(interval_seconds: int) -> None:
    log(f"Office 自动保存已启动，间隔 {interval_seconds} 秒。按 Ctrl+C 停止。")
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
        description="Office 自动保存脚本（Word/Excel/PowerPoint）"
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
#     if args.interval < 5:
#        raise ValueError("间隔不能小于 5 秒，建议 10 秒以上。")
    run(args.interval)
import argparse
import datetime as _dt
import os
import sys
import time
from dataclasses import dataclass

import pythoncom
import win32com.client


@dataclass(frozen=True)
class SaveResult:
    app: str
    name: str
    source: str
    backup: str


def _now_stamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def _safe_filename(s: str) -> str:
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        s = s.replace(ch, "_")
    return s.strip().strip(".") or "untitled"


def _ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)


def _backup_path(base_dir: str, source_fullname: str, stamp: str) -> str:
    # Mirror drive+path structure under base_dir to avoid collisions.
    # Example: C:\a\b.docx -> <base_dir>\C\a\b\20260327_120000__b.docx
    drive, tail = os.path.splitdrive(source_fullname)
    drive = drive.rstrip(":\\/")
    rel_dir, filename = os.path.split(tail.lstrip("\\/"))
    filename = _safe_filename(filename)
    out_dir = os.path.join(base_dir, drive, rel_dir)
    _ensure_dir(out_dir)
    return os.path.join(out_dir, f"{stamp}__{filename}")


def _try_get_active(prog_id: str):
    try:
        return win32com.client.GetActiveObject(prog_id)
    except Exception:
        return None


def _word_backups(word_app, backup_dir: str, stamp: str) -> list[SaveResult]:
    results: list[SaveResult] = []
    try:
        docs = word_app.Documents
        count = docs.Count
    except Exception:
        return results

    for i in range(1, count + 1):
        try:
            d = docs.Item(i)
            fullname = str(d.FullName or "")
            name = str(d.Name or "WordDocument")
            if not fullname:
                continue  # never saved -> would pop "Save As"
            if bool(getattr(d, "ReadOnly", False)):
                continue
            # Only back up when there are unsaved changes to reduce churn.
            if bool(getattr(d, "Saved", True)):
                continue

            target = _backup_path(backup_dir, fullname, stamp)
            d.SaveCopyAs(target)
            results.append(SaveResult(app="Word", name=name, source=fullname, backup=target))
        except Exception:
            continue
    return results


def _excel_backups(excel_app, backup_dir: str, stamp: str) -> list[SaveResult]:
    results: list[SaveResult] = []
    try:
        wbs = excel_app.Workbooks
        count = wbs.Count
    except Exception:
        return results

    for i in range(1, count + 1):
        try:
            wb = wbs.Item(i)
            fullname = str(wb.FullName or "")
            name = str(wb.Name or "ExcelWorkbook")
            if not fullname:
                continue
            if bool(getattr(wb, "ReadOnly", False)):
                continue
            if bool(getattr(wb, "Saved", True)):
                continue

            target = _backup_path(backup_dir, fullname, stamp)
            wb.SaveCopyAs(target)
            results.append(SaveResult(app="Excel", name=name, source=fullname, backup=target))
        except Exception:
            continue
    return results


def _ppt_backups(ppt_app, backup_dir: str, stamp: str) -> list[SaveResult]:
    results: list[SaveResult] = []
    try:
        pres = ppt_app.Presentations
        count = pres.Count
    except Exception:
        return results

    for i in range(1, count + 1):
        try:
            p = pres.Item(i)
            fullname = str(p.FullName or "")
            name = str(p.Name or "PowerPointPresentation")
            if not fullname:
                continue
            if bool(getattr(p, "ReadOnly", False)):
                continue

            # PowerPoint doesn't always expose a reliable "Saved" flag like Word/Excel.
            # Still try to reduce churn when possible.
            saved_flag = getattr(p, "Saved", None)
            if saved_flag is True:
                continue

            target = _backup_path(backup_dir, fullname, stamp)
            # SaveCopyAs is available for presentations; keep format (no conversion).
            p.SaveCopyAs(target)
            results.append(SaveResult(app="PowerPoint", name=name, source=fullname, backup=target))
        except Exception:
            continue
    return results


def run_loop(backup_dir: str, interval: int) -> int:
    print(f"[office_autosave] backup_dir={backup_dir}")
    print(f"[office_autosave] interval={interval}s (only backs up when file has unsaved changes)")
    print("[office_autosave] press Ctrl+C to stop")
    _ensure_dir(backup_dir)

    while True:
        stamp = _now_stamp()
        any_saved = False
        pythoncom.CoInitialize()
        try:
            word = _try_get_active("Word.Application")
            excel = _try_get_active("Excel.Application")
            ppt = _try_get_active("PowerPoint.Application")

            results: list[SaveResult] = []
            if word is not None:
                results.extend(_word_backups(word, backup_dir, stamp))
            if excel is not None:
                results.extend(_excel_backups(excel, backup_dir, stamp))
            if ppt is not None:
                results.extend(_ppt_backups(ppt, backup_dir, stamp))

            if results:
                any_saved = True
                for r in results:
                    print(f"[{stamp}] {r.app}: {r.name} -> {r.backup}")
            else:
                print(f"[{stamp}] no changes to back up")
        finally:
            pythoncom.CoUninitialize()

        # Keep the loop stable even if the work took time.
        sleep_s = max(1, int(interval))
        if any_saved:
            time.sleep(sleep_s)
        else:
            time.sleep(sleep_s)


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Auto-backup open Office documents (Word/Excel/PowerPoint) periodically.",
    )
    parser.add_argument(
        "--backup-dir",
        default=os.path.join(os.path.dirname(__file__), "backups"),
        help="Backup output directory (default: ./backups)",
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=30,
        help="Interval seconds (default: 30)",
    )
    args = parser.parse_args(argv)
    return run_loop(os.path.abspath(args.backup_dir), args.interval)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
