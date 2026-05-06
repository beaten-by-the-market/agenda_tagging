import argparse
import ctypes
import sys
import threading
import time
from pathlib import Path

import pythoncom
import win32com.client


BM_CLICK = 0x00F5

user32 = ctypes.windll.user32


def _window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    buffer = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buffer, length + 1)
    return buffer.value


def _class_name(hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buffer, len(buffer))
    return buffer.value


def _normalize_button_text(text: str) -> str:
    return text.replace("&", "").replace(" ", "").replace("\t", "").strip()


class HwpAccessPromptApprover:
    def __init__(self) -> None:
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, daemon=True)

    def start(self) -> None:
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        self._thread.join(timeout=1)

    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                self._scan_once()
            except Exception:
                pass
            time.sleep(0.25)

    def _scan_once(self) -> None:
        enum_windows_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def enum_children(parent_hwnd: int) -> None:
            def child_callback(child_hwnd: int, _lparam: int) -> bool:
                if _class_name(child_hwnd).lower() != "button":
                    return True

                text = _normalize_button_text(_window_text(child_hwnd))
                if text in {"모두접근허용", "모두허용", "전체허용"}:
                    user32.SendMessageW(child_hwnd, BM_CLICK, 0, 0)
                return True

            user32.EnumChildWindows(parent_hwnd, enum_windows_proc(child_callback), 0)

        def window_callback(hwnd: int, _lparam: int) -> bool:
            if user32.IsWindowVisible(hwnd):
                enum_children(hwnd)
            return True

        user32.EnumWindows(enum_windows_proc(window_callback), 0)


def iter_hwp_files(folder: Path, recurse: bool) -> list[Path]:
    pattern = "**/*.hwp" if recurse else "*.hwp"
    return sorted(path for path in folder.glob(pattern) if path.is_file() and path.suffix.lower() == ".hwp")


def convert_file(hwp, input_path: Path, overwrite: bool) -> None:
    output_path = input_path.with_suffix(".hwpx")
    if output_path.exists() and not overwrite:
        print(f"[SKIP] Exists: {output_path}")
        return

    print(f"[OPEN] {input_path}")
    opened = hwp.Open(str(input_path), "HWP", "forceopen:true")
    if not opened:
        raise RuntimeError(f"Hwp.Open failed: {input_path}")

    print(f"[SAVE] {output_path}")
    saved = hwp.SaveAs(str(output_path), "HWPX", "")
    if not saved:
        raise RuntimeError(f"Hwp.SaveAs HWPX failed: {output_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert .hwp files to .hwpx using installed Hancom Office.")
    parser.add_argument("folder", nargs="?", default=".", help="Folder containing .hwp files. Defaults to current folder.")
    parser.add_argument("-r", "--recurse", action="store_true", help="Convert files in subfolders too.")
    parser.add_argument("-o", "--overwrite", action="store_true", help="Overwrite existing .hwpx files.")
    parser.add_argument("--visible", action="store_true", help="Show the Hancom Office window while converting.")
    parser.add_argument("--no-auto-approve", action="store_true", help="Do not auto-click the access approval prompt.")
    args = parser.parse_args()

    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        print(f"Folder not found: {folder}", file=sys.stderr)
        return 1

    files = iter_hwp_files(folder, args.recurse)
    if not files:
        print(f"No .hwp files found: {folder}")
        return 0

    approver = None
    if not args.no_auto_approve:
        approver = HwpAccessPromptApprover()
        approver.start()

    pythoncom.CoInitialize()
    hwp = None
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

        try:
            hwp.XHwpWindows.Item(0).Visible = bool(args.visible)
        except Exception:
            pass

        try:
            registered = hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            print(f"[INFO] FilePathCheckerModule registered: {registered}")
        except Exception as exc:
            print(f"[INFO] FilePathCheckerModule registration failed: {exc}")

        for file_path in files:
            convert_file(hwp, file_path, args.overwrite)

        print(f"[DONE] Converted: {len(files)}")
        return 0
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass
        if approver is not None:
            approver.stop()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
