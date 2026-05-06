"""
HWP 파일 내 민감정보(회사명, 인명, 주소, 날짜) 익명화 + 재무수치 증감 스크립트.

사용법:
    python anonymize_hwp.py [입력.hwp] [-o 출력.hwp] [--overwrite] [--visible]

    입력 파일을 지정하지 않으면 현재 폴더의 모든 .hwp 파일을 처리합니다.
    출력 경로를 지정하지 않으면 원본 옆에 _익명 접미사를 붙여 저장합니다.
    --overwrite 를 주면 원본을 덮어씁니다.

처리 순서:
    1. HWP COM으로 텍스트 치환 (회사명/인명/주소/날짜)
    2. 임시 HWPX로 저장 → Python으로 재무수치 +FINANCIAL_INCREMENT
    3. 수정된 HWPX를 HWP COM으로 열어 최종 HWP로 저장
"""

import argparse
import ctypes
import io
import re
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path

import pythoncom
import win32com.client


BM_CLICK = 0x00F5
user32 = ctypes.windll.user32

# ─────────────────────────────────────────────
# 치환 목록 (여기를 수정하세요)
# ─────────────────────────────────────────────

# 1) 완전 일치 치환 (회사명, 인명, 주소 등)
LITERAL_REPLACEMENTS: list[tuple[str, str]] = [
    # 회사명
    ("비디아이", "대박제조"),
    ("팍스글로벌", "대박물류"),
    ("엘리슨", "대박바이오"),
    # 인명
    ("예경남", "홍길동"),
    ("안승만", "임꺽정"),
    ("김일강", "황진이"),
    # 감사법인명
    ("송강", "회계법인A"),
    ("성운", "회계법인B"),
    # 주소
    ("경기도 화성시 팔탄면 서해로 1155", "서울특별시 XX구 XXX로 123"),
]

# 2) HWP 와일드카드 치환 (날짜)
#    HWP 와일드카드:  # = 숫자 1자리,  ? = 임의 1글자,  * = 임의 여러 글자
#    날짜 앞 따옴표는 U+2018(‘)과 U+2019(‘) 두 종류가 혼용됩니다.
#    처리 순서가 중요합니다 — 긴 패턴을 먼저 처리해야 짧은 패턴이 남은 부분만 잡습니다.
WILDCARD_REPLACEMENTS: list[tuple[str, str]] = [
    # YYYY. MM. DD 형식  예) 2021. 01. 29
    ("####. ##. ##",      "20XX. XX. XX"),
    # ‘YY.MM.DD 형식  (U+2019 오른쪽 작은따옴표)
    ("’##.##.##",    "’XX.XX.XX"),
    # ‘YY.MM.DD 형식  (U+2018 왼쪽 작은따옴표)
    ("‘##.##.##",    "’XX.XX.XX"),
    # ‘YY.MM 형식  (U+2019)
    ("’##.##",       "’XX.XX"),
    # ‘YY.MM 형식  (U+2018)
    ("‘##.##",       "’XX.XX"),
    # ‘YY 단독 연도  (U+2019)  예) ‘17, ‘60, ‘62
    ("’##",          "’XX"),
    # ‘YY 단독 연도  (U+2018)
    ("‘##",          "’XX"),
]


# ─────────────────────────────────────────────
# HWP 접근 허용 팝업 자동 클릭 (convert 스크립트와 동일)
# ─────────────────────────────────────────────

def _window_text(hwnd: int) -> str:
    length = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, len(buf))
    return buf.value


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
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def enum_children(parent: int) -> None:
            def child_cb(child: int, _: int) -> bool:
                if _class_name(child).lower() == "button":
                    t = _window_text(child).replace("&", "").replace(" ", "").strip()
                    if t in {"모두접근허용", "모두허용", "전체허용"}:
                        user32.SendMessageW(child, BM_CLICK, 0, 0)
                return True
            user32.EnumChildWindows(parent, EnumWindowsProc(child_cb), 0)

        def win_cb(hwnd: int, _: int) -> bool:
            if user32.IsWindowVisible(hwnd):
                enum_children(hwnd)
            return True

        user32.EnumWindows(EnumWindowsProc(win_cb), 0)


# 재무수치에 더할 값 (0 이면 미적용)
FINANCIAL_INCREMENT: int = 10

_PURE_INT_RE = re.compile(r"^([△]?)(\d{1,3}(?:,\d{3})*)$")


# ─────────────────────────────────────────────
# 핵심 치환 함수
# ─────────────────────────────────────────────

def _find_replace(hwp, find: str, replace: str, *, use_wildcard: bool = False) -> None:
    hwp.HAction.GetDefault("RepFindReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = find
    hwp.HParameterSet.HFindReplace.ReplaceString = replace
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HParameterSet.HFindReplace.UseWildCard = 1 if use_wildcard else 0
    hwp.HParameterSet.HFindReplace.Direction = 0
    hwp.HParameterSet.HFindReplace.WholeWordOnly = 0
    hwp.HParameterSet.HFindReplace.MatchCase = 0
    hwp.HAction.Execute("RepFindReplace", hwp.HParameterSet.HFindReplace.HSet)


def _apply_financial_increment_to_hwpx(hwpx_path: Path) -> None:
    """HWPX XML 텍스트 노드에서 순수 정수만 골라 FINANCIAL_INCREMENT를 더한다."""
    if FINANCIAL_INCREMENT == 0:
        return

    def increment(text: str) -> str:
        m = _PURE_INT_RE.fullmatch(text.strip())
        if not m:
            return text
        sign, digits = m.group(1), m.group(2)
        val = int(digits.replace(",", ""))
        if sign == "△":
            val = -val
        new_val = val + FINANCIAL_INCREMENT
        if new_val < 0:
            return f"△{-new_val:,}"
        return f"{new_val:,}"

    buf = io.BytesIO()
    with zipfile.ZipFile(hwpx_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                xml_text = data.decode("utf-8")
                def replace_node(m: re.Match) -> str:
                    return m.group(1) + increment(m.group(2)) + m.group(3)
                xml_text = re.sub(
                    r"(<hp:t(?:\s[^>]*)?>)(.*?)(</hp:t>)",
                    replace_node,
                    xml_text,
                    flags=re.DOTALL,
                )
                data = xml_text.encode("utf-8")
            zout.writestr(item, data)
    hwpx_path.write_bytes(buf.getvalue())


def anonymize_file(hwp, input_path: Path, output_path: Path) -> None:
    print(f"[OPEN] {input_path}")
    if not hwp.Open(str(input_path), "HWP", "forceopen:true"):
        raise RuntimeError(f"열기 실패: {input_path}")

    # 1단계: 텍스트 치환 (회사명/인명/주소)
    for find, replace in LITERAL_REPLACEMENTS:
        _find_replace(hwp, find, replace)
        print(f"  [치환] {find!r:30s} → {replace!r}")

    # 2단계: 날짜 와일드카드 치환
    for find, replace in WILDCARD_REPLACEMENTS:
        _find_replace(hwp, find, replace, use_wildcard=True)
        print(f"  [날짜] {find!r:30s} → {replace!r}")

    if FINANCIAL_INCREMENT != 0:
        # 3단계: 임시 HWPX로 저장 → 재무수치 증감 → HWPX 재로드 → HWP 저장
        with tempfile.NamedTemporaryFile(suffix=".hwpx", delete=False) as tf:
            temp_hwpx = Path(tf.name)
        try:
            print(f"  [재무] 임시 HWPX 저장 중...")
            if not hwp.SaveAs(str(temp_hwpx), "HWPX", ""):
                raise RuntimeError("임시 HWPX 저장 실패")
            _apply_financial_increment_to_hwpx(temp_hwpx)
            print(f"  [재무] +{FINANCIAL_INCREMENT} 적용 완료")
            if not hwp.Open(str(temp_hwpx), "HWPX", "forceopen:true"):
                raise RuntimeError("임시 HWPX 열기 실패")
        finally:
            temp_hwpx.unlink(missing_ok=True)

    print(f"[SAVE] {output_path}")
    if not hwp.SaveAs(str(output_path), "HWP", ""):
        raise RuntimeError(f"저장 실패: {output_path}")
    print(f"[DONE] {output_path.name}")


# ─────────────────────────────────────────────
# CLI 진입점
# ─────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="HWP 파일 민감정보 익명화")
    parser.add_argument("input", nargs="?", help="처리할 .hwp 파일 또는 폴더 (기본: 현재 폴더)")
    parser.add_argument("-o", "--output", help="출력 파일 경로 (단일 파일 처리 시)")
    parser.add_argument("--overwrite", action="store_true", help="원본 파일 덮어쓰기")
    parser.add_argument("-r", "--recurse", action="store_true", help="하위 폴더 포함")
    parser.add_argument("--visible", action="store_true", help="HWP 창 표시")
    args = parser.parse_args()

    # 처리할 파일 목록 결정
    target = Path(args.input).resolve() if args.input else Path(".").resolve()
    if target.is_file():
        files = [target]
    elif target.is_dir():
        pattern = "**/*.hwp" if args.recurse else "*.hwp"
        files = sorted(p for p in target.glob(pattern) if p.suffix.lower() == ".hwp")
    else:
        print(f"경로를 찾을 수 없습니다: {target}", file=sys.stderr)
        return 1

    if not files:
        print("처리할 .hwp 파일이 없습니다.")
        return 0

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
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass

        for file_path in files:
            if args.overwrite:
                out_path = file_path
            elif args.output and len(files) == 1:
                out_path = Path(args.output).resolve()
            else:
                out_path = file_path.with_stem(file_path.stem + "_익명")

            try:
                anonymize_file(hwp, file_path, out_path)
            except Exception as e:
                print(f"[ERROR] {file_path.name}: {e}", file=sys.stderr)

        return 0
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass
        approver.stop()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
