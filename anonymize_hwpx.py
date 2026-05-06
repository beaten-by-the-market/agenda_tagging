"""
HWPX 파일 내 민감정보(회사명, 인명, 주소, 날짜) 익명화 스크립트.
한컴 오피스 없이 순수 Python으로 동작합니다.

사용법:
    python anonymize_hwpx.py [입력.hwpx] [-o 출력.hwpx] [--overwrite] [-r]

    입력 파일을 지정하지 않으면 현재 폴더의 모든 .hwpx 파일을 처리합니다.
    출력 경로를 지정하지 않으면 원본 옆에 _익명 접미사를 붙여 저장합니다.
    --overwrite 를 주면 원본을 덮어씁니다.

결과물을 HWP로도 저장하려면:
    anonymize_hwp.py (한컴 오피스 설치 필요) 또는
    생성된 _익명.hwpx 를 HWP에서 열어 다른 이름으로 저장 → HWP 선택
"""

import argparse
import io
import re
import shutil
import sys
import zipfile
from pathlib import Path


# ─────────────────────────────────────────────
# 치환 목록 (여기를 수정하세요)
# ─────────────────────────────────────────────

# 날짜 앞에 오는 따옴표 패턴:
#   U+0027 ‘ ASCII 아포스트로피  (재무 테이블 헤더 등)
#   U+2018 ‘ 왼쪽 작은따옴표
#   U+2019 ‘ 오른쪽 작은따옴표
_Q = "[\u0027\u2018\u2019]"  # ASCII apostrophe / LEFT / RIGHT single quote

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

# 2) 정규식 치환 (날짜)
#    패턴 길이가 긴 것(구체적인 것)을 먼저 처리해야 짧은 패턴이 남은 부분만 잡습니다.
#    날짜 앞 따옴표는 U+2018(‘)과 U+2019(‘) 두 종류가 혼용되므로 _Q 패턴으로 통합합니다.
DATE_REGEX_REPLACEMENTS: list[tuple[str, str]] = [
    # YYYY. MM. DD 형식  예) 2021. 01. 29  →  20XX. XX. XX
    (r"[12]\d{3}\.\s*\d{2}\.\s*\d{2}", "20XX. XX. XX"),
    # ‘YY.MM.DD 형식  예) ‘20.06.26  →  ‘XX.XX.XX
    (_Q + r"\d{2}\.\d{2}\.\d{2}", "’XX.XX.XX"),
    # ‘YY.MM 형식  예) ‘20.3Q  →  ‘XX.XQ  (월 숫자만 마스킹)
    (_Q + r"(\d{2})\.(\d+)", lambda m: f"’XX.{m.group(2)}"),
    # ‘YY 단독 연도  예) ‘17, ‘60  →  ‘XX
    (_Q + r"\d{2}(?=\D|$)", "’XX"),
]


# 재무수치에 더할 값 (0 이면 미적용)
FINANCIAL_INCREMENT: int = 10


# ─────────────────────────────────────────────
# 핵심 치환 함수
# ─────────────────────────────────────────────

_PURE_INT_RE = re.compile(r"^([△]?)(\d{1,3}(?:,\d{3})*)$")


def _increment_financial(text: str) -> str:
    """순수 정수 텍스트 노드(△포함)에 FINANCIAL_INCREMENT를 더한다."""
    if FINANCIAL_INCREMENT == 0:
        return text
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


def _apply_replacements(text: str) -> str:
    # 완전 일치 치환
    for find, replace in LITERAL_REPLACEMENTS:
        text = text.replace(find, replace)

    # 날짜 정규식 치환 (길고 구체적인 패턴 먼저)
    for pattern, replacement in DATE_REGEX_REPLACEMENTS:
        text = re.sub(pattern, replacement, text)

    # 재무수치 증감
    text = _increment_financial(text)

    return text


def anonymize_file(input_path: Path, output_path: Path) -> None:
    print(f"[OPEN] {input_path}")

    # 원본 HWPX를 메모리로 읽어 새 ZIP 생성
    buf = io.BytesIO()
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)

            # XML 파일의 텍스트 노드만 치환
            if item.filename.startswith("Contents/") and item.filename.endswith(".xml"):
                xml_text = data.decode("utf-8")
                # hp:t 텍스트 노드 내용만 선택적으로 치환
                def replace_text_node(m: re.Match) -> str:
                    tag_open = m.group(1)
                    content  = m.group(2)
                    tag_close = m.group(3)
                    return tag_open + _apply_replacements(content) + tag_close

                xml_text = re.sub(
                    r"(<hp:t(?:\s[^>]*)?>)(.*?)(</hp:t>)",
                    replace_text_node,
                    xml_text,
                    flags=re.DOTALL,
                )
                data = xml_text.encode("utf-8")

            # Preview 텍스트도 치환
            elif item.filename == "Preview/PrvText.txt":
                text = data.decode("utf-8")
                text = _apply_replacements(text)
                # PrvText.txt는 <숫자> 형태 셀만 치환 (예: <1,480>, <△81>)
                # 넓은 패턴을 쓰면 제3호, 주1) 같은 것도 변경되므로 꺽쇠 내부만 타겟
                if FINANCIAL_INCREMENT != 0:
                    def _inc_in_brackets(m: re.Match) -> str:
                        sign = m.group(1)
                        val = int(m.group(2).replace(",", ""))
                        if sign == "△":
                            val = -val
                        new_val = val + FINANCIAL_INCREMENT
                        if new_val < 0:
                            return f"△{-new_val:,}"
                        return f"{new_val:,}"
                    text = re.sub(
                        r"(?<=<)(△?)(\d{1,3}(?:,\d{3})*)(?=>)",
                        _inc_in_brackets,
                        text,
                    )
                data = text.encode("utf-8")

            zout.writestr(item, data)

    print(f"[SAVE] {output_path}")
    output_path.write_bytes(buf.getvalue())
    print(f"[DONE] {output_path.name}")


# ─────────────────────────────────────────────
# CLI 진입점
# ─────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="HWPX 파일 민감정보 익명화 (한컴 오피스 불필요)")
    parser.add_argument("input", nargs="?", help="처리할 .hwpx 파일 또는 폴더 (기본: 현재 폴더)")
    parser.add_argument("-o", "--output", help="출력 파일 경로 (단일 파일 처리 시)")
    parser.add_argument("--overwrite", action="store_true", help="원본 파일 덮어쓰기")
    parser.add_argument("-r", "--recurse", action="store_true", help="하위 폴더 포함")
    args = parser.parse_args()

    target = Path(args.input).resolve() if args.input else Path(".").resolve()
    if target.is_file():
        files = [target]
    elif target.is_dir():
        pattern = "**/*.hwpx" if args.recurse else "*.hwpx"
        files = sorted(p for p in target.glob(pattern) if p.suffix.lower() == ".hwpx")
    else:
        print(f"경로를 찾을 수 없습니다: {target}", file=sys.stderr)
        return 1

    if not files:
        print("처리할 .hwpx 파일이 없습니다.")
        return 0

    for file_path in files:
        if args.overwrite:
            out_path = file_path
        elif args.output and len(files) == 1:
            out_path = Path(args.output).resolve()
        else:
            out_path = file_path.with_stem(file_path.stem + "_익명")

        try:
            anonymize_file(file_path, out_path)
        except Exception as e:
            print(f"[ERROR] {file_path.name}: {e}", file=sys.stderr)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
