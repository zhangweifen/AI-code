# -*- coding: utf-8 -*-
"""
质检加班日期比对。

把 A 表（质检汇总）里「一个单元格多个加班日期」拆开，统一成 YYYY-MM-DD，
再按「姓名 + 加班日期」到 B 表（加班邮件数据）里查找。

判定：B 表只要出现过该人员和该日期，这一天即加班成立。
不看 B 表「是否同意加班」。

会生成两个结果文件：
- 独立结果：按原行汇总、比对明细、说明、B表提示。
  「加班是否成立」取「是 / 否 / 空」。日期为空时不比对，这一格留空。
- 原表追加：复制质检表，并在右侧追加「加班是否成立」「缺失的日期」「不成立的原因」。
  取值是「成立 / 不成立 / 空」。

用法：
    py -3 compare_overtime.py
    质检表和结果位置运行时输入。
    加班邮件数据表直接回车，使用当前目录下的 ISV自检\\加班邮件数据表.xlsx。

    py -3 compare_overtime.py --a "质检表.xlsx" --b "加班邮件数据表.xlsx" --output "结果文件夹"
    结果文件名为「质检表文件名_加班对比结果.xlsx」。
    py -3 compare_overtime.py --self-test
"""

from __future__ import annotations

import argparse
import math
import re
import sys
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MAIL_TABLE_NAME = "加班邮件数据表.xlsx"
DEFAULT_MAIL_RELATIVE = Path("ISV自检") / MAIL_TABLE_NAME
RESULT_NAME_SUFFIX = "加班对比结果"
APPEND_HEADERS = ["加班是否成立", "缺失的日期", "不成立的原因"]

# 区间最多展开的天数，避免脏数据把整年都铺开。
MAX_RANGE_DAYS = 62

A_ALIASES = {
    "name": ["姓名", "员工姓名", "人员姓名", "名字", "加班人员名称", "加班人员"],
    "phone": ["手机号码", "手机号", "手机", "联系电话", "移动电话", "电话号码"],
    "date": ["加班日期", "加班时间"],
}
B_ALIASES = {
    "name": ["加班人员名称", "姓名", "员工姓名", "人员姓名", "加班人员", "名字"],
    "date": ["加班日期", "加班时间"],
}

SUMMARY_COLUMNS = [
    "A表行号",
    "姓名",
    "手机号码",
    "加班日期",
    "加班是否成立",
    "加班缺失日期",
    "解析说明",
]
DETAIL_COLUMNS = [
    "A表行号",
    "姓名",
    "手机号码",
    "加班日期",
    "加班是否成立",
    "加班缺失日期",
    "原始加班日期",
    "解析说明",
]
JUDGEMENT_COLUMNS = ["A表行号", "追加状态", "缺失的日期", "不成立的原因"]

_FULLWIDTH_DIGITS = str.maketrans("０１２３４５６７８９", "0123456789")
_PUNCT_MAP = str.maketrans(
    {
        "／": "/",
        "．": ".",
        "－": "-",
        "﹣": "-",
        "—": "-",
        "–": "-",
        "～": "~",
        "，": ",",
        "；": ";",
        "（": "(",
        "）": ")",
        "　": " ",
    }
)

_MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
_MONTH_PATTERN = (
    r"Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|"
    r"Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|"
    r"Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?"
)

# 各类日期写法。后面会按位置合并，重叠时保留更长的那段。
_PATTERNS = [
    re.compile(
        r"(?:(?P<cy>\d{4})\s*年\s*)?(?P<cm>\d{1,2})\s*月\s*"
        r"(?P<cdays>\d{1,2}(?:\s*[、,]\s*\d{1,2})+)\s*[日号]"
    ),
    re.compile(
        r"(?P<fy>\d{4})\s*[年/.\-]\s*(?P<fm>\d{1,2})\s*[月/.\-]\s*(?P<fd>\d{1,2})\s*[日号]?"
    ),
    re.compile(r"(?P<sy>\d{2})\s*年\s*(?P<sm>\d{1,2})\s*月\s*(?P<sd>\d{1,2})\s*[日号]?"),
    re.compile(r"(?<!\d)(?P<y8>\d{4})(?P<m8>\d{2})(?P<d8>\d{2})(?!\d)"),
    re.compile(r"(?P<mm>\d{1,2})\s*[月/.\-]\s*(?P<md>\d{1,2})\s*[日号]?"),
    re.compile(r"(?<![\d月/.\-])(?P<od>\d{1,2})\s*[日号]"),
    re.compile(
        rf"(?P<emn>{_MONTH_PATTERN})\.?\s+(?P<ed>\d{{1,2}})(?:st|nd|rd|th)?\,?\s+(?P<ey>\d{{4}})",
        re.IGNORECASE,
    ),
    re.compile(
        rf"(?P<ed2>\d{{1,2}})(?:st|nd|rd|th)?\s+(?P<emn2>{_MONTH_PATTERN})\.?\s+(?P<ey2>\d{{4}})",
        re.IGNORECASE,
    ),
]
_RANGE_ONLY = re.compile(r"\s*(?:至|到|~|-)\s*")
_RANGE_TO_DAY = re.compile(r"\s*(?:至|到|~|-)\s*(\d{1,2})\s*[日号]?\s*[、,;]?\s*")
_TIME_TEXT = re.compile(r"\b\d{1,2}:\d{2}(?::\d{2})?\b")
# 时长说明不是日期。前面不能是日期分隔符，避免把“10/29天”里的日吃掉。
_DURATION_NOTES = (
    re.compile(r"(?<![/\-.年\d])\d+(?:\.\d+)?\s*个\s*(?:小时|钟头)"),
    re.compile(r"(?<![/\-.年\d])\d+(?:\.\d+)?\s*(?:小时|钟头)"),
    re.compile(r"(?<![/\-.年\d])\d+(?:\.\d+)?\s*天"),
    re.compile(r"加班"),
)


@dataclass
class DatePiece:
    """一个已解析或待补年份的日期。"""

    year: int | None
    month: int
    day: int
    raw: str
    notes: list[str] = field(default_factory=list)

    def iso_or_none(self) -> str | None:
        if self.year is None:
            return None
        try:
            return date(self.year, self.month, self.day).isoformat()
        except ValueError:
            return None


@dataclass
class CellParse:
    pieces: list[DatePiece]
    failures: list[str]
    notes: list[str]


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8")
            except Exception:
                pass


def normalize_text(value: object) -> str:
    """把全角数字和常见中文标点收成半角，方便同一套规则解析。"""
    text = "" if value is None else str(value)
    text = text.replace("\u3000", " ").translate(_FULLWIDTH_DIGITS).translate(_PUNCT_MAP)
    text = text.replace("、", ",")
    return text.strip()


def name_key(value: object) -> str:
    """比对用姓名：去空白、去括号备注，英文忽略大小写。"""
    text = normalize_text(value)
    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"\s+", "", text)
    return text.casefold()


def display_name(value: object) -> str:
    text = normalize_text(value)
    return re.sub(r"\s+", " ", text).strip()


def format_phone(value: object) -> str:
    """手机号按文本保留，去掉 Excel 数值带出的 .0。"""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    if isinstance(value, int) and not isinstance(value, bool):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return format(value, "f").rstrip("0").rstrip(".")
    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    if re.fullmatch(r"\d+(?:\.\d+)?[eE][+\-]?\d+", text):
        try:
            number = float(text)
        except ValueError:
            return text
        if number.is_integer():
            return str(int(number))
    return text


def _blank(value: object) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    return normalize_text(value) == ""


def _header_text(value: object) -> str:
    return re.sub(r"\s+", "", normalize_text(value))


def _match_headers(cells: list[object], aliases: dict[str, list[str]]) -> dict[str, int]:
    normalized = [(index, _header_text(cell)) for index, cell in enumerate(cells)]
    mapping: dict[str, int] = {}
    for field_name, names in aliases.items():
        for alias in sorted(names, key=len, reverse=True):
            target = _header_text(alias)
            for index, text in normalized:
                if text == target:
                    mapping[field_name] = index
                    break
            if field_name in mapping:
                break
    return mapping


def _sheet_score(mapping: dict[str, int], required: tuple[str, ...]) -> int:
    if any(name not in mapping for name in required):
        return -1
    return len(mapping)


def load_sheet_frame(path: Path, aliases: dict[str, list[str]], required: tuple[str, ...]) -> tuple[pd.DataFrame, dict[str, int], str, int]:
    """读取最像目标表的工作表，保留原始类型，并带上 Excel 行号。"""
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    except FileNotFoundError:
        raise SystemExit(f"找不到文件：{path}")
    except PermissionError:
        raise SystemExit(f"文件被占用，请先关闭 Excel 后重试：{path}")
    except Exception as exc:
        raise SystemExit(f"读取 Excel 失败：{path}\n{exc}") from exc

    best: tuple[int, str, int, dict[str, int]] | None = None
    for sheet_name, frame in sheets.items():
        if frame is None or frame.empty:
            continue
        limit = min(15, len(frame))
        for row_index in range(limit):
            cells = frame.iloc[row_index].tolist()
            mapping = _match_headers(cells, aliases)
            score = _sheet_score(mapping, required)
            if score < 0:
                continue
            if best is None or score > best[0]:
                best = (score, sheet_name, row_index, mapping)
    if best is None:
        names = "、".join(f"{item}({'/'.join(aliases[item])})" for item in required)
        raise SystemExit(f"{path.name} 里没有找到必需列：{names}")

    _, sheet_name, header_index, mapping = best
    frame = sheets[sheet_name].iloc[header_index + 1 :].copy()
    frame["_excel_row"] = frame.index + 1
    return frame, mapping, sheet_name, header_index + 1


def _excel_serial_to_date(number: float) -> date | None:
    if isinstance(number, bool) or math.isnan(number):
        return None
    # 约 1995-09 到 2064-04，覆盖加班日期，避开普通数量和手机号。
    if not 35000 <= number <= 60000:
        return None
    whole = int(number)
    try:
        return date(1899, 12, 30) + timedelta(days=whole)
    except OverflowError:
        return None


def _safe_piece(year: int | None, month: int, day: int, raw: str, notes: list[str] | None = None) -> DatePiece | None:
    # month 为 0 表示单元格里只有“几日”，月份留给后面沿用。
    if month != 0 and not 1 <= month <= 12:
        return None
    if not 1 <= day <= 31:
        return None
    if year is not None and not 1990 <= year <= 2100:
        return None
    return DatePiece(year=year, month=month, day=day, raw=raw, notes=list(notes or []))


def _pieces_from_match(match: re.Match[str]) -> list[DatePiece]:
    raw = match.group(0)
    groups = match.groupdict()
    if groups.get("cdays"):
        year = int(groups["cy"]) if groups.get("cy") else None
        month = int(groups["cm"])
        pieces = []
        for day_text in re.split(r"\s*[、,]\s*", groups["cdays"]):
            piece = _safe_piece(year, month, int(day_text), raw)
            if piece is not None:
                pieces.append(piece)
        return pieces
    if groups.get("fy"):
        piece = _safe_piece(int(groups["fy"]), int(groups["fm"]), int(groups["fd"]), raw)
        return [piece] if piece else []
    if groups.get("sy"):
        piece = _safe_piece(2000 + int(groups["sy"]), int(groups["sm"]), int(groups["sd"]), raw)
        return [piece] if piece else []
    if groups.get("y8"):
        piece = _safe_piece(int(groups["y8"]), int(groups["m8"]), int(groups["d8"]), raw)
        return [piece] if piece else []
    if groups.get("mm"):
        piece = _safe_piece(None, int(groups["mm"]), int(groups["md"]), raw)
        return [piece] if piece else []
    if groups.get("od"):
        piece = _safe_piece(None, 0, int(groups["od"]), raw)
        return [piece] if piece else []
    if groups.get("emn"):
        month = _MONTHS[groups["emn"].lower()]
        piece = _safe_piece(int(groups["ey"]), month, int(groups["ed"]), raw)
        return [piece] if piece else []
    if groups.get("emn2"):
        month = _MONTHS[groups["emn2"].lower()]
        piece = _safe_piece(int(groups["ey2"]), month, int(groups["ed2"]), raw)
        return [piece] if piece else []
    return []


@dataclass
class _Span:
    start: int
    end: int
    pieces: list[DatePiece]


def _collect_spans(text: str) -> list[_Span]:
    spans: list[_Span] = []
    for pattern in _PATTERNS:
        for match in pattern.finditer(text):
            pieces = _pieces_from_match(match)
            if pieces:
                spans.append(_Span(match.start(), match.end(), pieces))
    spans.sort(key=lambda item: (item.start, -(item.end - item.start)))
    chosen: list[_Span] = []
    cursor = -1
    for span in spans:
        if span.start < cursor:
            continue
        chosen.append(span)
        cursor = span.end
    return chosen


def _inherit_calendar(pieces: list[DatePiece]) -> None:
    """没有写年、月时，沿用同一单元格里前一个明确日期。"""
    last_year: int | None = None
    last_month: int | None = None
    for piece in pieces:
        if piece.month == 0:
            if last_month is None:
                piece.notes.append("只有“日”，前面没有月份，无法确定日期")
                continue
            piece.month = last_month
            piece.notes.append("未写月份，已沿用同一单元格中的前一个月份")
        if piece.year is None and last_year is not None:
            piece.year = last_year
            piece.notes.append("未写年份，已沿用同一单元格中的前一个年份")
        if piece.year is not None:
            last_year = piece.year
        if piece.month:
            last_month = piece.month


def _expand_range(start: DatePiece, end: DatePiece) -> list[DatePiece]:
    """把“2日至5日”展开成每一天。两边都没写年份时，展开结果也先不填年份。"""
    known_year = start.year or end.year
    synthetic_year = known_year or 2000
    try:
        start_date = date(start.year or synthetic_year, start.month, start.day)
        end_date = date(end.year or synthetic_year, end.month, end.day)
    except ValueError:
        start.notes.append("区间中有不存在的日期，已按两个独立日期处理")
        end.notes.append("区间中有不存在的日期，已按两个独立日期处理")
        return [start, end]
    if end_date < start_date:
        start.notes.append("区间结束早于开始，已按两个独立日期处理")
        end.notes.append("区间结束早于开始，已按两个独立日期处理")
        return [start, end]
    if (end_date - start_date).days > MAX_RANGE_DAYS:
        start.notes.append(f"日期间隔超过{MAX_RANGE_DAYS}天，未自动展开")
        end.notes.append(f"日期间隔超过{MAX_RANGE_DAYS}天，未自动展开")
        return [start, end]

    expanded: list[DatePiece] = []
    current = start_date
    while current <= end_date:
        if current == start_date:
            notes = list(start.notes)
        elif current == end_date:
            notes = list(end.notes)
        else:
            notes = ["由日期区间自动展开"]
        expanded.append(
            DatePiece(
                year=current.year if known_year is not None else None,
                month=current.month,
                day=current.day,
                raw=f"{start.raw}至{end.raw}",
                notes=notes,
            )
        )
        current += timedelta(days=1)
    return expanded


def _range_end_from_day(start: DatePiece, day_text: str, raw: str) -> DatePiece:
    return DatePiece(year=start.year, month=start.month, day=int(day_text), raw=raw.strip())


def _apply_ranges(text: str, spans: list[_Span]) -> tuple[list[DatePiece], list[tuple[int, int]]]:
    """识别“至 / 到 / ~ / -”区间。返回日期，以及被短区间吃掉的原文位置。"""
    pieces: list[DatePiece] = []
    consumed: list[tuple[int, int]] = []
    index = 0
    while index < len(spans):
        current = spans[index].pieces
        if index + 1 < len(spans):
            gap = text[spans[index].end : spans[index + 1].start]
            nxt = spans[index + 1].pieces
            if current and nxt and _RANGE_ONLY.fullmatch(gap):
                pieces.extend(current[:-1])
                pieces.extend(_expand_range(current[-1], nxt[0]))
                pieces.extend(nxt[1:])
                index += 2
                continue
            day_match = _RANGE_TO_DAY.fullmatch(gap)
            if day_match and current and current[-1].month:
                end = _range_end_from_day(current[-1], day_match.group(1), day_match.group(0))
                pieces.extend(current[:-1])
                pieces.extend(_expand_range(current[-1], end))
                index += 1
                continue
        else:
            tail = text[spans[index].end :]
            day_match = _RANGE_TO_DAY.fullmatch(tail)
            if day_match and current and current[-1].month:
                end = _range_end_from_day(current[-1], day_match.group(1), day_match.group(0))
                pieces.extend(current[:-1])
                pieces.extend(_expand_range(current[-1], end))
                consumed.append((spans[index].end, len(text)))
                index += 1
                continue
        pieces.extend(current)
        index += 1
    return pieces, consumed


def show_date(iso: str) -> str:
    """对外展示和写回原表时使用 YYYY/MM/DD。"""
    year, month, day = iso.split("-")
    return f"{int(year):04d}/{int(month):02d}/{int(day):02d}"


def _remove_duration_notes(text: str) -> str:
    """去掉“加班5个小时，0.625天”这类说明，只留日期。"""
    for pattern in _DURATION_NOTES:
        text = pattern.sub(" ", text)
    return text


def _unparsed_residue(residue: str) -> str:
    """时长、括号和标点不算解析失败。剩下的内容仍像日期时才返回。"""
    cleaned = _remove_duration_notes(residue)
    cleaned = re.sub(r"[\(（][^)）]*[\)）]", " ", cleaned)
    cleaned = re.sub(r"[\s,;|\\至到~\-.:：]+", " ", cleaned).strip()
    if not cleaned or not re.search(r"\d", cleaned):
        return ""
    if re.search(r"\d{4}", cleaned) or re.search(r"\d+\s*[月年/日号]", cleaned):
        return cleaned
    return ""


def parse_date_cell(value: object) -> CellParse:
    """把一个加班日期单元格拆成若干标准日期。解析不了的片段记入 failures，不中断整表。"""
    if _blank(value):
        return CellParse(pieces=[], failures=[], notes=["加班日期为空"])

    if isinstance(value, datetime):
        return CellParse(
            pieces=[DatePiece(value.year, value.month, value.day, value.strftime("%Y-%m-%d"))],
            failures=[],
            notes=[],
        )
    if isinstance(value, date):
        return CellParse(
            pieces=[DatePiece(value.year, value.month, value.day, value.isoformat())],
            failures=[],
            notes=[],
        )
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = _excel_serial_to_date(float(value))
        if converted is not None:
            return CellParse(
                pieces=[DatePiece(converted.year, converted.month, converted.day, str(value), ["Excel日期序列号已转换"])],
                failures=[],
                notes=[],
            )

    text = normalize_text(value)
    text = _TIME_TEXT.sub(" ", text)
    text = _remove_duration_notes(text)
    spans = _collect_spans(text)
    _inherit_calendar([piece for span in spans for piece in span.pieces])
    pieces, extra_spans = _apply_ranges(text, spans)

    covered = [(span.start, span.end) for span in spans]
    covered.extend(extra_spans)
    residue = text
    for start, end in sorted(covered, reverse=True):
        residue = residue[:start] + " " + residue[end:]
    residue = _RANGE_ONLY.sub(" ", residue)
    residue = re.sub(r"[\s,;|/\\至到~.-]+", " ", residue).strip()
    failures: list[str] = []
    notes: list[str] = []
    unparsed = _unparsed_residue(residue)
    if unparsed:
        failures.append(unparsed)
        notes.append(f"有日期片段无法解析：{unparsed}")

    valid: list[DatePiece] = []
    for piece in pieces:
        if piece.month == 0:
            failures.append(piece.raw)
            notes.append("、".join(piece.notes) or "无法确定月份")
            continue
        if piece.year is not None:
            try:
                date(piece.year, piece.month, piece.day)
            except ValueError:
                failures.append(piece.raw)
                notes.append(f"日期不存在：{piece.raw}")
                continue
        valid.append(piece)
    if text and not valid and not failures:
        failures.append(text)
        notes.append("没有识别到加班日期，请人工核对")
    return CellParse(pieces=valid, failures=failures, notes=notes)


def original_date_text(value: object) -> str:
    if _blank(value):
        return ""
    if isinstance(value, datetime):
        return show_date(value.strftime("%Y-%m-%d"))
    if isinstance(value, date):
        return show_date(value.isoformat())
    if isinstance(value, float) and value.is_integer():
        converted = _excel_serial_to_date(value)
        if converted is not None:
            return show_date(converted.isoformat())
    return normalize_text(value)


def _dedupe_pieces(pieces: list[DatePiece]) -> list[DatePiece]:
    seen: set[tuple[int | None, int, int]] = set()
    result: list[DatePiece] = []
    for piece in pieces:
        key = (piece.year, piece.month, piece.day)
        if key in seen:
            continue
        seen.add(key)
        result.append(piece)
    return result


def build_b_index(frame: pd.DataFrame, mapping: dict[str, int]) -> tuple[set[tuple[str, str]], dict[str, list[int]], list[str]]:
    """B 表索引：(姓名键, YYYY-MM-DD)。同意与否都不影响是否入索引。"""
    index: set[tuple[str, str]] = set()
    years_by_name: dict[str, list[int]] = {}
    warnings: list[str] = []
    name_col = mapping["name"]
    date_col = mapping["date"]
    for _, row in frame.iterrows():
        excel_row = int(row["_excel_row"])
        raw_name = row.iloc[name_col] if name_col < len(row) else None
        raw_date = row.iloc[date_col] if date_col < len(row) else None
        person = name_key(raw_name)
        if not person or _blank(raw_date):
            continue
        parsed = parse_date_cell(raw_date)
        if parsed.failures or parsed.notes:
            warnings.append(f"B表第{excel_row}行：{'；'.join(parsed.notes + parsed.failures)}")
        for piece in _dedupe_pieces(parsed.pieces):
            if piece.year is None:
                warnings.append(f"B表第{excel_row}行的日期没有年份，已跳过：{piece.raw}")
                continue
            iso = piece.iso_or_none()
            if iso is None:
                warnings.append(f"B表第{excel_row}行的日期不存在，已跳过：{piece.raw}")
                continue
            index.add((person, iso))
            years_by_name.setdefault(person, []).append(piece.year)
    return index, years_by_name, warnings


def _infer_year(person: str, month: int, day: int, years_by_name: dict[str, list[int]], b_index: set[tuple[str, str]]) -> tuple[int | None, str]:
    years = years_by_name.get(person) or [int(iso[:4]) for _person, iso in b_index]
    if not years:
        return None, "单元格和B表都没有可参考的年份"
    unique_years = sorted(set(years))
    matched = []
    for year in unique_years:
        try:
            iso = date(year, month, day).isoformat()
        except ValueError:
            continue
        if (person, iso) in b_index:
            matched.append(year)
    if len(matched) == 1:
        return matched[0], f"未写年份，已按B表匹配到{matched[0]}年"
    if len(matched) > 1:
        chosen = max(matched)
        return chosen, f"未写年份，B表有多个年份匹配，已采用{chosen}年"
    chosen = max(unique_years)
    return chosen, f"未写年份，B表没有同月日记录，已暂按{chosen}年判断"


def _blank_judgement(excel_row: int) -> dict[str, object]:
    return {"A表行号": excel_row, "追加状态": "", "缺失的日期": "", "不成立的原因": ""}


def _reason_text(person: str, missing_dates: list[str], failures: list[str], notes: list[str]) -> str:
    reasons: list[str] = []
    if not person:
        reasons.append("姓名为空，无法按人员与加班邮件数据表比对")
    if missing_dates:
        reasons.append("加班邮件数据表中找不到这些日期：" + "、".join(missing_dates))
    if failures:
        reasons.append("无法识别为日期：" + "、".join(failures))
    if not reasons:
        extra = "；".join(dict.fromkeys(note for note in notes if note and note != "加班日期为空"))
        if extra:
            reasons.append(extra)
    return "；".join(reasons)


def compare(
    a_frame: pd.DataFrame,
    a_mapping: dict[str, int],
    b_index: set[tuple[str, str]],
    years_by_name: dict[str, list[int]],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    summary_rows: list[dict[str, object]] = []
    detail_rows: list[dict[str, object]] = []
    judgement_rows: list[dict[str, object]] = []
    name_col = a_mapping["name"]
    date_col = a_mapping["date"]
    phone_col = a_mapping.get("phone")

    for _, row in a_frame.iterrows():
        excel_row = int(row["_excel_row"])
        raw_name = row.iloc[name_col] if name_col < len(row) else None
        raw_phone = row.iloc[phone_col] if phone_col is not None and phone_col < len(row) else None
        raw_date = row.iloc[date_col] if date_col < len(row) else None
        if _blank(raw_name) and _blank(raw_phone) and _blank(raw_date):
            continue

        person_text = display_name(raw_name)
        person = name_key(raw_name)
        phone = format_phone(raw_phone)
        original = original_date_text(raw_date)

        if _blank(raw_date):
            summary_rows.append(
                {
                    "A表行号": excel_row,
                    "姓名": person_text,
                    "手机号码": phone,
                    "加班日期": "",
                    "加班是否成立": "",
                    "加班缺失日期": "",
                    "解析说明": "",
                }
            )
            judgement_rows.append(_blank_judgement(excel_row))
            continue

        try:
            parsed = parse_date_cell(raw_date)
            pieces = _dedupe_pieces(parsed.pieces)
            resolved: list[tuple[str, str, list[str]]] = []
            cell_notes = [note for note in parsed.notes if note != "加班日期为空"]
            for piece in pieces:
                notes = list(piece.notes)
                year = piece.year
                if year is None:
                    year, infer_note = _infer_year(person, piece.month, piece.day, years_by_name, b_index)
                    notes.append(infer_note)
                if year is None:
                    resolved.append((piece.raw, "", notes + ["无法确定年份"]))
                    continue
                try:
                    iso = date(year, piece.month, piece.day).isoformat()
                except ValueError:
                    resolved.append((piece.raw, "", notes + [f"日期不存在：{year}/{piece.month:02d}/{piece.day:02d}"]))
                    continue
                resolved.append((show_date(iso), iso, notes))

            missing_dates: list[str] = []
            missing_for_file: list[str] = []
            established_flags: list[bool] = []
            for display, iso, notes in resolved:
                note_text = "；".join(dict.fromkeys(note for note in notes if note))
                if not iso:
                    ok = False
                    missing_text = display
                    missing_for_file.append(display)
                else:
                    shown = display
                    ok = (person, iso) in b_index if person else False
                    missing_text = "" if ok else shown
                    if not ok:
                        missing_dates.append(shown)
                        missing_for_file.append(shown)
                        if not person:
                            note_text = "；".join(part for part in (note_text, "姓名为空，无法比对") if part)
                established_flags.append(ok)
                detail_rows.append(
                    {
                        "A表行号": excel_row,
                        "姓名": person_text,
                        "手机号码": phone,
                        "加班日期": display,
                        "加班是否成立": "是" if ok else "否",
                        "加班缺失日期": missing_text,
                        "原始加班日期": original,
                        "解析说明": note_text,
                    }
                )

            for failure in parsed.failures:
                missing_for_file.append(f"无法解析:{failure}")
                detail_rows.append(
                    {
                        "A表行号": excel_row,
                        "姓名": person_text,
                        "手机号码": phone,
                        "加班日期": failure,
                        "加班是否成立": "否",
                        "加班缺失日期": f"无法解析:{failure}",
                        "原始加班日期": original,
                        "解析说明": "该片段不是可识别的日期，请人工核对",
                    }
                )

            all_ok = bool(established_flags) and all(established_flags) and not parsed.failures
            if not resolved and not parsed.failures:
                all_ok = False
            reason = "" if all_ok else _reason_text(person, missing_dates, parsed.failures, cell_notes)
            summary_rows.append(
                {
                    "A表行号": excel_row,
                    "姓名": person_text,
                    "手机号码": phone,
                    "加班日期": original,
                    "加班是否成立": "是" if all_ok else "否",
                    "加班缺失日期": "、".join(dict.fromkeys(missing_for_file)),
                    "解析说明": "；".join(dict.fromkeys(cell_notes)),
                }
            )
            judgement_rows.append(
                {
                    "A表行号": excel_row,
                    "追加状态": "成立" if all_ok else "不成立",
                    "缺失的日期": "、".join(dict.fromkeys(missing_dates)),
                    "不成立的原因": reason,
                }
            )
        except Exception as exc:
            message = f"该行处理失败：{exc}"
            summary_rows.append(
                {
                    "A表行号": excel_row,
                    "姓名": person_text,
                    "手机号码": phone,
                    "加班日期": original,
                    "加班是否成立": "否",
                    "加班缺失日期": "",
                    "解析说明": message,
                }
            )
            detail_rows.append(
                {
                    "A表行号": excel_row,
                    "姓名": person_text,
                    "手机号码": phone,
                    "加班日期": original,
                    "加班是否成立": "否",
                    "加班缺失日期": "",
                    "原始加班日期": original,
                    "解析说明": message,
                }
            )
            judgement_rows.append(
                {
                    "A表行号": excel_row,
                    "追加状态": "不成立",
                    "缺失的日期": "",
                    "不成立的原因": message,
                }
            )

    summary = pd.DataFrame(summary_rows, columns=SUMMARY_COLUMNS)
    detail = pd.DataFrame(detail_rows, columns=DETAIL_COLUMNS)
    judgements = pd.DataFrame(judgement_rows, columns=JUDGEMENT_COLUMNS)
    return summary, detail, judgements


def _style_workbook(path: Path) -> None:
    workbook = load_workbook(path)
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    yes_fill = PatternFill("solid", fgColor="C6EFCE")
    no_fill = PatternFill("solid", fgColor="FFC7CE")
    wrap = Alignment(vertical="center", wrap_text=True)
    for sheet in workbook.worksheets:
        if sheet.max_row < 1:
            continue
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        sheet.auto_filter.ref = sheet.dimensions
        sheet.freeze_panes = "A2"
        sheet.row_dimensions[1].height = 22
        status_col = None
        for cell in sheet[1]:
            if cell.value == "加班是否成立":
                status_col = cell.column
        if status_col is not None:
            for row in range(2, sheet.max_row + 1):
                cell = sheet.cell(row, status_col)
                if cell.value == "是":
                    cell.fill = yes_fill
                elif cell.value == "否":
                    cell.fill = no_fill
        for column in sheet.columns:
            letter = get_column_letter(column[0].column)
            width = 12
            for cell in column:
                cell.alignment = wrap
                length = len(str(cell.value)) if cell.value is not None else 0
                width = max(width, min(length + 2, 42))
            sheet.column_dimensions[letter].width = width
    workbook.save(path)


def result_file_path(a_path: Path, directory: Path) -> Path:
    """独立结果：原表文件名 + 加班对比结果。"""
    return directory / f"{a_path.stem}_{RESULT_NAME_SUFFIX}.xlsx"


def appended_output_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_原表追加{output_path.suffix}")


def write_original_with_results(
    a_path: Path,
    a_sheet: str,
    header_row: int,
    judgements: pd.DataFrame,
    output_path: Path,
) -> None:
    """复制质检表，按 Excel 行号把三列结果追加到右侧。"""
    try:
        workbook = load_workbook(a_path)
    except PermissionError:
        raise SystemExit(f"质检表被占用，请先关闭 Excel 后重试：{a_path}")
    if a_sheet not in workbook.sheetnames:
        raise SystemExit(f"质检表中没有工作表：{a_sheet}")
    sheet = workbook[a_sheet]
    start_col = sheet.max_column + 1
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    yes_fill = PatternFill("solid", fgColor="C6EFCE")
    no_fill = PatternFill("solid", fgColor="FFC7CE")
    alignment = Alignment(vertical="center", wrap_text=True)
    for offset, title in enumerate(APPEND_HEADERS):
        cell = sheet.cell(header_row, start_col + offset, title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
    by_row = {int(row["A表行号"]): row for _, row in judgements.iterrows()}
    for excel_row, row in by_row.items():
        status = row["追加状态"]
        status_cell = sheet.cell(excel_row, start_col, status)
        missing_cell = sheet.cell(excel_row, start_col + 1, row["缺失的日期"])
        reason_cell = sheet.cell(excel_row, start_col + 2, row["不成立的原因"])
        for cell in (status_cell, missing_cell, reason_cell):
            cell.alignment = alignment
            cell.number_format = "@"
        if status == "成立":
            status_cell.fill = yes_fill
        elif status == "不成立":
            status_cell.fill = no_fill
    for offset, width in enumerate((18, 36, 56)):
        sheet.column_dimensions[get_column_letter(start_col + offset)].width = width
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(output_path)
    except PermissionError:
        raise SystemExit(f"结果文件被占用，请关闭后重试：{output_path}")


def write_output(
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    output_path: Path,
    a_path: Path,
    b_path: Path,
    a_sheet: str,
    b_sheet: str,
    b_warnings: list[str],
    append_path: Path,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    established = int((summary["加班是否成立"] == "是").sum()) if not summary.empty else 0
    missing_rows = int((summary["加班是否成立"] == "否").sum()) if not summary.empty else 0
    info = pd.DataFrame(
        {
            "项目": [
                "A表",
                "A表工作表",
                "B表",
                "B表工作表",
                "A表有效行数",
                "日期全部成立的行数",
                "存在缺失或无法解析的行数",
                "明细行数",
                "判定规则",
                "汇总表说明",
                "明细表说明",
                "原表追加文件",
                "B表解析提示条数",
            ],
            "内容": [
                str(a_path),
                a_sheet,
                str(b_path),
                b_sheet,
                len(summary),
                established,
                missing_rows,
                len(detail),
                "按姓名和加班日期匹配。B表出现过即成立，不看是否同意加班。",
                "一行对应A表一行。加班日期是原单元格文本。全部拆分日期都匹配上，才是「是」。",
                "一行对应一个拆分后的日期。加班日期是YYYY/MM/DD；解析失败时保留原片段。",
                str(append_path),
                len(b_warnings),
            ],
        }
    )
    warning_frame = pd.DataFrame({"B表解析提示": b_warnings or ["无"]})
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="按原行汇总")
            detail.to_excel(writer, index=False, sheet_name="比对明细")
            info.to_excel(writer, index=False, sheet_name="说明")
            warning_frame.to_excel(writer, index=False, sheet_name="B表提示")
        _style_workbook(output_path)
    except PermissionError:
        raise SystemExit(f"结果文件被占用，请关闭后重试：{output_path}")


def _clean_typed_path(raw: str) -> str:
    return raw.strip().strip('"').strip("'")


def ask_existing_excel(prompt: str) -> Path:
    """运行时输入质检表路径，不使用写死的目录。"""
    while True:
        raw = _clean_typed_path(input(prompt))
        if not raw:
            print("路径不能为空，请重新输入。")
            continue
        path = Path(raw).expanduser().resolve()
        if not path.is_file():
            print(f"文件不存在：{path}")
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            print("请输入 Excel 文件（.xlsx、.xlsm 或 .xls）。")
            continue
        return path


def ask_output_directory(prompt: str, default: Path) -> Path:
    """只选择保存文件夹。文件名固定为原表名称加「加班对比结果」。"""
    while True:
        raw = _clean_typed_path(input(prompt))
        if not raw:
            return default
        path = Path(raw).expanduser().resolve()
        if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            path = path.parent
        if not path.is_dir():
            print(f"目录不存在：{path}")
            continue
        return path


def default_mail_table_path() -> Path:
    """回车确认时使用的加班邮件数据表：当前目录下的 ISV自检\\加班邮件数据表.xlsx。"""
    return (Path.cwd() / DEFAULT_MAIL_RELATIVE).resolve()


def _as_excel_file(path: Path) -> Path | None:
    if not path.is_file():
        print(f"文件不存在：{path}")
        return None
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        print("请输入 Excel 文件（.xlsx、.xlsm 或 .xls）。")
        return None
    return path


def ask_mail_table() -> Path:
    """直接回车使用默认路径，也可以改填其他位置。"""
    default = default_mail_table_path()
    prompt = f"请输入加班邮件数据表路径（直接回车使用 {DEFAULT_MAIL_RELATIVE}）："
    while True:
        raw = _clean_typed_path(input(prompt))
        if not raw:
            path = default
            print(f"使用默认路径：{path}")
        else:
            typed = Path(raw).expanduser()
            path = typed if typed.is_absolute() else (Path.cwd() / typed).resolve()
        checked = _as_excel_file(path)
        if checked is not None:
            return checked


def run(a_path: Path, b_path: Path, output_path: Path) -> None:
    print(f"A表：{a_path}")
    print(f"B表：{b_path}")
    a_frame, a_mapping, a_sheet, a_header_row = load_sheet_frame(a_path, A_ALIASES, ("name", "date"))
    b_frame, b_mapping, b_sheet, _b_header_row = load_sheet_frame(b_path, B_ALIASES, ("name", "date"))
    if "phone" not in a_mapping:
        print("提示：A表没有找到手机号列，结果里的手机号码会留空。")
    b_index, years_by_name, b_warnings = build_b_index(b_frame, b_mapping)
    print(f"B表可匹配的人员+日期：{len(b_index)} 条")
    summary, detail, judgements = compare(a_frame, a_mapping, b_index, years_by_name)
    append_path = appended_output_path(output_path)
    write_output(summary, detail, output_path, a_path, b_path, a_sheet, b_sheet, b_warnings, append_path)
    write_original_with_results(a_path, a_sheet, a_header_row, judgements, append_path)
    yes_count = int((summary["加班是否成立"] == "是").sum()) if not summary.empty else 0
    no_count = int((summary["加班是否成立"] == "否").sum()) if not summary.empty else 0
    skip_count = int(summary["加班是否成立"].fillna("").eq("").sum()) if not summary.empty else 0
    print(f"汇总 {len(summary)} 行：成立 {yes_count} 行，不成立 {no_count} 行，日期为空未判断 {skip_count} 行")
    print(f"明细 {len(detail)} 行")
    print(f"已输出：{output_path}")
    print(f"已输出：{append_path}")


def self_test() -> None:
    """用内存中的小表验证拆分、区间、跨格式和“不同意也算成立”。"""
    samples = {
        "多日枚举": parse_date_cell("2026年1月2、3、4日"),
        "混合分隔": parse_date_cell("2026-01-05、2026/1/6, 2026.01.07"),
        "沿用年份": parse_date_cell("2026年1月8日 1月9日"),
        "区间": parse_date_cell("2026-02-01至2026-02-03"),
        "短区间": parse_date_cell("2026-02-01~03"),
        "中文区间": parse_date_cell("2026年1月2日至1月5日"),
        "英文": parse_date_cell("Jan 2, 2026、3 Jan 2026"),
        "全角": parse_date_cell("２０２６－０５－０１、2026-05-02"),
        "带时长": parse_date_cell("2025/10/28、2025/10/29(加班5个小时，0.625天)"),
        "无效": parse_date_cell("13月40日、2026-03-01"),
        "空": parse_date_cell(""),
        "序列号": parse_date_cell(46023),
    }
    expected = {
        "多日枚举": ["2026-01-02", "2026-01-03", "2026-01-04"],
        "混合分隔": ["2026-01-05", "2026-01-06", "2026-01-07"],
        "沿用年份": ["2026-01-08", "2026-01-09"],
        "区间": ["2026-02-01", "2026-02-02", "2026-02-03"],
        "短区间": ["2026-02-01", "2026-02-02", "2026-02-03"],
        "中文区间": ["2026-01-02", "2026-01-03", "2026-01-04", "2026-01-05"],
        "英文": ["2026-01-02", "2026-01-03"],
        "全角": ["2026-05-01", "2026-05-02"],
        "带时长": ["2025-10-28", "2025-10-29"],
    }
    for label, dates in expected.items():
        got = [piece.iso_or_none() for piece in samples[label].pieces]
        if got != dates:
            raise AssertionError(f"{label} 解析为 {got}，期望 {dates}")
    if samples["短区间"].failures:
        raise AssertionError(f"短区间不应有残留失败片段：{samples['短区间'].failures}")
    if samples["带时长"].failures:
        raise AssertionError(f"时长说明不应被当成无法解析的日期：{samples['带时长'].failures}")
    if not samples["无效"].failures or samples["无效"].pieces[0].iso_or_none() != "2026-03-01":
        raise AssertionError("无效日期应保留失败片段，同时保留能解析的日期")
    if samples["空"].notes != ["加班日期为空"]:
        raise AssertionError("空日期提示不正确")
    serial_expect = (date(1899, 12, 30) + timedelta(days=46023)).isoformat()
    if samples["序列号"].pieces[0].iso_or_none() != serial_expect:
        raise AssertionError("Excel序列号转换不正确")

    a_frame = pd.DataFrame(
        [
            ["姓名", "手机号", "加班日期"],
            ["张三", 13800000000, "2026年1月2、3、4日"],
            ["李四", "13900000000.0", "2026-01-05、2026/1/6"],
            ["王五", "13700000000", "1月8日、1月9日"],
            ["赵六", "13600000000", "2026-02-01至2026-02-03"],
            ["吴十", "13200000000", "2026-06-01"],
            ["郑一", "13100000000", "完全不是日期"],
            ["孙八", "13400000000", None],
            ["钱七", "13500000000", "2025/10/29(加班5个小时，0.625天)"],
        ]
    )
    b_frame = pd.DataFrame(
        [
            ["加班人员名称", "加班日期", "是否同意加班"],
            ["张三", "2026-01-02", "同意"],
            ["张三", "2026-01-03", "不同意"],
            ["张三", "2026-01-04", "同意"],
            ["李四", datetime(2026, 1, 5), "不同意"],
            ["李四", datetime(2026, 1, 6), "同意"],
            ["王五", "2026-01-08", "同意"],
            ["王五", "2026-01-09", "同意"],
            ["赵六", "2026-02-01", "同意"],
            ["赵六", "2026-02-02", "不同意"],
            ["赵六", "2026-02-03", "同意"],
            ["钱七", "2025/10/29", "同意"],
        ]
    )
    a_mapping = {"name": 0, "phone": 1, "date": 2}
    b_mapping = {"name": 0, "date": 1}
    a_data = a_frame.iloc[1:].copy()
    a_data["_excel_row"] = a_data.index + 1
    b_data = b_frame.iloc[1:].copy()
    b_data["_excel_row"] = b_data.index + 1
    b_index, years_by_name, _ = build_b_index(b_data, b_mapping)
    summary, detail, judgements = compare(a_data, a_mapping, b_index, years_by_name)
    by_name = {row["姓名"]: row for _, row in summary.iterrows()}
    if by_name["张三"]["加班是否成立"] != "是" or by_name["张三"]["加班缺失日期"] != "":
        raise AssertionError("不同意加班的日期也应判定成立")
    if by_name["李四"]["手机号码"] != "13900000000":
        raise AssertionError("手机号未去掉 .0")
    if by_name["王五"]["加班是否成立"] != "是":
        raise AssertionError("缺年份时应借用B表年份完成匹配")
    if by_name["赵六"]["加班是否成立"] != "是":
        raise AssertionError("日期区间未完整匹配")
    if by_name["吴十"]["加班是否成立"] != "否" or by_name["吴十"]["加班缺失日期"] != "2026/06/01":
        raise AssertionError("B表不存在的日期应记入缺失")
    if by_name["孙八"]["加班是否成立"] != "":
        raise AssertionError("加班日期为空时不应作出是否成立的判断")
    if by_name["钱七"]["加班是否成立"] != "是":
        raise AssertionError("带时长说明的日期应能提取出来并匹配成功")
    judged = {row["A表行号"]: row for _, row in judgements.iterrows()}
    if judged[int(by_name["钱七"]["A表行号"])]["追加状态"] != "成立":
        raise AssertionError("原表追加结果应写成“成立”")
    if judged[int(by_name["孙八"]["A表行号"])]["追加状态"] != "":
        raise AssertionError("空日期追加到原表时应留空")
    wushi = judged[int(by_name["吴十"]["A表行号"])]
    if wushi["追加状态"] != "不成立" or wushi["缺失的日期"] != "2026/06/01" or "2026/06/01" not in wushi["不成立的原因"]:
        raise AssertionError("不成立时应列出缺失日期和原因")
    if by_name["郑一"]["加班是否成立"] != "否" or "无法解析" not in str(by_name["郑一"]["加班缺失日期"]):
        raise AssertionError("无法解析的文本应进入缺失日期")
    zhang_details = detail[detail["姓名"] == "张三"]
    if list(zhang_details["加班是否成立"]) != ["是", "是", "是"]:
        raise AssertionError("明细里的每一天都应成立")

    import tempfile

    with tempfile.TemporaryDirectory() as folder:
        folder_path = Path(folder)
        a_file = folder_path / "质检汇总_张伟芬_重复数据.xlsx"
        b_file = folder_path / "加班邮件数据表.xlsx"
        out_file = folder_path / "质检汇总_张伟芬_重复数据_加班对比结果.xlsx"
        pd.DataFrame(a_frame.iloc[1:].values, columns=a_frame.iloc[0].tolist()).to_excel(a_file, index=False)
        pd.DataFrame(b_frame.iloc[1:].values, columns=b_frame.iloc[0].tolist()).to_excel(b_file, index=False)
        exit_code = main(["--a", str(a_file), "--b", str(b_file), "--output", str(folder_path)])
        if exit_code != 0:
            raise AssertionError("传入质检表、加班邮件数据表和结果路径后比对失败")
        saved = pd.read_excel(out_file, sheet_name="按原行汇总")
        saved_detail = pd.read_excel(out_file, sheet_name="比对明细")
        appended = pd.read_excel(out_file.with_name("质检汇总_张伟芬_重复数据_加班对比结果_原表追加.xlsx"))
    saved_by_name = {row["姓名"]: row for _, row in saved.iterrows()}
    if saved_by_name["张三"]["加班是否成立"] != "是":
        raise AssertionError("写入 Excel 后，不同意加班的日期没有保持成立")
    if saved_by_name["吴十"]["加班缺失日期"] != "2026/06/01":
        raise AssertionError("写入 Excel 后缺失日期不正确")
    appended_by_name = {row["姓名"]: row for _, row in appended.iterrows()}
    if list(appended.columns[-3:]) != APPEND_HEADERS:
        raise AssertionError(f"原表追加列不正确：{list(appended.columns)}")
    if appended_by_name["钱七"]["加班是否成立"] != "成立":
        raise AssertionError("原表追加文件中，匹配成功应写成“成立”")
    if appended_by_name["吴十"]["加班是否成立"] != "不成立" or appended_by_name["吴十"]["缺失的日期"] != "2026/06/01":
        raise AssertionError("原表追加文件中的缺失日期不正确")
    sun_status = appended_by_name["孙八"]["加班是否成立"]
    if not (pd.isna(sun_status) or sun_status == ""):
        raise AssertionError("原表追加文件中，空日期的是否成立应为空")
    if int(saved_by_name["张三"]["A表行号"]) != 2:
        raise AssertionError("A表行号应是 Excel 原始行号")
    if saved_detail[saved_detail["姓名"] == "赵六"].shape[0] != 3:
        raise AssertionError("区间展开后的明细行数不正确")
    print("自检通过。")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按姓名和加班日期比对质检表与加班邮件表")
    parser.add_argument("--a", type=Path, help="质检汇总表路径。不传则运行时输入")
    parser.add_argument("--b", type=Path, help="加班邮件数据表路径。不传则运行时输入，直接回车使用 ISV自检\\加班邮件数据表.xlsx")
    parser.add_argument("--output", type=Path, help="结果保存文件夹。不传则运行时输入；文件名仍按原表名称生成")
    parser.add_argument("--self-test", action="store_true", help="只运行内置样例自检")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv if argv is not None else sys.argv[1:])
    if args.self_test:
        self_test()
        return 0

    if args.a:
        a_path = args.a.expanduser().resolve()
        if not a_path.is_file():
            print(f"质检汇总表不存在：{a_path}")
            return 2
    else:
        a_path = ask_existing_excel("请输入质检汇总表路径：")

    if args.b:
        b_path = args.b.expanduser().resolve()
        checked = _as_excel_file(b_path)
        if checked is None:
            return 2
        b_path = checked
    else:
        b_path = ask_mail_table()
    print(f"加班邮件数据表：{b_path}")

    print(f"结果文件：{a_path.stem}_{RESULT_NAME_SUFFIX}.xlsx")
    print(f"原表追加：{a_path.stem}_{RESULT_NAME_SUFFIX}_原表追加.xlsx")
    if args.output:
        output_dir = args.output.expanduser().resolve()
        if output_dir.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
            output_dir = output_dir.parent
        if not output_dir.is_dir():
            print(f"目录不存在：{output_dir}")
            return 2
    else:
        output_dir = ask_output_directory(
            f"请输入结果保存文件夹（直接回车保存到 {a_path.parent}）：",
            a_path.parent,
        )
    output_path = result_file_path(a_path, output_dir)

    try:
        run(a_path, b_path, output_path)
    except SystemExit:
        raise
    except Exception:
        print("比对过程中出现未预计的错误：")
        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
