# -*- coding: utf-8 -*-
"""
简历数据批量提取工具（PDF / DOCX）

提取字段：姓名、手机号码、毕业时间、工作经历（多段拆分 + 工作/实习类型标注）
输出：Excel（每段工作经历一行，另附“文件汇总”“排除文件”两个工作表）

版式：
    单栏（含日期写在行尾、左右并排的表格行）按视觉行从左到右读取，日期和单位留在同一行。
    双栏（左栏个人信息 / 右栏经历，或报纸式分栏）先读完左栏再读右栏；横跨两栏的标题单独成行，用来切开上下区域。
    同一页里既有单栏又有双栏时按上下区域分段读：单栏区域保持整行，双栏区域才按栏读，避免整页拆栏把日期和单位拆开。

用法：
    python resume_extractor.py
    python resume_extractor.py -i ./简历 -o ./结果.xlsx
    python resume_extractor.py -i ./简历 -n ./人员信息表.xlsx
    python resume_extractor.py -i ./简历 --dump-text ./简历文本

依赖：pip install pymupdf pdfplumber python-docx openpyxl
"""

from __future__ import annotations

import argparse
import bisect
import logging
import os
import re
import sys
import traceback
import warnings
from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

warnings.filterwarnings("ignore", message="The `fitz` API is deprecated")

try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
except ImportError:
    docx = None
    qn = None
    Table = None

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================================
# 配置
# ============================================================================

SUPPORTED_EXTS = {".pdf", ".docx"}

EXCLUDE_FILENAME_KEYWORDS = [
    "保密协议", "合同协议", "合规", "鸿蒙证书", "验收",
    "周报", "加班", "安全", "承诺书",
]

# 工作经历排序方式："asc" 按开始时间升序；"doc" 保持简历中的出现顺序
WORK_SORT = "asc"

ROSTER_FILENAMES = ("人员信息表.xlsx", "人员名单.xlsx")
ROSTER_NAME_HEADERS = ["姓名", "名字", "人员姓名", "员工姓名", "name"]

OUTPUT_COLUMNS = [
    "文件名", "文件路径", "文件格式", "版式", "姓名", "手机号码", "毕业时间",
    "工作序号", "工作类型", "工作开始时间", "工作结束时间", "提取状态", "错误信息",
]

CURRENT_YEAR = datetime.now().year
MIN_YEAR = 1970

SECTION_KEYWORDS: List[Tuple[str, List[str]]] = [
    ("intern", ["实习经", "实习背景", "实习工作", "实习履历", "internship",
                "社会实践", "实践经"]),
    ("project", ["项目经", "项目背景", "项目介绍", "核心项目", "项目案例",
                 "主要项目", "参与项目", "项目履历", "project"]),
    ("work", ["工作经", "职业经", "任职经", "工作履历", "从业经", "工作背景",
              "工作简历", "工作情况", "work experience", "professional experience",
              "work history", "employment", "career"]),
    ("edu", ["教育经", "教育背景", "教育情况", "教育信息", "学习经历", "education",
             "academic"]),
    ("other", ["自我评价", "个人评价", "自我介绍", "自我描述", "专业技能", "技能",
               "证书", "获奖", "荣誉", "培训经", "语言能力", "兴趣爱好", "基本信息",
               "个人信息", "求职意向", "个人优势", "个人总结", "校园经", "在校经",
               "附加信息", "skills", "certificat", "self-evaluation", "summary",
               "awards", "personal information"]),
]

EDU_EXACT_HEADINGS = {"学历", "学历背景", "学历信息", "学历情况", "学历教育"}

# 这些词出现在日期行上，说明该行是项目条目，而不是一段新工作。
# 不用“项目经”：它会误伤“项目经理 / 项目经验”这种职位。
PROJECT_LINE_RE = re.compile(
    r"(项目名称|项目时间|项目周期|项目描述|项目职责|项目简介|项目内容|项目角色|所属项目|"
    r"项目介绍|项目背景|project\s*name)\s*[:：]?"
    r"|^\s*(?:项目|project)\s*[:：]"
    r"|^\s*项目\s*[\d一二三四五六七八九十]",
    re.I,
)

ORG_RE = re.compile(r"有限责任公司|股份有限公司|有限公司|集团|银行|研究院|研究所|事务所|医院|公司")
TITLE_RE = re.compile(
    r"工程师|经理|总监|主管|专员|顾问|助理|实习生|研究员|分析师|设计师|架构师|负责人|"
    r"主任|总裁|教师|老师|律师|会计|开发|运维|产品经理|运营"
)
SCHOOL_RE = re.compile(r"大学|学院|学校|中学|高中|University|College|\bSchool\b", re.I)
DEGREE_RE = re.compile(r"本科|硕士|博士|学士|大专|专科|研究生|学位|MBA|Bachelor|Master|PhD", re.I)

NAME_TITLE_WORDS = {
    "基本信息", "个人简历", "个人信息", "简历", "求职简历", "个人资料", "联系方式",
    "求职意向", "教育背景", "教育经历", "工作经历", "工作经验", "项目经验", "项目经历",
    "自我评价", "个人优势", "专业技能", "应聘简历", "个人履历", "履历", "姓名",
    "男", "女", "汉族", "本科", "硕士", "博士", "大专", "中共党员", "共青团员", "群众",
    "微信", "支付宝", "钉钉", "抖音", "快手", "微博", "美团", "淘宝", "京东", "拼多多",
    "小红书", "腾讯", "阿里", "百度",
}
NAME_BAD_SUBSTRINGS = [
    "简历", "信息", "经历", "经验", "工程师", "开发", "测试", "求职", "意向", "电话",
    "手机", "邮箱", "性别", "年龄", "技能", "评价", "教育", "项目", "公司", "岗位",
    "职位", "应聘", "实习", "工作", "专业", "大学", "学院", "学校", "出生", "民族",
    "籍贯", "地址", "联系", "资料", "初级", "中级", "高级", "资深", "经理", "主管",
    "专员", "助理", "运维", "设计", "产品", "销售", "前端", "后端", "算法", "架构",
    "本科", "硕士", "博士", "党员", "汉族", "履历", "优势", "证书", "合同",
]
NAME_STOP_WORDS = [
    "性别", "年龄", "出生", "民族", "电话", "手机", "籍贯", "学历", "邮箱", "政治",
    "婚姻", "身高", "体重", "联系", "住址", "现居", "求职", "工作", "专业", "毕业",
    "户口", "户籍", "身份", "居住", "意向", "照片",
]
LABEL_WORDS = {
    "姓名", "名字", "电话", "手机", "手机号", "手机号码", "联系电话", "联系方式",
    "邮箱", "邮件", "电子邮箱", "性别", "年龄", "生日", "出生", "出生年月", "民族",
    "籍贯", "学历", "政治面貌", "现居", "居住地", "地址", "求职意向", "意向岗位",
    "工作年限", "身高", "体重", "微信", "qq",
}
_LABEL_KEYS = {re.sub(r"[\s:：]", "", word).lower() for word in LABEL_WORDS}

_TRANS = str.maketrans({
    "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": "-", "\u2014": "-",
    "\u2015": "-", "\u2212": "-", "\uff0d": "-", "\u2500": "-", "\u2501": "-",
    "\uff5e": "~", "\u301c": "~", "\u223c": "~",
    "\uff0e": ".", "\uff0f": "/", "\u3000": " ",
    "\xa0": " ", "\r": "\n", "\u200b": "", "\ufeff": "",
    **{chr(0xFF10 + i): str(i) for i in range(10)},
})


# ============================================================================
# 日期
# ============================================================================

_Y = r"(?:19|20)\d{2}(?!\d)"
_M = r"(?:0?[1-9]|1[0-2])(?!\d)"
_D = r"(?:\s*[./\-]\s*(?:0?[1-9]|[12]\d|3[01])(?!\d)\s*日?|\s*(?:0?[1-9]|[12]\d|3[01])\s*日)?"
_NOW_ALT = r"至今|现在|目前|今(?![天年月日])|now|present|till\s*now|to\s*date"


def _date_pattern(prefix: str) -> str:
    # 末尾的“年”不能吃掉“年度 / 年级”里的年
    return (rf"(?<!\d)(?P<{prefix}y>{_Y})\s*"
            rf"(?:[年./\-]\s*(?P<{prefix}m>{_M})\s*月?{_D})?\s*"
            rf"(?:年(?![度份级获底初末轻限]))?")


# 分隔符两侧最多允许一个换行，避免把上下两段经历粘成一个区间
_SEP = r"[ \t]*(?:\n[ \t]*)?(?:[-~至到]+|\bto\b)[ \t]*(?:\n[ \t]*)?"
_NOW = rf"(?P<now>{_NOW_ALT})"

RANGE_RE = re.compile(_date_pattern("s") + _SEP + rf"(?:{_date_pattern('e')}|{_NOW})", re.I)
DATE_RE = re.compile(_date_pattern("g"), re.I)
SHORT_RANGE_RE = re.compile(
    rf"(?<!\d)(?P<sy>{_Y})\s*[年./\-]\s*(?P<sm>{_M})\s*月?"
    rf"(?:\s*[./\-]\s*(?:0?[1-9]|[12]\d|3[01])\s*日?)?"
    rf"\s*(?:[-~至到]+)\s*(?P<em2>{_M})\s*月?(?!\d)"
    rf"(?!\s*[年./\-]\s*(?:19|20)\d{{2}})"
)
RANGE_MY_RE = re.compile(
    rf"(?<!\d)(?<![.\-年])(?P<sm>{_M})\s*[./]\s*(?P<sy>{_Y})\s*(?:[-~至到]+|\bto\b)\s*"
    rf"(?:(?P<em>{_M})\s*[./]\s*(?P<ey>{_Y})|(?P<now>{_NOW_ALT}))",
    re.I,
)
_MON = r"jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?"
RANGE_EN_RE = re.compile(
    rf"\b(?P<sm>{_MON})\.?\s*(?P<sy>{_Y})\s*(?:[-~至到]+|\bto\b)\s*"
    rf"(?:(?P<em>{_MON})\.?\s*(?P<ey>{_Y})|(?P<now>{_NOW_ALT}))",
    re.I,
)
_MONTH_INDEX = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

GRAD_LABEL_RE = re.compile(
    r"(?:毕业时间|毕业日期|毕业年份|毕业年月|毕业于|结业时间|结业日期|获得学位时间|"
    r"学位获得时间|学位授予时间|获得学位|预计毕业|结业|毕业(?!生))\s*[:：]?\s*"
)
GRAD_SUFFIX_RE = re.compile(
    _date_pattern("g") + r"\s*(?:毕业(?!生)|结业|获得学位|获学位)"
)
STUDY_LABEL_RE = re.compile(
    r"(?:就读时间|在校时间|就读期间|在读时间|学习时间|就读)\s*[:：]?\s*"
)

PHONE_RE = re.compile(r"(?<!\d)(1[3-9]\d)[\s\-]?(\d{4})[\s\-]?(\d{4})(?!\d)")
COUNTRY_CODE_RE = re.compile(
    r"(?<!\d)(?:(?:\+|00)\s*86[\s\-]*|86[\s\-]+)(?=1[3-9]\d)"
)

_BASIC_VALUE_RE = re.compile(
    r"^\s*(?:\d{1,2}(?:\.\d)?\s*(?:年|个月|years?)(?:以上|以下|左右)?|应届\S{0,3}|无|暂无|"
    r"[一二三四五六七八九十]+年(?:以上|以下|左右)?)\s*$",
    re.I,
)
_DECOR_RE = re.compile(
    r"^[\s\d一二三四五六七八九十、.．:：#*■□●◆◇▶►•·\-_|【】\[\]()（）<>《》]+"
)
_DECOR_END_RE = re.compile(r"[\s:：|】\]）)>》·•]+$")


def _year_ok(year: int, max_offset: int = 1) -> bool:
    return MIN_YEAR <= year <= CURRENT_YEAR + max_offset


def _month_name(token: str) -> int:
    return _MONTH_INDEX[token.lower()[:3]]


def fmt_date(year: int, month: Optional[int]) -> str:
    return f"{year}.{month:02d}" if month else f"{year}"


def date_key(value: str) -> Tuple[int, int]:
    if value in ("至今",):
        return (9999, 12)
    parts = value.split(".")
    return (int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)


def _rng(sy: int, sm: Optional[int], ey: Optional[int], em: Optional[int], now: bool) -> Optional[Dict]:
    if not _year_ok(sy, 6):
        return None
    if now:
        return {"sy": sy, "sm": sm, "ey": None, "em": None, "now": True}
    if ey is None or not _year_ok(ey, 6):
        return None
    if (ey, em or 12) < (sy, sm or 1):
        return None
    return {"sy": sy, "sm": sm, "ey": ey, "em": em, "now": False}


def _rng_from_year_first(m: re.Match) -> Optional[Dict]:
    sy = int(m.group("sy"))
    sm = int(m.group("sm")) if m.group("sm") else None
    if m.group("now"):
        return _rng(sy, sm, None, None, True)
    return _rng(sy, sm, int(m.group("ey")), int(m.group("em")) if m.group("em") else None, False)


def iter_ranges(text: str) -> List[Tuple[int, int, Dict]]:
    """返回互不重叠的时间段 (start, end, rng)。同一位置取更长的匹配。"""
    found: List[Tuple[int, int, Dict]] = []
    for m in RANGE_RE.finditer(text):
        rng = _rng_from_year_first(m)
        if rng:
            found.append((m.start(), m.end(), rng))
    for m in SHORT_RANGE_RE.finditer(text):
        sy, sm, em = int(m.group("sy")), int(m.group("sm")), int(m.group("em2"))
        ey = sy if em >= sm else sy + 1
        rng = _rng(sy, sm, ey, em, False)
        if rng:
            found.append((m.start(), m.end(), rng))
    for m in RANGE_MY_RE.finditer(text):
        now = bool(m.group("now"))
        rng = _rng(int(m.group("sy")), int(m.group("sm")),
                   None if now else int(m.group("ey")),
                   None if now else int(m.group("em")), now)
        if rng:
            found.append((m.start(), m.end(), rng))
    for m in RANGE_EN_RE.finditer(text):
        now = bool(m.group("now"))
        rng = _rng(int(m.group("sy")), _month_name(m.group("sm")),
                   None if now else int(m.group("ey")),
                   None if now else _month_name(m.group("em")), now)
        if rng:
            found.append((m.start(), m.end(), rng))
    found.sort(key=lambda item: (item[0], -(item[1] - item[0])))
    accepted: List[Tuple[int, int, Dict]] = []
    for item in found:
        if any(item[0] < prev[1] and item[1] > prev[0] for prev in accepted):
            continue
        accepted.append(item)
    return accepted


def _stitch_dangling_jin(text: str) -> str:
    """把被版式拆开的“至 / 今”接回“至今”。中间最多隔一行单位名称。"""
    lines = text.split("\n")
    output: List[str] = []
    for line in lines:
        if line.strip() == "今" and output:
            joined = False
            for back in (1, 2):
                if len(output) >= back and re.search(r"至\s*$", output[-back]):
                    output[-back] = re.sub(r"至\s*$", "至今", output[-back])
                    joined = True
                    break
            if not joined:
                output.append(line.strip())
        else:
            output.append(line)
    return "\n".join(output)


_RANGE_GLUE_RE = re.compile(r"(?:于|在|从|自|至|到|和|与|及)\s*$")


def _reflow_long_lines(text: str) -> str:
    """把粘在一起的超长行按时间段切开，让每一段经历从日期开始，单位跟在日期后面。"""
    pieces: List[str] = []
    for line in text.split("\n"):
        ranges = iter_ranges(line)
        buried = bool(ranges) and ranges[0][0] > 12
        if not ranges or (len(ranges) == 1 and not (len(line) > 120 and buried)):
            pieces.append(line)
            continue
        pending = ""
        cursor = 0
        for index, (start, _, _) in enumerate(ranges):
            prefix = line[cursor:start]
            nxt = ranges[index + 1][0] if index + 1 < len(ranges) else len(line)
            chunk = line[start:nxt]
            if _RANGE_GLUE_RE.search(prefix):
                pending += prefix + chunk
            else:
                if pending.strip():
                    pieces.append(pending.strip())
                pending = ""
                if prefix.strip():
                    pieces.append(prefix.strip())
                pending = chunk
            cursor = nxt
        if pending.strip():
            pieces.append(pending.strip())
    return "\n".join(piece for piece in pieces if piece.strip())


def normalize_text(text: str) -> str:
    text = text.translate(_TRANS)
    text = text.replace("\x00", "")
    text = re.sub(r"(?<=\d)[ \t]+(?=[年月日])", "", text)
    text = re.sub(r"(?<=[年月])[ \t]+(?=\d)", "", text)
    text = _stitch_dangling_jin(text)
    # “2020年3月 - 至”换行后紧跟“今……”，或被拼进同一行
    text = re.sub(r"至[ \t]*\n[ \t]*今(?!天)", "至今", text)
    text = re.sub(r"至[ \t]+今(?!天)", "至今", text)
    text = re.sub(r"[ \t\f\v]+", " ", text)
    lines = [line.strip() for line in text.split("\n")]
    text = "\n".join(line for line in lines if line)
    return _reflow_long_lines(text)


# ============================================================================
# 标题
# ============================================================================

def detect_heading(line: str) -> Optional[str]:
    raw = line.strip()
    if not raw or len(raw) > 48:
        return None
    if re.search(r"[，。；;！？!?]", raw):
        return None
    if re.search(r"(?:19|20)\d{2}", raw):
        return None
    clean = _DECOR_END_RE.sub("", _DECOR_RE.sub("", raw).strip()).strip()
    if not clean:
        return None
    parts = re.split(r"[:：]", clean, maxsplit=1)
    head = parts[0].strip()
    rest = parts[1].strip() if len(parts) > 1 else ""
    if not head or re.search(r"\d", head):
        return None
    if re.search(r"(丰富|较强|扎实|精通|多年)$", re.sub(r"\s+", "", head)):
        return None
    if rest and (_BASIC_VALUE_RE.match(rest) or re.match(r"^(?:\d|应届|无|暂无)", rest)):
        return None
    if rest and (len(rest) > 18 or re.search(r"[，,、]", rest)):
        return None
    ascii_ratio = sum(ch.isascii() for ch in head) / max(len(head), 1)
    if len(head) > (36 if ascii_ratio > 0.5 else 18):
        return None
    compact = re.sub(r"\s+", "", head)
    lower = re.sub(r"\s+", " ", head).strip().lower()
    for sec_type, keywords in SECTION_KEYWORDS:
        for kw in keywords:
            target = lower if kw.isascii() else compact
            needle = kw.lower() if kw.isascii() else kw
            idx = target.find(needle)
            if idx < 0 or idx > 4:
                continue
            if kw == "学历" and (compact not in EDU_EXACT_HEADINGS or rest):
                continue
            return sec_type
    return None


def line_is_project(line: str) -> bool:
    return bool(PROJECT_LINE_RE.search(line or ""))


def _has_org(text: str) -> bool:
    return bool(ORG_RE.search(text or ""))


def _has_title(text: str) -> bool:
    return bool(TITLE_RE.search(text or ""))


def _line_says_intern(*parts: str) -> bool:
    blob = " ".join(parts)
    return "实习" in blob or "intern" in blob.lower()


def _looks_like_education(text: str) -> bool:
    school = bool(SCHOOL_RE.search(text or ""))
    degree = bool(DEGREE_RE.search(text or ""))
    if school and not _has_org(text) and (degree or "专业" in text or "毕业" in text):
        return True
    return bool(school and degree and not _has_org(text))


def _is_date_text(text: str) -> bool:
    raw = (text or "").strip()
    if not raw or len(raw) > 40:
        return False
    compact = re.sub(r"\s+", "", raw)
    if not 4 <= len(compact) <= 26:
        return False
    if not re.search(r"(?:19|20)\d{2}|至今|现在|目前", compact):
        return False
    stripped = RANGE_RE.sub("", compact, count=1)
    stripped = SHORT_RANGE_RE.sub("", stripped, count=1)
    stripped = RANGE_MY_RE.sub("", stripped, count=1)
    stripped = DATE_RE.sub("", stripped, count=1)
    stripped = re.sub(r"至今|现在|目前|present|now", "", stripped, flags=re.I)
    stripped = re.sub(r"[年月日./~\-至到:：()（）\[\] ]", "", stripped)
    return len(stripped) <= 1


# ============================================================================
# 版式：单栏按行，双栏按栏
# ============================================================================

@dataclass
class VLine:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float = 12.0
    page: int = 0

    @property
    def xc(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def yc(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def width(self) -> float:
        return max(0.0, self.x1 - self.x0)

    @property
    def height(self) -> float:
        return max(1.0, self.y1 - self.y0)


def _drop_noise(lines: Sequence[VLine], page_height: float) -> List[VLine]:
    kept: List[VLine] = []
    for line in lines:
        text = line.text.strip()
        if page_height and (
            re.fullmatch(r"(?:第\s*)?\d{1,3}\s*页?", text)
            or re.fullmatch(r"[-–—]\s*\d{1,3}\s*[-–—]", text)
        ):
            if line.y0 < page_height * 0.08 or line.y1 > page_height * 0.92:
                continue
        kept.append(line)
    return kept


def _typical_height(items: Sequence[VLine]) -> float:
    """普通正文行高。段落文本框的外框很高，不能拿来当行距。"""
    small = sorted(item.height for item in items if item.height <= 28)
    if not small:
        small = [min(item.height for item in items)]
    return max(8.0, small[len(small) // 2])


def _date_fragment(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if not compact or len(compact) > 18:
        return False
    return bool(re.search(r"(?:19|20)\d{2}|至今|现在|目前", compact))


def _same_band(group: Sequence[VLine], line: VLine) -> bool:
    """同一视觉行。

    只认同一基线，或单元格里折成两行的日期。段落文本框不能把旁边整栏吸进来，
    否则左右栏会被拼成一条几百字的超长行。
    """
    height = _typical_height(list(group) + [line])
    top = min(item.y0 for item in group)
    bottom = max(item.y1 for item in group)
    band_top = min(top, line.y0)
    band_bottom = max(bottom, line.y1)
    if band_bottom - band_top > height * 2.6:
        return False
    center = sum(item.y0 + min(item.height, height) * 0.5 for item in group) / len(group)
    line_center = line.y0 + min(line.height, height) * 0.5
    if abs(line_center - center) <= height * 0.7:
        return True
    # 只有同一栏里折成两行的日期才向下并，不能把下一行职位或职责吸进来
    gap = line.y0 - bottom
    continued = _date_fragment(line.text) and any(
        _date_fragment(item.text) and abs(item.x0 - line.x0) <= 10 for item in group
    )
    return bool(continued and -2 <= gap <= height * 0.55)


def _cluster_rows(lines: Sequence[VLine]) -> List[List[VLine]]:
    ordered = sorted(lines, key=lambda line: (line.y0, line.x0))
    if not ordered:
        return []
    rows: List[List[VLine]] = [[ordered[0]]]
    for line in ordered[1:]:
        if _same_band(rows[-1], line):
            rows[-1].append(line)
        else:
            rows.append([line])
    return rows


# 同一基线上，间隙大于这个值就是分栏或日期右对齐，不能把两边句子粘成一条
_ROW_GAP = 32.0


def _flush_bucket(bucket: List[VLine]) -> Optional[VLine]:
    parts = [line.text.strip() for line in bucket if line.text.strip()]
    if not parts:
        return None
    return VLine(
        text=" ".join(parts),
        x0=min(line.x0 for line in bucket),
        y0=min(line.y0 for line in bucket),
        x1=max(line.x1 for line in bucket),
        y1=max(line.y1 for line in bucket),
        size=max(line.size for line in bucket),
        page=bucket[0].page,
    )


def _keep_across_gap(bucket: Sequence[VLine], line: VLine) -> bool:
    """日期和单位分列两侧时仍保持同一行；两边都是正文则切开。"""
    left = " ".join(item.text.strip() for item in bucket if item.text.strip())
    right = line.text.strip()
    if _is_date_text(left) or _is_date_text(right) or _date_fragment(left) or _date_fragment(right):
        return True
    return _is_pure_label(left) or _is_pure_label(right)


def _merge_rows(lines: Sequence[VLine]) -> List[VLine]:
    merged: List[VLine] = []
    for row in _cluster_rows(lines):
        row = [line for line in sorted(row, key=lambda item: (item.x0, item.y0)) if line.text.strip()]
        if not row:
            continue
        bucket: List[VLine] = [row[0]]
        for line in row[1:]:
            gap = line.x0 - max(item.x1 for item in bucket)
            if gap >= _ROW_GAP and not _keep_across_gap(bucket, line):
                flushed = _flush_bucket(bucket)
                if flushed:
                    merged.append(flushed)
                bucket = [line]
            else:
                bucket.append(line)
        flushed = _flush_bucket(bucket)
        if flushed:
            merged.append(flushed)
    return merged


def _side_kind(side: Sequence[VLine]) -> str:
    if len(side) < 2:
        return "short"
    dates = labels = headings = long_lines = 0
    for line in side:
        compact = re.sub(r"\s+", "", line.text)
        if _is_date_text(line.text):
            dates += 1
        if compact.lower() in LABEL_WORDS:
            labels += 1
        if detect_heading(line.text):
            headings += 1
        if len(compact) >= 22:
            long_lines += 1
    count = len(side)
    if headings == 0 and long_lines == 0 and dates / count >= 0.65:
        return "dates"
    if headings == 0 and long_lines == 0 and labels / count >= 0.6:
        return "labels"
    return "content"


def _aligned_ratio(side: Sequence[VLine], other: Sequence[VLine]) -> float:
    if not side or not other:
        return 0.0
    hit = 0
    for line in side:
        limit = max(line.height, 4) * 0.7
        if any(abs(line.yc - item.yc) <= max(limit, item.height * 0.7) for item in other):
            hit += 1
    return hit / len(side)


def _y_overlap_ok(left: Sequence[VLine], right: Sequence[VLine]) -> bool:
    left_top, left_bottom = min(line.y0 for line in left), max(line.y1 for line in left)
    right_top, right_bottom = min(line.y0 for line in right), max(line.y1 for line in right)
    overlap = max(0.0, min(left_bottom, right_bottom) - max(left_top, right_top))
    shorter = min(left_bottom - left_top, right_bottom - right_top)
    return shorter >= 36 and overlap / shorter >= 0.35


def _split_sides(lines: Sequence[VLine], split: float) -> Tuple[List[VLine], List[VLine]]:
    left: List[VLine] = []
    right: List[VLine] = []
    for line in lines:
        if line.x0 < split - 16 and line.x1 > split + 16:
            continue
        if line.xc < split:
            left.append(line)
        else:
            right.append(line)
    return left, right


def _enough(side: Sequence[VLine]) -> bool:
    text = "".join(re.sub(r"\s+", "", line.text) for line in side)
    if len(text) < 8:
        return False
    if len(side) >= 3:
        return True
    # 整栏被做成一个高文本框时，不能因为只有一行就放弃分栏
    return len(text) >= 24 and max(line.height for line in side) >= 36


def _classify_split(lines: Sequence[VLine], split: float) -> str:
    left, right = _split_sides(lines, split)
    if not _enough(left) or not _enough(right) or not _y_overlap_ok(left, right):
        return "none"
    left_kind, right_kind = _side_kind(left), _side_kind(right)
    if "dates" in (left_kind, right_kind) or "labels" in (left_kind, right_kind):
        narrow = left if left_kind in ("dates", "labels") else right
        other = right if narrow is left else left
        if _aligned_ratio(narrow, other) >= 0.5 or "dates" in (left_kind, right_kind):
            return "gutter"
    if left_kind == "content" and right_kind == "content":
        return "columns"
    return "none"


def _is_pure_label(text: str) -> bool:
    compact = re.sub(r"[\s:：]", "", text or "").lower()
    return bool(compact) and len(compact) <= 12 and compact in _LABEL_KEYS


def _valleys(lines: Sequence[VLine], page_width: float) -> List[float]:
    if len(lines) < 6 or page_width < 250:
        return []
    x0 = min(line.x0 for line in lines)
    x1 = max(line.x1 for line in lines)
    width = x1 - x0
    if width < max(180.0, page_width * 0.45):
        return []
    body = [line for line in lines if line.width < width * 0.80]
    if len(body) < 6:
        return []
    bins = 120
    bin_width = width / bins
    hist = [0] * bins
    for line in body:
        start = max(0, min(bins - 1, int((line.x0 - x0) / bin_width)))
        end = max(start, min(bins - 1, int((max(line.x1 - 0.1, line.x0) - x0) / bin_width)))
        for index in range(start, end + 1):
            hist[index] += 1
    peak = max(hist) or 1
    threshold = max(0.5, peak * 0.12)
    lo, hi = int(bins * 0.12), int(bins * 0.88)
    found: List[Tuple[float, float]] = []
    index = lo
    while index < hi:
        if hist[index] <= threshold:
            end = index
            while end < hi and hist[end] <= threshold:
                end += 1
            gap = (end - index) * bin_width
            if gap >= 12:
                found.append((gap, x0 + (index + end) / 2 * bin_width))
            index = end
        else:
            index += 1
    found.sort(key=lambda item: -item[0])
    return [split for _, split in found]


def _row_parts(row: Sequence[VLine], split: float) -> Tuple[List[VLine], List[VLine], List[VLine]]:
    spans: List[VLine] = []
    left: List[VLine] = []
    right: List[VLine] = []
    for line in row:
        if line.x0 < split - 16 and line.x1 > split + 16:
            spans.append(line)
        elif line.xc < split:
            left.append(line)
        else:
            right.append(line)
    return spans, left, right


def _row_kinds(rows: Sequence[Sequence[VLine]], split: float) -> List[str]:
    """每一视觉行：span 横跨中缝，gutter 是日期/标签和正文并排，columns 是左右都有正文，side 只有一侧。"""
    kinds: List[str] = []
    for row in rows:
        spans, left, right = _row_parts(row, split)
        if spans or not left or not right:
            kinds.append("span" if spans else "side")
            continue
        left_text = " ".join(line.text for line in left)
        right_text = " ".join(line.text for line in right)
        if (_is_date_text(left_text) or _is_date_text(right_text)
                or _is_pure_label(left_text) or _is_pure_label(right_text)):
            kinds.append("gutter")
        else:
            kinds.append("columns")
    return _smooth_gutter_rows(rows, kinds, split)


def _nearest_structural(kinds: Sequence[str], index: int, step: int) -> Optional[int]:
    cursor = index + step
    while 0 <= cursor < len(kinds) and kinds[cursor] == "side":
        cursor += step
    return cursor if 0 <= cursor < len(kinds) else None


def _row_gap(rows: Sequence[Sequence[VLine]], earlier: int, later: int) -> float:
    return min(line.y0 for line in rows[later]) - max(line.y1 for line in rows[earlier])


def _paired_gutter_row(row: Sequence[VLine], split: float) -> bool:
    """日期和单位（或标签和值）写在同一行的两侧，读的时候不能拆开。"""
    spans, left, right = _row_parts(row, split)
    if spans or not left or not right:
        return False
    left_text = " ".join(line.text.strip() for line in left if line.text.strip())
    right_text = " ".join(line.text.strip() for line in right if line.text.strip())
    if _is_pure_label(left_text) or _is_pure_label(right_text):
        return True
    if _is_date_text(left_text) and (_has_org(right_text) or _has_title(right_text)):
        return True
    return bool(_is_date_text(right_text) and (_has_org(left_text) or _has_title(left_text)))


def _smooth_gutter_rows(rows: Sequence[Sequence[VLine]], kinds: Sequence[str], split: float) -> List[str]:
    """双栏中间偶尔有一行是日期，不能因此把整栏拆开。日期对面就是公司名时，这一行保持整行。"""
    updated = list(kinds)
    for index, kind in enumerate(kinds):
        if kind != "gutter" or _paired_gutter_row(rows[index], split):
            continue
        before = _nearest_structural(kinds, index, -1)
        after = _nearest_structural(kinds, index, 1)
        if before is None or after is None:
            continue
        if kinds[before] != "columns" or kinds[after] != "columns":
            continue
        if _row_gap(rows, before, index) > 48 or _row_gap(rows, index, after) > 48:
            continue
        updated[index] = "columns"
    return updated


def _column_spans(rows: Sequence[Sequence[VLine]], kinds: Sequence[str]) -> List[Tuple[int, int]]:
    """连续的双栏行（含夹在中间的单侧行）。返回闭区间。"""
    spans: List[Tuple[int, int]] = []
    index = 0
    count = len(kinds)
    while index < count:
        if kinds[index] != "columns":
            index += 1
            continue
        end = index
        cursor = index + 1
        while cursor < count:
            if kinds[cursor] == "columns":
                if _row_gap(rows, end, cursor) > 48:
                    break
                end = cursor
                cursor += 1
                continue
            if kinds[cursor] == "side":
                probe = cursor
                while probe < count and kinds[probe] == "side":
                    probe += 1
                if probe < count and kinds[probe] == "columns" and _row_gap(rows, end, probe) <= 48:
                    cursor = probe
                    continue
            break
        spans.append((index, end))
        index = end + 1
    return spans


def _band_lines(rows: Sequence[Sequence[VLine]], start: int, end: int) -> List[VLine]:
    lines: List[VLine] = []
    for row in rows[start:end + 1]:
        lines.extend(row)
    return lines


def _accepted_column_spans(rows: Sequence[Sequence[VLine]], split: float) -> List[Tuple[int, int]]:
    kinds = _row_kinds(rows, split)
    accepted: List[Tuple[int, int]] = []
    for start, end in _column_spans(rows, kinds):
        band = _band_lines(rows, start, end)
        if _classify_split(band, split) == "columns":
            accepted.append((start, end))
    return accepted


def _best_split(lines: Sequence[VLine], page_width: float) -> Optional[float]:
    """选出确实切开了一块双栏区域的中缝。纯日期栏不会让整页被当成双栏。"""
    valleys = _valleys(lines, page_width)
    if not valleys:
        return None
    rows = _cluster_rows(lines)
    best: Optional[float] = None
    best_score = 0
    for split in valleys:
        score = 0
        for start, end in _accepted_column_spans(rows, split):
            left, right = _split_sides(_band_lines(rows, start, end), split)
            score += min(len(left), len(right))
        if score > best_score:
            best_score = score
            best = split
    return best


def _read_columns(lines: Sequence[VLine], split: float) -> List[VLine]:
    left: List[VLine] = []
    right: List[VLine] = []
    spanning: List[VLine] = []
    for line in lines:
        if line.x0 < split - 16 and line.x1 > split + 16:
            spanning.append(line)
        elif line.xc < split:
            left.append(line)
        else:
            right.append(line)
    left_rows = _merge_rows(left)
    right_rows = _merge_rows(right)
    spans = _merge_rows(spanning)
    bounds = [-1e9] + [line.yc for line in spans] + [1e9]
    left_bins: List[List[VLine]] = [[] for _ in range(len(bounds) - 1)]
    right_bins: List[List[VLine]] = [[] for _ in range(len(bounds) - 1)]

    def place(items: Sequence[VLine], bins: List[List[VLine]]) -> None:
        for line in items:
            for index in range(len(bounds) - 1):
                if bounds[index] <= line.yc < bounds[index + 1]:
                    bins[index].append(line)
                    break

    place(left_rows, left_bins)
    place(right_rows, right_bins)
    ordered: List[VLine] = []
    for index in range(len(bounds) - 1):
        ordered.extend(sorted(left_bins[index], key=lambda line: (line.y0, line.x0)))
        ordered.extend(sorted(right_bins[index], key=lambda line: (line.y0, line.x0)))
        if index < len(spans):
            ordered.append(spans[index])
    return ordered


def _mark_column_rows(rows: Sequence[Sequence[VLine]], split: float) -> Tuple[List[bool], bool]:
    """双栏行，以及紧贴双栏、只是这一行恰好写成日期的边界行。"""
    kinds = _row_kinds(rows, split)
    in_column = [False] * len(rows)
    for start, end in _accepted_column_spans(rows, split):
        for index in range(start, end + 1):
            in_column[index] = True
    for index, flagged in enumerate(list(in_column)):
        if not flagged:
            continue
        for neighbor in (index - 1, index + 1):
            if neighbor < 0 or neighbor >= len(rows) or in_column[neighbor]:
                continue
            if kinds[neighbor] != "gutter" or _paired_gutter_row(rows[neighbor], split):
                continue
            earlier, later = (neighbor, index) if neighbor < index else (index, neighbor)
            if _row_gap(rows, earlier, later) <= 24:
                in_column[neighbor] = True
    # 只有横跨中缝的标题，或日期/正文并排的单栏行，才把这一页算成「单栏+双栏」。
    # 双栏其中一栏比较长、多出来的单侧行不算。
    structural_single = any(
        not flagged and kinds[index] in ("gutter", "span")
        for index, flagged in enumerate(in_column)
    )
    return in_column, structural_single


def _read_mixed(lines: Sequence[VLine], split: float) -> Tuple[List[VLine], str]:
    """单栏区域按行读，双栏区域先左后右。日期写在行尾的单栏不会被双栏中缝拆开。"""
    rows = _cluster_rows(lines)
    in_column, structural_single = _mark_column_rows(rows, split)
    ordered: List[VLine] = []
    column_count = 0
    index = 0
    while index < len(rows):
        if not in_column[index]:
            block: List[VLine] = []
            while index < len(rows) and not in_column[index]:
                block.extend(rows[index])
                index += 1
            ordered.extend(_merge_rows(block))
            continue
        block = []
        while index < len(rows) and in_column[index]:
            block.extend(rows[index])
            index += 1
        ordered.extend(_read_columns(block, split))
        column_count += len(block)
    if column_count and structural_single:
        label = "单栏+双栏"
    elif column_count:
        label = "双栏"
    else:
        label = "单栏"
    return ordered, label


def order_lines(lines: Sequence[VLine], page_width: float, page_height: float = 0) -> Tuple[List[VLine], str]:
    """单栏返回按行合并的文本；双栏先左后右；同一页两种版式并存时分段读取。"""
    usable = _drop_noise(list(lines), page_height)
    split = _best_split(usable, page_width)
    if split is None:
        return _merge_rows(usable), "单栏"
    return _read_mixed(usable, split)


def _layout_label(modes: Sequence[str]) -> str:
    kinds = set(modes)
    if not kinds or kinds <= {"单栏"}:
        return "单栏"
    if kinds <= {"双栏"}:
        return "双栏"
    return "单栏+双栏"


# ============================================================================
# PDF / DOCX 文本
# ============================================================================

def _join_words(words: Sequence[dict]) -> str:
    out = ""
    previous = None
    for word in words:
        piece = word["text"]
        if previous is not None:
            gap = float(word["x0"]) - float(previous["x1"])
            height = max(float(word.get("bottom", 0)) - float(word.get("top", 0)), 8.0)
            prev_ch = out[-1:] if out else ""
            cjk = bool(re.match(r"[\u4e00-\u9fa5]$", prev_ch)) or bool(re.match(r"^[\u4e00-\u9fa5]", piece))
            if (gap > height * 0.35 and not cjk) or gap > height * 0.8:
                out += " "
        out += piece
        previous = word
    return out.strip()


def _span_groups(spans: Sequence[dict]) -> List[List[dict]]:
    """同一条 PDF 线里，水平跳空超过栏距的 span 拆开，避免左右栏粘成一行。"""
    groups: List[List[dict]] = []
    current: List[dict] = []
    prev_x1 = None
    prev_size = 12.0
    for span in spans:
        text = span.get("text") or ""
        if not text.strip():
            continue
        x0, _, x1, _ = span["bbox"]
        size = float(span.get("size") or prev_size or 12.0)
        gap_limit = max(size * 1.8, _ROW_GAP)
        if current and prev_x1 is not None and x0 - prev_x1 >= gap_limit:
            groups.append(current)
            current = []
        current.append(span)
        prev_x1 = x1
        prev_size = size
    if current:
        groups.append(current)
    return groups


def _pymupdf_lines(page, page_no: int) -> List[VLine]:
    lines: List[VLine] = []
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans") or []
            for group in _span_groups(spans):
                text = "".join(span.get("text", "") for span in group).strip()
                if not text:
                    continue
                size = max((float(span.get("size") or 0) for span in group), default=12.0) or 12.0
                lines.append(VLine(
                    text,
                    min(span["bbox"][0] for span in group),
                    min(span["bbox"][1] for span in group),
                    max(span["bbox"][2] for span in group),
                    max(span["bbox"][3] for span in group),
                    size,
                    page_no,
                ))
    return lines


def _plumber_lines(page, page_no: int) -> List[VLine]:
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False, extra_attrs=["size"]) or []
    words = sorted(words, key=lambda word: (float(word["top"]), float(word["x0"])))
    baselines: List[List[dict]] = []
    for word in words:
        if baselines and abs(float(word["top"]) - float(baselines[-1][0]["top"])) <= 3:
            baselines[-1].append(word)
        else:
            baselines.append([word])
    lines: List[VLine] = []
    for group in baselines:
        group.sort(key=lambda word: float(word["x0"]))
        fragment = [group[0]]

        def emit(items: List[dict]) -> None:
            text = _join_words(items)
            if not text:
                return
            size = max(float(item.get("size") or 12) for item in items)
            lines.append(VLine(
                text,
                min(float(item["x0"]) for item in items),
                min(float(item["top"]) for item in items),
                max(float(item["x1"]) for item in items),
                max(float(item["bottom"]) for item in items),
                size,
                page_no,
            ))

        for word in group[1:]:
            if float(word["x0"]) - float(fragment[-1]["x1"]) > 18:
                emit(fragment)
                fragment = [word]
            else:
                fragment.append(word)
        emit(fragment)
    return lines


def extract_pdf_document(path: str) -> Tuple[str, str, List[VLine]]:
    """返回 (原始视觉文本, 版式, 带坐标的原始行)。"""
    errors: List[str] = []
    if pymupdf is not None:
        try:
            document = pymupdf.open(path)
        except Exception as exc:
            errors.append(f"PyMuPDF: {exc}")
            document = None
        if document is not None:
            try:
                if document.needs_pass:
                    raise RuntimeError("PDF 已加密，需要密码")
                ordered: List[VLine] = []
                raw: List[VLine] = []
                modes: List[str] = []
                for page_no, page in enumerate(document):
                    page_lines = _pymupdf_lines(page, page_no)
                    raw.extend(page_lines)
                    page_ordered, mode = order_lines(page_lines, page.rect.width, page.rect.height)
                    ordered.extend(page_ordered)
                    modes.append(mode)
                    logging.debug("第 %d 页版式：%s", page_no + 1, mode)
                text = "\n".join(line.text for line in ordered if line.text.strip())
                if text.strip():
                    return text, _layout_label(modes), raw
            except Exception as exc:
                errors.append(f"PyMuPDF: {exc}")
            finally:
                document.close()
    if pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                ordered = []
                raw = []
                modes = []
                for page_no, page in enumerate(pdf.pages):
                    page_lines = _plumber_lines(page, page_no)
                    raw.extend(page_lines)
                    width = float(page.width or 600)
                    height = float(page.height or 800)
                    page_ordered, mode = order_lines(page_lines, width, height)
                    ordered.extend(page_ordered)
                    modes.append(mode)
                text = "\n".join(line.text for line in ordered if line.text.strip())
                if text.strip():
                    return text, _layout_label(modes), raw
        except Exception as exc:
            errors.append(f"pdfplumber: {exc}")
    if pymupdf is None and pdfplumber is None:
        raise RuntimeError("未安装 PDF 解析库，请执行 pip install pymupdf pdfplumber")
    if errors:
        raise RuntimeError("；".join(errors))
    return "", "单栏", []


def extract_pdf_text(path: str) -> str:
    text, _, _ = extract_pdf_document(path)
    return text


def _inside_textbox(node) -> bool:
    parent = node
    while parent is not None:
        if parent.tag == qn("w:txbxContent"):
            return True
        parent = parent.getparent()
    return False


def _paragraph_text(paragraph) -> str:
    return "".join(
        node.text or ""
        for node in paragraph.iter(qn("w:t"))
        if node.text and not _inside_textbox(node)
    ).strip()


def _cell_text(cell) -> str:
    lines: List[str] = []
    for child in cell._tc.iterchildren():
        if child.tag == qn("w:p"):
            text = _paragraph_text(child)
            if text:
                lines.append(text)
        elif child.tag == qn("w:tbl"):
            _, nested_lines = _emit_table(Table(child, cell))
            lines.extend(nested_lines)
    return "\n".join(lines)

def _table_matrix(table) -> List[List[str]]:
    rows: List[List[str]] = []
    for row in table.rows:
        seen = set()
        values: List[str] = []
        try:
            cells = row.cells
        except Exception:
            continue
        for cell in cells:
            marker = id(cell._tc)
            if marker in seen:
                continue
            seen.add(marker)
            values.append(_cell_text(cell).strip())
        rows.append(values)
    return rows


def _column_date_ratio(rows: Sequence[Sequence[str]], index: int) -> float:
    values = [row[index] for row in rows if len(row) > index and row[index].strip()]
    if not values:
        return 0.0
    hits = sum(1 for value in values if _is_date_text(value.replace("\n", " ")))
    return hits / len(values)


def _is_label_value_table(rows: Sequence[Sequence[str]]) -> bool:
    if len(rows) < 2:
        return False
    usable = 0
    matched = 0
    for row in rows:
        cells = [cell.strip() for cell in row if cell and cell.strip()]
        if len(cells) < 2:
            continue
        usable += 1
        label = re.sub(r"[\s:：]", "", cells[0])
        if 1 <= len(label) <= 10 and "\n" not in cells[0] and len(cells[1]) <= 40:
            matched += 1
    return usable >= 2 and matched / usable >= 0.6


def _is_two_column_table(rows: Sequence[Sequence[str]]) -> bool:
    if not rows or max(len(row) for row in rows) != 2:
        return False
    if _column_date_ratio(rows, 0) >= 0.6 or _column_date_ratio(rows, 1) >= 0.6:
        return False
    if _is_label_value_table(rows):
        return False
    left = [row[0] for row in rows if row and row[0].strip()]
    right = [row[1] for row in rows if len(row) > 1 and row[1].strip()]
    if not left or not right:
        return False
    left_text = re.sub(r"\s+", "", "\n".join(left))
    right_text = re.sub(r"\s+", "", "\n".join(right))
    if len(left_text) < 20 or len(right_text) < 20:
        return False
    if len(rows) == 1 and (left[0].count("\n") < 2 or right[0].count("\n") < 2):
        return False
    labelish = 0
    for value in left:
        first = re.sub(r"\s+", "", value.split("\n")[0])
        if len(first) <= 8:
            labelish += 1
    if labelish / len(left) >= 0.7 and sum(value.count("\n") for value in left) < 3:
        return False
    return True


def _emit_table(table) -> Tuple[str, List[str]]:
    rows = _table_matrix(table)
    return _emit_table_rows(rows)


def _emit_table_rows(rows: Sequence[Sequence[str]]) -> Tuple[str, List[str]]:
    if not rows:
        return "单栏", []
    if _is_two_column_table(rows):
        lines: List[str] = []
        width = max(len(row) for row in rows)
        for column in range(width):
            for row in rows:
                if column < len(row) and row[column].strip():
                    lines.extend(part.strip() for part in row[column].split("\n") if part.strip())
        return "双栏", lines
    if _is_label_value_table(rows):
        lines = []
        for row in rows:
            cells = [re.sub(r"\s*\n\s*", " ", cell.strip()) for cell in row if cell and cell.strip()]
            if cells:
                lines.append(" ".join(cells))
        return "单栏", lines
    lines = []
    for row in rows:
        cells = [cell.strip() for cell in row if cell and cell.strip()]
        if not cells:
            continue
        if len(cells) == 2 and _is_date_text(cells[0].replace("\n", " ")):
            detail = [part.strip() for part in cells[1].split("\n") if part.strip()]
            if detail:
                lines.append(f"{cells[0].replace(chr(10), ' ')} {detail[0]}")
                lines.extend(detail[1:])
            else:
                lines.append(cells[0].replace("\n", " "))
            continue
        if all("\n" not in cell and len(cell) <= 36 for cell in cells):
            lines.append(" ".join(cells))
            continue
        for cell in cells:
            lines.extend(part.strip() for part in cell.split("\n") if part.strip())
    return "单栏", lines


def _floating_blocks(root) -> List[Tuple[int, int, str]]:
    blocks: List[Tuple[int, int, str]] = []
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for drawing in root.iter(qn("w:drawing")):
        horizontal = vertical = None
        for node in drawing.iter(f"{wp}positionH"):
            offset = node.find(f"{wp}posOffset")
            if offset is not None and offset.text and offset.text.lstrip("-").isdigit():
                horizontal = int(offset.text)
                break
        for node in drawing.iter(f"{wp}positionV"):
            offset = node.find(f"{wp}posOffset")
            if offset is not None and offset.text and offset.text.lstrip("-").isdigit():
                vertical = int(offset.text)
                break
        paragraphs = []
        for paragraph in drawing.iter(qn("w:p")):
            text = "".join(node.text or "" for node in paragraph.iter(qn("w:t"))).strip()
            if text:
                paragraphs.append(text)
        if paragraphs:
            blocks.append((horizontal if horizontal is not None else 0,
                           vertical if vertical is not None else 0,
                           "\n".join(paragraphs)))
    return blocks


def _order_floats(blocks: Sequence[Tuple[int, int, str]]) -> Tuple[str, str]:
    if not blocks:
        return "", "单栏"
    if len(blocks) == 1:
        return blocks[0][2], "单栏"
    xs = sorted(block[0] for block in blocks)
    gaps = [(xs[index + 1] - xs[index], (xs[index] + xs[index + 1]) / 2) for index in range(len(xs) - 1)]
    gap, mid = max(gaps, key=lambda item: item[0])
    if gap < 400000:
        ordered = sorted(blocks, key=lambda block: (block[1], block[0]))
        return "\n".join(block[2] for block in ordered), "单栏"
    left = sorted((block for block in blocks if block[0] < mid), key=lambda block: block[1])
    right = sorted((block for block in blocks if block[0] >= mid), key=lambda block: block[1])
    if not left or not right:
        ordered = sorted(blocks, key=lambda block: (block[1], block[0]))
        return "\n".join(block[2] for block in ordered), "单栏"
    left_span = (min(block[1] for block in left), max(block[1] for block in left))
    right_span = (min(block[1] for block in right), max(block[1] for block in right))
    overlap = min(left_span[1], right_span[1]) - max(left_span[0], right_span[0])
    if overlap <= 0:
        ordered = sorted(blocks, key=lambda block: (block[1], block[0]))
        return "\n".join(block[2] for block in ordered), "单栏"
    text = "\n".join(block[2] for block in left + right)
    return text, "双栏"


def extract_docx_text(path: str) -> Tuple[str, List[List[List[str]]], str]:
    if docx is None:
        raise RuntimeError("未安装 python-docx，请执行 pip install python-docx")
    document = docx.Document(path)
    lines: List[str] = []
    tables: List[List[List[str]]] = []
    layout = "单栏"
    for section in document.sections:
        try:
            for paragraph in section.header.paragraphs:
                text = paragraph.text.strip()
                if text:
                    lines.append(text)
        except Exception:
            pass
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            text = _paragraph_text(child)
            if text:
                lines.append(text)
        elif child.tag == qn("w:tbl"):
            table = Table(child, document)
            rows = _table_matrix(table)
            tables.append(rows)
            mode, table_lines = _emit_table_rows(rows)
            if mode == "双栏":
                layout = "双栏"
            lines.extend(table_lines)
    floats = _floating_blocks(document.element.body)
    float_text, float_layout = _order_floats(floats)
    if float_text.strip():
        if float_layout == "双栏":
            layout = "双栏"
        body_len = len(re.sub(r"\s+", "", "\n".join(lines)))
        float_lines = [line for line in float_text.split("\n") if line.strip()]
        if body_len < 40 or float_layout == "双栏" or min(block[0] for block in floats) < 3_200_000:
            lines = float_lines + lines
        else:
            lines.extend(float_lines)
    return "\n".join(lines), tables, layout


# ============================================================================
# 分区
# ============================================================================

def _job_resume_line(lines: Sequence[str], index: int) -> bool:
    line = lines[index]
    if not iter_ranges(line) or line_is_project(line):
        return False
    prev_line = lines[index - 1] if index else ""
    next_line = lines[index + 1] if index + 1 < len(lines) else ""
    # 上一行还是“项目职责”时，只要本行或下一行已经出现公司，就视为下一段工作
    if line_is_project(prev_line) and not _has_org(line) and not _has_org(next_line):
        return False
    window = "\n".join((prev_line, line, next_line))
    if _has_org(line) or _has_org(next_line):
        return True
    return bool(_has_title(line) and "项目" not in line and _has_org(window))


class SectionMap:
    """字符偏移 → 所在行，以及该行所属栏目。"""

    def __init__(self, text: str):
        self.text = text
        self.lines = text.split("\n")
        self.line_starts: List[int] = []
        pos = 0
        for line in self.lines:
            self.line_starts.append(pos)
            pos += len(line) + 1
        self.sections: List[Optional[str]] = []
        mode: Optional[str] = None
        nested_project = False
        for index, line in enumerate(self.lines):
            heading = detect_heading(line)
            following = self.lines[index + 1] if index + 1 < len(self.lines) else ""
            if heading and _BASIC_VALUE_RE.match(following):
                heading = None
            if heading == "project" and mode in ("work", "intern"):
                nested_project = True
                self.sections.append("project")
                continue
            if heading:
                mode = heading
                nested_project = False
                self.sections.append(mode)
                continue
            if nested_project and _job_resume_line(self.lines, index):
                nested_project = False
            self.sections.append("project" if nested_project else mode)

    def line_index(self, pos: int) -> int:
        return max(bisect.bisect_right(self.line_starts, pos) - 1, 0)

    def section_at(self, pos: int) -> Optional[str]:
        return self.sections[self.line_index(pos)]

    def line_at(self, pos: int) -> str:
        return self.lines[self.line_index(pos)]

    def prev_line(self, pos: int) -> str:
        index = self.line_index(pos)
        return self.lines[index - 1] if index else ""

    def next_line(self, pos: int) -> str:
        index = self.line_index(pos)
        return self.lines[index + 1] if index + 1 < len(self.lines) else ""


# ============================================================================
# 字段
# ============================================================================

def is_valid_name(name: Optional[str]) -> bool:
    if not name:
        return False
    name = name.strip()
    if re.fullmatch(r"[A-Z][a-z]+(?:[ '\-][A-Z][a-z]+){1,2}", name):
        return name.lower() not in {"work experience", "project experience", "education background"}
    if "·" in name:
        if not 2 <= len(name.replace("·", "")) <= 12 or len(name) > 20:
            return False
    elif not 2 <= len(name) <= 6:
        return False
    if not re.search(r"[\u4e00-\u9fa5]", name):
        return False
    if re.search(r"\d", name):
        return False
    if name in NAME_TITLE_WORDS:
        return False
    return not any(bad in name for bad in NAME_BAD_SUBSTRINGS)


def _cut_name(candidate: str) -> str:
    for word in NAME_STOP_WORDS:
        index = candidate.find(word)
        if index > 0:
            candidate = candidate[:index]
    return candidate.strip("·").strip()


def _boundary_pos(text: str, name: str) -> int:
    start = 0
    while True:
        index = text.find(name, start)
        if index < 0:
            return -1
        before = text[index - 1] if index else ""
        after = text[index + len(name)] if index + len(name) < len(text) else ""
        before_ok = not before or not re.match(r"[\u4e00-\u9fa5·]", before)
        after_ok = not after or not re.match(r"[\u4e00-\u9fa5·]", after)
        if before_ok and after_ok:
            return index
        start = index + 1


def name_from_label(text: str) -> Optional[str]:
    for match in re.finditer(r"姓\s*名\s*[:：]?\s*([\u4e00-\u9fa5·]{2,20})", text):
        candidate = _cut_name(match.group(1))
        if is_valid_name(candidate):
            return candidate
    for match in re.finditer(
        r"姓\s*名\s*[:：]?\s*([\u4e00-\u9fa5](?: [\u4e00-\u9fa5]){1,3})(?![\u4e00-\u9fa5])",
        text,
    ):
        candidate = match.group(1).replace(" ", "")
        if is_valid_name(candidate):
            return candidate
    return None


def name_from_docx_tables(tables: List[List[List[str]]]) -> Optional[str]:
    labels = {"姓名", "名字", "姓名name", "name"}
    for rows in tables:
        for row_index, row in enumerate(rows):
            for col_index, cell in enumerate(row):
                for piece_index, piece in enumerate(cell.split("\n")):
                    key = re.sub(r"[\s:：]", "", piece).lower()
                    if key not in labels:
                        continue
                    following = cell.split("\n")[piece_index + 1:]
                    for value in following:
                        candidate = _cut_name(re.sub(r"\s+", "", value))
                        if candidate and is_valid_name(candidate):
                            return candidate
                        if candidate:
                            break
                key = re.sub(r"[\s:：]", "", cell.split("\n")[0]).lower()
                if key not in labels:
                    continue
                for right in row[col_index + 1:]:
                    value = re.sub(r"\s+", "", right.split("\n")[0])
                    if value and value.lower() not in labels:
                        candidate = _cut_name(value)
                        if is_valid_name(candidate):
                            return candidate
                        break
                if row_index + 1 < len(rows) and col_index < len(rows[row_index + 1]):
                    value = re.sub(r"\s+", "", rows[row_index + 1][col_index].split("\n")[0])
                    if is_valid_name(value):
                        return value
    return None


def _filename_tokens(filename: str) -> List[str]:
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r"(?:的?个人简历|的?简历|求职|应聘|resume|cv)", " ", stem, flags=re.I)
    return [token for token in re.split(r"[_\-\s()（）\[\]【】、,，.+#&@=~]+", stem) if token]


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _pick_earliest(candidates: Sequence[str], compact_text: str) -> Optional[str]:
    ranked = []
    for candidate in candidates:
        pos = _boundary_pos(compact_text, candidate)
        if pos >= 0:
            ranked.append((pos, -len(candidate), candidate))
    if not ranked:
        return None
    ranked.sort()
    return ranked[0][2]


def name_from_filename(filename: str, text: str) -> Optional[str]:
    compact_text = _compact(text)
    seen = set()
    both: List[str] = []
    for token in _filename_tokens(filename):
        for candidate in re.findall(r"[\u4e00-\u9fa5·]{2,12}", token):
            if candidate in seen or not is_valid_name(candidate):
                continue
            seen.add(candidate)
            if _boundary_pos(compact_text, candidate) >= 0:
                both.append(candidate)
    return _pick_earliest(both, compact_text)


def load_roster_names(path: str) -> List[str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.active
        rows = sheet.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            return []
        column = None
        for index, cell in enumerate(header):
            key = re.sub(r"\s+", "", str(cell or "")).lower()
            if any(name.lower() == key or name.lower() in key for name in ROSTER_NAME_HEADERS):
                column = index
                break
        if column is None:
            column = 0
            logging.warning("人员信息表未找到“姓名”列，改用第一列：%s", header[0])
        names: List[str] = []
        seen = set()
        for row in rows:
            if not row or column >= len(row) or row[column] is None:
                continue
            name = re.sub(r"\s+", "", str(row[column]).strip())
            if name and name not in seen:
                seen.add(name)
                names.append(name)
        return names
    finally:
        workbook.close()


def name_from_roster(text: str, filename: str, roster: Sequence[str]) -> Optional[str]:
    if not roster:
        return None
    compact_text = _compact(text)
    compact_name = _compact(os.path.splitext(filename)[0])
    in_both: List[str] = []
    in_text: List[str] = []
    in_file: List[str] = []
    for name in sorted(roster, key=len, reverse=True):
        if not name:
            continue
        text_pos = _boundary_pos(compact_text, name)
        file_pos = _boundary_pos(compact_name, name)
        if text_pos >= 0 and file_pos >= 0:
            in_both.append(name)
        elif text_pos >= 0:
            in_text.append(name)
        elif file_pos >= 0:
            in_file.append(name)
    picked = _pick_earliest(in_both, compact_text) or _pick_earliest(in_text, compact_text)
    if picked:
        return picked
    if in_file:
        in_file.sort(key=lambda name: (-len(name), compact_name.find(name)))
        return in_file[0]
    return None


def name_from_first_lines(text: str, max_lines: int = 6) -> Optional[str]:
    for line in text.split("\n")[:max_lines]:
        line = line.strip()
        if not line or re.search(r"[。；;！？]", line):
            continue
        if re.fullmatch(r"[\u4e00-\u9fa5](?: [\u4e00-\u9fa5]){1,3}", line):
            candidate = line.replace(" ", "")
            if is_valid_name(candidate):
                return candidate
        if re.fullmatch(r"[A-Z][a-z]+(?:\s+[A-Z][a-z]+){1,2}", line) and is_valid_name(line):
            return line
        first = re.split(r"[\s|｜/,，:：]+", line)[0]
        first = re.sub(r"\s+", "", first)
        if not (re.fullmatch(r"[\u4e00-\u9fa5·]{2,6}", first) and is_valid_name(first)):
            continue
        # 页眉经常是“王芳 | 女 | 1997.08 | 邮箱”，名字在最前面
        if len(line) <= 40 or re.match(r"^[\u4e00-\u9fa5·]{2,6}\s*[|｜/]", line):
            return first
    return None


def name_from_font(lines: Sequence[VLine]) -> Optional[str]:
    if not lines:
        return None
    first_page = lines[0].page
    page_lines = [line for line in lines if line.page == first_page] or list(lines)
    sizes = sorted(line.size for line in page_lines)
    median = sizes[len(sizes) // 2]
    bottom = max(line.y1 for line in page_lines)
    candidates = []
    for line in page_lines:
        if line.y0 > bottom * 0.48:
            continue
        if line.size < max(14.0, median * 1.25):
            continue
        raw = line.text.translate(_TRANS).strip()
        raw = re.sub(r"^(?:个人简历|求职简历|应聘简历|简历)\s*", "", raw)
        first = re.split(r"[\s|｜/,，]+", raw)[0]
        first = re.sub(r"\s+", "", first)
        if is_valid_name(first):
            candidates.append((-line.size, line.y0, first))
    if not candidates:
        return None
    candidates.sort()
    return candidates[0][2]


def extract_name(text: str, filename: str, ext: str, tables,
                 roster: Optional[Sequence[str]] = None,
                 font_lines: Optional[Sequence[VLine]] = None) -> Optional[str]:
    if roster:
        name = name_from_roster(text, filename, roster)
        if name:
            return name
    name = name_from_label(text)
    if not name and ext == ".docx":
        name = name_from_docx_tables(tables or [])
    if not name and font_lines:
        name = name_from_font(font_lines)
    if not name:
        name = name_from_filename(filename, text)
    if not name:
        name = name_from_first_lines(text)
    return name


def _phone_inside_long_run(line: str, phone: str) -> bool:
    for run in re.findall(r"\d{12,}", line):
        if phone in run:
            return True
    return False


def extract_phones(text: str) -> List[str]:
    cleaned = COUNTRY_CODE_RE.sub(" ", text)
    phones: List[str] = []
    for match in PHONE_RE.finditer(cleaned):
        phone = "".join(match.groups())
        line_start = cleaned.rfind("\n", 0, match.start()) + 1
        line_end = cleaned.find("\n", match.end())
        line = cleaned[line_start: len(cleaned) if line_end < 0 else line_end]
        prefix = line[: match.start() - line_start]
        if _phone_inside_long_run(line, phone):
            continue
        if re.search(r"身份证|证件号", prefix[-12:]) and not re.search(r"手机|电话|联系", line):
            continue
        if re.search(r"QQ|qq|扣扣", prefix[-6:]) and not re.search(r"手机|电话", line):
            continue
        if phone not in phones:
            phones.append(phone)
    return phones


def _format_graduations(pairs: Iterable[Tuple[int, Optional[int]]]) -> List[str]:
    by_year: Dict[int, set] = {}
    for year, month in pairs:
        if _year_ok(year, 6):
            by_year.setdefault(year, set()).add(month if month else None)
    values: List[str] = []
    for year, months in by_year.items():
        concrete = sorted(month for month in months if month)
        if concrete:
            values.extend(fmt_date(year, month) for month in concrete)
        else:
            values.append(fmt_date(year, None))
    return sorted(set(values), key=date_key)


def extract_graduation(text: str, smap: SectionMap) -> List[str]:
    found: List[Tuple[int, Optional[int]]] = []

    def add(year: int, month: Optional[int]) -> None:
        found.append((year, month))

    def add_range(rng: Optional[Dict]) -> None:
        if rng and not rng["now"]:
            add(rng["ey"], rng["em"])

    for match in GRAD_LABEL_RE.finditer(text):
        window = text[match.end(): match.end() + 48]
        ranges = iter_ranges(window)
        if ranges:
            add_range(ranges[0][2])
            continue
        date_match = DATE_RE.match(text, match.end())
        if date_match:
            add(int(date_match.group("gy")), int(date_match.group("gm")) if date_match.group("gm") else None)
    for match in GRAD_SUFFIX_RE.finditer(text):
        add(int(match.group("gy")), int(match.group("gm")) if match.group("gm") else None)
    for match in STUDY_LABEL_RE.finditer(text):
        ranges = iter_ranges(text[match.end(): match.end() + 48])
        if ranges:
            add_range(ranges[0][2])

    for start, _, rng in iter_ranges(text):
        if rng["now"]:
            continue
        section = smap.section_at(start)
        line = smap.line_at(start)
        prev = smap.prev_line(start)
        nxt = smap.next_line(start)
        if section in ("work", "intern", "project", "other"):
            continue
        around = "\n".join((prev, line, nxt)) if _is_date_text(line) else f"{prev}\n{line}"
        # 出生年月紧挨着教育经历时，不能把学校那一行一起丢掉
        if (re.search(r"出生|生日", line) or re.search(r"出生|生日", prev)) and not _looks_like_education(around):
            continue
        if section == "edu" or _looks_like_education(around):
            add_range(rng)
    return _format_graduations(found)


def _focus_span(line: str, offset: int, end_offset: int) -> Tuple[str, int, int]:
    """超长行里只看日期前后这一小段，行内更早的句号不能把这段经历一起丢掉。"""
    if len(line) <= 120:
        return line, offset, end_offset
    left = max(0, offset - 36)
    for mark in "。；;！？":
        found = line.rfind(mark, left, offset)
        if found >= left:
            left = found + 1
    next_start = len(line)
    for start, _, _ in iter_ranges(line):
        if start > offset:
            next_start = start
            break
    if next_start < len(line):
        right = next_start
    else:
        right = min(len(line), end_offset + 80)
    segment = line[left:right]
    return segment, offset - left, end_offset - left


def _skip_prose(line: str, offset: int, end_offset: int, rng: Dict, section: Optional[str], after: str) -> bool:
    """描述句里的年份（“参与 2022-2023 年度……”）不是一段经历。

    只看日期紧挨着的几个字。整行很长、或者更前面有句号，不能把这段日期丢掉。
    """
    if re.match(r"\s*(?:年[度份级获]|期间)", after):
        return True
    before = line[:offset]
    if _RANGE_GLUE_RE.search(before):
        return True
    has_month = bool(rng.get("sm") or rng.get("em") or rng.get("now"))
    nearby = line[max(0, offset - 36): min(len(line), end_offset + 40)]
    anchored = has_month and (_has_org(nearby) or _has_title(nearby) or section in ("work", "intern"))
    tail = before[-16:]
    if re.search(r"[。；;！？，,]\s*$", tail) and not anchored:
        return True
    return False


def _org_snippet(text: str) -> str:
    match = re.search(r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,24}(?:有限责任公司|股份有限公司|有限公司|集团|银行|研究院|研究所|事务所)", text)
    if match:
        return match.group(0)
    match = re.search(r"[\u4e00-\u9fa5A-Za-z0-9（）()]{2,18}公司", text)
    return match.group(0) if match else ""


def _work_plausible(rng: Dict) -> bool:
    if rng["now"]:
        return True
    return rng["ey"] <= CURRENT_YEAR + 1


def _explicit_job(line: str, nxt: str) -> bool:
    """左栏读完「项目经历 / 技能」后，右栏或下一页仍可能是一段工作。

    两种写法都算：日期单独成行且下一行是公司；同一行里同时有公司和日期。
    """
    if line_is_project(line) or line_is_project(nxt):
        return False
    if _is_date_text(line) and _has_org(nxt):
        return True
    return bool(_has_org(line) and iter_ranges(line))


def _stop_neighbor(other: str) -> bool:
    if detect_heading(other) or _is_date_text(other):
        return True
    return bool(iter_ranges(other) and (_has_org(other) or _has_title(other)))


def _scan_toward(lines: Sequence[str], index: int, step: int, limit: int = 3) -> List[str]:
    """向一侧找最近的公司名。遇到下一段经历或栏目标题就停。"""
    found: List[str] = []
    for dist in range(1, limit + 1):
        cursor = index + step * dist
        if cursor < 0 or cursor >= len(lines):
            break
        other = lines[cursor].strip()
        if not other:
            continue
        if _stop_neighbor(other):
            break
        found.append(other)
        if _has_org(other):
            break
    return found


def _neighbor_context(lines: Sequence[str], index: int) -> List[str]:
    if _has_org(lines[index]) or _looks_like_education(lines[index]):
        return []
    forward = _scan_toward(lines, index, 1)
    if any(_has_org(line) for line in forward):
        return forward
    backward = _scan_toward(lines, index, -1)
    if any(_has_org(line) for line in backward):
        return backward
    return forward or backward


def _date_org_pairs(lines: Sequence[str]) -> Dict[int, str]:
    """左栏全是日期、右栏全是公司时，按出现顺序把同一批日期和公司配上。"""
    pairs: Dict[int, str] = {}
    count = len(lines)
    index = 0
    while index < count:
        if not _is_date_text(lines[index]):
            index += 1
            continue
        end = index
        while end < count and _is_date_text(lines[end]):
            end += 1
        dates = list(range(index, end))
        index = end
        if len(dates) < 2:
            continue
        orgs: List[str] = []
        cursor = end
        limit = end + len(dates) * 3 + 2
        while cursor < count and cursor < limit and len(orgs) < len(dates):
            other = lines[cursor].strip()
            if other and (detect_heading(other) or _is_date_text(other)):
                break
            if other and iter_ranges(other) and not _has_org(other):
                break
            if other and _has_org(other) and not _looks_like_education(other):
                orgs.append(other)
            cursor += 1
        if len(orgs) == len(dates):
            for date_index, org in zip(dates, orgs):
                pairs[date_index] = org
    return pairs


def extract_work(text: str, smap: SectionMap) -> List[Dict]:
    found: List[Dict] = []
    pairs = _date_org_pairs(smap.lines)
    for start, end, rng in iter_ranges(text):
        if not _work_plausible(rng):
            continue
        section = smap.section_at(start)
        raw = smap.line_at(start)
        index = smap.line_index(start)
        prev = smap.prev_line(start)
        nxt = smap.next_line(start)
        offset = start - smap.line_starts[index]
        end_offset = offset + (end - start)
        line, local_offset, local_end = _focus_span(raw, offset, end_offset)
        if _skip_prose(line, local_offset, local_end, rng, section, text[end:end + 8]):
            continue
        if STUDY_LABEL_RE.search(raw) or STUDY_LABEL_RE.search(prev):
            continue
        if re.search(r"大学|学院", line) and re.search(r"进修|培训|课程|在读|就读", line) and not _has_org(line):
            continue
        paired = pairs.get(index, "")
        neighbors = _neighbor_context(smap.lines, index)
        # 日期单独成行时，单位可能在下一行；双栏被拆开时，公司可能再隔一两行
        entry = "\n".join(part for part in (raw, nxt) if part) if _is_date_text(raw) else line
        own_parts = [entry]
        if paired:
            own_parts.append(paired)
        own_text = "\n".join(own_parts)
        context_parts = [own_text]
        if not _has_org(own_text):
            context_parts.extend(neighbors)
        context = "\n".join(part for part in context_parts if part)
        org_near = paired or next((part for part in neighbors if _has_org(part)), "")
        intern_bits = [entry, prev, paired]
        # 下一行如果已经是另一段经历或栏目标题，不能把“实习”算到本段上
        if nxt and not detect_heading(nxt) and not iter_ranges(nxt):
            intern_bits.append(nxt)
        intern_bits.extend(part for part in neighbors if part and not iter_ranges(part))
        intern = _line_says_intern(*intern_bits)
        if line_is_project(raw) or line_is_project(line):
            explicit_job = False
        elif _is_date_text(raw) and (_has_org(nxt) or _has_org(org_near)):
            explicit_job = True
        else:
            explicit_job = _explicit_job(line, nxt)
        if section == "edu" and not (intern and _has_org(context)):
            continue
        if section == "other" and not explicit_job:
            continue
        if not explicit_job and (
            section == "project" or line_is_project(line)
            or (line_is_project(prev) and not _has_org(entry) and not _has_org(nxt) and not _has_org(org_near))
        ):
            continue
        if section in ("work", "intern"):
            work_type = "实习" if section == "intern" or intern else "工作"
        else:
            if _looks_like_education(context) and not (intern and _has_org(context)):
                continue
            if not _has_org(context) and not (_has_title(line) and "项目" not in line and "大学" not in context):
                continue
            work_type = "实习" if intern else "工作"
        found.append({
            "type": work_type,
            "start": fmt_date(rng["sy"], rng["sm"]),
            "end": "至今" if rng["now"] else fmt_date(rng["ey"], rng["em"]),
            "start_k": (rng["sy"], rng["sm"] or 0),
            "end_k": (9999, 12) if rng["now"] else (rng["ey"], rng["em"] or 0),
            "org": _org_snippet(own_text) or _org_snippet(context),
            "line_org": _has_org(own_text) or bool(paired),
            "idx": start,
        })
    return [_public_work(item) for item in _sort_works(_dedup_works(_drop_contained(found)))]


def _drop_contained(items: List[Dict]) -> List[Dict]:
    kept = []
    for item in items:
        drop = False
        for other in items:
            if other is item:
                continue
            same = item["start_k"] == other["start_k"] and item["end_k"] == other["end_k"]
            inside = other["start_k"] <= item["start_k"] and item["end_k"] <= other["end_k"]
            # 包在某段工作里的纯日期，多半是项目周期；带单位或标成实习的另算
            if inside and not same and item["type"] != "实习" and not item["line_org"] and other["line_org"]:
                drop = True
                break
        if not drop:
            kept.append(item)
    return kept


def _dedup_works(items: List[Dict]) -> List[Dict]:
    groups: Dict[Tuple[str, str, str], List[Dict]] = {}
    for item in items:
        groups.setdefault((item["type"], item["start"], item["end"]), []).append(item)
    chosen: List[Dict] = []
    for group in groups.values():
        with_org = [item for item in group if item["org"]]
        pool = with_org or group
        seen = set()
        for item in pool:
            if item["org"] in seen:
                continue
            seen.add(item["org"])
            chosen.append(item)
    return chosen


def _sort_works(items: List[Dict]) -> List[Dict]:
    if WORK_SORT == "asc":
        return sorted(items, key=lambda item: (item["start_k"], item["end_k"], item["idx"]))
    return sorted(items, key=lambda item: item["idx"])


def _public_work(item: Dict) -> Dict:
    return {"type": item["type"], "start": item["start"], "end": item["end"]}


# ============================================================================
# 单文件
# ============================================================================

def analyze_resume(text: str, filename: str = "简历.pdf", ext: str = ".pdf",
                   tables=None, roster: Optional[Sequence[str]] = None,
                   font_lines: Optional[Sequence[VLine]] = None,
                   layout: str = "单栏") -> Dict:
    text = normalize_text(text)
    result = {
        "文件名": filename,
        "文件路径": "",
        "文件格式": ext.lstrip(".").upper(),
        "版式": layout,
        "姓名": "",
        "手机号码": "",
        "毕业时间": "",
        "works": [],
        "提取状态": "",
        "错误信息": "",
    }
    if not text.strip():
        result["提取状态"] = "失败"
        result["错误信息"] = "未提取到文本（可能是扫描件/图片型PDF，需要OCR）"
        return result
    smap = SectionMap(text)
    result["姓名"] = extract_name(text, filename, ext, tables, roster, font_lines) or ""
    result["手机号码"] = ",".join(extract_phones(text))
    result["毕业时间"] = "；".join(extract_graduation(text, smap))
    result["works"] = extract_work(text, smap)
    missing = []
    if not result["姓名"]:
        missing.append("未识别到姓名")
    if not result["手机号码"]:
        missing.append("未识别到手机号")
    if not result["毕业时间"]:
        missing.append("未识别到毕业时间")
    if not result["works"]:
        missing.append("未识别到工作经历")
    result["提取状态"] = "成功" if result["姓名"] and result["手机号码"] else "部分成功"
    result["错误信息"] = "；".join(missing)
    return result


def process_file(path: str, dump_dir: Optional[str] = None,
                 roster: Optional[List[str]] = None) -> Dict:
    filename = os.path.basename(path)
    ext = os.path.splitext(filename)[1].lower()
    result = {
        "文件名": filename, "文件路径": os.path.abspath(path), "文件格式": ext.lstrip(".").upper(),
        "版式": "", "姓名": "", "手机号码": "", "毕业时间": "", "works": [],
        "提取状态": "", "错误信息": "",
    }
    try:
        tables: List = []
        font_lines: List[VLine] = []
        layout = "单栏"
        if ext == ".pdf":
            raw, layout, font_lines = extract_pdf_document(path)
        else:
            raw, tables, layout = extract_docx_text(path)
        analyzed = analyze_resume(raw, filename, ext, tables, roster, font_lines, layout)
        analyzed["文件路径"] = os.path.abspath(path)
        analyzed["版式"] = layout
        if dump_dir:
            os.makedirs(dump_dir, exist_ok=True)
            with open(os.path.join(dump_dir, filename + ".txt"), "w", encoding="utf-8") as handle:
                handle.write(f"# 版式: {layout}\n")
                handle.write(normalize_text(raw))
        return analyzed
    except Exception as exc:
        result["提取状态"] = "失败"
        result["错误信息"] = f"{type(exc).__name__}: {exc}"
        logging.debug(traceback.format_exc())
        return result


# ============================================================================
# 收集与 Excel
# ============================================================================

def collect_files(input_path: str, recursive: bool = True) -> Tuple[List[str], List[Tuple[str, str]]]:
    files: List[str] = []
    excluded: List[Tuple[str, str]] = []

    def consider(path: str) -> None:
        name = os.path.basename(path)
        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED_EXTS or name.startswith("~$"):
            return
        hit = next((keyword for keyword in EXCLUDE_FILENAME_KEYWORDS if keyword in name), None)
        if hit:
            excluded.append((path, f"文件名包含“{hit}”"))
        else:
            files.append(path)

    if os.path.isfile(input_path):
        consider(input_path)
    elif recursive:
        for root, _, names in os.walk(input_path):
            for name in sorted(names):
                consider(os.path.join(root, name))
    else:
        for name in sorted(os.listdir(input_path)):
            path = os.path.join(input_path, name)
            if os.path.isfile(path):
                consider(path)
    return files, excluded


def clean_cell(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _style_sheet(sheet, widths: Dict[str, int], text_cols: List[int]) -> None:
    header_fill = PatternFill("solid", fgColor="4F81BD")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for index, cell in enumerate(sheet[1], start=1):
        sheet.column_dimensions[get_column_letter(index)].width = widths.get(cell.value, 14)
    for column in text_cols:
        for row in sheet.iter_rows(min_row=2, min_col=column, max_col=column):
            row[0].number_format = "@"


def write_excel(results: List[Dict], excluded: List[Tuple[str, str]], output_path: str) -> str:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "简历提取结果"
    sheet.append(OUTPUT_COLUMNS)
    for result in results:
        base = [result["文件名"], result["文件路径"], result["文件格式"], result.get("版式", ""),
                result["姓名"], result["手机号码"], result["毕业时间"]]
        if result["works"]:
            for index, work in enumerate(result["works"], start=1):
                sheet.append([clean_cell(value) for value in base + [
                    f"第{index}段", work["type"], work["start"], work["end"],
                    result["提取状态"], result["错误信息"]]])
        else:
            sheet.append([clean_cell(value) for value in base + [
                "", "", "", "", result["提取状态"], result["错误信息"]]])

    widths = {"文件名": 36, "文件路径": 50, "文件格式": 9, "版式": 12, "姓名": 10, "手机号码": 16,
              "毕业时间": 22, "工作序号": 9, "工作类型": 9, "工作开始时间": 14,
              "工作结束时间": 14, "提取状态": 10, "错误信息": 40}
    text_cols = [OUTPUT_COLUMNS.index(name) + 1
                 for name in ("手机号码", "毕业时间", "工作开始时间", "工作结束时间")]
    _style_sheet(sheet, widths, text_cols)

    summary = workbook.create_sheet("文件汇总")
    summary_cols = ["文件名", "文件路径", "版式", "姓名", "手机号码", "毕业时间",
                    "工作段数", "实习段数", "提取状态", "错误信息"]
    summary.append(summary_cols)
    for result in results:
        work_count = sum(1 for work in result["works"] if work["type"] == "工作")
        intern_count = sum(1 for work in result["works"] if work["type"] == "实习")
        summary.append([clean_cell(value) for value in [
            result["文件名"], result["文件路径"], result.get("版式", ""), result["姓名"],
            result["手机号码"], result["毕业时间"], work_count, intern_count,
            result["提取状态"], result["错误信息"]]])
    _style_sheet(summary, {**widths, "工作段数": 9, "实习段数": 9}, [5, 6])

    excluded_sheet = workbook.create_sheet("排除文件")
    excluded_sheet.append(["文件名", "文件路径", "排除原因"])
    for path, reason in excluded:
        excluded_sheet.append([
            clean_cell(os.path.basename(path)), clean_cell(os.path.abspath(path)), reason,
        ])
    _style_sheet(excluded_sheet, {"文件名": 36, "文件路径": 60, "排除原因": 24}, [])

    try:
        workbook.save(output_path)
    except PermissionError:
        root, ext = os.path.splitext(output_path)
        output_path = f"{root}_{datetime.now():%H%M%S}{ext}"
        logging.warning("目标文件被占用（可能已在 Excel 中打开），改存为：%s", output_path)
        workbook.save(output_path)
    return output_path


def resolve_output_path(output: Optional[str], input_path: str) -> str:
    default_name = f"简历提取结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    base_dir = input_path if os.path.isdir(input_path) else os.path.dirname(os.path.abspath(input_path))
    if not output:
        return os.path.join(base_dir, default_name)
    if os.path.isdir(output) or output.endswith(("\\", "/")):
        os.makedirs(output, exist_ok=True)
        return os.path.join(output, default_name)
    if not output.lower().endswith(".xlsx"):
        output += ".xlsx"
    parent = os.path.dirname(os.path.abspath(output))
    if parent:
        os.makedirs(parent, exist_ok=True)
    return output


def _find_roster(input_path: str) -> Optional[str]:
    base = input_path if os.path.isdir(input_path) else os.path.dirname(os.path.abspath(input_path))
    for name in ROSTER_FILENAMES:
        path = os.path.join(base, name)
        if os.path.isfile(path):
            return path
    return None


def _ask(prompt: str) -> str:
    return input(prompt).strip().strip('"').strip("'")


def main() -> None:
    parser = argparse.ArgumentParser(description="简历数据批量提取（PDF/DOCX → Excel）")
    parser.add_argument("-i", "--input", help="简历文件夹或单个简历文件路径")
    parser.add_argument("-n", "--names", help="人员信息表.xlsx 路径（有则优先用名单中的姓名）")
    parser.add_argument("-o", "--output", help="输出 Excel 路径或输出文件夹（默认保存在输入目录）")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子文件夹")
    parser.add_argument("--dump-text", metavar="DIR", help="把识别出的纯文本导出到该目录，便于排查")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出调试日志")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    interactive = not args.input
    input_path = args.input or _ask("请输入简历文件夹路径（或单个简历文件路径）：")
    while not input_path or not os.path.exists(input_path):
        if not interactive:
            logging.error("输入路径不存在：%s", input_path)
            sys.exit(1)
        input_path = _ask("路径不存在，请重新输入：")

    output = args.output
    if output is None and interactive:
        output = _ask("请输入输出 Excel 路径或文件夹（直接回车 = 保存到输入目录）：") or None
    output_path = resolve_output_path(output, input_path)

    names_path = args.names
    if names_path is None and interactive:
        names_path = _ask("请输入人员信息表.xlsx路径（直接回车 = 不使用名单 / 自动寻找）：") or None
    if not names_path:
        names_path = _find_roster(input_path)
        if names_path:
            logging.info("在输入目录找到人员信息表：%s", names_path)

    roster: List[str] = []
    if names_path:
        if not os.path.isfile(names_path):
            logging.warning("人员信息表不存在，将按原文规则提取姓名：%s", names_path)
        else:
            try:
                roster = load_roster_names(names_path)
                logging.info("已加载人员名单 %d 人：%s", len(roster), names_path)
            except Exception as exc:
                logging.warning("读取人员信息表失败（%s），将按原文规则提取姓名", exc)

    files, excluded = collect_files(input_path, recursive=not args.no_recursive)
    logging.info("找到简历 %d 份，排除非简历文件 %d 份", len(files), len(excluded))
    for path, reason in excluded:
        logging.info("  排除：%s（%s）", os.path.basename(path), reason)

    results = []
    for index, path in enumerate(files, start=1):
        result = process_file(path, args.dump_text, roster)
        results.append(result)
        logging.info(
            "[%d/%d] %s → %s | 版式=%s 姓名=%s 手机=%s 毕业=%s 工作段数=%d %s",
            index, len(files), os.path.basename(path), result["提取状态"], result["版式"] or "-",
            result["姓名"] or "-", result["手机号码"] or "-", result["毕业时间"] or "-",
            len(result["works"]), f"({result['错误信息']})" if result["错误信息"] else "",
        )

    saved = write_excel(results, excluded, output_path)
    ok = sum(result["提取状态"] == "成功" for result in results)
    partial = sum(result["提取状态"] == "部分成功" for result in results)
    failed = sum(result["提取状态"] == "失败" for result in results)
    logging.info("完成：成功 %d，部分成功 %d，失败 %d。结果已保存：%s", ok, partial, failed, saved)
    if interactive:
        input("按回车键退出...")


if __name__ == "__main__":
    main()