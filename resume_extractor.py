# -*- coding: utf-8 -*-
"""
简历数据批量提取工具（PDF / DOCX）

提取字段：姓名、手机号码、毕业时间、工作经历（多段拆分 + 工作/实习类型标注）
输出：Excel（每段工作经历一行，另附“文件汇总”“排除文件”两个工作表）

用法：
    python resume_extractor.py                       # 交互式输入路径
    python resume_extractor.py -i D:\\简历 -o D:\\结果.xlsx
    python resume_extractor.py -i D:\\简历 -n D:\\人员信息表.xlsx
    python resume_extractor.py -i D:\\简历 --dump-text D:\\简历文本   # 同时导出识别文本便于排查

依赖：pip install pymupdf pdfplumber python-docx openpyxl
"""

import argparse
import bisect
import logging
import os
import re
import sys
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import docx
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    docx = None

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

# 无 section 标题时，向前查看的上下文长度
CONTEXT_WINDOW = 200

# 人员名单表默认文件名（可在简历目录下自动寻找）
DEFAULT_ROSTER_NAME = "人员信息表.xlsx"
ROSTER_NAME_HEADERS = ["姓名", "名字", "人员姓名", "员工姓名", "name"]

OUTPUT_COLUMNS = [
    "文件名", "文件路径", "文件格式", "姓名", "手机号码", "毕业时间",
    "工作序号", "工作类型", "工作开始时间", "工作结束时间", "提取状态", "错误信息",
]

CURRENT_YEAR = datetime.now().year
MIN_YEAR = 1970

# ---------------------------------------------------------------- section 关键词
# 匹配顺序即优先级：例如“实习工作经历”应归为实习，而不是工作
SECTION_KEYWORDS: List[Tuple[str, List[str]]] = [
    ("intern", ["实习经", "实习背景", "实习工作", "实习履历", "internship",
                "社会实践", "实践经"]),
    ("project", ["项目经", "项目背景", "项目介绍", "核心项目", "项目案例",
                 "主要项目", "参与项目", "项目履历", "project"]),
    ("work", ["工作经", "职业经", "任职经", "工作履历", "从业经", "工作背景",
              "工作简历", "工作情况", "work experience", "professional experience",
              "work history", "employment", "career"]),
    ("edu", ["教育经", "教育背景", "教育情况", "教育信息", "学习经历", "education",
             "学历"]),
    ("other", ["自我评价", "个人评价", "自我介绍", "自我描述", "专业技能", "技能",
               "证书", "获奖", "荣誉", "培训经", "语言能力", "兴趣爱好", "基本信息",
               "个人信息", "求职意向", "个人优势", "个人总结", "校园经", "在校经",
               "附加信息", "skills", "certificat", "self-evaluation", "summary",
               "awards", "personal information"]),
]

# “学历：本科”属于基本信息字段，只有整行是这些词时才视为教育标题
EDU_EXACT_HEADINGS = {"学历", "学历背景", "学历信息", "学历情况", "学历教育"}

COMPANY_KEYWORDS = [
    "有限公司", "有限责任公司", "集团", "公司", "科技", "工程师", "股份", "银行",
    "研究院", "研究所", "事务所", "工作室", "中心", "co.,", "ltd", "inc.", "corp",
    "任职", "就职", "职位", "岗位", "经理", "主管", "专员", "开发", "运维",
]
PROJECT_KEYWORDS = [
    "项目经", "项目名称", "项目描述", "项目职责", "项目简介", "项目内容", "项目时间",
    "项目角色", "所属项目", "项目周期", "项目介绍", "核心项目", "project",
]
# 工作 section 内部仍需跳过的项目子条目（只看当前行）
PROJECT_LINE_KEYWORDS = [
    "项目名称", "项目时间", "项目周期", "项目经", "所属项目", "项目:", "项目：",
    "project:", "project name",
]
SCHOOL_KEYWORDS = [
    "大学", "学院", "学校", "中学", "高中", "中专", "技校", "职校", "分校",
    "university", "college", "school", "institute",
]
EDU_LINE_KEYWORDS = SCHOOL_KEYWORDS + [
    "本科", "硕士", "博士", "研究生", "专科", "大专", "学士", "学位", "专业", "就读",
]

# ---------------------------------------------------------------- 姓名校验
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
# “姓名：张三性别：男”这类粘连文本，在这些词处截断
NAME_STOP_WORDS = [
    "性别", "年龄", "出生", "民族", "电话", "手机", "籍贯", "学历", "邮箱", "政治",
    "婚姻", "身高", "体重", "联系", "住址", "现居", "求职", "工作", "专业", "毕业",
    "户口", "户籍", "身份", "居住", "意向", "照片",
]

# ============================================================================
# 日期正则
# ============================================================================

_Y = r"(?:19|20)\d{2}(?!\d)"
_M = r"(?:0?[1-9]|1[0-2])(?!\d)"
_D = r"(?:\s*[./\-]\s*(?:0?[1-9]|[12]\d|3[01])(?!\d)\s*日?|\s*(?:0?[1-9]|[12]\d|3[01])\s*日)?"


def _date_pattern(prefix: str) -> str:
    return (rf"(?<!\d)(?P<{prefix}y>{_Y})\s*"
            rf"(?:[年./\-]\s*(?P<{prefix}m>{_M})\s*月?{_D})?\s*年?")


_SEP = r"\s*(?:[-~至到]|\bto\b)+\s*"
_NOW = r"(?P<now>至今|今|现在|目前|now|present|till\s*now|to\s*date)"

RANGE_RE = re.compile(_date_pattern("s") + _SEP + rf"(?:{_date_pattern('e')}|{_NOW})", re.I)
DATE_RE = re.compile(_date_pattern("g"), re.I)

GRAD_LABEL_RE = re.compile(
    r"(?:毕业时间|毕业日期|毕业年份|毕业年月|毕业于|结业时间|结业日期|获得学位时间|"
    r"学位获得时间|学位授予时间|获得学位|毕业|结业)\s*[:：]?\s*"
)
GRAD_SUFFIX_RE = re.compile(_date_pattern("g") + r"\s*(?:毕业|结业|获得学位|获学位)")
STUDY_LABEL_RE = re.compile(
    r"(?:就读时间|在校时间|就读期间|在读时间|学习时间|就读)\s*[:：]?\s*"
)

PHONE_RE = re.compile(r"(?<!\d)(1[3-9]\d)[\s\-]?(\d{4})[\s\-]?(\d{4})(?!\d)")
LONG_DIGITS_RE = re.compile(r"(?<!\d)\d{15,}[Xx]?")
COUNTRY_CODE_RE = re.compile(r"(?<!\d)(?:\+|00)?\s*86[\s\-]*(?=1[3-9]\d)")

_TRANS = str.maketrans({
    "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": "-", "\u2014": "-",
    "\u2015": "-", "\u2212": "-", "\uff0d": "-", "\u2500": "-", "\u2501": "-",
    "\uff5e": "~", "\u301c": "~", "\uff0e": ".", "\uff0f": "/", "\u3000": " ",
    "\xa0": " ", "\r": "\n", "\u200b": "", "\ufeff": "",
    **{chr(0xFF10 + i): str(i) for i in range(10)},
})


# ============================================================================
# 文本抽取
# ============================================================================

def extract_pdf_text(path: str) -> str:
    errors = []
    if fitz is not None:
        try:
            with fitz.open(path) as doc:
                if doc.needs_pass:
                    raise RuntimeError("PDF 已加密，需要密码")
                text = "\n".join(page.get_text("text") for page in doc)
            if text.strip():
                return text
        except Exception as e:
            errors.append(f"PyMuPDF: {e}")
    if pdfplumber is not None:
        try:
            with pdfplumber.open(path) as pdf:
                text = "\n".join((p.extract_text() or "") for p in pdf.pages)
            if text.strip():
                return text
        except Exception as e:
            errors.append(f"pdfplumber: {e}")
    if fitz is None and pdfplumber is None:
        raise RuntimeError("未安装 PDF 解析库，请执行 pip install pymupdf pdfplumber")
    if errors:
        raise RuntimeError("；".join(errors))
    return ""


def _iter_block_items(document):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield Table(child, document)


def _textbox_texts(element, seen: set) -> List[str]:
    out = []
    for txbx in element.iter(qn("w:txbxContent")):
        for p in txbx.iter(qn("w:p")):
            t = "".join(node.text or "" for node in p.iter(qn("w:t"))).strip()
            if t and t not in seen:
                seen.add(t)
                out.append(t)
    return out


def _table_rows(table, all_tables: List[List[List[str]]]) -> List[List[str]]:
    """返回去重合并单元格后的行，同时把嵌套表格也收集进 all_tables。"""
    rows = []
    for row in table.rows:
        try:
            cells = row.cells
        except Exception:
            continue
        seen_tc = set()
        row_texts = []
        for cell in cells:
            if id(cell._tc) in seen_tc:
                continue
            seen_tc.add(id(cell._tc))
            row_texts.append(cell.text.strip())
            for nested in cell.tables:
                _table_rows(nested, all_tables)
        rows.append(row_texts)
    all_tables.append(rows)
    return rows


def extract_docx_text(path: str) -> Tuple[str, List[List[List[str]]]]:
    if docx is None:
        raise RuntimeError("未安装 python-docx，请执行 pip install python-docx")
    document = docx.Document(path)
    lines: List[str] = []
    tables: List[List[List[str]]] = []
    textbox_seen: set = set()

    for section in document.sections:
        try:
            for p in section.header.paragraphs:
                if p.text.strip():
                    lines.append(p.text)
        except Exception:
            pass

    for block in _iter_block_items(document):
        if isinstance(block, Paragraph):
            lines.append(block.text)
            lines.extend(_textbox_texts(block._p, textbox_seen))
        else:
            # 每个单元格单独成行，便于识别“工作经历”这类标题单元格
            for row in _table_rows(block, tables):
                for cell_text in row:
                    lines.extend(cell_text.split("\n"))
    return "\n".join(lines), tables


def normalize_text(text: str) -> str:
    text = text.translate(_TRANS)
    text = re.sub(r"[ \t\f\v]+", " ", text)
    lines = (line.strip() for line in text.split("\n"))
    return "\n".join(line for line in lines if line)


# ============================================================================
# section 识别
# ============================================================================

_BASIC_VALUE_RE = re.compile(
    r"^\s*(?:\d{1,2}(?:\.\d)?\s*(?:年|个月|years?)(?:以上|以下|左右)?|应届\S{0,3}|无|暂无|"
    r"[一二三四五六七八九十]+年(?:以上|以下|左右)?)\s*$", re.I)

_HEADING_PREFIX_RE = re.compile(
    r"^[\s\d一二三四五六七八九十、.:：#*■□●◆◇▶►•·\-_|【】\[\]()（）<>《》]+"
)


def detect_heading(line: str) -> Optional[str]:
    clean = _HEADING_PREFIX_RE.sub("", line).strip()
    if not clean:
        return None
    parts = re.split(r"[:：\t|｜]", clean, maxsplit=1)
    head = re.sub(r"\s+", "", parts[0]).rstrip("】]）)>》")
    head_lower = re.sub(r"\s+", " ", parts[0]).strip().lower()
    rest = parts[1].strip() if len(parts) > 1 else ""

    ascii_ratio = sum(c.isascii() for c in head) / max(len(head), 1)
    if len(head) > (30 if ascii_ratio > 0.6 else 15):
        return None
    # “工作经验：3年”“工作经验：应届”是基本信息字段，不是标题
    if rest and re.match(r"^(?:\d{1,2}(?:\.\d)?\s*(?:年|个月|years?)|应届|无|暂无|[一二三四五六七八九十]+年)", rest, re.I):
        return None

    for sec_type, keywords in SECTION_KEYWORDS:
        for kw in keywords:
            target = head_lower if kw.isascii() else head
            idx = target.find(kw)
            if idx < 0 or idx > 4:
                continue
            if kw == "学历" and head not in EDU_EXACT_HEADINGS:
                continue
            return sec_type
    return None


class SectionMap:
    """把字符偏移映射到所在行及其所属 section。"""

    def __init__(self, text: str):
        self.text = text
        self.lines = text.split("\n")
        self.line_starts = []
        pos = 0
        for line in self.lines:
            self.line_starts.append(pos)
            pos += len(line) + 1
        self.sections: List[Optional[str]] = []
        current = None
        for i, line in enumerate(self.lines):
            heading = detect_heading(line)
            next_line = self.lines[i + 1] if i + 1 < len(self.lines) else ""
            # 表格中“工作经验 | 5年”拆成两行时，这仍是基本信息字段
            if heading and _BASIC_VALUE_RE.match(next_line):
                heading = None
            if heading:
                current = heading
            self.sections.append(current)

    def line_index(self, pos: int) -> int:
        return max(bisect.bisect_right(self.line_starts, pos) - 1, 0)

    def section_at(self, pos: int) -> Optional[str]:
        return self.sections[self.line_index(pos)]

    def line_at(self, pos: int) -> str:
        return self.lines[self.line_index(pos)]

    def prev_line(self, pos: int) -> str:
        idx = self.line_index(pos)
        return self.lines[idx - 1] if idx > 0 else ""


# ============================================================================
# 日期工具
# ============================================================================

def _year_ok(year: int, max_offset: int = 1) -> bool:
    return MIN_YEAR <= year <= CURRENT_YEAR + max_offset


def fmt_date(year: int, month: Optional[int]) -> str:
    return f"{year}.{month:02d}" if month else f"{year}"


def date_key(s: str) -> Tuple[int, int]:
    if s in ("至今",):
        return (9999, 12)
    parts = s.split(".")
    return (int(parts[0]), int(parts[1]) if len(parts) > 1 else 0)


def parse_range(m: re.Match) -> Optional[Dict]:
    sy = int(m.group("sy"))
    sm = int(m.group("sm")) if m.group("sm") else None
    if m.group("now"):
        if not _year_ok(sy):
            return None
        return {"sy": sy, "sm": sm, "ey": None, "em": None, "now": True}
    ey = int(m.group("ey"))
    em = int(m.group("em")) if m.group("em") else None
    if not (_year_ok(sy, 6) and _year_ok(ey, 6)):
        return None
    if (ey, em or 12) < (sy, sm or 1):
        return None
    return {"sy": sy, "sm": sm, "ey": ey, "em": em, "now": False}


# ============================================================================
# 字段提取
# ============================================================================

def is_valid_name(name: Optional[str]) -> bool:
    if not name:
        return False
    name = name.strip()
    if not (2 <= len(name) <= 6):
        return False
    if not re.search(r"[\u4e00-\u9fa5]", name):
        return False
    if re.search(r"\d", name):
        return False
    if name in NAME_TITLE_WORDS:
        return False
    return not any(bad in name for bad in NAME_BAD_SUBSTRINGS)


def _cut_name(cand: str) -> str:
    for w in NAME_STOP_WORDS:
        idx = cand.find(w)
        if idx > 0:
            cand = cand[:idx]
    return cand.strip("·").strip()


def name_from_label(text: str) -> Optional[str]:
    for m in re.finditer(r"姓\s*名\s*[:：]?\s*([\u4e00-\u9fa5·]{2,6})", text):
        cand = _cut_name(m.group(1))
        if is_valid_name(cand):
            return cand
    # PDF 中常见的“姓名：张 三”
    for m in re.finditer(r"姓\s*名\s*[:：]?\s*([\u4e00-\u9fa5](?: [\u4e00-\u9fa5]){1,3})(?![\u4e00-\u9fa5])", text):
        cand = m.group(1).replace(" ", "")
        if is_valid_name(cand):
            return cand
    return None


def name_from_docx_tables(tables: List[List[List[str]]]) -> Optional[str]:
    label_set = {"姓名", "名字", "姓名name", "name"}
    for rows in tables:
        for r, row in enumerate(rows):
            for c, cell in enumerate(row):
                key = re.sub(r"[\s:：]", "", cell).lower()
                if key not in label_set:
                    continue
                # 右侧单元格
                for right in row[c + 1:]:
                    val = re.sub(r"\s+", "", right)
                    if val and val.lower() not in label_set:
                        if is_valid_name(_cut_name(val)):
                            return _cut_name(val)
                        break
                # 下方单元格（表头式布局）
                if r + 1 < len(rows) and c < len(rows[r + 1]):
                    val = re.sub(r"\s+", "", rows[r + 1][c])
                    if is_valid_name(val):
                        return val
    return None


def _filename_tokens(filename: str) -> List[str]:
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r"(?:的?个人简历|的?简历|求职|应聘|resume|cv)", " ", stem, flags=re.I)
    return [t for t in re.split(r"[_\-\s()（）\[\]【】、,，.+#&@=~]+", stem) if t]


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def _pick_earliest_in_text(cands: List[str], compact_text: str) -> Optional[str]:
    """多个候选同时出现在正文时，取正文中最先出现的；起点相同则取更长的。"""
    ranked = []
    for cand in cands:
        pos = compact_text.find(cand)
        if pos >= 0:
            ranked.append((pos, -len(cand), cand))
    if not ranked:
        return None
    ranked.sort()
    return ranked[0][2]


def name_from_filename(filename: str, text: str) -> Optional[str]:
    compact_text = _compact(text)
    seen = set()
    both: List[str] = []
    for tok in _filename_tokens(filename):
        for cand in re.findall(r"[\u4e00-\u9fa5·]{2,6}", tok):
            if cand in seen:
                continue
            seen.add(cand)
            if is_valid_name(cand) and cand in compact_text:
                both.append(cand)
    return _pick_earliest_in_text(both, compact_text)


def load_roster_names(path: str) -> List[str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if not header:
            return []
        col = None
        for i, h in enumerate(header):
            key = re.sub(r"\s+", "", str(h or "")).lower()
            if any(n.lower() == key or n.lower() in key for n in ROSTER_NAME_HEADERS):
                col = i
                break
        if col is None:
            col = 0
            logging.warning("人员信息表未找到“姓名”列，改用第一列：%s", header[0])
        names, seen = [], set()
        for row in rows:
            if not row or col >= len(row) or row[col] is None:
                continue
            name = re.sub(r"\s+", "", str(row[col]).strip())
            if name and name not in seen:
                seen.add(name)
                names.append(name)
        return names
    finally:
        wb.close()


def name_from_roster(text: str, filename: str, roster: List[str]) -> Optional[str]:
    """优先用人员名单中的姓名。同时出现在文件名和正文中的优先，多个时取正文最先出现的。"""
    if not roster:
        return None
    compact_text = _compact(text)
    compact_fn = _compact(os.path.splitext(filename)[0])
    in_both, in_text, in_file = [], [], []
    for name in roster:
        if not name:
            continue
        t_pos = compact_text.find(name)
        f_pos = compact_fn.find(name)
        if t_pos >= 0 and f_pos >= 0:
            in_both.append(name)
        elif t_pos >= 0:
            in_text.append(name)
        elif f_pos >= 0:
            in_file.append(name)
    picked = _pick_earliest_in_text(in_both, compact_text)
    if picked:
        return picked
    picked = _pick_earliest_in_text(in_text, compact_text)
    if picked:
        return picked
    if in_file:
        in_file.sort(key=lambda n: (-len(n), compact_fn.find(n)))
        return in_file[0]
    return None


def name_from_first_lines(text: str, max_lines: int = 3) -> Optional[str]:
    for line in text.split("\n")[:max_lines]:
        line = line.strip()
        if re.fullmatch(r"[\u4e00-\u9fa5](?: [\u4e00-\u9fa5]){1,3}", line):
            cand = line.replace(" ", "")
            if is_valid_name(cand):
                return cand
        first = re.split(r"[\s|｜/,，:：]+", line)[0]
        if re.fullmatch(r"[\u4e00-\u9fa5·]{2,6}", first) and is_valid_name(first):
            return first
    return None


def extract_name(text: str, filename: str, ext: str, tables,
                 roster: Optional[List[str]] = None) -> Optional[str]:
    if roster:
        name = name_from_roster(text, filename, roster)
        if name:
            return name
    name = name_from_label(text)
    if not name and ext == ".docx":
        name = name_from_docx_tables(tables)
    if not name:
        name = name_from_filename(filename, text)
    if not name and ext == ".pdf":
        name = name_from_first_lines(text)
    return name


def extract_phones(text: str) -> List[str]:
    long_digit_runs = LONG_DIGITS_RE.findall(text)
    cleaned = COUNTRY_CODE_RE.sub(" ", text)
    phones: List[str] = []
    for m in PHONE_RE.finditer(cleaned):
        phone = "".join(m.groups())
        if any(phone in run for run in long_digit_runs):
            continue  # 身份证号 / 长数字片段
        if phone not in phones:
            phones.append(phone)
    return phones


def extract_graduation(text: str, smap: SectionMap) -> List[str]:
    found = set()

    def add_end(rng: Dict):
        if rng and not rng["now"] and _year_ok(rng["ey"], 6):
            found.add(fmt_date(rng["ey"], rng["em"] or 6))

    # ① 毕业/结业/获得学位 标签
    for m in GRAD_LABEL_RE.finditer(text):
        rm = RANGE_RE.match(text, m.end())
        if rm:
            add_end(parse_range(rm))
            continue
        dm = DATE_RE.match(text, m.end())
        if dm:
            y = int(dm.group("gy"))
            if _year_ok(y, 6):
                found.add(fmt_date(y, int(dm.group("gm")) if dm.group("gm") else 6))
    for m in GRAD_SUFFIX_RE.finditer(text):
        y = int(m.group("gy"))
        if _year_ok(y, 6):
            found.add(fmt_date(y, int(m.group("gm")) if m.group("gm") else 6))

    # ② 就读时间：2018.09-2024.06
    for m in STUDY_LABEL_RE.finditer(text):
        rm = RANGE_RE.match(text, m.end())
        if rm:
            add_end(parse_range(rm))

    # ③④⑤⑥ 教育 section 内的时间段 / 学校名称同行或上一行的时间段（含中文年月格式）
    for m in RANGE_RE.finditer(text):
        rng = parse_range(m)
        if not rng or rng["now"]:
            continue
        section = smap.section_at(m.start())
        line_lower = smap.line_at(m.start()).lower()
        prev_lower = smap.prev_line(m.start()).lower()
        in_school_line = any(k in line_lower for k in SCHOOL_KEYWORDS)
        after_school = (any(k in prev_lower for k in SCHOOL_KEYWORDS)
                        and not any(k in line_lower for k in COMPANY_KEYWORDS)
                        and section not in ("work", "intern", "project"))
        if section == "edu" or in_school_line or after_school:
            add_end(rng)

    return sorted(found, key=date_key)


def _nearest_type(text: str, pos: int) -> str:
    before = text[:pos]
    line_start = before.rfind("\n") + 1
    line_end = text.find("\n", pos)
    line = text[line_start: line_end if line_end >= 0 else len(text)]
    if "实习" in line or "intern" in line.lower():
        return "实习"
    lower = before.lower()
    intern_pos = max(lower.rfind(k) for k in ("实习", "internship", "社会实践"))
    work_pos = max(lower.rfind(k) for k in ("工作", "任职", "就职", "work experience", "employment"))
    return "实习" if intern_pos > work_pos else "工作"


def extract_work(text: str, smap: SectionMap) -> List[Dict]:
    works: List[Dict] = []
    seen = set()
    for m in RANGE_RE.finditer(text):
        rng = parse_range(m)
        if not rng:
            continue
        section = smap.section_at(m.start())
        line = smap.line_at(m.start())
        line_lower = line.lower()

        if section in ("edu", "project", "other"):
            continue
        if section in ("work", "intern"):
            if any(k in line_lower for k in PROJECT_LINE_KEYWORDS):
                continue
            if STUDY_LABEL_RE.search(line) or STUDY_LABEL_RE.search(smap.prev_line(m.start())):
                continue
            work_type = "实习" if section == "intern" else "工作"
        else:
            context = (text[max(0, m.start() - CONTEXT_WINDOW): m.start()] + line).lower()
            if not any(k in context for k in COMPANY_KEYWORDS):
                continue
            if any(k in context for k in PROJECT_KEYWORDS):
                continue
            if any(k in line_lower for k in EDU_LINE_KEYWORDS):
                continue
            work_type = _nearest_type(text, m.start())

        start = fmt_date(rng["sy"], rng["sm"])
        end = "至今" if rng["now"] else fmt_date(rng["ey"], rng["em"])
        if (start, end) in seen:
            continue
        seen.add((start, end))
        works.append({"type": work_type, "start": start, "end": end})

    if WORK_SORT == "asc":
        works.sort(key=lambda w: (date_key(w["start"]), date_key(w["end"])))
    return works


# ============================================================================
# 单文件处理
# ============================================================================

def process_file(path: str, dump_dir: Optional[str] = None,
                 roster: Optional[List[str]] = None) -> Dict:
    filename = os.path.basename(path)
    ext = os.path.splitext(filename)[1].lower()
    result = {
        "文件名": filename, "文件路径": os.path.abspath(path), "文件格式": ext.lstrip(".").upper(),
        "姓名": "", "手机号码": "", "毕业时间": "", "works": [],
        "提取状态": "", "错误信息": "",
    }
    try:
        tables: List = []
        if ext == ".pdf":
            raw = extract_pdf_text(path)
        else:
            raw, tables = extract_docx_text(path)
        text = normalize_text(raw)

        if dump_dir:
            os.makedirs(dump_dir, exist_ok=True)
            with open(os.path.join(dump_dir, filename + ".txt"), "w", encoding="utf-8") as f:
                f.write(text)

        if not text.strip():
            result["提取状态"] = "失败"
            result["错误信息"] = "未提取到文本（可能是扫描件/图片型PDF，需要OCR）"
            return result

        smap = SectionMap(text)
        result["姓名"] = extract_name(text, filename, ext, tables, roster) or ""
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
    except Exception as e:
        result["提取状态"] = "失败"
        result["错误信息"] = f"{type(e).__name__}: {e}"
        logging.debug(traceback.format_exc())
    return result


# ============================================================================
# 文件收集与输出
# ============================================================================

def collect_files(input_path: str, recursive: bool = True) -> Tuple[List[str], List[Tuple[str, str]]]:
    files, excluded = [], []

    def consider(path: str):
        name = os.path.basename(path)
        ext = os.path.splitext(name)[1].lower()
        if ext not in SUPPORTED_EXTS or name.startswith("~$"):
            return
        hit = next((k for k in EXCLUDE_FILENAME_KEYWORDS if k in name), None)
        if hit:
            excluded.append((path, f"文件名包含“{hit}”"))
        else:
            files.append(path)

    if os.path.isfile(input_path):
        consider(input_path)
    elif recursive:
        for root, _, names in os.walk(input_path):
            for n in sorted(names):
                consider(os.path.join(root, n))
    else:
        for n in sorted(os.listdir(input_path)):
            p = os.path.join(input_path, n)
            if os.path.isfile(p):
                consider(p)
    return files, excluded


def clean_cell(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    return value


def _style_sheet(ws, widths: Dict[str, int], text_cols: List[int]):
    header_fill = PatternFill("solid", fgColor="4F81BD")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, cell in enumerate(ws[1], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(cell.value, 14)
    for col in text_cols:
        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            row[0].number_format = "@"


def write_excel(results: List[Dict], excluded: List[Tuple[str, str]], output_path: str) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "简历提取结果"
    ws.append(OUTPUT_COLUMNS)
    for r in results:
        base = [r["文件名"], r["文件路径"], r["文件格式"], r["姓名"], r["手机号码"], r["毕业时间"]]
        if r["works"]:
            for i, w in enumerate(r["works"], start=1):
                ws.append([clean_cell(v) for v in base + [
                    f"第{i}段", w["type"], w["start"], w["end"], r["提取状态"], r["错误信息"]]])
        else:
            ws.append([clean_cell(v) for v in base + ["", "", "", "", r["提取状态"], r["错误信息"]]])

    widths = {"文件名": 36, "文件路径": 50, "文件格式": 9, "姓名": 10, "手机号码": 16,
              "毕业时间": 22, "工作序号": 9, "工作类型": 9, "工作开始时间": 13,
              "工作结束时间": 13, "提取状态": 10, "错误信息": 40}
    text_col_idx = [OUTPUT_COLUMNS.index(c) + 1 for c in ("手机号码", "毕业时间", "工作开始时间", "工作结束时间")]
    _style_sheet(ws, widths, text_col_idx)

    ws2 = wb.create_sheet("文件汇总")
    summary_cols = ["文件名", "文件路径", "姓名", "手机号码", "毕业时间", "工作段数", "实习段数", "提取状态", "错误信息"]
    ws2.append(summary_cols)
    for r in results:
        work_n = sum(1 for w in r["works"] if w["type"] == "工作")
        intern_n = sum(1 for w in r["works"] if w["type"] == "实习")
        ws2.append([clean_cell(v) for v in [r["文件名"], r["文件路径"], r["姓名"], r["手机号码"],
                                            r["毕业时间"], work_n, intern_n, r["提取状态"], r["错误信息"]]])
    _style_sheet(ws2, {**widths, "工作段数": 9, "实习段数": 9}, [4, 5])

    ws3 = wb.create_sheet("排除文件")
    ws3.append(["文件名", "文件路径", "排除原因"])
    for path, reason in excluded:
        ws3.append([clean_cell(os.path.basename(path)), clean_cell(os.path.abspath(path)), reason])
    _style_sheet(ws3, {"文件名": 36, "文件路径": 60, "排除原因": 24}, [])

    try:
        wb.save(output_path)
    except PermissionError:
        root, ext = os.path.splitext(output_path)
        output_path = f"{root}_{datetime.now():%H%M%S}{ext}"
        logging.warning("目标文件被占用（可能已在 Excel 中打开），改存为：%s", output_path)
        wb.save(output_path)
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
    os.makedirs(parent, exist_ok=True)
    return output


def _ask(prompt: str) -> str:
    return input(prompt).strip().strip('"').strip("'")


def main():
    parser = argparse.ArgumentParser(description="简历数据批量提取（PDF/DOCX → Excel）")
    parser.add_argument("-i", "--input", help="简历文件夹或单个简历文件路径")
    parser.add_argument("-n", "--names", help="人员信息表.xlsx 路径（有则优先用名单中的姓名）")
    parser.add_argument("-o", "--output", help="输出 Excel 路径或输出文件夹（默认保存在输入目录）")
    parser.add_argument("--no-recursive", action="store_true", help="不递归子文件夹")
    parser.add_argument("--dump-text", metavar="DIR", help="把识别出的纯文本导出到该目录，便于排查")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出调试日志")
    args = parser.parse_args()

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")

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
        auto = os.path.join(input_path if os.path.isdir(input_path)
                            else os.path.dirname(os.path.abspath(input_path)), DEFAULT_ROSTER_NAME)
        if os.path.isfile(auto):
            names_path = auto
            logging.info("在输入目录找到人员信息表：%s", names_path)

    roster: List[str] = []
    if names_path:
        if not os.path.isfile(names_path):
            logging.warning("人员信息表不存在，将按原文规则提取姓名：%s", names_path)
        else:
            try:
                roster = load_roster_names(names_path)
                logging.info("已加载人员名单 %d 人：%s", len(roster), names_path)
            except Exception as e:
                logging.warning("读取人员信息表失败（%s），将按原文规则提取姓名", e)

    files, excluded = collect_files(input_path, recursive=not args.no_recursive)
    logging.info("找到简历 %d 份，排除非简历文件 %d 份", len(files), len(excluded))
    for path, reason in excluded:
        logging.info("  排除：%s（%s）", os.path.basename(path), reason)

    results = []
    for idx, path in enumerate(files, start=1):
        r = process_file(path, args.dump_text, roster)
        results.append(r)
        logging.info("[%d/%d] %s → %s | 姓名=%s 手机=%s 毕业=%s 工作段数=%d %s",
                     idx, len(files), os.path.basename(path), r["提取状态"], r["姓名"] or "-",
                     r["手机号码"] or "-", r["毕业时间"] or "-", len(r["works"]),
                     f"({r['错误信息']})" if r["错误信息"] else "")

    saved = write_excel(results, excluded, output_path)
    ok = sum(r["提取状态"] == "成功" for r in results)
    partial = sum(r["提取状态"] == "部分成功" for r in results)
    failed = sum(r["提取状态"] == "失败" for r in results)
    logging.info("完成：成功 %d，部分成功 %d，失败 %d。结果已保存：%s", ok, partial, failed, saved)

    if interactive:
        input("按回车键退出...")


if __name__ == "__main__":
    main()
