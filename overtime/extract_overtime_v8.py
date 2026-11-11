#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
加班邮件提取 V8。

从非结构化邮件正文、回复链、邮件名或 OCR 文本中提取加班申请与审批结果，并导出 Excel。
支持 .eml 和 Outlook .msg。读取 .msg 需要 extract-msg。
仓库里本版文件名是 overtime/extract_overtime_v8.py。
v1、v2、v3、v6、v7 不覆盖。
「加班审批」只保留姓名和日期都有的记录；没识别到、没提取到，或姓名、日期缺一的邮件记在「未完整提取」。

规则摘要
--------
1. 回复链按「最新在上」切分。只有位于申请上方的明确同意
   （同意 / 同意加班 / +1 / 辛苦了 等）才记为 APPROVED；
   明确拒绝记为 REJECTED；否则 PENDING。
2. 加班人必须出现在「人员名单.xlsx」中。正文、邮件名（主题和文件名）、
   表格附件和图片 OCR 都会对名单。不再用宽泛正则乱猜姓名。
3. 只提取与「加班」相关的日期，忽略发送时间、写道时间等噪声日期。
   正文没有日期时，再看批复段、文件名和附件。
   「日」和「号」同等看待，月日之间认「-」，例如 11-22号、11月22号。
   Content-ID、cid: 引用里的连字符数字不当成日期。
4. 邮件内嵌图片、HTML 里的图片会做 OCR；xlsx/xls/csv/docx 附件会读出文字，
   识别出的姓名和日期与正文一起提取。
   Content-Type 写成 application/octet-stream 的 PNG/JPEG 也会识别。
5. 「回复同意加班的人」优先取批复段发件人；没有则取申请邮件收件人。
6. Excel 列：加班人员名称、加班日期、是否同意加班、回复同意加班的人、邮件所在路径、文件名。

用法
----
    py -3 extract_overtime_v8.py --check
    py -3 extract_overtime_v8.py --dir 邮件文件夹 --roster 人员名单.xlsx -o 结果.xlsx
    py -3 extract_overtime_v8.py       # 交互式：邮件目录 + 人员名单 + 导出路径
"""

from __future__ import annotations

__version__ = "8"

import base64
import hashlib
import html
import io
import json
import os
import queue
import re
import sys
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from itertools import product
from pathlib import Path
from typing import Any, Iterable

# 单封邮件参与解析的最大字符数（回复链头部已含审批与申请）
_MAX_PARSE_CHARS = 120_000
# 快速探测是否含「加班」时读取的字节数
_PEEK_BYTES = 24_000
_DEFAULT_WORKERS = min(32, max(4, (os.cpu_count() or 4) * 4))


# ---------------------------------------------------------------------------
# 词表
# ---------------------------------------------------------------------------

# 强同意：出现在批复区即可。不要用「同意.*?加班」，会误伤「请同意…加班」
_STRONG_APPROVE = re.compile(
    r"(?<!请)(?<!不)同意加班|(?<!请)(?<!不)同意申请|(?<!请)(?<!不)同意审批|"
    r"审批通过|审核通过|批准加班|可以加班|准予加班|已批准|已审批|准予"
)
# 「同意」单独出现。排除 不同意 / 请同意 / 未同意 / 申请同意 / 待同意
_BARE_APPROVE = re.compile(r"(?<!不)(?<!请)(?<!未)(?<!待)(?<!申请)同意")
_SOFT_APPROVE = re.compile(
    r"(?:^|[\s，,。！!])(?:好的|可以|没问题|辛苦了|辛苦啦|批准|已阅|\+1|OK|ok)(?:[，,。！!\s]|$)"
)
_REJECT = re.compile(r"不予批准|不予同意|不同意|不批准|驳回|拒绝")
# 短批复行：允许「同意，请知悉」「辛苦了！」这类
_SHORT_APPROVE_LINE = re.compile(
    r"^(?=.*(?:同意|辛苦了|辛苦啦|没问题|批准|好的|可以|\+1|OK|ok|已批|已阅))"
    r".{0,40}$"
)

_HEADER_LINE = re.compile(
    r"^\s*(?:"
    r"发件人|发送人|发件时间|发送时间|收件人|抄送|密送|主题|"
    r"From|Sender|To|Cc|CC|Bcc|BCC|Sent|Date|Subject|Reply-To"
    r")\s*[:：]",
    re.IGNORECASE,
)
_SUBJECT_LINE = re.compile(r"^(?:主题|Subject)\s*[:：]\s*(.+)$", re.IGNORECASE)
_PARTICIPANT_HEADER = re.compile(
    r"^\s*(发件人|发送人|收件人|抄送|密送|From|Sender|To|Cc|CC|Bcc|BCC)\s*[:：]\s*(.+)$",
    re.IGNORECASE,
)
_FROM_LABELS = {"发件人", "发送人", "from", "sender"}
_TO_LABELS = {"收件人", "to"}
_SENT_TIME = re.compile(
    r"(?:发送时间|发件时间|Sent|Date)\s*[:：]\s*"
    r"(20\d{2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*(\d{1,2})",
    re.IGNORECASE,
)
_STATUS_LABEL = {
    "APPROVED": "是",
    "REJECTED": "否",
    "PENDING": "待审批",
}
_EXCEL_COLUMNS = (
    "加班人员名称",
    "加班日期",
    "是否同意加班",
    "回复同意加班的人",
    "邮件所在路径",
    "文件名",
)
_GAP_SHEET = "未完整提取"
_GAP_COLUMNS = (
    "文件链接",
    "原因",
    "加班人员名称",
    "加班日期",
    "文件名",
)
_TEXT_LIKE_EXTENSIONS = {
    ".eml",
    ".msg",
    ".txt",
    ".html",
    ".htm",
    ".mime",
    ".email",
}
_SKIP_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".xlsm",
    ".csv",
    ".doc",
    ".docx",
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".zip",
    ".rar",
    ".7z",
    ".exe",
    ".dll",
    ".pyc",
    ".py",
}
_NAME_KEYWORDS = ("加班", "服务时长")

_CHAIN_SPLIT = re.compile(
    r"(?m)^[ \t]*(?:"
    r"-{3,}[^\n]{0,80}(?:原始邮件|转发邮件|Original Message|Forwarded message)?[^\n]*|"
    r"_{5,}[ \t]*|"
    r"在[ \t].{0,80}?写道[：:][ \t]*|"
    r"On .{0,180}?wrote:[ \t]*"
    r")$",
    re.IGNORECASE,
)
_META_DATE_LINE = re.compile(
    r"^[ \t]*(?:发送时间|发件时间|Sent|Date)\s*[:：].*$|"
    r"^[ \t]*在[ \t].{0,80}?写道[：:].*$|"
    r"^[ \t]*On .{0,180}?wrote:[ \t]*$",
    re.IGNORECASE | re.MULTILINE,
)

_TITLE_NAME = re.compile(
    r"^[\u4e00-\u9fff]{1,3}(?:老师|经理|主任|总监|主管|领导|部长|处长|科长|总工|总)$"
)
_SALUTATION = re.compile(
    r"(?m)(?:^|[，,。\s])([\u4e00-\u9fff]{1,3}(?:老师|经理|主任|总监|主管|领导|部长|处长|科长|总))"
    r"\s*(?:[：:，,]|好|您好)"
)

_NAME = r"[\u4e00-\u9fff]{2,8}"
_NAME_LIST = rf"{_NAME}(?:\s*[、,，/／和与及]\s*{_NAME})*"
# 姓名不能从汉字中间切开，避免「安排麻孟博」再匹配出「排麻孟博」
_NAME_LEFT = r"(?<![\u4e00-\u9fff])"
_COUNT = r"(?:[0-9０-９一二两三四五六七八九十]+\s*人|等(?:人)?)"

_STOPWORDS = {
    "加班", "申请", "审批", "安排", "同意", "批准", "驳回", "拒绝", "辛苦",
    "可以", "需要", "进行", "参加", "负责", "确认", "收到", "回复", "邮件",
    "各位", "大家", "同事", "同学", "领导", "老师", "部门", "公司", "团队",
    "全体", "相关", "人员", "名单", "日期", "时间", "谢谢", "您好", "你好",
    "工作", "内容", "如下", "详见", "具体", "分别", "其中", "因为", "由于",
    "为了", "本次", "此次", "当天", "今日", "明日", "后天", "本人", "我们",
    "你们", "他们", "她们", "自己", "其他", "另外", "同时", "以及", "或者",
    "如果", "但是", "所以", "这个", "那个", "什么", "怎么", "已经", "没有",
    "不是", "不能", "不要", "麻烦", "帮忙", "告知", "周知", "知悉", "查收",
    "项目", "上线", "发布", "附件", "原始", "发件", "收件", "抄送", "主题",
    "周六", "周日", "周一", "周二", "周三", "周四", "周五", "周末", "本周",
    "下周", "这周", "今天", "明天", "后天", "一天", "两天", "三天", "半天",
    "全天", "一日", "三人", "两人", "四人", "五人", "六人", "十人", "多人",
    "等人", "数人", "两日", "三日", "工时", "小时", "分钟", "左右", "以上",
    "以下", "之间", "之前", "之后", "截止", "请您", "请帮", "谢谢", "感谢",
    "辛苦了", "没问题", "早上好", "下午好", "晚上好",
}

_WEEKDAY = {"一": 0, "二": 1, "三": 2, "四": 3, "五": 4, "六": 5, "日": 6, "天": 6}

_MAX_RANGE_DAYS = 62

# 日期相关正则预编译，避免每封邮件重复 compile
_DATE_CROSS = re.compile(
    r"(?:(20\d{2})\s*[年./\-]\s*)?(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*[日号]?"
    r"\s*(?:至|到|~|～|—|–|－|-)\s*"
    r"(?:(20\d{2})\s*[年./\-]\s*)?(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*[日号]?"
)
_DATE_SAME_MONTH = re.compile(
    r"(?:(20\d{2})\s*年\s*)?(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?\s*[-~～—–至到]\s*(\d{1,2})\s*[日号]"
)
_DATE_LISTED = re.compile(
    r"(?:(20\d{2})\s*年\s*)?(\d{1,2})\s*月\s*"
    r"(\d{1,2}\s*[日号](?:\s*[、,，和与及]\s*\d{1,2}\s*[日号])+)"
)
_DATE_NUM_LISTED = re.compile(
    r"(20\d{2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*"
    r"(\d{1,2})(?!\d)"
    r"((?:\s*[、,，和与及]\s*(?!\d{4})\d{1,2}(?!\d))+)"
)
_DATE_FULL = re.compile(r"(20\d{2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*[日号]?")
_DATE_MONTH_DAY = re.compile(r"(?<!\d)(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]")
# 1.24 / 10.18日 / 9.21 / 10.18号。排除 2025.09.13 内部、IP 和 1.875 这类长小数。
_DATE_DOT = re.compile(
    r"(?<![\d.])(\d{1,2})\s*[.．]\s*(\d{1,2})\s*([日号])?(?![\d.])"
)
_DATE_DOT_RANGE = re.compile(
    r"(?<![\d.])(\d{1,2})\s*[.．]\s*(\d{1,2})\s*[日号]?\s*(?:至|到|~|～|—|–|－|-)\s*"
    r"(\d{1,2})\s*[.．]\s*(\d{1,2})\s*[日号]?(?![\d.])"
)
# 11-22 / 11-22号 / 11-22日。不从 2025-11-22 中间再切一刀，
# 两侧也不能连着英文字母或数字，避免 Content-ID 里的 F6-16A 被当成 6月16日。
_DATE_HYPHEN = re.compile(
    r"(?<![A-Za-z0-9.\-])(\d{1,2})\s*-\s*(\d{1,2})\s*([日号])?(?![A-Za-z0-9.\-])"
)
# cid:image@01DC.F6-16AB、Content-ID: <...>。解析日期前整段丢掉。
_CID_REF = re.compile(
    r"(?i)(?:\bcid:[^\s<>\"']+|content-id\s*[:=]\s*<?[^>\s\"']+>?)"
)
# 文件名里的 20251231
_DATE_COMPACT = re.compile(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d)")
_DATE_META_LINE = re.compile(r"发送时间|发件时间|写道|Sent|Date\s*[:：]", re.I)
_DATE_LABEL_LINE = re.compile(r"加班日期|加班时间|加班日|值班日期")
_DATE_OT_WORD = re.compile(r"加班|值班")
# 「本周六周日」「周六、周日」整段一起吃掉，避免只留下第一天
_WEEK_MENTION = re.compile(
    r"(本周|这周|下周|(?<![本这下])周)"
    r"([一二三四五六日天](?:(?:[、,，和与及]|周)*[一二三四五六日天])*)"
)
_WEEKEND_MENTION = re.compile(r"(本周|这周|下周)末")
_APP_CUT = re.compile(
    r"(加班人员|加班人|加班名单|申请|安排|需要).{0,10}加班|于.{0,20}加班"
)
_CONTEXT_NEAR_NAME = re.compile(r"加班|值班|安排|申请|需要|人员|名单|姓名|名字|员工")
_SUPPLEMENT_MARK = re.compile(r"【(?:附件|图片识别)】")
_TITLE_TAIL = re.compile(
    r"(?:老师|经理|主任|总监|主管|领导|部长|处长|科长|总工|先生|女士|同学)$"
)
_DURATION_TAIL = re.compile(r"\s*(?:人天|小时|天|个小时|倍)")
_EN_MONTH = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}
_SENT_TIME_EN_DMY = re.compile(
    r"(?:发送时间|发件时间|Sent|Date)\s*[:：]\s*"
    r"(?:[A-Za-z]{3,9},?\s+)?(\d{1,2})\s+([A-Za-z]{3,9})\s+(20\d{2})",
    re.IGNORECASE,
)
_SENT_TIME_EN_MDY = re.compile(
    r"(?:发送时间|发件时间|Sent|Date)\s*[:：]\s*"
    r"([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(20\d{2})",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# 文本规范化
# ---------------------------------------------------------------------------

def _to_halfwidth(text: str) -> str:
    out: list[str] = []
    for ch in text:
        code = ord(ch)
        if code == 0x3000:
            out.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        else:
            out.append(ch)
    return "".join(out)


def _looks_like_spaced_ocr(text: str) -> bool:
    """OCR 常把每个汉字都用空格隔开。正常邮件几乎不会这样。"""
    gaps = len(re.findall(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", text))
    han = len(re.findall(r"[\u4e00-\u9fff]", text))
    return gaps >= 6 and han > 0 and gaps / han >= 0.3


def normalize_text(text: str) -> str:
    text = text.replace("\x00", "")
    text = text.replace("\u00a0", " ").replace("\u200b", "").replace("\ufeff", "")
    text = _to_halfwidth(text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if _looks_like_spaced_ocr(text):
        text = re.sub(r"(?<=[\u4e00-\u9fff])[ \t]+(?=[\u4e00-\u9fff])", "", text)
        text = re.sub(r"(?<=\d)[ \t]+(?=\d)", "", text)
        text = re.sub(r"(?<=\d)[ \t]+(?=[年月日])", "", text)
        text = re.sub(r"(?<=[年月日])[ \t]+(?=\d)", "", text)
    return text.strip()


def _tighten_cjk(text: str) -> str:
    """只用于批复判定，把「同 意」合成「同意」。"""
    return re.sub(r"(?<=[\u4e00-\u9fff])[ \t]+(?=[\u4e00-\u9fff])", "", text)


# ---------------------------------------------------------------------------
# 回复链与邮件头
# ---------------------------------------------------------------------------

def split_reply_chain(text: str) -> list[str]:
    """从上到下返回各段，第 0 段是最新回复。"""
    parts = [p.strip() for p in _CHAIN_SPLIT.split(text) if p and p.strip()]
    return parts or ([text.strip()] if text.strip() else [])


def _people_from_header_value(value: str) -> list[str]:
    """从邮件头字段值里抽出中文姓名/称呼，去重保序。"""
    value = re.sub(r"<[^>]*>", " ", value)
    value = re.sub(r"\S+@\S+", " ", value)
    people: list[str] = []
    seen: set[str] = set()
    for token in re.split(r"[;；,，、/／]", value):
        token = token.strip()
        if not token:
            continue
        candidates: list[str]
        if re.fullmatch(r"[\u4e00-\u9fff]{2,8}", token):
            candidates = [token]
        else:
            candidates = re.findall(r"[\u4e00-\u9fff]{2,4}", token)
        for name in candidates:
            if name in seen or name in _STOPWORDS:
                continue
            seen.add(name)
            people.append(name)
    return people


def _header_field_people(raw: str, labels: set[str]) -> list[str]:
    people: list[str] = []
    seen: set[str] = set()
    for line in raw.splitlines():
        matched = _PARTICIPANT_HEADER.match(line.strip())
        if not matched:
            continue
        if matched.group(1).lower() not in labels:
            continue
        for name in _people_from_header_value(matched.group(2)):
            if name in seen:
                continue
            seen.add(name)
            people.append(name)
    return people


def _header_people(raw: str) -> set[str]:
    people: set[str] = set()
    for line in raw.splitlines():
        matched = _PARTICIPANT_HEADER.match(line.strip())
        if not matched:
            continue
        people.update(_people_from_header_value(matched.group(2)))
    for matched in _SALUTATION.finditer(raw):
        people.add(matched.group(1))
    return people


def _signature_approver(decision_body: str) -> str | None:
    """批复正文末尾签名，如「冯老师」或「—— 王总」。"""
    lines = [ln.strip() for ln in decision_body.splitlines() if ln.strip()]
    if not lines:
        return None
    for line in reversed(lines[-3:]):
        cleaned = re.sub(r"^[-—–=~\s]+", "", line).strip()
        if _SHORT_APPROVE_LINE.fullmatch(cleaned):
            continue
        if _TITLE_NAME.fullmatch(cleaned) or re.fullmatch(r"[\u4e00-\u9fff]{2,4}", cleaned):
            if cleaned not in _STOPWORDS:
                return cleaned
    return None


def _approver_from_application(application_raw: str | None) -> str | None:
    if not application_raw:
        return None
    recipients = _header_field_people(application_raw, _TO_LABELS)
    if recipients:
        titled = [n for n in recipients if _TITLE_NAME.fullmatch(n)]
        return (titled or recipients)[0]
    for matched in _SALUTATION.finditer(application_raw):
        return matched.group(1)
    return None


def extract_approver_name(
    decision_raw: str | None,
    application_raw: str | None,
    status: str,
    *,
    same_segment: bool = False,
) -> str | None:
    """
    仅在 APPROVED 时填写「回复同意加班的人」。

    - 批复与申请不在同一段：批复发件人 > 批复签名 > 申请收件人
    - 同一段（正文顶部直接写「同意」）：不能用该段发件人（通常是申请人），
      优先取申请收件人 / 称呼
    """
    if status != "APPROVED":
        return None

    if same_segment:
        return _approver_from_application(application_raw)

    if decision_raw:
        senders = _header_field_people(decision_raw, _FROM_LABELS)
        if senders:
            return senders[0]
        body = _strip_header_lines(decision_raw)
        signed = _signature_approver(leading_decision_text(body) or body)
        if signed:
            return signed

    return _approver_from_application(application_raw)


def _strip_header_lines(raw: str) -> str:
    kept = [ln for ln in raw.splitlines() if not _HEADER_LINE.match(ln)]
    return "\n".join(kept).strip()


def _sent_date(raw: str) -> datetime | None:
    matched = _SENT_TIME.search(raw)
    if matched:
        try:
            return datetime(int(matched.group(1)), int(matched.group(2)), int(matched.group(3)))
        except ValueError:
            return None
    dmy = _SENT_TIME_EN_DMY.search(raw)
    if dmy:
        month = _EN_MONTH.get(dmy.group(2).lower())
        if month:
            try:
                return datetime(int(dmy.group(3)), month, int(dmy.group(1)))
            except ValueError:
                return None
    mdy = _SENT_TIME_EN_MDY.search(raw)
    if mdy:
        month = _EN_MONTH.get(mdy.group(1).lower())
        if month:
            try:
                return datetime(int(mdy.group(3)), month, int(mdy.group(2)))
            except ValueError:
                return None
    return None


def _is_application(body: str) -> bool:
    text = body.strip()
    if not text:
        return False
    # 纯批复句不是申请
    if classify_decision(text) and not re.search(r"(申请|安排|需要|人员|名单|于)", text):
        return False
    if re.search(r"加班人员|加班人|加班名单|加班日期", text):
        return True
    if re.search(r"(申请|安排|需要|需).{0,40}加班", text):
        return True
    if re.search(r"于\s*.{0,40}加班", text):
        return True
    for matched in re.finditer(rf"{_NAME_LEFT}({_NAME_LIST})\s*(?:{_COUNT}\s*)?加班", text):
        names = _split_name_chunk(matched.group(1))
        if any(_accept_name(n, set()) for n in names):
            return True
    return False


def _pure_approval_line(line: str) -> bool:
    return bool(_SHORT_APPROVE_LINE.fullmatch(line.strip()))


def leading_decision_text(body: str) -> str:
    """同一段里，批复通常写在申请句之前。申请句本身不拿来判同意。"""
    kept: list[str] = []
    for line in body.splitlines():
        stripped = line.strip()
        if not stripped:
            kept.append(line)
            continue
        if _pure_approval_line(stripped):
            kept.append(line)
            continue
        if re.search(r"加班|申请|安排", stripped):
            break
        kept.append(line)
    return "\n".join(kept).strip()


def classify_decision(decision_text: str) -> str | None:
    """返回 APPROVED / REJECTED；这段文字没有明确裁决时返回 None。"""
    if not decision_text or not decision_text.strip():
        return None
    text = _tighten_cjk(decision_text)

    if _REJECT.search(text):
        return "REJECTED"
    if _STRONG_APPROVE.search(text):
        return "APPROVED"

    # 申请里的「请同意 / 烦请同意」本身不是批复
    request_only = bool(re.search(r"请(?:您)?同意|烦请同意|申请同意|待同意", text))

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    for line in lines[:8]:
        tight = _tighten_cjk(line)
        if _REJECT.search(tight):
            return "REJECTED"
        if re.search(r"请(?:您)?同意|烦请同意|申请同意|待同意", tight):
            continue
        if len(tight) <= 40 and (
            _BARE_APPROVE.search(tight)
            or _SOFT_APPROVE.search(tight)
            or re.fullmatch(r"\+1[。！!]?", tight)
            or re.fullmatch(r"(?:同意|辛苦了|辛苦啦|没问题|好的|可以|批准)[，,。！!\s]*", tight)
        ):
            return "APPROVED"

    if request_only:
        return None

    # 整段较短时，裸「同意」也算（且不是「请同意」）
    if (
        len(text) <= 80
        and _BARE_APPROVE.search(text)
        and not re.search(r"请(?:您)?同意|烦请同意", text)
    ):
        return "APPROVED"
    return None


def detect_approval_status(
    segments: list[tuple[str, str]],
    roster: list[str] | None = None,
) -> tuple[str, str, str | None, str | None, bool]:
    """
    segments: (raw, body_without_headers)，最新在前。
    返回 (status, application_body, decision_raw, application_raw, same_segment)。
    """
    del roster  # 审批判定不依赖名单；保留参数兼容旧调用
    bodies = [body for _, body in segments]
    raws = [raw for raw, _ in segments]
    app_indexes = [i for i, body in enumerate(bodies) if _is_application(body)]

    decision_index: int | None = None
    application_index: int | None = None
    same_segment = False

    if app_indexes:
        newest_app = min(app_indexes)
        application_index = newest_app
        # 申请段及其下方的历史回复。上方批复段里的日期在解析时另行回退。
        application = "\n".join(bodies[newest_app:])
    else:
        application = "\n".join(bodies).strip()

    status: str | None = None
    if application_index is not None:
        for idx in range(application_index):
            status = classify_decision(bodies[idx])
            if not status:
                status = classify_decision(leading_decision_text(bodies[idx]))
            if not status:
                head = "\n".join(
                    ln for ln in raws[idx].splitlines()[:12] if not _HEADER_LINE.match(ln)
                )
                status = classify_decision(head)
            if status:
                decision_index = idx
                break
    if status is None:
        whole_head_lines = []
        for ln in "\n".join(bodies).splitlines()[:15]:
            if ln.strip() and not _HEADER_LINE.match(ln):
                whole_head_lines.append(ln)
        whole_head = "\n".join(whole_head_lines)
        cut = _APP_CUT.search(whole_head)
        head_only = whole_head[: cut.start()] if cut else leading_decision_text(application)
        if not head_only.strip():
            head_only = leading_decision_text(application)
        status = (
            classify_decision(head_only)
            or classify_decision(leading_decision_text(application))
            or "PENDING"
        )
        if status != "PENDING" and application_index is not None:
            decision_index = application_index
            same_segment = True
        elif status != "PENDING" and application_index is None:
            decision_index = 0
            same_segment = True

    joined = "\n".join(bodies)
    if status == "APPROVED" and "加班" not in application and "加班" not in joined:
        status = "PENDING"
        decision_index = None
        same_segment = False

    decision_raw = raws[decision_index] if decision_index is not None else None
    application_raw = raws[application_index] if application_index is not None else (
        raws[-1] if raws else None
    )
    return status, application, decision_raw, application_raw, same_segment


# ---------------------------------------------------------------------------
# 日期
# ---------------------------------------------------------------------------

def _valid_date(year: int, month: int, day: int) -> str | None:
    try:
        return datetime(year, month, day).strftime("%Y-%m-%d")
    except ValueError:
        return None


def _expand_range(start: datetime, end: datetime) -> list[str]:
    if end < start:
        start, end = end, start
    days = (end - start).days
    if days > _MAX_RANGE_DAYS:
        return [start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")]
    out: list[str] = []
    cur = start
    while cur <= end:
        out.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    return out


def _is_deadline(text: str, end: int) -> bool:
    tail = text[end : end + 8]
    if re.match(r"\s*(?:之前|前|截止)", tail):
        window = text[max(0, end - 12) : end + 8]
        if "加班" not in window:
            return True
    return False


def _in_supplement_region(text: str, pos: int) -> bool:
    """附件表格、图片 OCR 是加班材料，里面的日期直接收下。"""
    head = text[:pos]
    return max(head.rfind("【附件】"), head.rfind("【图片识别】")) >= 0


def _embedded_in_token(text: str, start: int, end: int) -> bool:
    """前后若还连着英文字母或数字，就是夹在十六进制片段里，不是独立日期。"""
    if start > 0:
        prev = text[start - 1]
        if prev.isascii() and prev.isalnum():
            return True
    if end < len(text):
        nxt = text[end]
        if nxt.isascii() and nxt.isalnum():
            return True
    return False


def _accept_dotted(
    text: str,
    end: int,
    month: int,
    day: int,
    has_ri: bool,
    start: int | None = None,
) -> bool:
    """点号或连字符日期。排除 1.5天、1.875 人天，保留 1.24、10.18日、9.21、11-22号。"""
    if start is not None and _embedded_in_token(text, start, end):
        return False
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return False
    tail = text[end : end + 4]
    if _DURATION_TAIL.match(tail):
        return False
    if has_ri:
        return True
    if month < 10 and day < 10:
        return False
    return True


def _date_in_overtime_context(text: str, start: int, end: int) -> bool:
    """日期必须靠近加班语境，避免把发送时间、写道时间当成加班日。"""
    line_start = text.rfind("\n", 0, start) + 1
    line_end = text.find("\n", end)
    if line_end < 0:
        line_end = len(text)
    line = text[line_start:line_end]
    if _DATE_META_LINE.search(line):
        return False
    if _in_supplement_region(text, start):
        return True
    if _DATE_LABEL_LINE.search(line):
        return True
    # 文件名、主题通常整行都不长，行内有「加班」即可
    if _DATE_OT_WORD.search(line) and len(line) <= 160:
        return True
    window = text[max(0, start - 80) : min(len(text), end + 80)]
    return bool(_DATE_OT_WORD.search(window))


def _push(
    bucket: list[tuple[int, int, list[str]]],
    start: int,
    end: int,
    dates: Iterable[str],
    text: str,
    *,
    require_context: bool = True,
) -> None:
    cleaned = [d for d in dates if d]
    if not cleaned:
        return
    if require_context and not _date_in_overtime_context(text, start, end):
        return
    bucket.append((start, end, cleaned))


def extract_dates(text: str, default_year: int | None = None) -> list[str]:
    """提取加班日期，去重后按时间排序。自动忽略非加班语境日期。"""
    text = _CID_REF.sub(" ", text)
    text = _META_DATE_LINE.sub(" ", text)
    year = default_year or datetime.now().year
    candidates: list[tuple[int, int, list[str]]] = []

    for matched in _DATE_CROSS.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        y1 = int(matched.group(1) or year)
        y2 = int(matched.group(4) or y1)
        start = _valid_date(y1, int(matched.group(2)), int(matched.group(3)))
        end = _valid_date(y2, int(matched.group(5)), int(matched.group(6)))
        if start and end:
            _push(
                candidates,
                matched.start(),
                matched.end(),
                _expand_range(
                    datetime.strptime(start, "%Y-%m-%d"),
                    datetime.strptime(end, "%Y-%m-%d"),
                ),
                text,
            )

    for matched in _DATE_SAME_MONTH.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        y = int(matched.group(1) or year)
        month = int(matched.group(2))
        d1, d2 = int(matched.group(3)), int(matched.group(4))
        if d2 < d1:
            continue
        start = _valid_date(y, month, d1)
        end = _valid_date(y, month, d2)
        if start and end:
            _push(
                candidates,
                matched.start(),
                matched.end(),
                _expand_range(
                    datetime.strptime(start, "%Y-%m-%d"),
                    datetime.strptime(end, "%Y-%m-%d"),
                ),
                text,
            )

    for matched in _DATE_LISTED.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        y = int(matched.group(1) or year)
        month = int(matched.group(2))
        days = [int(n) for n in re.findall(r"(\d{1,2})\s*[日号]", matched.group(3))]
        dates = [_valid_date(y, month, d) for d in days]
        _push(candidates, matched.start(), matched.end(), dates, text)

    for matched in _DATE_NUM_LISTED.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        y, month = int(matched.group(1)), int(matched.group(2))
        days = [int(matched.group(3))]
        days.extend(int(n) for n in re.findall(r"(?<!\d)(\d{1,2})(?!\d)", matched.group(4)))
        dates = [_valid_date(y, month, d) for d in days]
        _push(candidates, matched.start(), matched.end(), dates, text)

    for matched in _DATE_FULL.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        _push(
            candidates,
            matched.start(),
            matched.end(),
            [_valid_date(int(matched.group(1)), int(matched.group(2)), int(matched.group(3)))],
            text,
        )

    for matched in _DATE_MONTH_DAY.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        _push(
            candidates,
            matched.start(),
            matched.end(),
            [_valid_date(year, int(matched.group(1)), int(matched.group(2)))],
            text,
        )

    for matched in _DATE_DOT_RANGE.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        m1, d1 = int(matched.group(1)), int(matched.group(2))
        m2, d2 = int(matched.group(3)), int(matched.group(4))
        if _embedded_in_token(text, matched.start(), matched.end()):
            continue
        if not _accept_dotted(text, matched.end(), m1, d1, False, matched.start()):
            continue
        if not _accept_dotted(text, matched.end(), m2, d2, False, matched.start()):
            continue
        start = _valid_date(year, m1, d1)
        end = _valid_date(year, m2, d2)
        if start and end:
            _push(
                candidates,
                matched.start(),
                matched.end(),
                _expand_range(
                    datetime.strptime(start, "%Y-%m-%d"),
                    datetime.strptime(end, "%Y-%m-%d"),
                ),
                text,
            )

    for matched in _DATE_DOT.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        month, day = int(matched.group(1)), int(matched.group(2))
        if not _accept_dotted(
            text, matched.end(), month, day, bool(matched.group(3)), matched.start()
        ):
            continue
        _push(
            candidates,
            matched.start(),
            matched.end(),
            [_valid_date(year, month, day)],
            text,
        )

    for matched in _DATE_HYPHEN.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        month, day = int(matched.group(1)), int(matched.group(2))
        if not _accept_dotted(
            text, matched.end(), month, day, bool(matched.group(3)), matched.start()
        ):
            continue
        _push(
            candidates,
            matched.start(),
            matched.end(),
            [_valid_date(year, month, day)],
            text,
        )

    for matched in _DATE_COMPACT.finditer(text):
        if _is_deadline(text, matched.end()):
            continue
        _push(
            candidates,
            matched.start(),
            matched.end(),
            [_valid_date(int(matched.group(1)), int(matched.group(2)), int(matched.group(3)))],
            text,
        )

    # 较长的区间、列举优先，避免被其中的单日再计一次
    candidates.sort(key=lambda item: (-(item[1] - item[0]), item[0]))
    chosen: list[tuple[int, int, list[str]]] = []
    for start, end, dates in candidates:
        if any(not (end <= s or start >= e) for s, e, _ in chosen):
            continue
        chosen.append((start, end, dates))
    chosen.sort(key=lambda item: item[0])

    ordered: list[str] = []
    seen: set[str] = set()
    for _, _, dates in chosen:
        for date in dates:
            if date not in seen:
                seen.add(date)
                ordered.append(date)
    ordered.sort()
    return ordered


def _weekday_offset(anchor: datetime, prefix: str, day_char: str) -> str:
    weekday = _WEEKDAY[day_char]
    monday = anchor - timedelta(days=anchor.weekday())
    if prefix.startswith("下"):
        monday += timedelta(days=7)
    return (monday + timedelta(days=weekday)).strftime("%Y-%m-%d")


def _relative_span_has_absolute(text: str, start: int, end: int) -> bool:
    span = text[max(0, start - 8) : end + 24]
    return bool(re.search(r"20\d{2}|\d{1,2}\s*月\s*\d{1,2}\s*日", span))


def _relative_weekdays(text: str, anchor: datetime | None) -> list[str]:
    """「本周六周日加班」「本周末加班」换算成具体日期。没有锚定日则放弃。"""
    if anchor is None:
        return []
    found: list[str] = []
    seen: set[str] = set()

    def add(prefix: str, day_char: str) -> None:
        date = _weekday_offset(anchor, prefix, day_char)
        if date not in seen:
            seen.add(date)
            found.append(date)

    for matched in _WEEK_MENTION.finditer(text):
        window = text[max(0, matched.start() - 32) : matched.end() + 32]
        if not _DATE_OT_WORD.search(window):
            continue
        if _relative_span_has_absolute(text, matched.start(), matched.end()):
            continue
        prefix = matched.group(1)
        for day_char in re.findall(r"[一二三四五六日天]", matched.group(2)):
            add(prefix, day_char)

    for matched in _WEEKEND_MENTION.finditer(text):
        window = text[max(0, matched.start() - 32) : matched.end() + 32]
        if not _DATE_OT_WORD.search(window):
            continue
        if _relative_span_has_absolute(text, matched.start(), matched.end()):
            continue
        add(matched.group(1), "六")
        add(matched.group(1), "日")
    return found


def _extract_overtime_dates(
    text: str,
    default_year: int,
    anchor: datetime | None = None,
) -> list[str]:
    dates = extract_dates(text, default_year)
    for date in _relative_weekdays(text, anchor):
        if date not in dates:
            dates.append(date)
    dates.sort()
    return dates


# ---------------------------------------------------------------------------
# 人员名单 / 姓名
# ---------------------------------------------------------------------------

def _name_pattern_piece(name: str) -> str:
    """中文名允许字间空格；英文名允许连续空白。"""
    if re.fullmatch(r"[\u4e00-\u9fff]+", name):
        return r"[ \t\u00a0\u3000]*".join(re.escape(ch) for ch in name)
    return re.sub(r"\\ ", r"[ \t]+", re.escape(name))


def _compact_name_key(name: str) -> str:
    return re.sub(r"[\s·•]+", "", name).casefold()


class RosterMatcher:
    """人员名单一次编译，多封邮件复用。"""

    __slots__ = ("names", "pattern", "canonical")

    def __init__(self, names: list[str]):
        # 长名优先，交替匹配时先吃长名
        uniq = sorted({n.strip() for n in names if n and n.strip()}, key=len, reverse=True)
        self.names = uniq
        self.canonical: dict[str, str] = {}
        for name in uniq:
            self.canonical.setdefault(_compact_name_key(name), name)
        if uniq:
            flags = re.IGNORECASE if any(re.search(r"[A-Za-z]", n) for n in uniq) else 0
            pieces = [f"(?:{_name_pattern_piece(n)})" for n in uniq]
            self.pattern: re.Pattern[str] | None = re.compile("|".join(pieces), flags)
        else:
            self.pattern = None

    def __len__(self) -> int:
        return len(self.names)

    def _canon(self, matched: str) -> str:
        return self.canonical.get(_compact_name_key(matched), re.sub(r"\s+", "", matched))

    def _hits(self, text: str) -> list[tuple[int, int, str]]:
        if not self.pattern or not text:
            return []
        hits: list[tuple[int, int, str]] = []
        last_end = -1
        for matched in self.pattern.finditer(text):
            start, end = matched.start(), matched.end()
            if start < last_end:
                continue
            hits.append((start, end, self._canon(matched.group(0))))
            last_end = end
        return hits

    def find_unbounded(self, text: str) -> list[str]:
        """文件名、附件、OCR 里只要完整出现名单姓名就收下。"""
        found: list[str] = []
        seen: set[str] = set()
        for _, _, name in self._hits(text):
            if name not in seen and name not in _STOPWORDS:
                seen.add(name)
                found.append(name)
        return found

    def _find_contextual(self, body: str) -> list[str]:
        hits = self._hits(body)
        found: list[str] = []
        seen: set[str] = set()
        for start, end, name in hits:
            if name in seen or name in _STOPWORDS:
                continue
            window = body[max(0, start - 48) : end + 48]
            line_start = body.rfind("\n", 0, start) + 1
            line_end = body.find("\n", end)
            if line_end < 0:
                line_end = len(body)
            line = body[line_start:line_end].strip()
            pure_line = bool(re.fullmatch(rf"{re.escape(name)}(?:\s*[、,，;；])?", line))
            behind = body[max(0, start - 800) : start]
            if _CONTEXT_NEAR_NAME.search(window) or (
                pure_line and _CONTEXT_NEAR_NAME.search(behind)
            ):
                seen.add(name)
                found.append(name)
        if not found and "加班" in body:
            for _, _, name in hits:
                if name not in seen and name not in _STOPWORDS:
                    seen.add(name)
                    found.append(name)
        return found

    def find(self, text: str) -> list[str]:
        if not self.pattern or not text:
            return []
        body = _strip_header_lines(text)
        main, extras = _split_supplements(body)
        found = self._find_contextual(main)
        seen = set(found)
        for extra in extras:
            for name in self.find_unbounded(extra):
                if name not in seen:
                    seen.add(name)
                    found.append(name)
        return found


def _split_supplements(text: str) -> tuple[str, list[str]]:
    marks = list(_SUPPLEMENT_MARK.finditer(text))
    if not marks:
        return text, []
    main = text[: marks[0].start()]
    extras: list[str] = []
    for index, mark in enumerate(marks):
        end = marks[index + 1].start() if index + 1 < len(marks) else len(text)
        extras.append(text[mark.end() : end])
    return main, extras


_LATIN_NAME = re.compile(
    r"[A-Za-z][A-Za-z·.'\-]{0,30}(?:[ \t]+[A-Za-z][A-Za-z·.'\-]{0,30}){0,3}"
)
_LATIN_STOP = {
    "name", "names", "email", "e-mail", "none", "null", "nan",
    "dept", "department", "team", "staff", "employee", "ok",
}


def _split_roster_cell(text: str) -> list[str]:
    """顿号拆人；空格只在两侧都已经是完整姓名时才拆，避免把「张 三」拆碎。"""
    pieces: list[str] = []
    for item in re.split(r"[、,，;；/／|]+", text):
        item = item.strip()
        if not item:
            continue
        bits = [bit for bit in re.split(r"\s+", item) if bit]
        cjk_bits = [bit for bit in bits if re.fullmatch(r"[\u4e00-\u9fff]{2,8}", bit)]
        if len(bits) >= 2 and len(cjk_bits) == len(bits):
            pieces.extend(cjk_bits)
        else:
            pieces.append(item)
    return pieces


def _normalize_roster_value(raw: str) -> list[str]:
    """
    姓名列放宽：去掉称谓和括号备注，保留 2-8 个汉字，
    也保留纯英文名。一个单元格里用顿号隔开的多个人会拆开。
    """
    text = str(raw).strip()
    if not text or text.lower() in {"nan", "none", "null"}:
        return []
    pieces = _split_roster_cell(text)
    names: list[str] = []
    for piece in pieces:
        piece = piece.strip()
        piece = re.sub(r"（[^）]*）|\([^)]*\)|\[[^\]]*\]|【[^】]*】", "", piece)
        piece = piece.strip(" \t\"'“”‘’")
        if not piece:
            continue
        compact = re.sub(r"(?<=[\u4e00-\u9fff])[\s·•]+(?=[\u4e00-\u9fff])", "", piece)
        titled = _TITLE_TAIL.sub("", compact).strip()
        source = titled or compact
        cjk_runs = re.findall(r"[\u4e00-\u9fff]{2,12}", source)
        if cjk_runs:
            # 整格几乎就是姓名时才收，避免从一句说明里抠出碎词
            cjk_len = len(re.findall(r"[\u4e00-\u9fff]", source))
            kept = [
                run
                for run in cjk_runs
                if run not in _STOPWORDS
                and not any(word in run for word in ("加班", "申请", "审批", "日期", "项目", "部门", "公司", "附件"))
            ]
            if kept and cjk_len <= 12:
                names.extend(kept[:3])
            continue
        latin = re.sub(r"\s+", " ", piece).strip()
        matched = _LATIN_NAME.fullmatch(latin)
        if matched and latin.casefold() not in _LATIN_STOP and len(latin.replace(" ", "")) >= 2:
            names.append(latin)
    return names


def load_roster_names(path: str | Path) -> list[str]:
    """
    从人员名单.xlsx读取姓名。
    优先读列名含 姓名/名字/人员/员工 的列；否则扫描全部单元格中的中文姓名。
    姓名允许 2-8 个汉字、字间空格、中英混合（取汉字部分）和纯英文名。
    """
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise SystemExit("读取人员名单需要 pandas：pip install pandas openpyxl") from exc

    file_path = Path(path).expanduser().resolve()
    if not file_path.is_file():
        raise FileNotFoundError(f"人员名单不存在：{file_path}")

    frame = pd.read_excel(file_path, dtype=str)
    if frame.empty:
        return []

    preferred = [
        col
        for col in frame.columns
        if re.search(r"姓名|名字|人员|员工|成员|name", str(col), re.I)
    ]
    values: list[str] = []
    loose = bool(preferred)
    columns = preferred or list(frame.columns)
    for col in columns:
        values.extend(str(v).strip() for v in frame[col].tolist() if pd.notna(v))

    names: list[str] = []
    seen: set[str] = set()
    for raw in values:
        if loose:
            candidates = _normalize_roster_value(raw)
        else:
            token = re.sub(r"\s+", "", raw)
            token = _TITLE_TAIL.sub("", token)
            candidates = [token] if re.fullmatch(r"[\u4e00-\u9fff]{2,12}", token) else []
        for name in candidates:
            if name in _STOPWORDS or name in seen:
                continue
            seen.add(name)
            names.append(name)
    names.sort(key=len, reverse=True)
    return names


def load_roster_matcher(path: str | Path) -> RosterMatcher:
    return RosterMatcher(load_roster_names(path))


def _cleanup_chunk(chunk: str) -> str:
    chunk = chunk.strip(" \t:：,，、；;。")
    chunk = re.sub(
        r"^(?:请同意|麻烦|请|不同意|同意|需要|需|安排|申请|由|员工|同事)+",
        "",
        chunk,
    )
    chunk = re.sub(r"(?:等)?(?:[0-9一二两三四五六七八九十]+\s*人)?\s*$", "", chunk)
    chunk = re.sub(r"(加班|申请|安排|审批|日期|时间|一天|半天|全天).*$", "", chunk)
    return chunk.strip(" \t:：,，、；;。")


def _split_name_chunk(chunk: str) -> list[str]:
    chunk = _cleanup_chunk(chunk)
    if not chunk:
        return []
    parts = re.split(r"[、,，/／和与及\s]+", chunk)
    return [p.strip() for p in parts if p.strip()]


def _accept_name(name: str, excluded: set[str]) -> bool:
    if name in _STOPWORDS or name in excluded:
        return False
    if not re.fullmatch(_NAME, name):
        return False
    if any(len(word) >= 2 and word in name for word in _STOPWORDS):
        return False
    if _TITLE_NAME.fullmatch(name):
        return False
    if any(name.endswith(suffix) for suffix in ("老师", "经理", "主任", "总监", "主管", "领导", "部长", "处长", "科长")):
        return False
    return True


def _collect_names(patterns: list[str], text: str, excluded: set[str]) -> list[str]:
    names: list[str] = []
    seen: set[str] = set()
    for pattern in patterns:
        for matched in re.finditer(pattern, text):
            raw = matched.group(1)
            if re.match(r"\s*人", text[matched.end(1) :]):
                raw = re.sub(r"[一二两三四五六七八九十]$", "", raw)
            for name in _split_name_chunk(raw):
                if not _accept_name(name, excluded) or name in seen:
                    continue
                seen.add(name)
                names.append(name)
    return names


def _extract_names(
    text: str,
    excluded: set[str] | None = None,
    roster: list[str] | RosterMatcher | None = None,
) -> list[str]:
    """
    有人员名单时：只返回名单∩正文。
    无名单时：回退到旧正则（仅用于自检/兼容）。
    """
    if isinstance(roster, RosterMatcher):
        return roster.find(text)
    if isinstance(roster, list):
        return RosterMatcher(roster).find(text)

    excluded = excluded or set()
    tail = rf"(?:{_COUNT}\s*)?(?:于|在|加班|，|,|本周|这周|下周)"
    patterns = [
        rf"(?:加班人员|加班人|加班名单|参加加班(?:人员|同事|同学)?|值班人员|值班人)\s*[:：为是]?\s*({_NAME_LIST})",
        rf"(?:请\s*)?安排\s*({_NAME_LIST})\s*{tail}",
        rf"(?:申请|由)\s*({_NAME_LIST})\s*{tail}",
        rf"(?<!不)同意\s*({_NAME_LIST})\s*{tail}",
        rf"需要\s*({_NAME_LIST})\s*(?:{_COUNT}\s*)?(?=安排|加班|于|在|本周|这周|下周|参加)",
        rf"{_NAME_LEFT}({_NAME_LIST})\s*(?:{_COUNT}\s*)?(?:申请|需要|需)\s*加班",
        rf"{_NAME_LEFT}({_NAME_LIST})\s*(?:{_COUNT}\s*)?于\s*[^。\n]{{0,48}}?加班",
        rf"请\s*({_NAME_LIST})\s*(?:{_COUNT}\s*)?加班",
    ]
    return _collect_names(patterns, text, excluded)


# ---------------------------------------------------------------------------
# 组装
# ---------------------------------------------------------------------------

def _reference_year(
    raw_segments: list[str],
    application: str,
    reference_year: int | None,
    source_file: str | None = None,
) -> tuple[int, datetime | None]:
    anchors = [d for d in (_sent_date(seg) for seg in raw_segments) if d is not None]
    anchor = anchors[-1] if anchors else None  # 最旧一封，通常是申请发出日
    if anchor is not None:
        return anchor.year, anchor
    years = [int(y) for y in re.findall(r"(20\d{2})\s*年", application)]
    if not years and source_file:
        compact = _DATE_COMPACT.search(source_file)
        if compact:
            years = [int(compact.group(1))]
        else:
            years = [int(y) for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", source_file)]
    if years:
        return years[0], None
    if reference_year:
        return reference_year, None
    return datetime.now().year, None


def _prepend_mtime_anchor(
    text: str,
    source_path: str | None,
    source_file: str | None,
) -> str:
    """正文没有发送时间时，用邮件文件修改时间给「本周六」这类说法做锚点。"""
    if _sent_date(text) is not None or not source_path or not source_file:
        return text
    file_path = Path(source_path) / source_file
    try:
        if not file_path.is_file():
            return text
        stamp = datetime.fromtimestamp(file_path.stat().st_mtime)
    except OSError:
        return text
    return f"发送时间：{stamp.strftime('%Y-%m-%d')}\n{text}"


def _mail_name_chunks(raw_segments: list[str], source_file: str | None) -> list[str]:
    """主题行、读信时标出的【邮件名】、以及 .eml 文件名。这些位置也会写加班人。"""
    chunks: list[str] = []
    for raw in raw_segments:
        for line in raw.splitlines():
            stripped = line.strip()
            if stripped.startswith("【邮件名】"):
                chunks.append(stripped[len("【邮件名】") :].strip())
                continue
            matched = _SUBJECT_LINE.match(stripped)
            if matched:
                chunks.append(matched.group(1).strip())
    if source_file:
        chunks.append(normalize_text(Path(source_file).stem))
    return [chunk for chunk in chunks if chunk]


def _names_in_mail_names(
    raw_segments: list[str],
    source_file: str | None,
    roster: list[str] | RosterMatcher | None,
) -> list[str]:
    if roster is None:
        return []
    matcher = roster if isinstance(roster, RosterMatcher) else RosterMatcher(roster)
    found: list[str] = []
    seen: set[str] = set()
    for chunk in _mail_name_chunks(raw_segments, source_file):
        for name in matcher.find_unbounded(chunk):
            if name not in seen:
                seen.add(name)
                found.append(name)
    return found


def _null_record(
    status: str,
    approver: str | None = None,
    *,
    source_path: str | None = None,
    source_file: str | None = None,
) -> dict[str, Any]:
    return {
        "member_name": None,
        "overtime_date": None,
        "approval_status": status,
        "approved": _STATUS_LABEL.get(status, "待审批"),
        "approver_name": approver if status == "APPROVED" else None,
        "source_path": source_path,
        "source_file": source_file,
    }


def build_records(
    names: list[str],
    dates: list[str],
    approval_status: str,
    approver_name: str | None = None,
    *,
    source_path: str | None = None,
    source_file: str | None = None,
) -> list[dict[str, Any]]:
    """每人每一天一条。姓名或日期缺失时对应字段为 None。"""
    approved = _STATUS_LABEL.get(approval_status, "待审批")
    approver = approver_name if approval_status == "APPROVED" else None
    if not names and not dates:
        return [
            _null_record(
                approval_status,
                approver,
                source_path=source_path,
                source_file=source_file,
            )
        ]
    name_axis: list[str | None] = list(names) or [None]
    date_axis: list[str | None] = list(dates) or [None]
    return [
        {
            "member_name": name,
            "overtime_date": date,
            "approval_status": approval_status,
            "approved": approved,
            "approver_name": approver,
            "source_path": source_path,
            "source_file": source_file,
        }
        for name, date in product(name_axis, date_axis)
    ]


def records_to_excel_rows(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """转成 Excel 中文列。"""
    rows: list[dict[str, Any]] = []
    for rec in records:
        rows.append(
            {
                "加班人员名称": rec.get("member_name"),
                "加班日期": rec.get("overtime_date"),
                "是否同意加班": rec.get("approved")
                or _STATUS_LABEL.get(rec.get("approval_status", ""), "待审批"),
                "回复同意加班的人": rec.get("approver_name"),
                "邮件所在路径": rec.get("source_path"),
                "文件名": rec.get("source_file"),
            }
        )
    return rows


def _file_link(source_path: str | None, source_file: str | None) -> str:
    if source_path and source_file:
        return str(Path(source_path) / source_file)
    return source_file or source_path or ""


def _gap_reason(name: str | None, date: str | None) -> str:
    if not name and not date:
        return "未提取到加班人员和加班日期"
    if not name:
        return "未提取到加班人员名称"
    return "未提取到加班日期"


def _gap_row(
    link: str,
    reason: str,
    *,
    name: str | None = None,
    date: str | None = None,
    filename: str | None = None,
) -> dict[str, Any]:
    return {
        "文件链接": link,
        "原因": reason,
        "加班人员名称": name,
        "加班日期": date,
        "文件名": filename,
    }


def split_records(
    records: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """姓名和日期都有的留下，缺任一字段的改记到未完整提取。"""
    complete: list[dict[str, Any]] = []
    gaps: list[dict[str, Any]] = []
    for rec in records:
        name = rec.get("member_name")
        date = rec.get("overtime_date")
        if name and date:
            complete.append(rec)
            continue
        gaps.append(
            _gap_row(
                _file_link(rec.get("source_path"), rec.get("source_file")),
                _gap_reason(name, date),
                name=name,
                date=date,
                filename=rec.get("source_file"),
            )
        )
    return complete, gaps


def _write_file_hyperlinks(path: Path, sheet_name: str) -> None:
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    book = load_workbook(path)
    sheet = book[sheet_name]
    font = Font(color="0563C1", underline="single")
    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1):
        cell = row[0]
        if not cell.value:
            continue
        try:
            cell.hyperlink = Path(str(cell.value)).as_uri()
        except ValueError:
            cell.hyperlink = str(cell.value)
        cell.font = font
    book.save(path)


def export_records_to_excel(
    records: list[dict[str, Any]],
    output_path: str | Path,
    gaps: list[dict[str, Any]] | None = None,
) -> Path:
    """写成两个 sheet：加班审批只含姓名和日期都有的行，其余进未完整提取。"""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise SystemExit("导出 Excel 需要安装 pandas 与 openpyxl：pip install pandas openpyxl") from exc

    path = Path(output_path).expanduser().resolve()
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)

    complete, derived = split_records(records)
    all_gaps = list(gaps or []) + derived
    rows = records_to_excel_rows(complete)
    approved = pd.DataFrame(rows, columns=list(_EXCEL_COLUMNS))
    missed = pd.DataFrame(all_gaps, columns=list(_GAP_COLUMNS))
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        approved.to_excel(writer, index=False, sheet_name="加班审批")
        missed.to_excel(writer, index=False, sheet_name=_GAP_SHEET)
    _write_file_hyperlinks(path, _GAP_SHEET)
    return path


def parse_overtime_email(
    email_text: str,
    reference_year: int | None = None,
    roster: list[str] | RosterMatcher | None = None,
    *,
    source_path: str | None = None,
    source_file: str | None = None,
) -> list[dict[str, Any]]:
    """
    解析一封邮件或一条回复链。

    roster: 人员名单 list 或预编译 RosterMatcher。
    """
    meta = {"source_path": source_path, "source_file": source_file}
    if not email_text or not email_text.strip():
        return [_null_record("PENDING", **meta)]

    normalized = normalize_text(email_text)
    if len(normalized) > _MAX_PARSE_CHARS:
        normalized = normalized[:_MAX_PARSE_CHARS]
    filename_text = normalize_text(source_file or "")
    if (
        "加班" not in normalized
        and "加班" not in _tighten_cjk(normalized)
        and "加班" not in filename_text
    ):
        return [_null_record("PENDING", **meta)]
    normalized = _prepend_mtime_anchor(normalized, source_path, source_file)

    raw_segments = split_reply_chain(normalized)
    excluded: set[str] = set()
    prepared: list[tuple[str, str]] = []
    for raw in raw_segments:
        excluded |= _header_people(raw)
        prepared.append((raw, _strip_header_lines(raw)))

    status, application, decision_raw, application_raw, same_segment = detect_approval_status(
        prepared,
        roster=None,
    )
    year, anchor = _reference_year(raw_segments, application, reference_year, source_file)
    names = _extract_names(application, excluded, roster=roster)
    full_body = "\n".join(body for _, body in prepared)
    # 人名只写在批复段、转发丢了正文时，再扫整封邮件
    if not names and full_body != application:
        names = _extract_names(full_body, excluded, roster=roster)
    # 邮件主题、文件名里的名单姓名始终并入，不因正文已经有别人而丢掉
    for name in _names_in_mail_names(raw_segments, source_file, roster):
        if name not in names:
            names.append(name)
    dates = _extract_overtime_dates(application, year, anchor)
    # 日期在上方批复段，或只写在文件名（1.24、9月20日、20251231）
    if not dates and full_body != application:
        dates = _extract_overtime_dates(full_body, year, anchor)
    if not dates:
        label_text = "\n".join(_mail_name_chunks(raw_segments, source_file))
        if label_text.strip():
            dates = _extract_overtime_dates(f"加班日期：\n{label_text}", year, anchor)
    approver = extract_approver_name(
        decision_raw,
        application_raw,
        status,
        same_segment=same_segment,
    )
    return build_records(
        names,
        dates,
        status,
        approver,
        source_path=source_path,
        source_file=source_file,
    )


def parse_overtime_email_json(
    email_text: str,
    reference_year: int | None = None,
    roster: list[str] | None = None,
) -> str:
    records = parse_overtime_email(
        email_text,
        reference_year=reference_year,
        roster=roster,
    )
    return json.dumps(records, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 文件读取 / 批量
# ---------------------------------------------------------------------------

_DATA_URI_RE = re.compile(
    r"data:image/(?:png|jpeg|jpg|bmp|gif|webp);base64,([A-Za-z0-9+/=\r\n]+)",
    re.IGNORECASE,
)
_IMG_SRC_RE = re.compile(
    r"""<img\b[^>]*?\bsrc\s*=\s*["']([^"']+)["']""",
    re.IGNORECASE,
)
_OCR_SLOTS = 2
_ocr_pool: queue.Queue = queue.Queue()
_ocr_made = 0
_ocr_make_lock = threading.Lock()
_ocr_cache: dict[str, str] = {}
_ocr_cache_lock = threading.Lock()
_ocr_unavailable = False
_ocr_warned = False


def _html_to_text(html_text: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", html_text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"(?i)</tr\s*>", "\n", text)
    text = re.sub(r"(?i)</t[dh]\s*>", " ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    for _ in range(2):
        unescaped = html.unescape(text)
        if unescaped == text:
            break
        text = unescaped
    text = text.replace("\xa0", " ").replace("\u3000", " ").replace("\u200b", "")
    text = re.sub(r"(?i)&nbsp;?", " ", text)
    return text


def _merge_plain_html(plain: str, html_text: str) -> str:
    """纯文本是通用模板、名单在 HTML 表格里时，把 HTML 文本补进来。"""
    html_text = html_text.strip()
    if not html_text:
        return plain
    if not plain.strip():
        return html_text
    sample = html_text[:240]
    if sample and sample in plain:
        return plain
    if len(plain) >= len(html_text) * 0.8 and plain[:240] in html_text:
        return plain
    return plain.rstrip() + "\n" + html_text


def _decode_data_uri_images(html: str) -> list[bytes]:
    blobs: list[bytes] = []
    for matched in _DATA_URI_RE.finditer(html or ""):
        try:
            raw = base64.b64decode(re.sub(r"\s+", "", matched.group(1)), validate=False)
        except Exception:
            continue
        if len(raw) >= 800:
            blobs.append(raw)
        if len(blobs) >= 8:
            break
    return blobs


def _make_rapid_ocr():
    try:
        from rapidocr_onnxruntime import RapidOCR

        return RapidOCR()
    except Exception:
        pass
    try:
        from rapidocr import RapidOCR

        return RapidOCR()
    except Exception:
        return None


def _warn_ocr_once() -> None:
    global _ocr_warned
    if _ocr_warned:
        return
    _ocr_warned = True
    print(
        "提示：图片 OCR 不可用，图片里的姓名和日期会丢失。"
        "可安装 rapidocr-onnxruntime 或 pytesseract（chi_sim）。"
    )


def _ocr_tesseract(data: bytes) -> str:
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return ""
    try:
        image = Image.open(io.BytesIO(data))
        return pytesseract.image_to_string(image, lang="chi_sim+eng") or ""
    except Exception:
        return ""


def _borrow_ocr():
    global _ocr_made, _ocr_unavailable
    if _ocr_unavailable:
        return None
    try:
        return _ocr_pool.get_nowait()
    except queue.Empty:
        pass
    with _ocr_make_lock:
        if _ocr_unavailable:
            return None
        if _ocr_made < _OCR_SLOTS:
            engine = _make_rapid_ocr()
            if engine is None:
                _ocr_unavailable = True
                return None
            _ocr_made += 1
            return engine
    return _ocr_pool.get()


def _return_ocr(engine) -> None:
    if engine is not None:
        _ocr_pool.put(engine)


def _prepare_ocr_image(data: bytes):
    """过小的图标、分割线不识别；过大图片先缩小。"""
    try:
        from PIL import Image
    except ImportError:
        return None
    try:
        image = Image.open(io.BytesIO(data))
        image = image.convert("RGB")
    except Exception:
        return None
    width, height = image.size
    if width < 40 or height < 20:
        return None
    if width * height < 5_000:
        return None
    if width / max(height, 1) > 40 or height / max(width, 1) > 40:
        return None
    longest = max(width, height)
    if longest > 1800:
        scale = 1800 / longest
        image = image.resize((int(width * scale), int(height * scale)))
    try:
        import numpy as np
    except ImportError:
        return None
    return np.asarray(image)


def ocr_image_bytes(data: bytes) -> str:
    """识别一张图片，返回按阅读顺序排列的文字。"""
    if not data:
        return ""
    digest = hashlib.md5(data).hexdigest()
    with _ocr_cache_lock:
        cached = _ocr_cache.get(digest)
    if cached is not None:
        return cached

    array = _prepare_ocr_image(data)
    if array is None:
        text = ""
    else:
        engine = _borrow_ocr()
        text = ""
        if engine is not None:
            try:
                result, _elapsed = engine(array)
                lines: list[tuple[float, float, str]] = []
                for item in result or []:
                    if len(item) < 3:
                        continue
                    box, content, score = item[0], str(item[1]), float(item[2])
                    if score < 0.45 or not content.strip():
                        continue
                    ys = [point[1] for point in box]
                    xs = [point[0] for point in box]
                    lines.append((min(ys), min(xs), content.strip()))
                lines.sort()
                text = "\n".join(line for _, _, line in lines)
            except Exception:
                text = ""
            finally:
                _return_ocr(engine)
        else:
            text = _ocr_tesseract(data)
            if not text.strip():
                _warn_ocr_once()
    with _ocr_cache_lock:
        if len(_ocr_cache) > 400:
            _ocr_cache.clear()
        _ocr_cache[digest] = text
    return text


def _append_ocr_text(body: str, images: list[bytes]) -> str:
    chunks: list[str] = []
    seen: set[str] = set()
    for blob in images:
        if len(chunks) >= 8:
            break
        digest = hashlib.md5(blob).hexdigest()
        if digest in seen:
            continue
        seen.add(digest)
        text = ocr_image_bytes(blob).strip()
        if text:
            chunks.append(text)
    if not chunks:
        return body
    return body.rstrip() + "\n\n【图片识别】\n" + "\n".join(chunks)


_IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"}


def _looks_like_image(data: bytes) -> bool:
    """按文件头判断是不是图片。发件人常把内嵌 PNG 标成 application/octet-stream。"""
    if len(data) < 12:
        return False
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return True
    if data.startswith(b"\xff\xd8\xff"):
        return True
    if data.startswith((b"GIF87a", b"GIF89a")):
        return True
    if data.startswith(b"BM"):
        return True
    if data.startswith(b"RIFF") and data[8:12] == b"WEBP":
        return True
    return False


def _part_filename(part) -> str:
    filename = part.get_filename()
    if filename:
        return str(filename)
    name = part.get_param("name")
    return str(name or "")


def _should_read_as_image(part) -> bool:
    ctype = (part.get_content_type() or "").lower()
    if ctype.startswith("image/"):
        return True
    if ctype in {"application/octet-stream", "application/octetstream", "binary/octet-stream"}:
        return True
    if Path(_part_filename(part)).suffix.lower() in _IMAGE_SUFFIXES:
        return True
    if part.get("content-id") and not ctype.startswith("text/"):
        return True
    return False


def _part_is_image(part, data: bytes) -> bool:
    ctype = (part.get_content_type() or "").lower()
    if ctype.startswith("image/"):
        return True
    if Path(_part_filename(part)).suffix.lower() in _IMAGE_SUFFIXES:
        return True
    return _looks_like_image(data)


def _message_images(msg) -> list[bytes]:
    blobs: list[bytes] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            if ctype == "text/html":
                try:
                    html_parts.append(str(part.get_content()))
                except Exception:
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        html_parts.append(payload.decode("utf-8", errors="ignore"))
                continue
            if not _should_read_as_image(part):
                continue
            try:
                data = part.get_payload(decode=True)
            except Exception:
                data = None
            if isinstance(data, bytes) and data and _part_is_image(part, data):
                blobs.append(data)
    for html_part in html_parts:
        blobs.extend(_decode_data_uri_images(html_part))
    return blobs


def _local_html_images(path: Path, html: str) -> list[bytes]:
    blobs = _decode_data_uri_images(html)
    for matched in _IMG_SRC_RE.finditer(html):
        src = matched.group(1).strip()
        if not src or src.lower().startswith(("data:", "http:", "https:", "cid:")):
            continue
        img_path = (path.parent / src).resolve()
        try:
            if img_path.is_file() and img_path.stat().st_size >= 800:
                blobs.append(img_path.read_bytes())
        except OSError:
            continue
        if len(blobs) >= 8:
            break
    return blobs


def _message_sent_line(msg) -> str:
    """把邮件头 Date 写成正文里的发送时间，供年份和「本周六」换算。"""
    raw = msg.get("date")
    if not raw:
        return ""
    try:
        stamp = parsedate_to_datetime(str(raw))
    except (TypeError, ValueError, IndexError, OverflowError):
        return ""
    if stamp is None:
        return ""
    return f"发送时间：{stamp.strftime('%Y-%m-%d %H:%M')}\n"


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none"}:
        return ""
    return re.sub(r"\s+00:00:00$", "", text)


def _table_bytes_to_text(data: bytes, suffix: str) -> str:
    suffix = suffix.lower()
    if suffix in {".csv", ".txt"}:
        for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
            try:
                return data.decode(enc)
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="ignore")
    if suffix == ".docx":
        return _docx_bytes_to_text(data)
    if suffix not in {".xlsx", ".xlsm", ".xls"}:
        return ""
    try:
        import pandas as pd
    except ImportError:
        return ""
    frame = None
    for engine in (None, "openpyxl"):
        try:
            kwargs: dict[str, Any] = {"header": None, "sheet_name": 0}
            if engine:
                kwargs["engine"] = engine
            frame = pd.read_excel(io.BytesIO(data), **kwargs)
            break
        except Exception:
            frame = None
    if frame is None or frame.empty:
        return ""
    lines: list[str] = []
    for row in frame.itertuples(index=False):
        cells = [_cell_text(cell) for cell in row]
        cells = [cell for cell in cells if cell]
        if cells:
            lines.append(" ".join(cells[:20]))
        if len(lines) >= 300:
            break
    return "\n".join(lines)


def _docx_bytes_to_text(data: bytes) -> str:
    import zipfile

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    except (KeyError, OSError, ValueError):
        return ""
    text = re.sub(r"</w:p>", "\n", xml)
    text = re.sub(r"<[^>]+>", "", text)
    return text


def _attachment_text(msg) -> str:
    """Excel / CSV / Word 附件里的加班名单。图片走 OCR，不在这里读。"""
    if not getattr(msg, "is_multipart", lambda: False)():
        return ""
    chunks: list[str] = []
    for part in msg.walk():
        filename = part.get_filename() or ""
        suffix = Path(filename).suffix.lower()
        ctype = (part.get_content_type() or "").lower()
        if suffix not in {".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".txt"}:
            if ctype in {"text/csv", "application/csv"}:
                suffix = ".csv"
            elif "spreadsheet" in ctype or "excel" in ctype:
                suffix = ".xlsx"
            elif "wordprocessingml" in ctype:
                suffix = ".docx"
            else:
                continue
        try:
            data = part.get_payload(decode=True)
        except Exception:
            data = None
        if not isinstance(data, bytes) or len(data) < 16:
            continue
        text = _table_bytes_to_text(data, suffix).strip()
        if text:
            chunks.append(text)
        if len(chunks) >= 4:
            break
    return "\n".join(chunks)


class MsgReadError(Exception):
    """Outlook .msg 读失败时带给未完整提取页的原因。"""


def _clean_msg_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        text = ""
        for enc in ("utf-8", "utf-16", "gb18030"):
            try:
                text = value.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = value.decode("utf-8", errors="ignore")
    else:
        text = str(value)
    return text.replace("\x00", "").strip()


def _render_outlook_message(message, depth: int = 0) -> str:
    """把 Outlook 邮件收成和 .eml 一样的文本：发送时间、邮件名、正文、图片、附件。"""
    subject = _clean_msg_text(getattr(message, "subject", ""))
    body = _clean_msg_text(getattr(message, "body", ""))
    html = _html_to_text(_clean_msg_text(getattr(message, "htmlBody", "")))
    merged = _merge_plain_html(body, html)
    sent = ""
    stamp = getattr(message, "date", None)
    if isinstance(stamp, datetime):
        sent = f"发送时间：{stamp.strftime('%Y-%m-%d %H:%M')}\n"
    sender = _clean_msg_text(getattr(message, "sender", ""))
    recipient = _clean_msg_text(getattr(message, "to", ""))
    headers = ""
    if sender:
        headers += f"发件人：{sender}\n"
    if recipient:
        headers += f"收件人：{recipient}\n"
    subject_line = f"【邮件名】{subject}\n" if subject else ""
    images: list[bytes] = []
    attach_chunks: list[str] = []
    for att in getattr(message, "attachments", []) or []:
        filename = ""
        try:
            filename = _clean_msg_text(
                att.longFilename or att.shortFilename or att.getFilename()
            )
        except Exception:
            filename = ""
        try:
            data = att.data
        except Exception:
            data = None
        if data is None:
            continue
        if depth < 2 and hasattr(data, "subject"):
            try:
                attach_chunks.append(_render_outlook_message(data, depth + 1))
            except Exception:
                pass
            continue
        if not isinstance(data, (bytes, bytearray)):
            continue
        blob = bytes(data)
        suffix = Path(filename).suffix.lower()
        mime = _clean_msg_text(getattr(att, "mimetype", "")).lower()
        if suffix in _IMAGE_SUFFIXES or mime.startswith("image/") or _looks_like_image(blob):
            images.append(blob)
            continue
        table = _table_bytes_to_text(blob, suffix).strip()
        if table:
            attach_chunks.append(table)
    text = _append_ocr_text(sent + headers + subject_line + merged, images)
    if attach_chunks:
        text = text.rstrip() + "\n\n【附件】\n" + "\n".join(attach_chunks)
    return text


def read_msg(path: Path) -> str:
    """读取 Outlook .msg 的主题、正文、发送时间和附件。"""
    try:
        import extract_msg
    except ImportError as exc:
        raise MsgReadError("读取 .msg 需要安装 extract-msg：pip install extract-msg") from exc
    try:
        message = extract_msg.openMsg(str(path))
    except Exception as exc:
        raise MsgReadError(f"无法解析 Outlook 邮件：{exc}") from exc
    try:
        return _render_outlook_message(message)
    finally:
        message.close()


def read_eml(path: Path) -> str:
    with path.open("rb") as fh:
        msg = BytesParser(policy=policy.default).parse(fh)
    plain_parts: list[str] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disposition = (part.get_content_disposition() or "").lower()
            if disposition == "attachment":
                continue
            if ctype == "text/plain":
                plain_parts.append(str(part.get_content()))
            elif ctype == "text/html":
                html_parts.append(_html_to_text(str(part.get_content())))
    else:
        try:
            content = str(msg.get_content())
        except Exception:
            payload = msg.get_payload(decode=True)
            content = payload.decode("utf-8", errors="ignore") if isinstance(payload, bytes) else ""
        if (msg.get_content_type() or "").lower() == "text/html":
            html_parts.append(_html_to_text(content))
        else:
            plain_parts.append(content)
    subject = str(msg.get("subject", "") or "").strip()
    subject_line = f"【邮件名】{subject}\n" if subject else ""
    merged = _merge_plain_html("\n".join(plain_parts), "\n".join(html_parts))
    body = _append_ocr_text(_message_sent_line(msg) + subject_line + merged, _message_images(msg))
    attachment = _attachment_text(msg)
    if attachment:
        body = body.rstrip() + "\n\n【附件】\n" + attachment
    return body


def read_html_file(path: Path) -> str:
    raw = read_text_file(path)
    body = _html_to_text(raw)
    return _append_ocr_text(body, _local_html_images(path, raw))


def read_text_file(path: Path) -> str:
    data = path.read_bytes()
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def load_file_text(path: Path) -> str | None:
    suffix = path.suffix.lower()
    try:
        if suffix == ".eml":
            return read_eml(path)
        if suffix == ".msg":
            return read_msg(path)
        if suffix in {".html", ".htm"}:
            return read_html_file(path)
        if suffix in _TEXT_LIKE_EXTENSIONS or suffix == "":
            return read_text_file(path)
        text = read_text_file(path)
        sample = text[:500]
        printable = sum(1 for c in sample if c.isprintable() or c.isspace())
        if len(text.strip()) >= 20 and printable > 400:
            return text
        return None
    except MsgReadError:
        raise
    except OSError:
        return None


def _is_candidate_file(path: Path) -> bool:
    suffix = path.suffix.lower()
    if suffix in _SKIP_EXTENSIONS:
        return False
    name = path.name
    if any(k in name for k in _NAME_KEYWORDS):
        return True
    return suffix in _TEXT_LIKE_EXTENSIONS or suffix == ""


def _peek_has_overtime(path: Path) -> bool:
    """先读一小段，没有「加班」则跳过整文件解析。"""
    if any(k in path.name for k in _NAME_KEYWORDS):
        return True
    if path.suffix.lower() == ".msg":
        # 正文在 OLE 流里，文件头看不出「加班」，必须整封解析。
        return True
    with path.open("rb") as fh:
        chunk = fh.read(_PEEK_BYTES)
    for enc in ("utf-8", "gb18030", "gbk", "latin-1"):
        try:
            text = chunk.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = chunk.decode("utf-8", errors="ignore")
    return "加班" in text


def _process_one_file(
    path: Path,
    *,
    reference_year: int | None,
    roster: list[str] | RosterMatcher | None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    resolved = path.resolve()
    link = str(resolved)
    try:
        has_overtime = _peek_has_overtime(path)
    except OSError as exc:
        return [], [_gap_row(link, f"文件无法读取：{exc}", filename=resolved.name)]
    if not has_overtime:
        return [], [_gap_row(link, "未识别到加班内容", filename=resolved.name)]
    try:
        text = load_file_text(path)
    except MsgReadError as exc:
        return [], [_gap_row(link, str(exc), filename=resolved.name)]
    if text is None:
        return [], [_gap_row(link, "文件无法读取", filename=resolved.name)]
    if "加班" not in text and "加班" not in path.name:
        return [], [_gap_row(link, "未识别到加班内容", filename=resolved.name)]
    parsed = parse_overtime_email(
        text,
        reference_year=reference_year,
        roster=roster,
        source_path=str(resolved.parent),
        source_file=resolved.name,
    )
    complete, gaps = split_records(parsed)
    if not complete and not gaps:
        gaps = [_gap_row(link, "未提取到加班人员和加班日期", filename=resolved.name)]
    return complete, gaps


def parse_overtime_directory(
    src_dir: str | Path,
    *,
    recursive: bool = True,
    reference_year: int | None = None,
    roster: list[str] | RosterMatcher | None = None,
    roster_path: str | Path | None = None,
    workers: int | None = None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """扫描目录。返回 (姓名日期都有的记录, 未完整提取及原因)。"""
    if roster is None and roster_path is not None:
        roster = load_roster_matcher(roster_path)
    elif isinstance(roster, list):
        roster = RosterMatcher(roster)

    root = Path(src_dir).expanduser().resolve()
    files = [
        p
        for p in (root.rglob("*") if recursive else root.iterdir())
        if p.is_file() and _is_candidate_file(p)
    ]
    if not files:
        return [], []

    worker_n = workers if workers and workers > 0 else min(_DEFAULT_WORKERS, len(files))
    all_rows: list[dict[str, Any]] = []
    all_gaps: list[dict[str, Any]] = []
    done = 0
    total = len(files)
    print(f"待处理候选文件：{total}，线程数：{worker_n}")

    with ThreadPoolExecutor(max_workers=worker_n) as pool:
        futures = {
            pool.submit(
                _process_one_file,
                path,
                reference_year=reference_year,
                roster=roster,
            ): path
            for path in files
        }
        for fut in as_completed(futures):
            done += 1
            path = futures[fut]
            try:
                rows, gaps = fut.result()
            except Exception as exc:  # noqa: BLE001
                print(f"[跳过] 处理失败：{path.name}：{exc}")
                rows = []
                gaps = [_gap_row(str(path.resolve()), f"处理失败：{exc}", filename=path.name)]
            all_rows.extend(rows)
            all_gaps.extend(gaps)
            if done % 50 == 0 or done == total:
                print(f"进度 {done}/{total}，完整 {len(all_rows)} 条，未完整 {len(all_gaps)} 条")
    return all_rows, all_gaps


# ---------------------------------------------------------------------------
# 自检
# ---------------------------------------------------------------------------

def _records(
    text: str,
    year: int | None = None,
    roster: list[str] | None = None,
) -> list[tuple[Any, Any, str, Any]]:
    return [
        (
            row["member_name"],
            row["overtime_date"],
            row["approval_status"],
            row["approver_name"],
        )
        for row in parse_overtime_email(text, reference_year=year, roster=roster)
    ]


def _sample_outlook_msg_bytes() -> bytes:
    """构造一封最小的 Outlook .msg，只给自检用。"""
    import struct

    end = 0xFFFFFFFE
    fatsect = 0xFFFFFFFD
    free = 0xFFFFFFFF
    nostream = 0xFFFFFFFF

    def entry(name: str, etype: int, left: int = nostream, right: int = nostream, child: int = nostream, start: int = 0, size: int = 0) -> bytes:
        raw = name.encode("utf-16le")
        buf = bytearray(128)
        buf[: len(raw)] = raw
        struct.pack_into("<H", buf, 64, len(raw) + 2)
        buf[66] = etype
        buf[67] = 1
        struct.pack_into("<III", buf, 68, left, right, child)
        struct.pack_into("<I", buf, 116, start)
        struct.pack_into("<Q", buf, 120, size)
        return bytes(buf)

    def prop(name: str, payload: bytes) -> bytes:
        return bytes.fromhex(name)[::-1] + b"\x00\x00\x00\x00" + payload[:8].ljust(8, b"\x00")

    def u16(text: str) -> bytes:
        return text.encode("utf-16le") + b"\x00\x00"

    ticks = int((datetime(2025, 9, 12, 10, 0) - datetime(1601, 1, 1)).total_seconds() * 10_000_000)
    streams = [
        ("__properties_version1.0", b"\x00" * 32 + prop("340D0003", (0x40000).to_bytes(4, "little")) + prop("00390040", ticks.to_bytes(8, "little"))),
        ("__substg1.0_001A001F", u16("IPM.Note")),
        ("__substg1.0_0037001F", u16("加班申请")),
        ("__substg1.0_1000001F", u16("请安排张三于2025年9月13日加班。")),
    ]

    def pad64(data: bytes) -> bytes:
        return data + b"\x00" * ((64 - len(data) % 64) % 64)

    mini = b""
    starts: list[int] = []
    for _, data in streams:
        starts.append(len(mini) // 64)
        mini += pad64(data)
    fat = [free] * 128
    fat[0] = fatsect
    fat[1] = 2
    fat[2] = end
    fat[3] = end
    fat[4] = end
    minifat = [end] * 128
    for index, (_, data) in enumerate(streams):
        count = len(pad64(data)) // 64
        start = starts[index]
        for step in range(count - 1):
            minifat[start + step] = start + step + 1
        minifat[start + count - 1] = end
    directory = b"".join(
        [
            entry("Root Entry", 5, child=3, start=4, size=len(mini)),
            entry(streams[0][0], 2, start=starts[0], size=len(streams[0][1])),
            entry(streams[1][0], 2, left=1, start=starts[1], size=len(streams[1][1])),
            entry(streams[2][0], 2, left=2, right=4, start=starts[2], size=len(streams[2][1])),
            entry(streams[3][0], 2, start=starts[3], size=len(streams[3][1])),
        ]
    )
    directory += b"\x00" * (1024 - len(directory))
    header = bytearray(512)
    header[0:8] = bytes.fromhex("D0CF11E0A1B11AE1")
    struct.pack_into("<HHHH", header, 0x18, 0x003E, 0x0003, 0xFFFE, 9)
    struct.pack_into("<H", header, 0x20, 6)
    struct.pack_into("<I", header, 0x2C, 1)
    struct.pack_into("<I", header, 0x30, 1)
    struct.pack_into("<I", header, 0x38, 4096)
    struct.pack_into("<I", header, 0x3C, 3)
    struct.pack_into("<I", header, 0x40, 1)
    struct.pack_into("<I", header, 0x44, end)
    struct.pack_into("<I", header, 0x4C, 0)
    fat_bytes = b"".join(struct.pack("<I", item) for item in fat)
    mini_bytes = b"".join(struct.pack("<I", item) for item in minifat)
    stream = mini + b"\x00" * ((512 - len(mini) % 512) % 512)
    return bytes(header) + fat_bytes + directory + mini_bytes + stream


def run_self_check() -> None:
    assert extract_dates("加班日期：2025年9月13日、2025年9月14日", 2026) == [
        "2025-09-13",
        "2025-09-14",
    ]
    assert extract_dates("加班日期：2025-09-13、14", 2026) == [
        "2025-09-13",
        "2025-09-14",
    ]
    assert extract_dates("9月13日-15日加班", 2025) == [
        "2025-09-13",
        "2025-09-14",
        "2025-09-15",
    ]
    assert extract_dates("请于2025年9月12日前审批，2025年9月13日加班", 2026) == [
        "2025-09-13"
    ]
    # 发送时间不应被当成加班日
    assert extract_dates("发送时间：2025-09-12\n请安排加班", 2026) == []
    # 文件名里的点号日期、带「日」、紧凑日期；人天小数不是日期
    assert extract_dates(
        "1.875人天-学而思培优-朱洋洋-1.24鸿蒙化项目支撑员工加班申请",
        2026,
    ) == ["2026-01-24"]
    assert extract_dates("加班1.5天", 2026) == []
    assert extract_dates("云视听小电视10.18日加班申请", 2025) == ["2025-10-18"]
    assert extract_dates("加班邮件_冯浩男_20251231_1", 2026) == ["2025-12-31"]
    assert extract_dates("9 月 20 日加班申请", 2025) == ["2025-09-20"]
    # ASCII 连字符和「号」
    assert extract_dates("11-22号加班", 2025) == ["2025-11-22"]
    assert extract_dates("加班日期：2025-11-22号", 2026) == ["2025-11-22"]
    assert extract_dates("11月22号加班", 2025) == ["2025-11-22"]
    assert extract_dates("11月22号、23号加班", 2025) == ["2025-11-22", "2025-11-23"]
    assert extract_dates("11月22号-11月24号加班", 2025) == [
        "2025-11-22",
        "2025-11-23",
        "2025-11-24",
    ]
    assert extract_dates("加班日期：11-22号-11-24号", 2025) == [
        "2025-11-22",
        "2025-11-23",
        "2025-11-24",
    ]
    # Content-ID 里的 6-16 不是日期；同一封邮件里的真日期还要留下
    assert extract_dates(
        "请安排加班 cid:image001.png@01DC4A2B.F6-16A3D4E5",
        2025,
    ) == []
    assert extract_dates(
        "Content-ID: <A1B2C3D4-E5F6-16A7-8901-23456789ABCD> 申请加班",
        2025,
    ) == []
    assert extract_dates(
        "申请11-22号加班 cid:image001.png@01DC4A2B.F6-16A3D4",
        2025,
    ) == ["2025-11-22"]

    roster = ["张三", "李四", "麻孟博", "郁金潮", "付权", "刘鸿威"]

    approved = """同意，辛苦了

-----原始邮件-----
发件人：小雨
收件人：冯老师
抄送：王经理
发送时间：2025-09-12 18:02

冯老师好：
因项目上线，申请张三、李四于2025年9月13日、2025年9月14日加班。
请审批。
"""
    got = _records(approved, 2026, roster)
    expect = [
        ("张三", "2025-09-13", "APPROVED", "冯老师"),
        ("张三", "2025-09-14", "APPROVED", "冯老师"),
        ("李四", "2025-09-13", "APPROVED", "冯老师"),
        ("李四", "2025-09-14", "APPROVED", "冯老师"),
    ]
    assert got == expect, got

    # 名单外的怪名字不应出现
    weird = """同意

-----原始邮件-----
发件人：小雨
收件人：冯老师
因项目上线，申请项目上线、张三于2025年9月13日加班。
"""
    assert _records(weird, 2026, roster) == [("张三", "2025-09-13", "APPROVED", "冯老师")]

    excel_rows = records_to_excel_rows(
        parse_overtime_email(
            approved,
            2026,
            roster,
            source_path=r"D:\mails",
            source_file="加班申请.eml",
        )
    )
    assert excel_rows[0] == {
        "加班人员名称": "张三",
        "加班日期": "2025-09-13",
        "是否同意加班": "是",
        "回复同意加班的人": "冯老师",
        "邮件所在路径": r"D:\mails",
        "文件名": "加班申请.eml",
    }

    pending = """发件人：小雨
收件人：冯老师
发送时间：2025-09-12 18:02

冯老师：
请同意张三于2025年9月13日加班。
"""
    assert _records(pending, 2026, roster) == [("张三", "2025-09-13", "PENDING", None)]

    rejected = """不同意，本周人手够了。

-----原始邮件-----
发件人：小雨
发送时间：2025-09-12 09:00

请安排麻孟博于2025年11月22日加班1天。
"""
    assert _records(rejected, 2026, roster) == [("麻孟博", "2025-11-22", "REJECTED", None)]

    from_header_approve = """发件人：王总
收件人：小雨

同意加班

-----原始邮件-----
发件人：小雨
收件人：王总
发送时间：2025-09-12

请安排张三于2025年9月13日加班。
"""
    assert _records(from_header_approve, 2026, roster) == [
        ("张三", "2025-09-13", "APPROVED", "王总")
    ]

    # 同意写在较长回复里也应识别
    polite = """您好，

同意加班，辛苦了，请知悉。

-----原始邮件-----
发件人：小雨
收件人：冯老师
请安排张三于2025年9月13日加班。
"""
    assert _records(polite, 2026, roster) == [("张三", "2025-09-13", "APPROVED", "冯老师")]

    plus_one = """+1

发件人：小雨 <xiaoyu@corp.com>
收件人：冯老师
发送时间：2025年9月12日

加班人员：张三、李四
加班日期：2025-09-13至2025-09-15
"""
    assert _records(plus_one, 2026, roster) == [
        ("张三", "2025-09-13", "APPROVED", "冯老师"),
        ("张三", "2025-09-14", "APPROVED", "冯老师"),
        ("张三", "2025-09-15", "APPROVED", "冯老师"),
        ("李四", "2025-09-13", "APPROVED", "冯老师"),
        ("李四", "2025-09-14", "APPROVED", "冯老师"),
        ("李四", "2025-09-15", "APPROVED", "冯老师"),
    ]

    need = """辛苦了

在 2025年9月12日，小雨写道：
需要郁金潮，付权，刘鸿威三人安排本周六（2025年9月13日）加班一天。
"""
    assert _records(need, 2026, roster) == [
        ("郁金潮", "2025-09-13", "APPROVED", None),
        ("付权", "2025-09-13", "APPROVED", None),
        ("刘鸿威", "2025-09-13", "APPROVED", None),
    ]

    override = """不同意

-----原始邮件-----
同意加班

-----原始邮件-----
发件人：小雨
加班人员：张三
加班日期：9月13日、14日
发送时间：2025-09-12
"""
    assert _records(override, 2026, roster) == [
        ("张三", "2025-09-13", "REJECTED", None),
        ("张三", "2025-09-14", "REJECTED", None),
    ]

    deadline = """同意加班

-----原始邮件-----
发件人：冯老师
发送时间：2025-09-10

请于2025年9月12日前审批。张三于2025年9月13日加班。
"""
    assert _records(deadline, 2026, roster) == [("张三", "2025-09-13", "APPROVED", None)]

    missing_name = """同意

-----原始邮件-----
发送时间：2025-09-12
申请于2025年9月13日加班，人员名单待定。
"""
    assert _records(missing_name, 2026, roster) == [(None, "2025-09-13", "APPROVED", None)]

    weekday_only = """同意

-----原始邮件-----
发件人：小雨
发送时间：2025-09-12 10:00

请安排张三本周六加班。
"""
    assert _records(weekday_only, 2026, roster) == [("张三", "2025-09-13", "APPROVED", None)]

    spaced = "同 意 加 班\n\n请 安 排 张 三 于 2 0 2 5 年 9 月 1 3 日 加 班"
    assert _records(spaced, 2026, roster) == [("张三", "2025-09-13", "APPROVED", None)]

    both_days = """发送时间：2025-09-12 10:00

请安排张三本周六周日加班。
"""
    assert _records(both_days, 2026, roster) == [
        ("张三", "2025-09-13", "PENDING", None),
        ("张三", "2025-09-14", "PENDING", None),
    ]

    weekend = """发送时间：2025-09-12 10:00

请安排张三本周末加班。
"""
    assert _records(weekend, 2026, roster) == [
        ("张三", "2025-09-13", "PENDING", None),
        ("张三", "2025-09-14", "PENDING", None),
    ]

    # 日期写在申请段上方的批复里
    upper_date = """同意，就按9月20日加班吧

-----原始邮件-----
发件人：小雨
收件人：冯老师
纵横小说鸿蒙版项目人员申请加班，请审批。
"""
    assert _records(upper_date, 2026, roster) == [(None, "2026-09-20", "APPROVED", "冯老师")]

    # 正文没有人名、日期只在文件名里
    from_name = parse_overtime_email(
        "各位好：\n云视听小电视项目组人员加班，请审批。\n",
        2026,
        ["朱洋洋", "冯浩男", "韩文金"],
        source_file="1.875人天-学而思培优-朱洋洋-1.24鸿蒙化项目支撑员工加班申请.eml",
    )
    assert [(row["member_name"], row["overtime_date"]) for row in from_name] == [
        ("朱洋洋", "2026-01-24")
    ]

    # 正文已经有加班人时，邮件文件名里的人名也要计入
    kept = parse_overtime_email(
        "请安排张三于2025年9月13日加班。",
        2026,
        ["张三", "李四"],
        source_file="李四加班申请.eml",
    )
    assert [(row["member_name"], row["overtime_date"]) for row in kept] == [
        ("张三", "2025-09-13"),
        ("李四", "2025-09-13"),
    ]

    subject_name = """主题：王五9月13日加班申请
请安排张三于2025年9月13日加班。
"""
    assert [row[0] for row in _records(subject_name, 2026, ["张三", "王五", "李四"])] == [
        "张三",
        "王五",
    ]

    long_name = "请安排阿不都热西提于2025年9月13日加班。"
    assert _records(long_name, 2026, ["阿不都热西提"]) == [
        ("阿不都热西提", "2025-09-13", "PENDING", None)
    ]

    assert _records("   ", roster=roster) == [(None, None, "PENDING", None)]
    assert _records("同意，这个方案可以。", roster=roster) == [
        (None, None, "PENDING", None)
    ]

    newest_wins = """同意

-----原始邮件-----
不同意

-----原始邮件-----
发件人：小雨
请安排李四于2025-09-13加班。
"""
    assert _records(newest_wins, 2026, roster) == [
        ("李四", "2025-09-13", "APPROVED", None)
    ]

    # 姓名列：空格、超过 4 个字、中英混合、称谓
    try:
        import pandas as pd
    except ImportError:
        pd = None
    if pd is not None:
        roster_path = Path(__file__).with_name("_tmp_roster_check.xlsx")
        try:
            pd.DataFrame(
                {
                    "姓名": [
                        " 张 三 ",
                        "欧阳·娜娜子",
                        "Tom Smith",
                        "韩文金Kevin",
                        "李四老师",
                        "张三 李四",
                    ]
                }
            ).to_excel(roster_path, index=False)
            loaded = load_roster_names(roster_path)
        finally:
            if roster_path.exists():
                roster_path.unlink()
        assert "张三" in loaded, loaded
        assert "欧阳娜娜子" in loaded, loaded
        assert "Tom Smith" in loaded, loaded
        assert "韩文金" in loaded, loaded
        assert "李四" in loaded, loaded
        assert "李四老师" not in loaded

        sheet = io.BytesIO()
        pd.DataFrame(
            {"姓名": ["张三", "李四"], "加班日期": ["2025-10-18", "2025-10-18"]}
        ).to_excel(sheet, index=False)
        from email.message import EmailMessage

        mail = EmailMessage()
        mail["Subject"] = "加班申请"
        mail["From"] = "小雨"
        mail["To"] = "冯老师"
        mail.set_content("各位好，加班名单见附件，请审批。")
        mail.add_attachment(
            sheet.getvalue(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="名单.xlsx",
        )
        attach_path = Path(__file__).with_name("_tmp_overtime_attach.eml")
        attach_path.write_bytes(mail.as_bytes())
        try:
            attach_text = load_file_text(attach_path) or ""
            assert "张三" in attach_text and "2025-10-18" in attach_text, attach_text
            attach_rows = _records(attach_text, 2026, ["张三", "李四", "冯老师"])
            assert ("张三", "2025-10-18", "PENDING", None) in attach_rows, attach_rows
            assert ("李四", "2025-10-18", "PENDING", None) in attach_rows, attach_rows
        finally:
            if attach_path.exists():
                attach_path.unlink()

    # 内嵌图片里的姓名和日期
    from email.message import EmailMessage
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (900, 220), "white")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(r"C:\Windows\Fonts\msyh.ttc", 36)
    draw.text((20, 40), "加班人员：张三、李四", fill="black", font=font)
    draw.text((20, 110), "加班日期：2025年9月13日", fill="black", font=font)
    png_buf = io.BytesIO()
    image.save(png_buf, format="PNG")
    mail = EmailMessage()
    mail["Subject"] = "加班申请"
    mail["From"] = "小雨"
    mail["To"] = "冯老师"
    mail.set_content("请领导审批，明细见附图。")
    mail.add_attachment(png_buf.getvalue(), maintype="image", subtype="png", filename="table.png")
    eml_path = Path(__file__).with_name("_tmp_overtime_image.eml")
    eml_path.write_bytes(mail.as_bytes())
    try:
        image_text = load_file_text(eml_path) or ""
        assert "张三" in image_text and "2025" in image_text, image_text
        image_rows = _records(image_text, 2026, roster)
        assert ("张三", "2025-09-13", "PENDING", None) in image_rows, image_rows
        assert ("李四", "2025-09-13", "PENDING", None) in image_rows, image_rows
    finally:
        if eml_path.exists():
            eml_path.unlink()

    # 内嵌 PNG 被标成 application/octet-stream 时也要 OCR
    octet = EmailMessage()
    octet["Subject"] = "11.22日加班邮件转发"
    octet["From"] = "小雨"
    octet["To"] = "冯老师"
    octet.set_content("人员见附图。")
    octet.add_attachment(
        png_buf.getvalue(),
        maintype="application",
        subtype="octet-stream",
        filename="image001.png",
    )
    octet_path = Path(__file__).with_name("_tmp_overtime_octet.eml")
    octet_path.write_bytes(octet.as_bytes())
    try:
        octet_text = load_file_text(octet_path) or ""
        assert "张三" in octet_text and "2025" in octet_text, octet_text
        octet_rows = _records(octet_text, 2026, roster)
        assert ("张三", "2025-09-13", "PENDING", None) in octet_rows, octet_rows
    finally:
        if octet_path.exists():
            octet_path.unlink()

    html_table = _html_to_text(
        "<table><tr><td>加班人员</td><td>张&nbsp;三</td><td>李&nbsp;四</td></tr>"
        "<tr><td>加班日期</td><td>11-22号</td></tr></table>"
    )
    assert "&nbsp;" not in html_table, html_table
    html_rows = _records("申请加班\n" + html_table, 2025, roster)
    assert ("张三", "2025-11-22", "PENDING", None) in html_rows, html_rows
    assert ("李四", "2025-11-22", "PENDING", None) in html_rows, html_rows

    msg_path = Path(__file__).with_name("_tmp_overtime_sample.msg")
    msg_path.write_bytes(_sample_outlook_msg_bytes())
    try:
        msg_text = load_file_text(msg_path) or ""
        assert "张三" in msg_text and "加班申请" in msg_text, msg_text
        assert _records(msg_text, 2026, roster) == [("张三", "2025-09-13", "PENDING", None)], _records(msg_text, 2026, roster)
    finally:
        if msg_path.exists():
            msg_path.unlink()

    tmp = Path(__file__).with_name("_tmp_overtime_check.xlsx")
    gap_tmp = Path(__file__).with_name("_tmp_overtime_gap.xlsx")
    try:
        out = export_records_to_excel(
            parse_overtime_email(
                approved,
                2026,
                roster,
                source_path=r"D:\mails",
                source_file="a.eml",
            ),
            tmp,
        )
        assert out.exists() and out.stat().st_size > 0
        from openpyxl import load_workbook

        book = load_workbook(out)
        assert book.sheetnames == ["加班审批", "未完整提取"]
        assert book["加班审批"]["A2"].value == "张三"
        assert book["未完整提取"].max_row == 1

        gap_out = export_records_to_excel(
            parse_overtime_email(
                missing_name,
                2026,
                roster,
                source_path=r"D:\mails",
                source_file="缺姓名.eml",
            ),
            gap_tmp,
        )
        gap_book = load_workbook(gap_out)
        assert gap_book["加班审批"].max_row == 1
        assert gap_book["未完整提取"]["A2"].value == r"D:\mails\缺姓名.eml"
        assert gap_book["未完整提取"]["B2"].value == "未提取到加班人员名称"
    finally:
        if tmp.exists():
            tmp.unlink()
        if gap_tmp.exists():
            gap_tmp.unlink()


def _ask_path(prompt: str, *, must_exist: bool, is_dir: bool = True) -> Path:
    while True:
        raw = input(prompt).strip().strip('"').strip("'")
        if not raw:
            print("路径不能为空，请重新输入。")
            continue
        path = Path(raw).expanduser().resolve()
        if must_exist:
            if is_dir and not path.is_dir():
                print(f"目录不存在：{path}")
                continue
            if not is_dir and not path.is_file():
                print(f"文件不存在：{path}")
                continue
        return path


def _parse_cli(argv: list[str]) -> dict[str, Any]:
    options: dict[str, Any] = {
        "check": False,
        "json": False,
        "input": None,
        "dir": None,
        "output": None,
        "roster": None,
        "recursive": True,
        "workers": None,
    }
    args = list(argv)
    i = 0
    while i < len(args):
        token = args[i]
        if token in {"--check", "-c"}:
            options["check"] = True
        elif token in {"--json", "-j"}:
            options["json"] = True
        elif token in {"--dir", "-d"}:
            i += 1
            options["dir"] = args[i]
        elif token in {"--output", "-o"}:
            i += 1
            options["output"] = args[i]
        elif token in {"--roster", "-r"}:
            i += 1
            options["roster"] = args[i]
        elif token in {"--workers", "-w"}:
            i += 1
            options["workers"] = int(args[i])
        elif token == "--no-recursive":
            options["recursive"] = False
        elif token in {"-", "--"}:
            options["input"] = "-"
        elif not token.startswith("-") and options["input"] is None and options["dir"] is None:
            options["input"] = token
        else:
            raise SystemExit(f"未知参数：{token}")
        i += 1
    return options


def main(argv: list[str] | None = None) -> int:
    options = _parse_cli(sys.argv[1:] if argv is None else argv)

    if options["check"]:
        run_self_check()
        print("self-check ok")
        return 0

    # 无参数：交互模式
    if not options["input"] and not options["dir"]:
        print("=" * 50)
        print("加班邮件提取 → Excel")
        print(
            "加班审批：姓名和日期都有的记录。"
            "未完整提取：没识别到、或姓名/日期缺失的邮件链接和原因。"
        )
        print("=" * 50)
        src = _ask_path("请输入加班邮件所在文件夹：", must_exist=True, is_dir=True)
        roster_file = _ask_path(
            "请输入人员名单.xlsx路径：",
            must_exist=True,
            is_dir=False,
        )
        out = _ask_path(
            "请输入 Excel 保存路径（例如 D:\\加班提取结果.xlsx）：",
            must_exist=False,
            is_dir=False,
        )
        recursive = input("是否递归子文件夹？(Y/n，默认 Y)：").strip().lower()
        recursive = recursive not in {"n", "no", "否"}
        roster = load_roster_matcher(roster_file)
        print(f"已加载人员名单 {len(roster)} 人")
        records, gaps = parse_overtime_directory(src, recursive=recursive, roster=roster)
        path = export_records_to_excel(records, out, gaps)
        print(f"完整记录 {len(records)} 条，未完整 {len(gaps)} 条 → {path}")
        return 0

    roster: RosterMatcher | None = None
    if options["roster"]:
        roster = load_roster_matcher(options["roster"])
        print(f"已加载人员名单 {len(roster)} 人")
    else:
        print("警告：未提供 --roster 人员名单，姓名可能不准确。")

    gaps: list[dict[str, Any]] = []
    if options["dir"]:
        records, gaps = parse_overtime_directory(
            options["dir"],
            recursive=options["recursive"],
            roster=roster,
            workers=options["workers"],
        )
    else:
        if options["input"] in {None, "-", "--"}:
            text = sys.stdin.read()
            source_path = None
            source_file = None
        else:
            input_path = Path(options["input"]).expanduser().resolve()
            text = load_file_text(input_path) or ""
            source_path = str(input_path.parent)
            source_file = input_path.name
        records, gaps = split_records(
            parse_overtime_email(
                text,
                roster=roster,
                source_path=source_path,
                source_file=source_file,
            )
        )

    if options["output"]:
        path = export_records_to_excel(records, options["output"], gaps)
        print(f"完整记录 {len(records)} 条，未完整 {len(gaps)} 条 → {path}")
        return 0

    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    if options["json"]:
        sys.stdout.write(json.dumps(records, ensure_ascii=False, indent=2))
    else:
        default_out = Path.cwd() / "加班审批提取结果.xlsx"
        path = export_records_to_excel(records, default_out, gaps)
        print(f"完整记录 {len(records)} 条，未完整 {len(gaps)} 条 → {path}")
        for row in records_to_excel_rows(records)[:20]:
            print(
                f"{row['加班人员名称']}\t{row['加班日期']}\t"
                f"{row['是否同意加班']}\t{row['回复同意加班的人']}\t"
                f"{row['文件名']}"
            )
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as exc:
        print(f"self-check failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
    except KeyboardInterrupt:
        print("\n已取消。")
        raise SystemExit(130)
