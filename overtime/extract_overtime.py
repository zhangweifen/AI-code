#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从非结构化邮件正文、回复链或 OCR 文本中提取加班申请与审批结果，并导出 Excel。

规则摘要
--------
1. 回复链按「最新在上」切分。只有位于申请上方的明确同意
   （同意 / 同意加班 / +1 / 辛苦了 等）才记为 APPROVED；
   明确拒绝记为 REJECTED；否则 PENDING。
2. 加班人必须出现在「人员名单.xlsx」中，并且出现在邮件正文里；
   不再用宽泛正则乱猜姓名。
3. 只提取与「加班」相关的日期，忽略发送时间、写道时间等噪声日期。
4. 邮件内嵌图片、HTML 里的图片会做 OCR，识别出的姓名和日期与正文一起提取。
5. 「回复同意加班的人」优先取批复段发件人；没有则取申请邮件收件人。
6. Excel 列：加班人员名称、加班日期、是否同意加班、回复同意加班的人、邮件所在路径、文件名。

用法
----
    py -3 extract_overtime.py --check
    py -3 extract_overtime.py --dir 邮件文件夹 --roster 人员名单.xlsx -o 结果.xlsx
    py -3 extract_overtime.py          # 交互式：邮件目录 + 人员名单 + 导出路径
"""

from __future__ import annotations

import base64
import hashlib
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
_TEXT_LIKE_EXTENSIONS = {
    ".eml",
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

_NAME = r"[\u4e00-\u9fff]{2,4}"
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
    r"(?:(20\d{2})\s*[年./\-]\s*)?(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*日?"
    r"\s*(?:至|到|~|～|—|–|－)\s*"
    r"(?:(20\d{2})\s*[年./\-]\s*)?(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*日?"
)
_DATE_SAME_MONTH = re.compile(
    r"(?:(20\d{2})\s*年\s*)?(\d{1,2})\s*月\s*(\d{1,2})\s*日?\s*[-~～—–至到]\s*(\d{1,2})\s*日"
)
_DATE_LISTED = re.compile(
    r"(?:(20\d{2})\s*年\s*)?(\d{1,2})\s*月\s*"
    r"(\d{1,2}\s*日(?:\s*[、,，和与及]\s*\d{1,2}\s*日)+)"
)
_DATE_NUM_LISTED = re.compile(
    r"(20\d{2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*"
    r"(\d{1,2})(?!\d)"
    r"((?:\s*[、,，和与及]\s*(?!\d{4})\d{1,2}(?!\d))+)"
)
_DATE_FULL = re.compile(r"(20\d{2})\s*[年./\-]\s*(\d{1,2})\s*[月./\-]\s*(\d{1,2})\s*日?")
_DATE_MONTH_DAY = re.compile(r"(?<!\d)(\d{1,2})\s*月\s*(\d{1,2})\s*日")
_DATE_META_LINE = re.compile(r"发送时间|发件时间|写道|Sent|Date\s*[:：]", re.I)
_DATE_LABEL_LINE = re.compile(r"加班日期|加班时间|加班日|值班日期")
_DATE_OT_WORD = re.compile(r"加班|值班")
_REL_WEEKDAY = re.compile(
    r"(?:(?:本周|这周|下周|周)([一二三四五六日天]).{0,24}加班|"
    r"加班.{0,24}(?:本周|这周|下周|周)([一二三四五六日天]))"
)
_APP_CUT = re.compile(
    r"(加班人员|加班人|加班名单|申请|安排|需要).{0,10}加班|于.{0,20}加班"
)
_HAS_DATE_TOKEN = re.compile(r"20\d{2}|[0-9]{1,2}\s*月")
_CONTEXT_NEAR_NAME = re.compile(r"加班|值班|安排|申请|需要|人员|名单")


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
    if not matched:
        return None
    try:
        return datetime(int(matched.group(1)), int(matched.group(2)), int(matched.group(3)))
    except ValueError:
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
        # 直接合并申请段及以下历史，避免二次全量姓名/日期抽取
        application = "\n".join(bodies[newest_app:])
        if not _HAS_DATE_TOKEN.search(application) and newest_app + 1 < len(bodies):
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


def _date_in_overtime_context(text: str, start: int, end: int) -> bool:
    """日期必须靠近加班语境，避免把发送时间、写道时间当成加班日。"""
    line_start = text.rfind("\n", 0, start) + 1
    line_end = text.find("\n", end)
    if line_end < 0:
        line_end = len(text)
    line = text[line_start:line_end]
    if _DATE_META_LINE.search(line):
        return False
    if _DATE_LABEL_LINE.search(line):
        return True
    window = text[max(0, start - 40) : min(len(text), end + 40)]
    if _DATE_OT_WORD.search(window):
        return True
    sentence = text[max(0, start - 80) : min(len(text), end + 80)]
    return bool(_DATE_OT_WORD.search(sentence))


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
        days = [int(n) for n in re.findall(r"(\d{1,2})\s*日", matched.group(3))]
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


def _relative_weekdays(text: str, anchor: datetime | None) -> list[str]:
    """仅当「本周六…加班」这类句式时才换算，避免「本周人手够了」造日期。"""
    if anchor is None:
        return []
    found: list[str] = []
    for matched in _REL_WEEKDAY.finditer(text):
        span = text[matched.start() : matched.end() + 24]
        if re.search(r"20\d{2}|\d{1,2}\s*月\s*\d{1,2}\s*日", span):
            continue
        day_char = matched.group(1) or matched.group(2)
        prefix = "下周" if "下周" in matched.group(0) else "本周"
        weekday = _WEEKDAY[day_char]
        monday = anchor - timedelta(days=anchor.weekday())
        if prefix == "下周":
            monday += timedelta(days=7)
        target = monday + timedelta(days=weekday)
        found.append(target.strftime("%Y-%m-%d"))
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

class RosterMatcher:
    """人员名单一次编译，多封邮件复用。"""

    __slots__ = ("names", "pattern")

    def __init__(self, names: list[str]):
        # 长名优先，交替匹配时先吃长名
        uniq = sorted(set(names), key=len, reverse=True)
        self.names = uniq
        if uniq:
            self.pattern = re.compile("|".join(re.escape(n) for n in uniq))
        else:
            self.pattern = None

    def __len__(self) -> int:
        return len(self.names)

    def find(self, text: str) -> list[str]:
        if not self.pattern or not text:
            return []
        body = _strip_header_lines(text)
        hits: list[tuple[int, int, str]] = []
        last_end = -1
        # finditer 按文本顺序；同位置长名已因 alternation 优先
        for matched in self.pattern.finditer(body):
            start, end = matched.start(), matched.end()
            if start < last_end:
                continue
            hits.append((start, end, matched.group(0)))
            last_end = end

        found: list[str] = []
        seen: set[str] = set()
        for start, end, name in hits:
            window = body[max(0, start - 36) : end + 36]
            if _CONTEXT_NEAR_NAME.search(window) and name not in seen:
                seen.add(name)
                found.append(name)
        if not found and "加班" in body:
            for _, _, name in hits:
                if name not in seen:
                    seen.add(name)
                    found.append(name)
        return found


def load_roster_names(path: str | Path) -> list[str]:
    """
    从人员名单.xlsx读取姓名。
    优先读列名含 姓名/名字/人员/员工 的列；否则扫描全部单元格中的中文姓名。
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
    if preferred:
        for col in preferred:
            values.extend(str(v).strip() for v in frame[col].tolist() if pd.notna(v))
    else:
        for col in frame.columns:
            values.extend(str(v).strip() for v in frame[col].tolist() if pd.notna(v))

    names: list[str] = []
    seen: set[str] = set()
    for raw in values:
        raw = re.sub(r"\s+", "", raw)
        raw = re.sub(r"(老师|经理|主任|总监|主管|总)$", "", raw)
        if not re.fullmatch(r"[\u4e00-\u9fff]{2,4}", raw):
            continue
        if raw in _STOPWORDS or raw in seen:
            continue
        seen.add(raw)
        names.append(raw)
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

def _reference_year(raw_segments: list[str], application: str, reference_year: int | None) -> tuple[int, datetime | None]:
    anchors = [d for d in (_sent_date(seg) for seg in raw_segments) if d is not None]
    anchor = anchors[-1] if anchors else None  # 最旧一封，通常是申请发出日
    if anchor is not None:
        return anchor.year, anchor
    years = [int(y) for y in re.findall(r"(20\d{2})\s*年", application)]
    if years:
        return years[0], None
    if reference_year:
        return reference_year, None
    return datetime.now().year, None


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


def export_records_to_excel(
    records: list[dict[str, Any]],
    output_path: str | Path,
) -> Path:
    """把提取结果写成 xlsx。依赖 openpyxl / pandas。"""
    try:
        import pandas as pd
    except ImportError as exc:  # pragma: no cover
        raise SystemExit("导出 Excel 需要安装 pandas 与 openpyxl：pip install pandas openpyxl") from exc

    path = Path(output_path).expanduser().resolve()
    if path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")
    path.parent.mkdir(parents=True, exist_ok=True)

    rows = records_to_excel_rows(records)
    frame = pd.DataFrame(rows, columns=list(_EXCEL_COLUMNS))
    frame.to_excel(path, index=False, sheet_name="加班审批")
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
    if "加班" not in normalized and "加班" not in _tighten_cjk(normalized):
        return [_null_record("PENDING", **meta)]

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
    year, anchor = _reference_year(raw_segments, application, reference_year)
    names = _extract_names(application, excluded, roster=roster)
    dates = _extract_overtime_dates(application, year, anchor)
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


def _html_to_text(html: str) -> str:
    text = re.sub(r"(?i)<br\s*/?>", "\n", html)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    return text


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
            try:
                from rapidocr_onnxruntime import RapidOCR

                engine = RapidOCR()
            except Exception:
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
    if width < 80 or height < 40:
        return None
    if width * height < 12_000:
        return None
    if width / max(height, 1) > 12 or height / max(width, 1) > 12:
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


def _message_images(msg) -> list[bytes]:
    blobs: list[bytes] = []
    html_parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            if ctype.startswith("image/"):
                try:
                    data = part.get_payload(decode=True)
                except Exception:
                    data = None
                if isinstance(data, bytes) and data:
                    blobs.append(data)
            elif ctype == "text/html":
                try:
                    html_parts.append(str(part.get_content()))
                except Exception:
                    payload = part.get_payload(decode=True)
                    if isinstance(payload, bytes):
                        html_parts.append(payload.decode("utf-8", errors="ignore"))
    for html in html_parts:
        blobs.extend(_decode_data_uri_images(html))
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


def read_eml(path: Path) -> str:
    with path.open("rb") as fh:
        msg = BytesParser(policy=policy.default).parse(fh)
    parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            if ctype == "text/plain":
                parts.append(str(part.get_content()))
            elif ctype == "text/html" and not parts:
                parts.append(_html_to_text(str(part.get_content())))
    else:
        try:
            content = str(msg.get_content())
        except Exception:
            payload = msg.get_payload(decode=True)
            content = payload.decode("utf-8", errors="ignore") if isinstance(payload, bytes) else ""
        if (msg.get_content_type() or "").lower() == "text/html":
            content = _html_to_text(content)
        parts.append(content)
    subject = msg.get("subject", "") or ""
    body = _append_ocr_text(subject + "\n" + "\n".join(parts), _message_images(msg))
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
    try:
        with path.open("rb") as fh:
            chunk = fh.read(_PEEK_BYTES)
    except OSError:
        return False
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
) -> list[dict[str, Any]]:
    if not _peek_has_overtime(path):
        return []
    text = load_file_text(path)
    if text is None or "加班" not in text:
        return []
    resolved = path.resolve()
    rows: list[dict[str, Any]] = []
    for rec in parse_overtime_email(
        text,
        reference_year=reference_year,
        roster=roster,
        source_path=str(resolved.parent),
        source_file=resolved.name,
    ):
        if rec["member_name"] is None and rec["overtime_date"] is None:
            continue
        rows.append(rec)
    return rows


def parse_overtime_directory(
    src_dir: str | Path,
    *,
    recursive: bool = True,
    reference_year: int | None = None,
    roster: list[str] | RosterMatcher | None = None,
    roster_path: str | Path | None = None,
    workers: int | None = None,
) -> list[dict[str, Any]]:
    """扫描目录下邮件/文本，合并全部加班记录（多线程）。"""
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
        return []

    worker_n = workers if workers and workers > 0 else min(_DEFAULT_WORKERS, len(files))
    all_rows: list[dict[str, Any]] = []
    done = 0
    total = len(files)
    print(f"待处理候选文件：{total}，线程数：{worker_n}")

    with ThreadPoolExecutor(max_workers=worker_n) as pool:
        futures = [
            pool.submit(
                _process_one_file,
                path,
                reference_year=reference_year,
                roster=roster,
            )
            for path in files
        ]
        for fut in as_completed(futures):
            done += 1
            try:
                rows = fut.result()
            except Exception as exc:  # noqa: BLE001
                print(f"[跳过] 处理失败：{exc}")
                rows = []
            all_rows.extend(rows)
            if done % 50 == 0 or done == total:
                print(f"进度 {done}/{total}，已提取 {len(all_rows)} 条")
    return all_rows


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

    tmp = Path(__file__).with_name("_tmp_overtime_check.xlsx")
    try:
        out = export_records_to_excel(
            parse_overtime_email(approved, 2026, roster, source_file="a.eml"),
            tmp,
        )
        assert out.exists() and out.stat().st_size > 0
    finally:
        if tmp.exists():
            tmp.unlink()


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
            "列：加班人员名称 | 加班日期 | 是否同意加班 | "
            "回复同意加班的人 | 邮件所在路径 | 文件名"
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
        records = parse_overtime_directory(src, recursive=recursive, roster=roster)
        path = export_records_to_excel(records, out)
        print(f"共导出 {len(records)} 条 → {path}")
        return 0

    roster: RosterMatcher | None = None
    if options["roster"]:
        roster = load_roster_matcher(options["roster"])
        print(f"已加载人员名单 {len(roster)} 人")
    else:
        print("警告：未提供 --roster 人员名单，姓名可能不准确。")

    if options["dir"]:
        records = parse_overtime_directory(
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
        records = parse_overtime_email(
            text,
            roster=roster,
            source_path=source_path,
            source_file=source_file,
        )

    if options["output"]:
        path = export_records_to_excel(records, options["output"])
        print(f"共导出 {len(records)} 条 → {path}")
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
        path = export_records_to_excel(records, default_out)
        print(f"共导出 {len(records)} 条 → {path}")
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
