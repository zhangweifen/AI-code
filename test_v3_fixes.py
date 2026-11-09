# -*- coding: utf-8 -*-
"""V3 关键修复点验证测试"""

import importlib
import sys

sys.path.insert(0, r"C:\Users\1\resume_extractor")
v3 = importlib.import_module("resume_extractor_v3")

passed = 0
failed = 0


def check(name, cond, extra=""):
    global passed, failed
    if cond:
        passed += 1
        print(f"  [PASS] {name}")
    else:
        failed += 1
        print(f"  [FAIL] {name}  {extra}")


print("=" * 60)
print("测试一：时间段分隔符（双横杠 / Unicode dash）")
print("=" * 60)

# 1.1 双横杠 2017.06 - - 2018.08（V1 无法匹配）
m = v3.RANGE_RE.search("2017.06 - - 2018.08")
rng = v3.parse_range(m) if m else None
check("双横杠 2017.06 - - 2018.08",
      rng is not None and rng["sy"] == 2017 and rng["ey"] == 2018,
      f"rng={rng}")

# 1.2 en-dash U+2013
m = v3.RANGE_RE.search("2017.06 \u2013 2018.08")
check("en-dash 2017.06 – 2018.08", m is not None)

# 1.3 em-dash U+2014
m = v3.RANGE_RE.search("2017.06 \u2014 2018.08")
check("em-dash 2017.06 — 2018.08", m is not None)

# 1.4 普通单横杠仍可用
m = v3.RANGE_RE.search("2017.06-2018.08")
check("单横杠 2017.06-2018.08", m is not None)

# 1.5 至今
m = v3.RANGE_RE.search("2017.06 至今")
rng = v3.parse_range(m) if m else None
check("至今 2017.06 至今",
      rng is not None and rng["now"] is True)

print()
print("=" * 60)
print("测试二：「工作经验:7年以上」不被误判为 work section")
print("=" * 60)

# V1 只匹配 "7年"/"应届"，遇到 "7年以上"/"应届生" 会漏判
check("工作经验:7年以上", v3.detect_heading("工作经验:7年以上") is None)
check("工作经验:7年以下", v3.detect_heading("工作经验:7年以下") is None)
check("工作经验:7年左右", v3.detect_heading("工作经验:7年左右") is None)
check("工作经验:应届生", v3.detect_heading("工作经验:应届生") is None)
check("工作经验:5.5年", v3.detect_heading("工作经验:5.5年") is None)
check("工作经验:三年以上", v3.detect_heading("工作经验:三年以上") is None)
# 对照：V1 已支持的
check("工作经验:7年 (V1已有)", v3.detect_heading("工作经验:7年") is None)
# 对照：真正的标题仍应识别为 work
check("工作经历 → work", v3.detect_heading("工作经历") == "work")
check("工作经验 → work", v3.detect_heading("工作经验") == "work")
check("工作背景 → work", v3.detect_heading("工作背景") == "work")

print()
print("=" * 60)
print("测试三：毕业时间 - 学校名离时间段较远（双栏 PDF 场景）")
print("=" * 60)

# 模拟周铜生场景：教育 section 为空，学校名在 project section，
# 时间段在 other section，相距 6 行。V1 的 after_school 只查前一行会漏。
text_grad = "\n".join([
    "姓名：周铜生",
    "手机：13800138000",
    "教育经历",        # edu section 标题，但后续无内容 → 教育 section 为空
    "项目经验",        # project section
    "某某系统开发",
    "技术栈：Java/Spring",
    "职责：后端开发",
    "数据库设计",
    "接口联调",
    "性能优化",
    "江西机电职业技术学院",  # 学校名（被双栏错位到 project section）
    "模块六",
    "模块七",
    "模块八",
    "其他信息",        # other section
    "2014.06-2016.03",  # 时间段，离学校名 6 行
])
text_grad = v3.normalize_text(text_grad)
smap_grad = v3.SectionMap(text_grad)
grads = v3.extract_graduation(text_grad, smap_grad)
check("毕业时间包含 2016.03（学校名远距离兜底）",
      "2016.03" in grads, f"grads={grads}")

print()
print("=" * 60)
print("测试四：work section 内项目时间段不被误提为工作")
print("=" * 60)

# V3 新增：前后 PROJECT_GUARD_LINES 行有项目关键词时跳过
# 第4行 2018.01-2018.06 当前行无项目关键词，但前一行有"项目名称"
text_work = "\n".join([
    "工作经历",
    "2017.06-2018.08 某有限公司 前端工程师",
    "项目名称：某某项目",
    "2018.01-2018.06 某功能模块",  # 项目时间段，前后行有项目关键词
])
text_work = v3.normalize_text(text_work)
smap_work = v3.SectionMap(text_work)
works = v3.extract_work(text_work, smap_work)
starts = [w["start"] for w in works]
check("合法工作 2017.06 被提取", "2017.06" in starts, f"works={works}")
check("项目时间段 2018.01 不被误提", "2018.01" not in starts, f"works={works}")

print()
print("=" * 60)
print("测试五：fallback 用强公司信号，教育时间段不被误提为工作")
print("=" * 60)

# 孤立时间段附近只有"工程师/开发"（弱信号），不应被提为工作
text_fallback = "\n".join([
    "姓名：王金祥",
    "手机：13900139000",
    "基本信息",
    "工作经验:7年以上",   # V3 不再误判为 work section
    "自我评价",
    "熟悉 Java 开发",     # "开发"是弱公司信号
    "2015.09-2019.06",   # 教育时间段，附近无强公司信号
])
text_fallback = v3.normalize_text(text_fallback)
smap_fb = v3.SectionMap(text_fallback)
works_fb = v3.extract_work(text_fallback, smap_fb)
check("教育时间段 2015.09 不被 fallback 误提为工作",
      "2015.09" not in [w["start"] for w in works_fb],
      f"works={works_fb}")

print()
print("=" * 60)
print(f"结果：{passed} 通过，{failed} 失败")
print("=" * 60)
sys.exit(0 if failed == 0 else 1)
