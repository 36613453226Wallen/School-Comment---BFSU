#!/usr/bin/env python3
"""Generate a searchable original DOCX from the 北外 MJC 知乎直答 screenshot."""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Cm, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "files" / "口令总表"))
from docx_style import (  # noqa: E402
    INK,
    NAVY,
    SLATE,
    add_bottom_border,
    compact,
    set_run_font,
)

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "mjc-北外新闻与传播专业硕士-知乎直答.docx"


def p(doc, text, *, font="宋体", size=12, bold=False, color=INK, after=6, first=None, align=None):
    para = doc.add_paragraph()
    compact(para, after=after, line=1.3, first_line=first, align=align)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    run = para.add_run(text)
    set_run_font(run, font, size, bold, color)


def h(doc, text):
    para = doc.add_paragraph()
    compact(para, before=10, after=6, line=1.15)
    add_bottom_border(para, "8A4B12", "8")
    run = para.add_run(text)
    set_run_font(run, "黑体", 13, True, NAVY)


def main():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)

    k = doc.add_paragraph()
    compact(k, after=2, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run = k.add_run("合集  ·  mjc 文档转写")
    set_run_font(run, "黑体", 10, True, NAVY)

    t = doc.add_paragraph()
    compact(t, after=8, line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_bottom_border(t)
    run = t.add_run("北外 MJC（新闻与传播专业硕士）知乎直答")
    set_run_font(run, "黑体", 16, True, NAVY)

    for line in [
        "来源：知乎搜索「北外mjc」页的「知乎直答」（22 篇内容 AI 总结）",
        "PDF 快照：2026-09-15 22:23 UTC（FireShot 元数据）",
        "这是检索页 AI 综述，不是学生就读原文，也不是学院官网。数字以截图为准，不另查网页补造。",
        "页脚可见学院站点线索：sijc.bfsu.edu.cn",
    ]:
        p(doc, line, size=9, color=SLATE, after=2, align=WD_ALIGN_PARAGRAPH.CENTER)

    p(
        doc,
        "北京外国语大学 MJC（新闻与传播专业硕士）是依托北外语言优势、以国际传播为特色的新闻传播专业硕士项目，隶属于国际新闻与传播学院，目前是新传考研中热度持续上升的报考选择。",
        first=22,
        after=8,
    )

    h(doc, "方向与学制设置")
    p(doc, "北外 MJC 目前共开设 4 个研究方向，学制和竞争程度各有不同：", first=22)
    p(doc, "新闻与传播｜2 年｜招生人数最多，竞争最激烈，分数线高于其他方向", after=4)
    p(doc, "翻译与国际传播｜3 年｜国际新闻与传播学院+高级翻译学院联合培养，要求较强英语翻译能力，2021 年首次招生", after=4)
    p(doc, "国际艺术传播｜2 年｜2022 年新增方向，报考人数较少", after=4)
    p(doc, "国际出版｜2 年｜2022 年新增方向，报考人数较少，院内有调剂名额", after=8)
    p(doc, "根据 2024 年最新专业目录，北外新传专硕呈现缩招趋势，报考需要注意最新招生计划变化。", first=22, after=8)

    h(doc, "考试特点")
    p(doc, "1. 初试要求", font="黑体", size=12, bold=True, color=NAVY, after=4)
    p(doc, "公共课外国语可选择多语种：除英语外，还可选俄语、日语、法语、德语、西班牙语，对小语种考生友好。", first=22)
    p(
        doc,
        "专业课特色：334 新闻与传播专业综合能力和 440 新闻与传播专业基础有部分题目需要英文作答，但近些年英文题比例在不断缩减，主要集中在名词解释和简答题；出题重点方向为研究方法、国际传播、跨文化传播。",
        first=22,
    )
    p(doc, "备考建议：整理核心理论和概念的英文词汇专题，背诵名词解释时同步记忆英文即可，不用过度担心。", first=22)
    p(doc, "2. 复试规则", font="黑体", size=12, bold=True, color=NAVY, after=4)
    p(doc, "北外复试初复试成绩各占 50%，复试表现对最终录取结果影响较大：", first=22)
    p(doc, "复试分数线：2022 年新闻与传播方向复试线 367 分，2023 年暴涨至 381 分，其他三个方向均为国家线 363 分。", first=22)
    p(doc, "复试流程：通常包括自我介绍、中英问题各一道（抽签抽题）、老师追问三个环节。", first=22)

    h(doc, "报考难度与适配人群")
    p(doc, "历年报录情况", font="黑体", size=12, bold=True, color=NAVY, after=4)
    p(
        doc,
        "2022 年新闻与传播方向拟招 17 人，实际统考录取 28 人（含保研 6 人），报录比约为 15:1，早期报录比约为 10:1，在新传考研中属于中等竞争水平。",
        first=22,
    )
    p(
        doc,
        "从分数线变化能看出，近年北外 MJC 报考人数持续增长，前三年复试线就是国家线，但从 2023 年开始热门方向分数线已经明显高于国家线，竞争强度提升。",
        first=22,
    )
    p(doc, "适合报考人群", font="黑体", size=12, bold=True, color=NAVY, after=4)
    p(
        doc,
        "更推荐英语能力较好，或者本科为英语专业、对国际传播/国际新闻方向感兴趣的考生报考，北外的语言特色能最大化考生的优势。",
        first=22,
    )
    p(
        doc,
        "不建议英语基础薄弱的考生报考，无论初试答题还是复试考核，都对英文能力有一定要求，备考难度会显著提升。",
        first=22,
    )

    h(doc, "培养特色")
    p(
        doc,
        "北外 MJC 实行双导师制：由校内学术导师+校外业界导师联合培养，业界导师多来自新华社、中央广播电视总台、中国日报、中新社等主流媒体，能够帮助学生提升实务能力，拓宽就业渠道。关于导师选择，上岸前辈建议越早联系越好，可以提前锁定心仪导师的名额，也能提前进入课题组适应研究节奏。",
        first=22,
        after=8,
    )

    tail = doc.add_paragraph()
    compact(tail, before=12, after=0, line=1.15)
    add_bottom_border(tail, "D6DEE8", "8")
    run = tail.add_run("检索页：")
    set_run_font(run, "黑体", 10, True, NAVY)
    run = tail.add_run("知乎搜索「北外mjc」（截图未给出可点击的单篇 URL）")
    set_run_font(run, "宋体", 10, color=RGBColor(0x2E, 0x74, 0xB5))

    doc.core_properties.title = "北外 MJC 知乎直答"
    doc.save(OUT)
    print(f"已生成：{OUT}")


if __name__ == "__main__":
    main()
