#!/usr/bin/env python3
"""Generate original and annotated DOCX for article 1 from the FireShot PDF."""

from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "口令总表"))
from docx_style import (  # noqa: E402
    INK,
    NAVY,
    SLATE,
    add_bottom_border,
    compact,
    shade_cell,
    set_cell_margins,
    set_run_font,
)

ROOT = Path(__file__).resolve().parent
SOURCE_URL = "https://www.zhihu.com/question/64391382/answer/1335732679"
TITLE = "考上北外啥感觉？外院香、校园土、回乡还是吹"
ORIGINAL_NAME = f"1{TITLE}.docx"
ANNOTATED_NAME = f"1{TITLE}.分维评注.docx"

PARAS = [
    "高三的时候心心念念，似乎北外的一切都是完美的，当时也是非北外不去的那种想法。",
    "但是真正来了北外，其实也没有特别兴奋与幸福吧。",
    "首先我是东北地区的考生，东北高考难度不大，考上北外也并不是很难，所以到北外比不过一些高考大省或者发达地区的学生也是情理之中的。",
    "有人说上大学就是解放，但是如果在北外，除了一些比较佛系的小众专业，在其他所有专业，都意味着比高中更忙碌和累人的磨炼开始了，特别是自己如果之前并没有许多语言和表达的底子。",
    "在北外读书，最大的体验就是累与无力，特别是对于我们这种读死书考上来的人来说，因为北外非常重视外语，所以输出的环节多且重要，社交恐惧症在这里会非常痛苦，课上的演讲自述，对话，还有每周好几个的录音作业，通选课英语课的pre，一切的一切都要求站在很多人面前开口，对于之前缺乏这种训练的人来说无疑是一种磨炼，但是时间长了，真的会有提高，自己的仪态可以得到矫正，自信也会有所提高（狭义的自信）。",
    "北外一个特点是地域差距，因为教育资源分配不均，高考竞争压力与难度不同，所以各省份考上来的人素质也会有不同，可能发达省区录取的最低分素质也要比不是那么发达的地区考上来的第一名还要优秀一些，这种差距主要体现在英语水平和知识面上，也可能会体现在专业学习上，北外地域歧视并不是很常见，但是因为地域差距形成的成绩差异是存在的。作为东北人，如果不拼命，真的很难拔尖（苦笑）。",
    "之于我自己，考上北外的第一个学期，最大的感受就是身边的同学怎么都这么强，都是各省份的大学霸，而且很多自称“发挥失常”来北外的，刚开始我有点不信，但是后来考了试，发现人家真的是大佬，而且是那种非常强的大佬，确实北外的一些招生手段很高明，以这么垃圾的硬件条件吸引到了这么多优秀的学生，我服！",
    "都说北外漂亮小姐姐很多，但是我自己并没有这种体会，相反，我觉得北外的帅哥还是不少的。北外又不是央戏北影，招上来的女生大多数还是灰头土脸千辛万苦攒实力考上来的，自然不会有什么特殊性。但是北外为什么美女又显得多一些呢？首先是因为基数和比例大，这个不多说，然后是因为北外每年会招一些外国语学校的保送生，这些保送生一般来说压力会比高考生小一些，而且外国语高中有很多家境殷实的高素质的女孩，形象气质自然不会差，这也是北外漂亮小姐姐的一大来源了。一般来说，北外的王牌专业，比如英语、德语、法语，都有非常大的女性拥趸基数，而这些进入北外的女生，很多也是家庭条件非常好，素质很高的女生，在北外学外语，然后出国留学，这也是一个典型的方向，最重要的是，北外毕业生的对口方向，一般来说，对形象气质有一定的要求，所以会发现很多进入北外的女孩子在大学四年会突然出落得很漂亮。",
    "相比之下男生就稍显逊色了，北外的男生配不上女生，这也是事实了，我自己也是男生，但是我也不护短，北外的男生确实没有北外的女生那么突出，也有很多非常油腻猥琐的人（可能我自己就算），但是大多数男生还是非常有教养有礼貌的，也是素质非常高的。",
    "其实考上北外的感觉，不同学院的一定会有不同的感受，因为不同学院的生活状态是不一样的，在这里我也不想分享自己在自己系里的感受，我所说的只是对学校的一种印象。",
    "说实话，来北外还是挺失望的，北外和我心目中的北外相比确实存在落差，北外校园看上去也根本不像是一个重点大学，甚至不像一个大学，非常非常小，而且北外确实存在一些问题，比如非专业课的质量、教务管理、生活管理等等，曾经的我也确实因为这些感到苦恼。",
    "但是，北外已经尽力了，一个211的语言院校，本来就经费有限，收入也不多，甚至只有中山大学的十分之一不到，能创造出这种学习条件已经非常可敬了，毕竟有抱怨的声音，已经说明该有的都有，并且条件已经很可以了。而我自己，虽然在学校过得不是很顺心，压力很大，每天被身边优秀的同学碾压，但是回到家乡，依然会信心满满不遗余力地宣传自己的母校，这，也许就是一种骄傲吧……",
]

POINTS = [
    {
        "name": "录取预期 vs 到校心情",
        "claim": "高三非北外不去，真来了却并不特别兴奋幸福。",
        "support": "开篇两句对照：「似乎北外的一切都是完美的」和「其实也没有特别兴奋与幸福吧」。",
        "reason": "用预期落差给整篇定调，后面所有吐槽都挂在这组对比上。",
    },
    {
        "name": "地域与录取难度",
        "claim": "东北高考相对不难，上北外比不过高考大省/发达地区，情理之中。",
        "support": "「东北高考难度不大，考上北外也并不是很难」。",
        "reason": "先给自己定位，避免把个人感受写成「北外普遍好考」。",
    },
    {
        "name": "学习强度：开口输出",
        "claim": "北外最大体验是累与无力，外语输出环节多，读死书上来的人尤其痛苦，但长期能提高仪态和狭义自信。",
        "support": "演讲自述、对话、每周多个录音作业、通选课英语 pre，「都要求站在很多人面前开口」。",
        "reason": "这是全文最具体的就读体验，可核对、可共鸣，信息密度高。",
    },
    {
        "name": "生源差距（不是地域歧视）",
        "claim": "各省教育资源不均造成素质/成绩差，主要在英语和知识面；歧视不常见，但成绩差存在。东北人不拼命很难拔尖。",
        "support": "「发达省区录取的最低分素质也可能比欠发达地区第一名还要优秀一些」；「地域歧视并不是很常见」。",
        "reason": "把「歧视」和「水平差」拆开，观察比口号细。不过「素质」一词偏笼统。",
    },
    {
        "name": "同学很强 vs 硬件很差",
        "claim": "第一学期同学是各省学霸，自称发挥失常的人考出来真强；硬件垃圾却招到优秀学生。",
        "support": "「以这么垃圾的硬件条件吸引到了这么多优秀的学生，我服！」",
        "reason": "个人观察加情绪判断，硬件差后面有校园描写承接。",
    },
    {
        "name": "性别结构、保送生与形象路径",
        "claim": "美女显得多，主要因为基数、外校保送、王牌语种女多、家境与对口行业对气质有要求；男生整体没女生突出。",
        "support": "外国语学校保送生、英语/德语/法语拥趸、出国留学典型方向、「大学四年突然出落得很漂亮」。",
        "reason": "有机制解释（保送、专业性别比、就业形象），也有外貌/家境概括，容易滑向刻板印象。",
    },
    {
        "name": "学院差异与校园硬件",
        "claim": "不同学院感受不同；作者只谈学校印象，不谈本系。校园小、不像重点大学；非专业课、教务和生活管理有问题。",
        "support": "「我也不想分享自己在自己系里的感受」；「非常非常小」；点名非专业课质量和两类管理。",
        "reason": "自觉限制外推范围，这点诚实。但后面仍用「北外」总称，读者容易忘掉限定。",
    },
    {
        "name": "经费约束与母校骄傲",
        "claim": "211语言院校经费有限（举例「只有中山大学十分之一不到」），条件已经可敬；自己过得不顺心，回乡仍宣传母校。",
        "support": "「北外已经尽力了」到结尾「也许就是一种骄傲吧」。",
        "reason": "把失望收成复杂感情，结构完整。中大经费对比没有出处，只能当作者印象。",
    },
]


def shade_run(run, fill):
    rpr = run._element.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    rpr.append(shd)


def add_blank_lines(doc, n=4):
    for _ in range(n):
        p = doc.add_paragraph("＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿")
        compact(p, after=2, line=1.4)
        for run in p.runs:
            set_run_font(run, "宋体", 11, color=RGBColor(0xBB, 0xBB, 0xBB))


def new_doc():
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    return doc


def add_kicker(doc, text):
    p = doc.add_paragraph()
    compact(p, after=2, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(text)
    set_run_font(run, "黑体", 10, True, NAVY)


def add_title(doc, text):
    p = doc.add_paragraph()
    compact(p, after=8, line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_bottom_border(p)
    run = p.add_run(text)
    set_run_font(run, "黑体", 16, True, NAVY)


def add_meta(doc):
    lines = [
        "作者：Mourad Bey",
        "知乎问题：考上北京外国语大学是什么感觉？",
        "编辑于：2020-07-30 11:38（截图页脚）",
        "首次发布：未知（截图只显示编辑时间，不补造）",
        "PDF 快照：2026-09-15 11:27 UTC（FireShot 元数据）",
        "681 人赞同该回答（界面数字，已从正文剔除）",
    ]
    for line in lines:
        p = doc.add_paragraph()
        compact(p, after=2, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER)
        run = p.add_run(line)
        set_run_font(run, "宋体", 9, color=SLATE)


def add_body(doc):
    for text in PARAS:
        p = doc.add_paragraph()
        compact(p, after=8, line=1.35, first_line=22)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        run = p.add_run(text)
        set_run_font(run, "宋体", 12, color=INK)


def add_link(doc):
    p = doc.add_paragraph()
    compact(p, before=12, after=0, line=1.15)
    add_bottom_border(p, "D6DEE8", "8")
    run = p.add_run("知乎原文：")
    set_run_font(run, "黑体", 10, True, NAVY)
    run = p.add_run(SOURCE_URL)
    set_run_font(run, "宋体", 10, color=RGBColor(0x2E, 0x74, 0xB5))


def add_analysis(doc):
    h = doc.add_paragraph()
    compact(h, before=16, after=8, line=1.1)
    add_bottom_border(h, "8A4B12", "12")
    run = h.add_run("分维评注（副本加写，原文未改）")
    set_run_font(run, "黑体", 14, True, NAVY)

    intro = doc.add_paragraph()
    compact(intro, after=8, line=1.2)
    run = intro.add_run(
        "作者大约从 8 个方面评价北外。下面每条先写分论点，再写支撑句和理由，最后单独看中肯程度。"
        "这是整理，不是替作者改口。"
    )
    set_run_font(run, "宋体", 11, color=INK)

    for i, point in enumerate(POINTS, start=1):
        p = doc.add_paragraph()
        compact(p, before=8, after=4, line=1.15)
        run = p.add_run(f"{i}. {point['name']}")
        set_run_font(run, "黑体", 12, True, NAVY)
        for label, key in (("分论点", "claim"), ("支撑", "support"), ("理由", "reason")):
            q = doc.add_paragraph()
            compact(q, after=3, line=1.2)
            run = q.add_run(f"{label}：")
            set_run_font(run, "黑体", 11, True, NAVY)
            run = q.add_run(point[key])
            set_run_font(run, "宋体", 11, color=INK)

    judge = doc.add_paragraph()
    compact(judge, before=10, after=6, line=1.2)
    run = judge.add_run("分论点是否中肯")
    set_run_font(run, "黑体", 13, True, NAVY)

    verdict = (
        "整体中肯程度：中等偏上，但是「个人体验」和「学校总评」缠在一起。"
        "学习强度、开口训练、同学很强、校园很小、管理有槽点，都有可核对的细节。"
        "「东北相对好考」「保送生家境与气质」「男生配不上女生」是作者视角，不能当成统计结论。"
        "「经费只有中山大学十分之一不到」没有出处，目录里不当事实。"
        "作者自己说不谈本系、只谈学校印象，这一点诚实；读的时候要把外院现象和全校现象分开。"
        "外貌、油腻、家境等判断带冒犯风险，保留原句是为了忠实原文，不是附和。"
    )
    p = doc.add_paragraph()
    compact(p, after=10, line=1.3)
    run = p.add_run(verdict)
    set_run_font(run, "宋体", 11, color=INK)

    blank_title = doc.add_paragraph()
    compact(blank_title, before=8, after=4, line=1.1)
    run = blank_title.add_run("讨论留白（给读者自己写）")
    set_run_font(run, "黑体", 13, True, NAVY)

    prompts = [
        "你觉得哪一条最像事实，哪一条只是心情？",
        "如果作者其实不在外院，哪些段落会失效？",
        "经费对比、外貌与家境，你想反驳还是补充？",
    ]
    for prompt in prompts:
        p = doc.add_paragraph()
        compact(p, after=2, line=1.15)
        run = p.add_run(prompt)
        set_run_font(run, "楷体", 11, color=SLATE)
        add_blank_lines(doc, 3)


def write_original():
    doc = new_doc()
    add_kicker(doc, "就读文章  ·  原文转写")
    add_title(doc, TITLE)
    add_meta(doc)
    add_body(doc)
    add_link(doc)
    path = ROOT / ORIGINAL_NAME
    doc.core_properties.title = TITLE
    doc.save(path)
    print(f"已生成：{path}")
    return path


def write_annotated():
    doc = new_doc()
    add_kicker(doc, "就读文章  ·  分维评注副本")
    add_title(doc, TITLE)
    add_meta(doc)
    add_body(doc)
    add_analysis(doc)
    add_link(doc)
    path = ROOT / ANNOTATED_NAME
    doc.core_properties.title = TITLE + "（分维评注）"
    doc.save(path)
    print(f"已生成：{path}")
    return path


if __name__ == "__main__":
    write_original()
    write_annotated()
