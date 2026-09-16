#!/usr/bin/env python3
"""Build 论坛风云: first-person weight folder, UG/PG splits, remaining-count Excel.

Does not rewrite 论坛纷纭 sample texts. Copies them into
「学院名本科部 / 学院名研究生部」subfolders (max 7 each).
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FENYUN = Path("/workspace/files/论坛纷纭")
ROOT = Path("/workspace/files/论坛风云")
WEIGHT = ROOT / "自述权重+"
RETRIEVED = "2026-09-16"
BAR = 5  # 业余印象分析门槛：每格合格第一人称就读条数
PEOPLE_BAR = 3

COLLEGES = ["英语学院", "俄语学院", "高级翻译学院", "国际商学院", "国际关系学院"]

# 自述权重+：只放第一人称就读现场。约稿/机构/备考不进此夹。
WEIGHT_SAMPLES = [
    {
        "college": "英语学院",
        "dept": "本科部",
        "weight": "合格",
        "title": "豆瓣-英院幸存者偏差",
        "handle": "匿名豆瓣楼主",
        "people": "豆瓣英院楼主",
        "src": "论坛纷纭/英语学院/03-本科-豆瓣小组.txt",
        "url": "https://www.douban.com/group/topic/269635697/",
        "note": "在读第一人称，分班考、保送生、2016课改、精读词数。",
    },
    {
        "college": "英语学院",
        "dept": "本科部",
        "weight": "半合格",
        "title": "金榜转载-英专大二课表",
        "handle": "自称大二，未署名",
        "people": "金榜英专大二",
        "src": "论坛纷纭/英语学院/04-本科-金榜教育转载问答.txt",
        "url": "https://m.360jibing.com/news/409412.html",
        "note": "课表细，但转载站、大三以后是观摩。只算半合格。",
    },
    {
        "college": "俄语学院",
        "dept": "本科部",
        "weight": "合格",
        "title": "就读文章5-Whever俄语后悔",
        "handle": "Whever",
        "people": "Whever",
        "src": "论坛纷纭/俄语学院/01-本科-知乎.txt",
        "url": "https://www.zhihu.com/question/64391382/answer/3331909814",
        "note": "仓库第5篇。短，但是俄语本科第一人称。",
    },
    {
        "college": "俄语学院",
        "dept": "本科部",
        "weight": "合格",
        "title": "就读文章6-Whever劝退俄语",
        "handle": "Whever",
        "people": "Whever",
        "src": "论坛纷纭/俄语学院/02-本科-知乎.txt",
        "url": "https://www.zhihu.com/question/597450083/answer/3000551110",
        "note": "同一人第二篇。推荐国关，不是国关学生自述。",
    },
    {
        "college": "俄语学院",
        "dept": "研究生部",
        "weight": "半合格",
        "title": "考研论坛-2018俄语MTI复试现场",
        "handle": "旧帖楼主（站点改版后用户名被导航淹没）",
        "people": "2018俄语MTI复试楼主",
        "src": "论坛纷纭/俄语学院/04-研究生部-考研论坛-kaoyan-com.txt",
        "url": "http://bbs.kaoyan.com/t8648738p1",
        "note": "考场第一人称，不是入学后课堂日记。",
    },
    {
        "college": "国际商学院",
        "dept": "本科部",
        "weight": "半合格",
        "title": "经管之家-曹渊国商大四考金专",
        "handle": "caoyuanibsbfsu / 曹渊",
        "people": "曹渊",
        "src": "论坛纷纭/国际商学院/01-本科-研究生部-经管之家-人大经济论坛.txt",
        "url": "https://bbs.pinggu.org/thread-4577562-1-1.html",
        "note": "身份是国商大四，正文几乎全是考研科目。",
    },
    {
        "college": "国际商学院",
        "dept": "本科部",
        "weight": "半合格",
        "title": "经管之家-lu-llaby66北外金融本科推免",
        "handle": "lu-llaby66",
        "people": "lu-llaby66",
        "src": "论坛纷纭/国际商学院/02-本科-经管之家.txt",
        "url": "https://bbs.pinggu.org/thread-4945124-1-1.html",
        "note": "「我本科就是在北外读金融」，有实习，仍偏保研流程。",
    },
    {
        "college": "国际商学院",
        "dept": "研究生部",
        "weight": "半合格",
        "title": "查字典转载-金专入学后午餐会宿舍CFA",
        "handle": "未署名，总分400",
        "people": "金专400分帖",
        "src": "论坛纷纭/国际商学院/04-研究生部-查字典公务员网转载考研经验.txt",
        "url": "https://gwy.chazidian.com/news385566/",
        "note": "「来到北外之后」几句生活，转载站可能美化。",
    },
    {
        "college": "国际商学院",
        "dept": "研究生部",
        "weight": "半合格",
        "title": "文书帮转载-国商学习两个多月",
        "handle": "未署名（与保研帖高度相似，或同一批转载）",
        "people": "文书帮国商两个月",
        "src": "",
        "url": "https://www.wenshubang.com/xindetihui/2895526.html",
        "quote": (
            "总之，来到了梦想中的北外还是挺高兴的。北外地理位置相当不错，在西三环上，旁边的院校有北理、中青政、民大等，"
            "离人大、北大、清华也很近，交通非常便利。北外校园小巧精致，硬件设施一直在改善，图书馆和国际大厦修建的很高大上。"
            "同时，北外学风浓厚，思想包容开放，社团活动相当丰富……"
            "在国商学习了两个多月后，我深刻地感受到老师很认真负责，为学生着想。"
            "学院的学术导师和外界聘请的企业导师都很牛，对今后的学习和工作都有很大的帮助。"
            "总之，来到北外是个不悔的选择。"
        ),
        "note": "续搜新增。入学后短评，转载站，可能与经管之家保研帖同源。",
    },
    {
        "college": "国际关系学院",
        "dept": "本科部",
        "weight": "合格",
        "title": "转载-外交系在读双专业",
        "handle": "未署名",
        "people": "njarts外交系在读",
        "src": "论坛纷纭/国际关系学院/02-本科-教育问答转载.txt",
        "url": "https://www.njarts.cn/a_jiaoyu/202106/216899.html",
        "note": "「我是北外外交系在讀的」。短，转载站。",
    },
    {
        "college": "国际关系学院",
        "dept": "研究生部",
        "weight": "半合格",
        "title": "考研帮-wjb0911外交学毕业回看",
        "handle": "wjb0911",
        "people": "wjb0911",
        "src": "论坛纷纭/国际关系学院/05-研究生部-考研帮.txt",
        "url": "https://yz.kaoyan.com/bfsu/jingyan/566fb7c0d6c56.html",
        "note": "自称国关研究生国家奖学金两年。抒情多、课业少。",
    },
]

EXCLUDED_PROMO = [
    ("英语学院", "本科部", "Insight 康惠阳、吴旨瑨", "学院采访/学生组织约稿，出色个例"),
    ("英语学院", "本科部", "今日头条就业实录", "媒体成稿"),
    ("英语学院", "研究生部", "北鼎/新祥旭英文学硕", "考研机构上岸稿，不是读研日常"),
    ("俄语学院", "本科部", "豆瓣小语种辩护", "学校层，未写俄语学院"),
    ("俄语学院", "研究生部", "新祥旭就业率100%", "机构推销"),
    ("高级翻译学院", "本科部", "推免方案、保研指南", "无普通本科就读对象"),
    ("高级翻译学院", "研究生部", "尚姝辰毕业生专访、李长栓访谈", "官网风采/教师，不是学生论坛自述"),
    ("高级翻译学院", "研究生部", "北鼎法语口译在读", "法语学院，不记入高翻"),
    ("国际商学院", "本科部", "iMAP郑宽、国际本科评价", "留学项目≠统招"),
    ("国际商学院", "本科部", "高顿倪子君双学位申哥大", "ACCA机构学霸稿"),
    ("国际商学院", "研究生部", "新祥旭MIB、Reddit申请帖", "备考/申请，不是就读"),
    ("国际关系学院", "本科部", "中青报2021级外交学、国关优秀毕业生风采", "党媒/学院风采"),
    ("国际关系学院", "本科部", "Whever国关养老", "俄语生外推，作者不在国关"),
    ("国际关系学院", "研究生部", "北鼎复试/真题", "机构备考"),
    ("校级未分院", "—", "就读文章1 MouradBey、7 成就你的梦想", "未写五院之一；第7篇是付费咨询"),
    ("他院", "—", "就读文章2 朝鲜语、3–4 MJC", "亚洲学院/国新，不占五院格子"),
]


def classify_fenyun(name: str) -> str | None:
    stem = name[:-4] if name.endswith(".txt") else name
    if stem.startswith("本夹") or "说明" in stem:
        return None
    if "研究生部" in stem and "本科" not in stem:
        return "研究生部"
    if "本科" in stem:
        return "本科部"
    if "研究生" in stem:
        return "研究生部"
    return None


def copy_fenyun_split():
    """Place existing 论坛纷纭 samples into 学院名本科部 / 学院名研究生部, max 7."""
    summary = {}
    for college in COLLEGES:
        src_dir = FENYUN / college
        buckets = {"本科部": [], "研究生部": []}
        for p in sorted(src_dir.glob("*.txt")):
            if p.name in ("本夹说明.txt",):
                continue
            dept = classify_fenyun(p.name)
            if dept:
                buckets[dept].append(p)
        summary[college] = {}
        for dept, files in buckets.items():
            dest = src_dir / f"{college}{dept}"
            dest.mkdir(parents=True, exist_ok=True)
            taken = files[:7]
            for old in dest.glob("*.txt"):
                if old.name != "本夹说明.txt":
                    old.unlink()
            for f in taken:
                shutil.copy2(f, dest / f.name)
            leftover = max(0, 7 - len(taken))
            (dest / "本夹说明.txt").write_text(
                f"{college}{dept}\n"
                f"从「论坛纷纭/{college}」原取样再放置，不删原文件。\n"
                f"本夹 {len(taken)} 条（上限 7，如有则放满；现有不足不编造）。\n"
                f"距每部 7 条还差 {leftover} 条公开取样。\n"
                f"其中约稿、机构稿、备考帖仍在，不代表自述权重+。\n"
                f"合格第一人称就读见 files/论坛风云/自述权重+/\n",
                encoding="utf-8",
            )
            summary[college][dept] = {"placed": len(taken), "short_of_7": leftover, "files": [f.name for f in taken]}
    return summary


def write_weight_files():
    for college in COLLEGES:
        for dept in ("本科部", "研究生部"):
            folder = WEIGHT / college / f"{college}{dept}"
            folder.mkdir(parents=True, exist_ok=True)
            for old in folder.glob("*.txt"):
                old.unlink()

    by_cell: dict[tuple[str, str], list] = {}
    for sample in WEIGHT_SAMPLES:
        key = (sample["college"], sample["dept"])
        by_cell.setdefault(key, []).append(sample)

    for college in COLLEGES:
        for dept in ("本科部", "研究生部"):
            folder = WEIGHT / college / f"{college}{dept}"
            items = by_cell.get((college, dept), [])[:7]
            for i, sample in enumerate(items, start=1):
                if sample.get("src"):
                    src = Path("/workspace/files") / sample["src"]
                    text = src.read_text(encoding="utf-8") if src.exists() else ""
                else:
                    text = (
                        f"学院：{sample['college']}\n"
                        f"层次：{dept}\n"
                        f"用户/署名：{sample['handle']}\n"
                        f"链接：{sample['url']}\n"
                        f"检索：{RETRIEVED} 续搜\n\n"
                        f"摘录：\n{sample.get('quote', '')}\n"
                    )
                header = (
                    f"自述权重+　{sample['weight']}\n"
                    f"标题：{sample['title']}\n"
                    f"独立叙述标记：{sample['people']}\n"
                    f"说明：{sample['note']}\n"
                    f"{'=' * 40}\n"
                )
                (folder / f"{i:02d}-{sample['weight']}-{sample['title']}.txt").write_text(
                    header + text, encoding="utf-8"
                )
            n = len(items)
            extra = ""
            if college == "高级翻译学院" and dept == "本科部":
                extra = "高翻公开建制以研究生为主，本科格按「不适用」计，不凑 7 条假自述。\n"
            (folder / "本夹说明.txt").write_text(
                f"{college}{dept}　自述权重+\n"
                f"合格/半合格第一人称就读：{n} 条（上限 7，如有才放）。\n"
                f"约稿出色个例、考研机构稿不进此夹。\n"
                f"{extra}"
                f"业余印象门槛 {BAR} 条合格；本夹合格条数见 Excel。\n",
                encoding="utf-8",
            )
    return by_cell


def cell_stats(by_cell):
    rows = []
    for college in COLLEGES:
        for dept, level_name in (("本科部", "本科部"), ("研究生部", "研究生院")):
            items = by_cell.get((college, dept), [])
            ok = [x for x in items if x["weight"] == "合格"]
            half = [x for x in items if x["weight"] == "半合格"]
            people_ok = {x["people"] for x in ok}
            na = college == "高级翻译学院" and dept == "本科部"
            bar = 0 if na else BAR
            remain = 0 if na else max(0, bar - len(ok))
            remain_p = 0 if na else max(0, PEOPLE_BAR - len(people_ok))
            enough = "不适用" if na else ("否" if remain else "是")
            called = "；".join(x["title"] for x in items) or "（无）"
            if na:
                gap = "无普通本科就读对象。余量记 0，不要用保研文件凑数。"
            elif not ok and not half:
                gap = f"合格 0。还差 {remain} 条独立就读自述才够业余印象门槛。"
            elif not ok:
                gap = f"只有半合格。还差 {remain} 条合格自述（门槛 {bar}）。"
            else:
                gap = f"合格 {len(ok)} 条 / {len(people_ok)} 人。还差 {remain} 条、{remain_p} 个独立叙述者。"
            rows.append(
                {
                    "学院": college,
                    "层次": level_name,
                    "业余印象门槛_条": bar,
                    "合格自述_条": len(ok),
                    "半合格_条": len(half),
                    "独立叙述人数_合格": len(people_ok),
                    "还剩合格条数": remain,
                    "还剩独立叙述者": remain_p,
                    "够业余印象吗": enough,
                    "已调用": called,
                    "缺口": gap,
                }
            )
    return rows


def write_excel(rows, fenyun_split):
    wb = Workbook()
    thin = Border(
        left=Side(style="thin", color="B8C4CE"),
        right=Side(style="thin", color="B8C4CE"),
        top=Side(style="thin", color="B8C4CE"),
        bottom=Side(style="thin", color="B8C4CE"),
    )
    head_fill = PatternFill("solid", fgColor="1F4E79")
    head_font = Font(name="微软雅黑", bold=True, color="FFFFFF", size=10)
    wrap = Alignment(wrap_text=True, vertical="center")

    ws = wb.active
    ws.title = "还剩多少"
    headers = [
        "学院",
        "层次",
        "业余印象门槛（合格条）",
        "已调用合格自述（条）",
        "半合格（条）",
        "合格独立叙述人数",
        "还剩合格条数",
        "还剩独立叙述者",
        "够一般业余印象分析吗",
        "已调用的自述类文章",
        "缺口说明",
    ]
    ws.append(headers)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(1, col, h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    keys = [
        "学院",
        "层次",
        "业余印象门槛_条",
        "合格自述_条",
        "半合格_条",
        "独立叙述人数_合格",
        "还剩合格条数",
        "还剩独立叙述者",
        "够业余印象吗",
        "已调用",
        "缺口",
    ]
    for r in rows:
        ws.append([r[k] for k in keys])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=11):
        for cell in row:
            cell.alignment = wrap
            cell.border = thin
            cell.font = Font(name="宋体", size=10)
            if cell.column in (7, 8) and isinstance(cell.value, int) and cell.value > 0:
                cell.fill = PatternFill("solid", fgColor="FCE4D6")
            if cell.column == 9 and cell.value == "否":
                cell.fill = PatternFill("solid", fgColor="F8CBAD")
            if cell.column == 9 and cell.value == "不适用":
                cell.fill = PatternFill("solid", fgColor="D9E2F3")
    widths = [14, 12, 18, 16, 12, 16, 14, 16, 18, 42, 46]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 28
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 48
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:K{ws.max_row}"

    ws2 = wb.create_sheet("论坛纷纭再放置")
    ws2.append(["学院", "层次文件夹", "已再放置条数", "距每部7条还差", "文件"])
    for col in range(1, 6):
        c = ws2.cell(1, col)
        c.fill = head_fill
        c.font = head_font
        c.border = thin
    for college, deps in fenyun_split.items():
        for dept, info in deps.items():
            ws2.append([college, f"{college}{dept}", info["placed"], info["short_of_7"], "；".join(info["files"])])
    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=5):
        for cell in row:
            cell.alignment = wrap
            cell.border = thin
            cell.font = Font(name="宋体", size=10)
    for i, w in enumerate([14, 22, 14, 16, 70], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    ws3 = wb.create_sheet("不计入自述权重")
    ws3.append(["学院", "层次", "材料", "为何不算第一人称就读自述"])
    for col in range(1, 5):
        c = ws3.cell(1, col)
        c.fill = head_fill
        c.font = head_font
        c.border = thin
    for row in EXCLUDED_PROMO:
        ws3.append(list(row))
    for r in ws3.iter_rows(min_row=2, max_row=ws3.max_row, max_col=4):
        for cell in r:
            cell.alignment = wrap
            cell.border = thin
            cell.font = Font(name="宋体", size=10)
    for i, w in enumerate([16, 12, 36, 40], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    ws4 = wb.create_sheet("口径")
    ws4["A1"] = "还剩多少第一人称视角の就读体验"
    ws4["A1"].font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
    notes = [
        f"检索日 {RETRIEVED}（含续搜）。门槛：每院每层次 {BAR} 条「合格」第一人称就读，且尽量 {PEOPLE_BAR} 个不同叙述者，才够做一般业余印象分析。",
        "合格：自称该院该层次在读/刚毕业，写课业、同学、压力、后悔等现场；非约稿、非机构广告。",
        "半合格：身份对得上，但正文是考研/复试/课表转载/入学后两三句。不减门槛。",
        "Insight 采访、中青报、学院风采、高顿学霸、iMAP、北鼎/新祥旭：用户已指出可能是宣传出色个例，全部不进自述权重+。",
        "就读文章目录：只有第5、6篇 Whever 进入俄语学院本科部。第1篇未写学院，第7篇是咨询号，第2–4篇不是这五院。",
        "高翻本科部不适用。续搜仍未找到英院研究生部、高翻研究生部的合格课堂日记。",
        "法语学院口译在读帖不记入高翻。",
        "本表在 files/论坛风云/。五个院的部级文件夹同时存在于 论坛纷纭（原取样再放置）和 论坛风云/自述权重+（过滤后）。",
    ]
    for i, line in enumerate(notes, start=3):
        ws4[f"A{i}"] = line
        ws4[f"A{i}"].alignment = wrap
        ws4.row_dimensions[i].height = 32
    ws4.column_dimensions["A"].width = 110

    out = ROOT / "还剩多少第一人称视角の就读体验.xlsx"
    wb.save(out)
    return out


def main():
    ROOT.mkdir(parents=True, exist_ok=True)
    WEIGHT.mkdir(parents=True, exist_ok=True)
    fenyun_split = copy_fenyun_split()
    by_cell = write_weight_files()
    rows = cell_stats(by_cell)
    xlsx = write_excel(rows, fenyun_split)
    (ROOT / "检索说明.txt").write_text(
        "论坛风云\n"
        "与「论坛纷纭」并列：纷纭保留原取样不动正文；风云做自述加权和缺口表。\n"
        "自述权重+/学院/学院名本科部|研究生部：只放合格或半合格第一人称就读，如有才放，不凑 7。\n"
        "论坛纷纭/学院/学院名本科部|研究生部：原取样再放置，每部最多 7 条。\n"
        "Excel：还剩多少第一人称视角の就读体验.xlsx\n",
        encoding="utf-8",
    )
    (WEIGHT / "口径.txt").write_text(
        "自述权重+ 口径\n"
        "进夹：第一人称、对该院该层次、写就读现场。\n"
        "半合格另标，不充合格条数。\n"
        "不进夹：采访、风采、咨询号、考研机构、留学预科、他院口译。\n",
        encoding="utf-8",
    )
    (ROOT / "目录.json").write_text(
        json.dumps(
            {
                "title": "论坛风云",
                "retrieved": RETRIEVED,
                "excel": "还剩多少第一人称视角の就读体验.xlsx",
                "bar": BAR,
                "cells": rows,
                "fenyun_split": fenyun_split,
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    print(xlsx)
    for r in rows:
        print(f"{r['学院']}{r['层次']}: 合格{r['合格自述_条']} 还剩{r['还剩合格条数']} 够吗={r['够业余印象吗']}")
    for college, deps in fenyun_split.items():
        for dept, info in deps.items():
            print(f"再放置 {college}{dept}: {info['placed']} 差{info['short_of_7']}")


if __name__ == "__main__":
    main()
