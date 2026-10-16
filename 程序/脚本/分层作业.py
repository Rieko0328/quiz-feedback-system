# -*- coding: utf-8 -*-
"""
分层作业生成器 —— 生成两份能直接发家长的 Word：

  背诵单（好生 / 中间生）
      · 本讲单词表（中英对照，留空默写栏）
      · 不规则动词表（原形 → 过去式 → 过去分词）
      · 重点句型（含例句 + 造句练习）

  专项练习单（差生）
      · 选择题 5 道
      · 句型转换 5 道
      · 附参考答案（单独一页，方便家长批）

用法：
  python 分层作业.py               # 处理 1-名单 里所有有学生的班
  python 分层作业.py --班级 周六10点
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_ANSWER = ROOT / "课测答案"
DIR_OUT = ROOT / "生成结果"

# 不规则动词（本讲涉及）
IRREGULAR = [
    ("return", "returned", "returned", "规则动词，加 -ed"),
    ("have", "had", "had", "不规则"),
    ("cost", "cost", "cost", "原形=过去式=过去分词"),
    ("fly", "flew", "flown", "易错：不是 flied"),
    ("sell", "sold", "sold", "易错：不是 selled"),
]

# 本讲重点句型
PATTERNS = [
    ("How long have/has + 主语 + 过去分词?",
     "How long has Jack been a teacher?",
     "提问「持续了多久」，答语用 for + 时间段"),
    ("How soon will + 主语 + 动词原形?",
     "How soon will they have a party?",
     "提问「多久以后」，答语用 in + 时间段"),
    ("主语 + have/has + 过去分词 + for + 时间段",
     "He has lived in Zhengzhou for ten years.",
     "for 后面跟「一段时间」"),
    ("主语 + have/has + 过去分词 + since + 时间点",
     "I have been in China since 2000.",
     "since 后面跟「时间点」"),
]

# 选择题（5 道，练本讲标志词）
CHOICE = [
    ("They ________ already ________ this film.",
     ["have; see", "have; seen", "has; saw", "has; seen"], "B",
     "already 是现在完成时标志词，have + 过去分词"),
    ("My uncle has lived in England ________ forty years.",
     ["in", "for", "since", "on"], "B",
     "forty years 是一段时间，用 for"),
    ("I have been in China ________ 2000.",
     ["for", "since", "in", "from"], "B",
     "2000 是一个时间点，用 since"),
    ("--- ________ will you come back?  --- In two weeks.",
     ["How long", "How soon", "How often", "How many"], "B",
     "答语 In two weeks 表示「多久以后」，用 How soon"),
    ("--- ________ have you been here?  --- For ten years.",
     ["How long", "What time", "When", "Where"], "A",
     "答语 For ten years 表示「持续多久」，用 How long"),
]

# 句型转换（5 道）
# (题干, 参考答案, 要划线的部分｜None 表示不划线)
TRANSFORM = [
    ("I will go to Shanghai next week.（改为一般疑问句并作否定回答）",
     "Will you go to Shanghai next week? No, I won't.", None),
    ("Lucy has already finished her homework.（改为否定句）",
     "Lucy hasn't finished her homework yet.", None),
    ("He had a holiday last month.（改为否定句）",
     "He didn't have a holiday last month.", None),
    ("They will have a party in five hours.（对划线部分提问）",
     "How soon will they have a party?", "in five hours"),
    ("Jack has been a teacher for ten years.（对划线部分提问）",
     "How long has Jack been a teacher?", "for ten years"),
]


def set_font(run, name="微软雅黑", size=10.5, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for a in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(a), name)


def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "微软雅黑"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    return doc


def title(doc, text, sub=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_font(p.add_run(text), size=16, bold=True, color=(0x1B, 0x5E, 0xA8))
    if sub:
        q = doc.add_paragraph()
        q.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_font(q.add_run(sub), size=10, color=(0x80, 0x80, 0x80))


def head(doc, text):
    doc.add_paragraph()
    p = doc.add_paragraph()
    set_font(p.add_run(text), size=12.5, bold=True, color=(0x1B, 0x5E, 0xA8))


def line(doc, text, size=10.5, indent=0.0, color=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    set_font(p.add_run(text), size=size, color=color)
    return p


def read_words(ans_path: Path):
    cand = ans_path.with_name(ans_path.stem + "_单词表.txt")
    words = []
    if not cand.exists():
        return words
    for raw in cand.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith(("#", "@")):
            continue
        parts = [x.strip() for x in s.split("|")]
        if len(parts) >= 2:
            words.append((parts[0], parts[1], parts[2] if len(parts) > 2 else ""))
    return words


def make_recite_sheet(path, cls, lesson, words, level="好"):
    """背诵单：不规则动词 + 重点句型（不背单词）"""
    doc = new_doc()
    tag = "巩固背诵单" if level == "好" else "基础背诵单"
    title(doc, f"{cls} {lesson} {tag}",
          "家长您好：这份材料是给孩子在家背的，不用写、不用打印，"
          "看着屏幕读或背即可。下节课我会抽查。")

    # 一、不规则动词
    head(doc, "一、动词的过去式和过去分词（本讲必背，下节课抽背）")
    line(doc, "要求：看着左边的原形，一口气说出过去式和过去分词。", size=10,
         color=(0x80, 0x80, 0x80))
    doc.add_paragraph()
    for i, (base, past, pp, note) in enumerate(IRREGULAR, 1):
        line(doc, f"{i}. {base}　→　{past}　→　{pp}"
                  + (f"　（{note}）" if note else ""), size=11)

    # 二、重点句型
    head(doc, "二、本讲重点句型（会背 + 会自己造句）")
    for i, (pat, ex, note) in enumerate(PATTERNS, 1):
        line(doc, f"{i}. {pat}", size=11, color=(0x1B, 0x5E, 0xA8))
        line(doc, f"　　例句：{ex}", size=10.5, indent=0.5)
        line(doc, f"　　要点：{note}", size=10, indent=0.5, color=(0x70, 0x70, 0x70))

    doc.add_paragraph()
    head(doc, "三、自己造句（每句写下来，下节课读给我听）")
    for i, (pat, _ex, _n) in enumerate(PATTERNS[:2], 1):
        line(doc, f"{i}. 用「{pat.split('+')[0].strip()}」造一句：", size=10.5)
        line(doc, "　　" + "_" * 46, size=10.5, color=(0xA0, 0xA0, 0xA0))

    doc.save(path)
    return path


def make_practice_sheet(path, cls, lesson):
    """差生专项练习单：5 选择 + 5 句型转换 + 答案页"""
    doc = new_doc()
    title(doc, f"{cls} {lesson} 专项练习",
          "家长您好：这份练习共 10 道题，15 分钟内能完成。"
          "做完拍照发我，我批完反馈给您。答案在最后一页，请做完再对照。")

    line(doc, "姓名：____________　　日期：____________　　用时：______ 分钟", size=10)

    head(doc, "一、单项选择（5 题，把答案写在题号前的括号里）")
    for i, (stem, opts, _ans, _why) in enumerate(CHOICE, 1):
        line(doc, f"(　　) {i}. {stem}", size=10.5)
        for letter, opt in zip("ABCD", opts):
            line(doc, f"　　　　{letter}. {opt}", size=10.5, indent=0.5)

    head(doc, "二、按要求改写句子（5 题）")
    for i, (stem, _ans, ul) in enumerate(TRANSFORM, 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        set_font(p.add_run(f"{i}. "), size=10.5)
        if ul and ul in stem:
            a, _, b = stem.partition(ul)
            set_font(p.add_run(a), size=10.5)
            r = p.add_run(ul)
            set_font(r, size=10.5)
            r.underline = True          # ← 关键：把划线部分真正画出来
            set_font(p.add_run(b), size=10.5)
        else:
            set_font(p.add_run(stem), size=10.5)
        line(doc, "　　" + "_" * 50, size=10.5, color=(0xA0, 0xA0, 0xA0))
        line(doc, "　　" + "_" * 50, size=10.5, color=(0xA0, 0xA0, 0xA0))

    # 答案页
    doc.add_page_break()
    head(doc, "参考答案（家长批改用，请孩子做完再看）")
    line(doc, "一、单项选择", size=11, color=(0xC0, 0x39, 0x2B))
    for i, (_s, _o, ans, why) in enumerate(CHOICE, 1):
        line(doc, f"{i}. {ans}　—— {why}", size=10)
    doc.add_paragraph()
    line(doc, "二、按要求改写句子", size=11, color=(0xC0, 0x39, 0x2B))
    for i, (_s, ans, _u) in enumerate(TRANSFORM, 1):
        line(doc, f"{i}. {ans}", size=10)

    doc.save(path)
    return path


def read_roster(path: Path):
    meta, names = {}, []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("@"):
            k, _, v = s[1:].partition(":")
            meta[k.strip()] = v.strip()
            continue
        p = s.split()
        if len(p) >= 2:
            try:
                names.append((p[0], float(p[1].rstrip("分"))))
            except ValueError:
                pass
    return meta, names


def main():
    ap = argparse.ArgumentParser(description="分层作业生成器")
    ap.add_argument("--班级", default=None)
    args = ap.parse_args()

    rosters = [f for f in sorted(DIR_ROSTER.glob("*.txt"))
               if not f.name.startswith(("_", "~$"))]
    if args.班级:
        rosters = [f for f in rosters if args.班级 in f.stem]
    if not rosters:
        print("没有找到名单文件。")
        return

    answers = [f for f in DIR_ANSWER.glob("*.txt") if not f.name.startswith(("_", "~$"))]
    for rp in rosters:
        meta, students = read_roster(rp)
        if not students:
            continue
        cls = meta.get("班级", rp.stem)
        lesson = meta.get("课次", "")
        full = float(meta.get("满分") or 100)
        try:
            focus = float(meta.get("重点线") or 30)
        except ValueError:
            focus = 30.0

        # 找对应答案文件
        ans = None
        for a in answers:
            if lesson and (lesson.replace(" ", "") in a.stem.replace(" ", "")
                           or a.stem.replace(" ", "") in lesson.replace(" ", "")):
                ans = a
                break
        if ans is None and answers:
            ans = answers[0]
        words = read_words(ans) if ans else []

        outdir = DIR_OUT / f"{cls}_{lesson}".replace("/", "_")
        outdir.mkdir(parents=True, exist_ok=True)

        low = [n for n, sc in students if sc < focus]
        mid_high = [n for n, sc in students if sc >= focus]

        p1 = make_recite_sheet(outdir / "背诵单_巩固版.docx", cls, lesson, words, "好")
        p2 = make_practice_sheet(outdir / "专项练习_10题.docx", cls, lesson)

        print(f"\n{cls} {lesson}")
        print(f"  背诵单 -> {p1.name}（发给 {len(mid_high)} 人：{'、'.join(mid_high)}）")
        print(f"  练习单 -> {p2.name}（发给 {len(low)} 人：{'、'.join(low)}）")
        print(f"  单词表词条：{len(words)} 个")


if __name__ == "__main__":
    main()
