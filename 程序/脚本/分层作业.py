# -*- coding: utf-8 -*-

"""

通用练习生成器 —— 全班一份，给缺考/临时补发的孩子用。



  背诵单：本次考到的考点，每个考点列要背的要点

  综合练习：每个考点抽 1-2 道题，凑 10 道，答案单独一页



内容全部来自 `题目/<班级>_<课次>/题库.json`，代码里不写死任何单元。

没题库就不产出 —— 宁可不发，也不发错的。



退出码：0 正常；3 没题库（跳过，不是错误）



用法：

  python 分层作业.py               # 处理 班级名单 里所有有学生的班

  python 分层作业.py --班级 周六10点

"""

from __future__ import annotations



import argparse

import json

import re

import sys

from pathlib import Path



from docx import Document

from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx.oxml import OxmlElement

from docx.oxml.ns import qn

from docx.shared import Cm, Pt, RGBColor



def _fix_stdout() -> None:

    """把 stdout 设成 UTF-8，Windows 控制台也不会因中文乱码崩。



    加了 try：stdout 不一定是 TextIOWrapper（被重定向、被包装、测试里替换过），

    那种情况下 reconfigure 不存在，直接调用会让整个脚本 import 就失败。

    """

    try:

        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    except (AttributeError, ValueError, OSError):

        pass



_fix_stdout()

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _ws import fast_remove, workspace_root, safe_name  # noqa: E402

from roster import strip_id  # noqa: E402

from fonts import docx_font_name  # noqa: E402



ROOT = workspace_root()

DIR_ROSTER = ROOT / "班级名单"

DIR_ANSWER = ROOT / "课测答案"

DIR_OUT = ROOT / "生成结果"



# 不规则动词（本讲涉及）

def set_font(run, name=None, size=10.5, bold=False, color=None):

    name = name or docx_font_name()

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

    st.font.name = docx_font_name()

    st.font.size = Pt(10.5)

    st.element.rPr.rFonts.set(qn("w:eastAsia"), docx_font_name())

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





def make_from_bank(bank: dict, cls: str, lesson: str, outdir: Path):

    """通用练习（全班一份，缺考的孩子补）：按本讲考点生成，不写死任何单元。



    之前这里写死了「现在完成时：return/have/cost + How long/How soon」，

    换个讲次就变成发错题。现在内容全部来自 题目/<班级>_<课次>/题库.json。

    题库为空就不产出 —— 宁可没有，也不发错的。

    """

    if not bank:

        return None, None



    kps = [k for k in bank if (bank[k] or {}).get("背诵") or (bank[k] or {}).get("题")]

    if not kps:

        return None, None



    # ---- 一、背诵要点 ----

    doc = new_doc()

    title(doc, f"{cls} {lesson} 本讲要背的", "按这次考到的考点整理，全班通用")

    n = 0

    for kp in kps:

        pts = (bank[kp] or {}).get("背诵") or []

        if not pts:

            continue

        n += 1

        head(doc, f"{n}. {kp}")

        for ptxt in pts:

            line(doc, f"· {ptxt}", size=10.5, indent=0.4)

    p1 = outdir / "背诵单_巩固版.docx"

    doc.save(str(p1))



    # ---- 二、综合练习：每个考点最多 2 道，凑 10 道 ----

    picked = []

    per = max(1, 10 // max(1, len(kps)))

    for kp in kps:

        for q in ((bank[kp] or {}).get("题") or [])[:per]:

            picked.append((kp, q))

        if len(picked) >= 10:

            break

    picked = picked[:10]

    if not picked:

        return p1, None



    doc2 = new_doc()

    title(doc2, f"{cls} {lesson} 综合练习",

          f"共 {len(picked)} 题　做完翻到最后一页对答案")

    for i, (kp, q) in enumerate(picked, 1):

        line(doc2, f"{i}. {q.get('题干', '')}", size=11)

        for op in (q.get("选项") or []):

            line(doc2, f"　　{op}", size=10.5)

        line(doc2, "我的答案：____________", size=10.5, indent=0.4)



    doc2.add_page_break()

    head(doc2, "答案与解析")

    for i, (kp, q) in enumerate(picked, 1):

        line(doc2, f"{i}. {q.get('答案', '')}　{q.get('解析', '')}", size=10.5)

    p2 = outdir / "专项练习_10题.docx"

    doc2.save(str(p2))

    return p1, p2





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

        p = strip_id(s.split())

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



    skipped = 0          # 因没题库而跳过的班数

    for rp in rosters:

        meta, students = read_roster(rp)

        if not students:

            continue

        cls = meta.get("班级", rp.stem)

        lesson = meta.get("课次", "")

        outdir = DIR_OUT / f"{safe_name(cls)}_{safe_name(lesson)}"

        outdir.mkdir(parents=True, exist_ok=True)



        # 内容全部来自本讲题库，写死在代码里的「现在完成时」那套已经删了

        bank = {}

        bp = ROOT / "题目" / f"{safe_name(cls)}_{safe_name(lesson)}" / "题库.json"

        if bp.exists():

            try:

                _raw = json.loads(bp.read_text(encoding="utf-8-sig"))

                bank = {k: v for k, v in (_raw.get("考点") or _raw).items()

                        if isinstance(v, dict)}

            except Exception:

                bank = {}



        p1, p2 = make_from_bank(bank, cls, lesson, outdir)



        print(f"\n{cls} {lesson}")

        if not p1:

            print("  [跳过] 没有题库（题目/<班级>_<课次>/题库.json），"

                  "通用练习不出——缺考的孩子这次没有东西补。")

            # 清掉上一讲留下的旧文件：归档只看文件在不在，

            # 不清理就会把上一次的通用练习当成本次的归档进去。

            _olds = [outdir / x for x in

                     ("背诵单_巩固版.docx", "专项练习_10题.docx")]

            _olds = [x for x in _olds if x.exists()]

            if _olds:

                fast_remove(*_olds)

                print(f"  （已清掉上一次的 {'、'.join(x.name for x in _olds)}）")

            skipped += 1

            continue

        _used = len([k for k in bank

                     if (bank[k] or {}).get("背诵")])

        print(f"  背诵要点 -> {p1.name}（{_used} 个考点）")

        if p2:

            print(f"  综合练习 -> {p2.name}")



    # 3 = 没题库跳过（不是错误，但要让上层知道「这次没有通用练习」）

    return 3 if skipped else 0





if __name__ == "__main__":

    sys.exit(main())

