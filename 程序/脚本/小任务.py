# -*- coding: utf-8 -*-
r"""
小任务生成器 —— 给每个孩子单独出一份「他这次要补的东西」。

输出：4-输出\<班级>_<课次>\小任务\<姓名>.docx
用法：
  python 小任务.py --班级 四年级
  python 小任务.py --班级 四年级 --知识点库 L11-18_lib

数据来源（有哪个用哪个，越全越准）：
  1. 名单里的「错题号」  → 精确到题，最准
  2. <答案>_大题扣分_<班级>.json → 只到 5 个大题，也能用
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_OUT = ROOT / "生成结果"
DIR_ANS = ROOT / "课测答案"

GREY = (0x80, 0x80, 0x80)
BLUE = (0x1B, 0x5E, 0xA8)
RED = (0xC0, 0x39, 0x2B)
GREEN = (0x1E, 0x7D, 0x4F)


def load_lib(name: str):
    """加载知识点库模块。"""
    p = Path(__file__).parent / f"知识点库_{name}.py"
    spec = importlib.util.spec_from_file_location(f"lib_{name}", p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def set_font(run, size=10.5, bold=False, color=None, name="微软雅黑"):
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


def para(doc, text="", size=10.5, bold=False, color=None, after=3,
         align=None, indent=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    if text:
        set_font(p.add_run(text), size, bold, color)
    return p


def read_roster(path: Path):
    meta, students = {}, []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("@"):
            k, _, v = s[1:].partition(":")
            meta[k.strip()] = v.strip()
            continue
        parts = re.split(r"[\s,，\t]+", s)
        if not parts:
            continue
        name = parts[0]
        score, wrong = None, []
        for p in parts[1:]:
            # 分数只取第一个数字；后面的都当错题号
            if score is None:
                try:
                    score = float(p.rstrip("分"))
                    continue
                except ValueError:
                    pass
                m = re.match(r"^(-?\d+(?:\.\d+)?)$", p)
                if m:
                    score = float(m.group(1))
                    continue
            wrong += [int(x) for x in re.findall(r"\d+", p)]
        students.append({"姓名": name, "分数": score, "错题号": wrong})
    return meta, students


def main():
    ap = argparse.ArgumentParser(description="每人一份小任务")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--库", default="L11-18", help="知识点库名（知识点库_<名>.py）")
    ap.add_argument("--试运行", action="store_true")
    args = ap.parse_args()

    rp = next((f for f in DIR_ROSTER.glob("*.txt") if args.班级 in f.stem), None)
    if not rp:
        print("找不到名单:", args.班级)
        return
    meta, students = read_roster(rp)
    # 知识点库优先用名单里的 @知识点库，命令行只是兜底
    lib = load_lib(meta.get("知识点库") or args.库)
    lesson = meta.get("课次", "")
    outdir = DIR_OUT / f"{meta.get('班级', rp.stem)}_{lesson}".replace("/", "_")
    taskdir = outdir / "小任务"
    if not args.试运行:
        taskdir.mkdir(parents=True, exist_ok=True)

    # 大题扣分
    sec_ded = {}
    for cand in DIR_ANS.glob(f"*_大题扣分_{meta.get('班级','')}.json"):
        try:
            raw = json.loads(cand.read_text(encoding="utf-8-sig"))
            sec_ded = {k: v for k, v in raw.items() if not k.startswith("_")}
            print("大题扣分:", cand.name)
        except Exception:
            pass
        break

    n = 0
    for st in students:
        name, score = st["姓名"], st["分数"]
        if score is None:
            continue
        # 该生要补的知识点（有共性知识点就全班共用一套）
        _common = [x.strip() for x in
                   re.split(r"[,，、]", meta.get("共性知识点", "")) if x.strip()]
        _per = {}
        for _it in re.split(r"[；;]", meta.get("逐人知识点", "")):
            if "=" in _it:
                _w, _v = _it.split("=", 1)
                _per[_w.strip()] = [x.strip() for x in re.split(r"[,，、]", _v) if x.strip()]
        _cap_all = None
        try:
            _cap_all = int(str(meta.get("任务件数", "")).strip())
        except ValueError:
            pass
        _cap_all = _cap_all or 5
        _mine = list(_per.get(name) or [])
        # 本人弱项不足 3 件时，用共性补足
        for _k in _common:
            if len(_mine) >= 3:
                break
            if _k not in _mine:
                _mine.append(_k)
        kps = []
        if _mine:
            kps = _mine[:_cap_all]
            src = "本人弱项" if _per.get(name) else "全班共性"
        elif st["错题号"]:
            kps = lib.kps_for_wrong(st["错题号"])
            src = "按错题号"
        if not kps and not _mine:
            d = sec_ded.get(name) or {}
            # 按大题扣分从多到少排，优先补失分最多的那块
            secs = sorted(("一二三四五"),
                          key=lambda k: -float(d.get(k, 0) or 0))
            for k in secs:
                if float(d.get(k, 0) or 0) > 0:
                    for kk in lib.kps_for_section(k):
                        if kk not in kps:
                            kps.append(kk)
            src = "按大题扣分"
        if not kps:
            kps = lib.kps_for_section("二")[:2]     # 兜底：给最基础的
            src = "兜底"

        # 控制数量：四年级孩子一次做不完太多，好的少给、弱的多给但也要能完成
        good = score >= float(meta.get("重点线") or 30)
        cap = 4 if good else 6
        try:
            cap = int(str(meta.get("任务件数", "")).strip()) or cap
        except ValueError:
            pass
        if len(kps) > cap:
            kps = kps[:cap]
            src += f"，取前{cap}件"

        short = name[-2:] if len(name) > 2 else name
        doc = Document()
        sec = doc.sections[0]
        sec.top_margin = Cm(1.5); sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(1.8); sec.right_margin = Cm(1.8)
        stl = doc.styles["Normal"]
        stl.font.name = "微软雅黑"; stl.font.size = Pt(10.5)
        stl.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

        para(doc, f"{short}的小任务", 18, True, BLUE, 2, WD_ALIGN_PARAGRAPH.CENTER)
        para(doc, f"{meta.get('课次', '')}　{meta.get('考试名称', '')}", 9.5, False, GREY, 10,
             WD_ALIGN_PARAGRAPH.CENTER)

        if good:
            para(doc, f"{short}你好！这次考得不错！下面只有 {len(kps)} 件小事要再记一下，"
                      f"过一遍就过关啦。", 11, False, None, 8)
        else:
            para(doc, f"{short}你好！这次错的地方老师都帮你整理出来了，"
                      f"一共就 {len(kps)} 件事。一件一件来，不难的。", 11, False, None, 8)

        for i, k in enumerate(kps, 1):
            title, items, _qs = lib.KP[k]
            para(doc, f"第 {i} 件：{title}", 12, True, RED, 3)
            for it in items:
                para(doc, f"　　{it}", 10.5, False, None, 2, indent=0.4)
            para(doc, "　　☐ 读三遍　☐ 让妈妈问一遍，能说出来", 9.5, False, GREEN, 10,
                 indent=0.4)

        para(doc, "————", 10, False, GREY, 6, WD_ALIGN_PARAGRAPH.CENTER)
        para(doc, "做完了让妈妈拍张照发给老师，老师下节课抽你哦～", 10.5, True, BLUE,
             4, WD_ALIGN_PARAGRAPH.CENTER)
        para(doc, "（每天五分钟就够，不用一次做完）", 9, False, GREY, 0,
             WD_ALIGN_PARAGRAPH.CENTER)

        if not args.试运行:
            doc.save(taskdir / f"{name}.docx")
        print(f"  ✓ {name:<6}{score:>6g}  {len(kps)} 件  ({src})")
        n += 1

    print(f"\n{'[试运行] ' if args.试运行 else ''}共 {n} 人 -> {taskdir}")


if __name__ == "__main__":
    main()
