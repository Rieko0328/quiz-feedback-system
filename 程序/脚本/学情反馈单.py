# -*- coding: utf-8 -*-
r"""学情反馈单.py —— 给每个孩子出一张图：试卷分析 + 薄弱考点 + 家长话术。

用法：
    python 学情反馈单.py --班级 五年级A班
    python 学情反馈单.py --班级 五年级A班 --课次 五年级第五讲
输出：
    生成结果\{班级}_{课次}\学情反馈单\{姓名}.png
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_OUT = ROOT / "生成结果"

W, H = 1240, 1754                     # A4 @150dpi
BLUE = (27, 94, 168)
TEAL = (30, 142, 90)
ORANGE = (201, 122, 16)
RED = (192, 57, 43)
GREY = (107, 118, 132)
DARK = (31, 41, 55)
LIGHT = (242, 245, 249)
L_BLUE = (232, 241, 251)
L_GREEN = (236, 247, 241)
L_ORANGE = (255, 248, 236)
L_RED = (253, 238, 236)
WHITE = (255, 255, 255)
LINE = (223, 230, 239)

FONT_PATHS = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
]


def font(size, bold=False):
    for p in ([r"C:\Windows\Fonts\msyhbd.ttc"] if bold else []) + FONT_PATHS:
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


LEVEL = [(0.90, "优秀", TEAL), (0.80, "良好", BLUE), (0.70, "合格", ORANGE),
         (0.60, "待提升", ORANGE), (0.0, "需重点关注", RED)]


def level_of(score, full):
    if not full:
        return "—", GREY
    r = score / full
    for th, name, col in LEVEL:
        if r >= th:
            return name, col
    return "需重点关注", RED


def parse_roster(p: Path):
    meta, studs = {}, []
    for raw in p.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("@"):
            for sep in (":", "："):
                if sep in s:
                    k, v = s[1:].split(sep, 1)
                    meta[k.strip()] = v.strip()
                    break
            continue
        parts = s.split()
        if not parts:
            continue
        try:
            sc = float(parts[1])
        except (IndexError, ValueError):
            sc = None
        studs.append({"姓名": parts[0], "分数": sc,
                      "错题": parts[2] if len(parts) > 2 else ""})
    return meta, studs


def wrap(draw, text, fnt, maxw):
    """按像素宽度折行（中文逐字）。"""
    lines, cur = [], ""
    for ch in text:
        if ch == "\n":
            lines.append(cur)
            cur = ""
            continue
        t = cur + ch
        if draw.textlength(t, font=fnt) > maxw and cur:
            lines.append(cur)
            cur = ch
        else:
            cur = t
    if cur:
        lines.append(cur)
    return lines


def draw_report(nm, info, meta, sections, ded, kps, scripts, out_path,
                classstat=None, advice=None):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    f_title = font(46, True)
    f_sub = font(22)
    f_name = font(64, True)
    f_score = font(60, True)
    f_small = font(22)
    f_h2 = font(30, True)
    f_body = font(25)
    f_body_b = font(25, True)
    f_body2 = font(23)
    f_tiny = font(20)

    # ── 顶部标题条 ──
    d.rectangle([0, 0, W, 150], fill=BLUE)
    d.text((60, 38), "课测学情反馈单", font=f_title, fill=WHITE)
    cls = meta.get("班级", "")
    les = meta.get("课次", "")
    d.text((60, 100), f"{cls}　·　{les}", font=f_sub, fill=(210, 228, 245))
    d.text((W - 60, 100), date.today().strftime("%Y-%m-%d"),
           font=f_sub, fill=(210, 228, 245), anchor="ra")

    # ── 学生 + 分数 ──
    y = 190
    d.text((60, y), nm, font=f_name, fill=DARK)
    sc = info["分数"]
    full = meta.get("满分", "40")
    lv, lvcol = level_of(sc or 0, float(full or 40))
    score_txt = f"{int(sc) if sc == int(sc) else sc}" if sc is not None else "缺考"
    d.text((W - 60, y - 6), score_txt, font=f_score, fill=lvcol, anchor="ra")
    d.text((W - 160, y + 46), f"/ {full}", font=f_sub, fill=GREY, anchor="ra")
    # 档位徽章
    bw = d.textlength(lv, font=f_body_b) + 40
    d.rounded_rectangle([W - 60 - bw, y + 108, W - 60, y + 152],
                        radius=22, fill=lvcol)
    d.text((W - 60 - bw / 2, y + 130), lv, font=f_body_b, fill=WHITE, anchor="mm")
    # 重点线
    focus = meta.get("重点线", "")
    if focus and sc is not None:
        diff = sc - float(focus)
        tip = f"重点线 {focus} 分　·　{'超出' if diff >= 0 else '差'} {abs(diff):.0f} 分"
        d.text((60, y + 110), tip, font=f_small,
               fill=TEAL if diff >= 0 else RED)

    y = 380
    d.line([60, y, W - 60, y], fill=LINE, width=2)

    # ── 试卷分析 ──
    y += 34
    d.text((60, y), "试卷分析", font=f_h2, fill=BLUE)
    y += 52
    maxded = max([ded.get(k, 0) for k, _ in sections] + [1])
    bar_x, bar_w = 300, 520
    for key, name in sections:
        v = ded.get(key, 0) or 0
        d.text((60, y + 4), name, font=f_body, fill=DARK)
        # 底槽
        d.rounded_rectangle([bar_x, y + 8, bar_x + bar_w, y + 34],
                            radius=13, fill=(237, 241, 246))
        if v > 0:
            w = max(24, bar_w * v / maxded)
            col = RED if v >= 4 else (ORANGE if v >= 2 else TEAL)
            d.rounded_rectangle([bar_x, y + 8, bar_x + w, y + 34],
                                radius=13, fill=col)
        txt = f"扣 {int(v) if v == int(v) else v} 分" if v else "全对"
        d.text((bar_x + bar_w + 24, y + 4), txt, font=f_body,
               fill=(RED if v else TEAL))
        y += 52

    y += 16
    d.line([60, y, W - 60, y], fill=LINE, width=2)

    # ── 薄弱考点 ──
    y += 34
    d.text((60, y), "需要补的点", font=f_h2, fill=ORANGE)
    y += 52
    if kps:
        for kp in kps:
            d.ellipse([64, y + 11, 76, y + 23], fill=ORANGE)
            d.text((92, y), kp, font=f_body, fill=DARK)
            y += 42
    else:
        d.text((60, y), "这次没有明显薄弱点，保持就好。", font=f_body, fill=TEAL)
        y += 42

    # ── 班级位置 ──
    if classstat:
        y += 12
        d.rounded_rectangle([60, y, W - 60, y + 100], radius=16, fill=L_BLUE)
        d.text((88, y + 18), "班级位置", font=f_body_b, fill=BLUE)
        d.text((88, y + 56), classstat, font=f_body, fill=DARK)
        y += 116

    y += 16
    d.line([60, y, W - 60, y], fill=LINE, width=2)

    # ── 家长配合建议（放话术前面，位置固定） ──
    if advice:
        y += 32
        d.text((60, y), "家长可以怎么配合", font=f_h2, fill=ORANGE)
        y += 50
        for i, line in enumerate(advice, 1):
            d.text((60, y), f"{i}.", font=f_body_b, fill=ORANGE)
            d.text((96, y), line, font=f_small, fill=DARK)
            y += 38
        y += 8
        d.line([60, y, W - 60, y], fill=LINE, width=2)

    # ── 给家长的话 ──
    y += 34
    d.text((60, y), "给家长的话", font=f_h2, fill=TEAL)
    y += 54
    box_top = y
    inner_w = W - 120 - 56
    lines = []
    for seg in scripts:
        seg = re.sub(r"<<<SPLIT\d>>>", "", seg).strip()
        seg = re.sub(r"—+\s*消息\d\s*—+", "", seg).strip()
        if not seg:
            continue
        for para in seg.split("\n"):
            if not para.strip():
                lines.append("")
                continue
            lines.extend(wrap(d, para.strip(), f_body2, inner_w))
        lines.append("")

    line_h = 34
    box_h = len(lines) * line_h + 56
    box_bottom = min(box_top + box_h, H - 150)
    d.rounded_rectangle([60, box_top, W - 60, box_bottom], radius=18, fill=L_GREEN)
    ty = box_top + 26
    truncated = False
    for ln in lines:
        if ty + line_h > box_bottom - 10:
            truncated = True
            break
        d.text((88, ty), ln, font=f_body2, fill=DARK)
        ty += line_h
    if truncated:
        d.text((W - 96, box_bottom - 40), "…", font=f_body2, fill=GREY, anchor="ra")

    # ── 页脚 ──
    d.line([60, H - 76, W - 60, H - 76], fill=LINE, width=2)
    d.text((60, H - 58), "课测反馈系统自动生成　·　分数和考点已人工核对",
           font=f_tiny, fill=GREY)

    img.save(out_path, "PNG", optimize=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    args = ap.parse_args()

    roster = DIR_ROSTER / f"{args.班级}.txt"
    if not roster.exists():
        print(f"[错] 找不到名单：{roster}")
        return 1
    meta, studs = parse_roster(roster)
    meta["班级"] = args.班级
    lesson = args.课次 or meta.get("课次", "")
    out_dir = DIR_OUT / f"{args.班级}_{lesson}"
    if not out_dir.exists():
        print(f"[错] 找不到生成结果：{out_dir}")
        return 1

    # 大题扣分
    ded_all = {}
    p = out_dir / "大题扣分.json"
    if p.exists():
        try:
            ded_all = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            ded_all = {}

    # 大题名称
    sec_names = [x.strip() for x in (meta.get("大题名称") or "").split(",") if x.strip()]
    keys = ["一", "二", "三", "四", "五"]
    sections = [(keys[i], n) for i, n in enumerate(sec_names[:5])]

    # 逐人知识点
    kp_map = {}
    for item in (meta.get("逐人知识点") or "").split("；"):
        if "=" in item:
            k, v = item.split("=", 1)
            kp_map[k.strip()] = [x.strip() for x in v.split(",") if x.strip()]

    # 家长话术
    script_map = {}
    sp = out_dir / "家长反馈话术.txt"
    if sp.exists():
        txt = sp.read_text(encoding="utf-8-sig")
        for blk in txt.split("------------------------------------------------------------"):
            m = re.search(r"#####\s*(\S+?)\s*\|", blk)
            if not m:
                continue
            name = m.group(1)
            segs = re.split(r"<<<SPLIT\d>>>", blk)
            parts = []
            for s in segs[1:] if len(segs) > 1 else []:
                parts.append(s)
            if not parts:
                parts = [blk]
            script_map[name] = parts

    outdir = out_dir / "学情反馈单"
    outdir.mkdir(parents=True, exist_ok=True)

    # 班级统计：平均分 / 排名
    scored = [x for x in studs if x["分数"] is not None]
    ranked = sorted(scored, key=lambda x: -x["分数"])
    total = len(ranked)
    avg = sum(x["分数"] for x in ranked) / total if total else 0
    stat_map = {}
    for i, x in enumerate(ranked, 1):
        diff = x["分数"] - avg
        trend = "高于" if diff >= 0 else "低于"
        stat_map[x["姓名"]] = (
            f"本次班级平均 {avg:.1f} 分，{x['姓名']}考了 "
            f"{int(x['分数']) if x['分数'] == int(x['分数']) else x['分数']} 分，"
            f"{trend}平均 {abs(diff):.1f} 分　·　"
            f"在 {total} 人里排第 {i} 名")

    ADVICE = [
        "每天陪着过一遍小任务，一次 5 分钟就够，不用额外加量",
        "读课文时让孩子大声念给你听，听得出磕绊的地方就是没掌握的",
        "错的地方不要立刻讲答案，先问一句「你当时怎么想的」",
        "连续两次课测同一个知识点还错，跟我说一声，我课上单独过",
    ]

    n_ok = 0
    for st in studs:
        nm = st["姓名"]
        if st["分数"] is None:
            continue
        ded = ded_all.get(nm) or {}
        if isinstance(ded, dict):
            ded = {k: v for k, v in ded.items() if not k.startswith("_")}
        else:
            ded = {}
        scripts = script_map.get(nm) or [f"{nm}这次考了 {st['分数']} 分。"]
        out = outdir / f"{nm}.png"
        draw_report(nm, st, meta, sections, ded, kp_map.get(nm, []), scripts, out,
                    classstat=stat_map.get(nm, ""), advice=ADVICE)
        n_ok += 1
        print(f"  [{n_ok}] {nm}.png")

    print()
    print(f"完成：{n_ok} 张　-> {outdir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
