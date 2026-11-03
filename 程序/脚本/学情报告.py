# -*- coding: utf-8 -*-
r"""学情报告.py —— 两页 A4 诊断报告（PNG）。

第一版：先出样张看效果。数据结构按最终形态设计，后面接真实数据只改 load() 。

    python 学情报告.py --演示
    python 学情报告.py --演示 --输出 "<输出目录>"
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

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

from PIL import Image, ImageDraw  # noqa: E402
from fonts import pil_font  # noqa: E402

W, H = 1240, 1754
M = 80                      # 页边距
CW = W - 2 * M              # 内容宽 1080

INK = (17, 24, 39)          # 主文字
SUB = (107, 114, 128)       # 次要文字
HINT = (156, 163, 175)      # 提示
LINE = (229, 231, 235)      # 细线
SOFT = (243, 244, 246)      # 极浅底
CARD = (248, 250, 252)      # 卡片底
BLUE = (27, 94, 168)
BLUE_L = (181, 212, 244)
GREEN = (29, 158, 117)
GREEN_L = (212, 240, 228)
AMBER = (239, 159, 39)
AMBER_L = (250, 238, 218)
RED = (192, 57, 43)
RED_L = (252, 235, 235)
GREY_BAR = (211, 209, 199)


class Pen:
    """一层薄封装：省掉每次写 anchor/joint 的重复。"""

    def __init__(self, img):
        self.d = ImageDraw.Draw(img)

    def t(self, xy, s, size=24, color=INK, bold=False, anchor="la"):
        self.d.text(xy, str(s), font=pil_font(size, bold), fill=color, anchor=anchor)

    def tmid(self, x, ymid, s, size=23, color=INK, bold=False):
        """按行的**垂直中点**写一行字。

        表格行高会被压到 44（内容多的时候），而文字原来固定画在 ry+19，
        23 号字实际占 27-28px，19+28 就超出了行高，被下一行那条分隔线切掉半截。
        按中点对齐，行高怎么变都居中。
        """
        self.d.text((x, ymid), str(s), font=pil_font(size, bold),
                    fill=color, anchor="lm")

    def line(self, xy1, xy2, color=LINE, w=2):
        self.d.line([xy1, xy2], fill=color, width=w)

    def rect(self, box, fill=None, outline=None, w=2, r=0):
        x0, y0, x1, y1 = box
        if x1 <= x0:
            x1 = x0 + 1
        if y1 <= y0:
            y1 = y0 + 1
        self.d.rounded_rectangle((x0, y0, x1, y1), radius=r, fill=fill,
                                 outline=outline, width=w)

    def width(self, s, size=24, bold=False):
        box = self.d.textbbox((0, 0), str(s), font=pil_font(size, bold))
        return box[2] - box[0]



def num(x):
    """36.0 -> 36；36.5 -> 36.5"""
    try:
        f = float(x)
        return f"{f:g}"
    except (TypeError, ValueError):
        return str(x)


def fit(s, pen, max_w, size=24, bold=False):
    """超宽就截断加省略号。"""
    s = str(s)
    if pen.width(s, size, bold) <= max_w:
        return s
    while s and pen.width(s + "…", size, bold) > max_w:
        s = s[:-1]
    return s + "…"


# ---------------------------------------------------------------- 演示数据
# 这段只在 --演示 模式用，绝对不是真实业务数据（下面 hc-demo 标记让扫描器跳过）
# hc-demo-start
DEMO = {
    "学校": "新东方郑州学校",
    "科目": "六年级英语",
    "卷名": "语言运用 II-B1 阶段诊断",
    "日期": "2026-10-22",
    "姓名": "陈思潼",
    "分数": 37,
    "满分": 40,
    "排名": 5,
    "总人数": 21,
    "班级平均": 34.2,
    "档位": "优秀",
    "模块": [
        {"名": "单选", "题数": [1, 15], "满分": 15, "得分": 14, "班级率": 0.88},
        {"名": "填空", "题数": [16, 25], "满分": 10, "得分": 9, "班级率": 0.80},
        {"名": "阅读", "题数": [26, 30], "满分": 5, "得分": 5, "班级率": 0.86},
        {"名": "选词", "题数": [31, 35], "满分": 5, "得分": 5, "班级率": 0.82},
        {"名": "补全对话", "题数": [36, 40], "满分": 5, "得分": 4, "班级率": 0.79},
    ],
    "分数段": [("25 以下", 0), ("26-30", 2), ("31-35", 6), ("36-40", 13)],
    "我段": 3,
    "失分": [
        ("第 1 题", "形容词比较级（than 提示）", "单选", 1, "高"),
        ("第 10 题", "形容词比较级（much 修饰）", "单选", 1, "中"),
        ("第 38 题", "交际用语 · 补全对话", "补全对话", 1, "中"),
    ],
    "结论": [
        "形容词和副词占了全卷 62.5% 的分值，思潼在这两块各丢 1 分，阅读和选词全对。",
        "接下来两周只盯一件事：比较级的构成规则，别铺开练。",
    ],
    "寄语": "思潼这次的基础很稳，丢的 3 分全在比较级和一个小对话上，属于记准了就能拿回来的分。"
            "课上我会带你把这几处再过一遍，回家不用加量。",
    "全班薄弱": [("形容词比较级", 12, 21), ("形容词最高级", 9, 21), ("补全对话", 8, 21)],
    "未定位": False,
}
# hc-demo-end


MOD_COLORS = [(27, 94, 168), (29, 158, 117), (239, 159, 39),
              (127, 119, 221), (136, 135, 128)]


def draw_segments(pen: Pen, d: dict, top: int, height: int = 210) -> int:
    """班级分数段柱状图，本人在的那一段标红。返回占用到的底部 y。"""
    pen.t((M, top), "班级分数段分布", 26, INK, True)
    bx, by = M + 40, top + 56
    bw = CW - 80
    bh = height
    pen.line((bx, by + bh), (bx + bw, by + bh), color=(203, 213, 225), w=2)
    segs = d["分数段"]
    mx = max(n for _, n in segs) or 1
    slot = bw / len(segs)
    mine = None
    for i, (label, n) in enumerate(segs):
        h = bh * (n / mx) * 0.82 if n else 0
        x0 = bx + slot * i + slot * 0.24
        x1 = bx + slot * (i + 1) - slot * 0.24
        is_mine = (i == d.get("我段", -1))
        if is_mine:
            mine = (x0, x1)
        if n:
            pen.rect((x0, by + bh - h, x1, by + bh - 2),
                     fill=RED if is_mine else (203, 213, 225), r=6)
        pen.t(((x0 + x1) / 2, by + bh - h - 34), f"{n} 人", 22,
              RED if is_mine else SUB, is_mine, anchor="ma")
        pen.t(((x0 + x1) / 2, by + bh + 18), label, 20, SUB, anchor="ma")
    if mine:
        # 写进柱子内部，避免和柱顶的「N 人」打架
        pen.t(((mine[0] + mine[1]) / 2, by + bh - 46),
              "本人", 20, (255, 255, 255), True, anchor="ma")
    pen.t((bx + bw, by + bh + 48),
          f"全班 {d['总人数']} 人　平均 {d['班级平均']} 分", 20, HINT, anchor="ra")
    return by + bh + 60


def page1(d: dict, pen: Pen) -> None:
    # ---- 抬头 ----
    # 标题不写死：优先用老师配的「考试名称」，没有就用课次，
    # 都没配才退成「诊断报告」。写死「2026 秋季」明年就会露馅。
    _title = (d.get("卷名") or "").strip()
    if not _title or _title.endswith("诊断"):
        _title = "诊断报告"
    else:
        _title = f"{_title} · 诊断报告"
    pen.t((M, 86), _title, 44, INK, True)
    bits = [d.get("学校") or "", d.get("科目") or "", d.get("课次") or ""]
    pen.t((M, 148), " · ".join([b for b in bits if b]), 22, SUB)
    pen.t((W - M, 96), d["日期"], 20, HINT, anchor="ra")
    pen.line((M, 196), (W - M, 196))

    # ---- 姓名 / 分数 / 排名 ----
    y = 240
    pen.t((M, y), d["姓名"], 52, INK, True)
    x = M + pen.width(d["姓名"], 52, True) + 40
    sc = num(d["分数"])
    pen.t((x, y), sc, 52, BLUE, True)
    pen.t((x + pen.width(sc, 52, True) + 10, y + 12), f"/ {num(d['满分'])}", 26, HINT)
    pen.t((M, y + 74), f"班级排名　第 {d['排名']} / {d['总人数']}", 22, SUB)

    # ---- 四个指标卡 ----
    kpis = [
        ("得分率", f"{(d['分数'] / d['满分'] * 100 if d['满分'] else 0):.0f}%",
         BLUE_L, BLUE),
        ("超过平均", f"{d['分数'] - d['班级平均']:+.1f} 分",
         GREEN_L if d["分数"] >= d["班级平均"] else RED_L,
         GREEN if d["分数"] >= d["班级平均"] else RED),
        ("档位", d["档位"], AMBER_L, AMBER),
        ("失分", f"{d['满分'] - d['分数']} 分", RED_L, RED),
    ]
    kw, kg = 168, 22
    kx = W - M - kw * 4 - kg * 3
    left_end = x + pen.width(sc, 52, True) + 60 + pen.width(f"/ {num(d['满分'])}", 26)
    if kx < left_end + 20:                      # 名字太长就缩卡片
        kw = max(132, (W - M - left_end - 40 - kg * 3) / 4)
        kx = W - M - kw * 4 - kg * 3
    for i, (label, val, bg, fg) in enumerate(kpis):
        bx = kx + i * (kw + kg)
        pen.rect((bx, 232, bx + kw, 330), fill=bg, r=16)
        pen.t((bx + kw / 2, 252), label, 20, fg, anchor="ma")
        pen.t((bx + kw / 2, 284), val, 30, fg, True, anchor="ma")

    pen.line((M, 372), (W - M, 372))

    # ---- 饼图：模块分值占比 ----
    pen.t((M, 404), "模块分值占比", 26, INK, True)
    cx, cy, r = M + 170, 610, 150
    # 颜色取模循环：大题超过 5 个时原来只画前 5 块，饼图会有缺口
    colors = [MOD_COLORS[i % len(MOD_COLORS)] for i in range(len(d["模块"]))]
    total = sum(m["满分"] for m in d["模块"])
    start = -90.0
    for m, c in zip(d["模块"], colors):
        sweep = (m["满分"] / total * 360.0) if total else 0.0
        pen.d.pieslice((cx - r, cy - r, cx + r, cy + r), start, start + sweep, fill=c)
        start += sweep
    pen.d.ellipse((cx - 78, cy - 78, cx + 78, cy + 78), fill=(255, 255, 255))
    pen.t((cx, cy - 22), "总分", 22, SUB, anchor="ma")
    pen.t((cx, cy + 22), num(d["满分"]), 36, INK, True, anchor="mm")

    # ---- 图例 ----
    lx = cx + r + 60
    for i, (m, c) in enumerate(zip(d["模块"], colors)):
        ly = cy - 150 + i * 62
        pen.rect((lx, ly, lx + 16, ly + 16), fill=c, r=4)
        pen.t((lx + 30, ly - 4), m["名"], 24, INK)
        pen.t((lx + 30, ly + 26),
              f"{num(m['满分'])} 分 · {m['满分'] / total * 100:.1f}%", 20, SUB)

    # ---- 右侧：得分率对比条 ----
    bx0 = 700
    pen.t((bx0, 404), "各模块得分率", 26, INK, True)
    pen.rect((bx0 + 200, 410, bx0 + 216, 426), fill=BLUE, r=4)
    pen.t((bx0 + 226, 406), "本人", 20, SUB)
    pen.rect((bx0 + 300, 410, bx0 + 316, 426), fill=GREY_BAR, r=4)
    pen.t((bx0 + 326, 406), "班级", 20, SUB)

    bar_x, bar_w = bx0 + 150, 230
    for i, m in enumerate(d["模块"]):
        ly = 470 + i * 66
        nsize = 22 if pen.width(m["名"], 22) <= 132 else 19
        pen.t((bx0, ly - 4), fit(m["名"], pen, 132, nsize), nsize, INK)
        rate = m["得分"] / m["满分"] if m["满分"] else 0
        pen.rect((bar_x, ly, bar_x + bar_w, ly + 18), fill=(238, 240, 243), r=9)
        pen.rect((bar_x, ly, bar_x + max(6, bar_w * rate), ly + 18), fill=BLUE, r=9)
        pen.rect((bar_x, ly + 24, bar_x + bar_w, ly + 34), fill=(238, 240, 243), r=5)
        pen.rect((bar_x, ly + 24, bar_x + max(4, bar_w * m["班级率"]), ly + 34),
                 fill=GREY_BAR, r=5)
        pen.t((bar_x + bar_w + 16, ly - 6), f"{rate * 100:.0f}%", 22, BLUE, True)
    pen.t((bx0, 470 + len(d["模块"]) * 66 + 6),
          "上条为本人，下条为班级平均", 18, HINT)

    pen.line((M, 888), (W - M, 888))

    # ---- 班级分数段 ----
    bottom = draw_segments(pen, d, 912, 292)

    # ---- 结论卡 ----
    cy0 = bottom + 26
    pen.rect((M, cy0, W - M, cy0 + 66 + 40 * len(d["结论"])), fill=CARD,
             outline=LINE, w=2, r=18)
    pen.t((M + 40, cy0 + 26), "一句话结论", 24, BLUE, True)
    for i, s in enumerate(d["结论"]):
        pen.t((M + 40, cy0 + 68 + i * 40), fit(s, pen, CW - 80, 24), 24, (55, 65, 81))

    # ---- 页脚 ----
    pen.line((M, 1640), (W - M, 1640))
    pen.t((M, 1660), f"{d['姓名']} · {d['卷名']}", 18, HINT)
    pen.t((W - M, 1660), "第 1 页 / 共 2 页", 18, HINT, anchor="ra")


def page2(d: dict, pen: Pen) -> None:
    pen.t((M, 86), "成绩与失分明细", 36, INK, True)
    pen.t((M, 134), f"{d['姓名']} · {d['卷名']} · {num(d['分数'])} / {num(d['满分'])}", 22, SUB)
    pen.line((M, 176), (W - M, 176))

    # 没有错题号、也没有全班薄弱 → 内容偏少，行高放大一点，别让下半页空着
    sparse = not d["失分"] and not d["全班薄弱"]
    RH = 84 if sparse else 66

    # ---- 页高预算 ----
    # 大题一多（听力/语音/单选/词汇/句型/完形/阅读/写作就是 8 个），
    # 固定行高会把「老师的话」一路顶到页脚外面，家长拿到的是一句被裁掉的话。
    # 做法：先按固定行高试算，放不下就压行高；压到下限还放不下，才砍「全班薄弱」最后几条。
    n1, n2 = len(d["模块"]), len(d["失分"])
    n3 = len(d["全班薄弱"])
    BOTTOM = 1610                     # 正文底线（页脚线在 1640）

    def _need(rh: int, k: int) -> int:
        h = 310                       # 表一标题(214) + 48 + 表头 48
        h += n1 * rh
        h += 102                      # 表二标题 54 + 表头 48
        h += n2 * rh
        if k:
            h += 176 + k * 62          # 全班薄弱：标题 54 + 副标题 38 + 起点 84
        h += 294                       # 老师的话 46 + 48 + 寄语盒预留 200
        return h

    while RH > 44 and _need(RH, n3) > BOTTOM:
        RH -= 4
    while n3 and _need(RH, n3) > BOTTOM:
        n3 -= 1
    if n3 < len(d["全班薄弱"]):
        d = dict(d, 全班薄弱=d["全班薄弱"][:n3])   # 不污染调用方传进来的 dict

    RX = [M + 24, M + 230, M + 410, M + 570, M + 720, M + 900]
    CX = [M + 24, M + 180, M + 650, M + 830, M + 940]

    drew_tbl = False                 # 表二到底画没画（决定收口线要不要补）

    # ---- 表一：模块明细 ----
    ty = 214
    pen.t((M, ty), "各模块得分", 26, INK, True)
    if d.get("模块估算"):
        # 没读到各大题扣分时，模块分数是按总分等比折算出来的，不是真数据。
        # 不标出来家长会当成卷面上的真分数。
        pen.t((W - M, ty + 4), "分项为折算值（缺各题扣分或分值不全）", 18, HINT, anchor="ra")
    hy = ty + 48
    pen.rect((M, hy, W - M, hy + 48), fill=SOFT, r=8)
    for x, name in zip(RX, ["模块", "得分 / 满分", "得分率", "失分", "班级平均", "差距"]):
        pen.t((x, hy + 13), name, 20, SUB)
    ry = hy + 48
    for i, m in enumerate(d["模块"]):
        h = RH
        if i % 2 == 0:
            pen.rect((M, ry, W - M, ry + h), fill=(252, 252, 253))
        rate = m["得分"] / m["满分"] if m["满分"] else 0
        lost = m["满分"] - m["得分"]
        # 班级平均取不到（别人也没分项数据）时画「—」。
        # 落 0 会画出「全班模块平均 0 分、我比平均高 13 分」，家长一看就知道假。
        _has = bool(m.get("班级有数"))
        avg = (m["班级率"] * m["满分"]) if _has else 0.0
        gap = (m["得分"] - avg) if _has else 0.0
        _m = ry + h / 2
        pen.tmid(RX[0], _m, m["名"], 23, INK)
        pen.tmid(RX[1], _m, f"{num(m['得分'])} / {num(m['满分'])}", 23, INK)
        pen.tmid(RX[2], _m, f"{rate * 100:.0f}%", 23,
                 RED if rate < 0.8 else INK, rate < 0.8)
        pen.tmid(RX[3], _m, f"-{num(round(lost, 1))}" if lost else "0", 23,
                 RED if lost else SUB)
        pen.tmid(RX[4], _m, f"{avg:.1f}" if _has else "—", 23, SUB)
        pen.tmid(RX[5], _m, f"{gap:+.1f}" if _has else "—", 23,
                 (GREEN if gap >= 0 else RED) if _has else SUB)
        ry += h
    pen.line((M, ry), (W - M, ry), color=(203, 213, 225))

    # ---- 表二：失分考点 ----
    # 没有错题号时整块换成「模块级丢分清单」——宁可说粗一档，也不留一张空表头
    ty = ry + 54
    if d["失分"]:
        pen.t((M, ty), "失分考点 · 复习优先级", 26, INK, True)
        hy = ty + 48
        pen.rect((M, hy, W - M, hy + 48), fill=SOFT, r=8)
        for x, name in zip(CX, ["题号", "考点", "模块", "失分", "优先级"]):
            pen.t((x, hy + 13), name, 20, SUB)
        ry = hy + 48
        for i, (no, kp, mod, lost, pri) in enumerate(d["失分"]):
            h = RH
            if i % 2 == 0:
                pen.rect((M, ry, W - M, ry + h), fill=(252, 252, 253))
            _m = ry + h / 2
            pen.tmid(CX[0], _m, no, 23, INK)
            pen.tmid(CX[1], _m, fit(kp, pen, 430, 23), 23, INK)
            pen.tmid(CX[2], _m, fit(mod, pen, 168, 20), 20, SUB)
            pen.tmid(CX[3], _m, f"-{num(round(lost, 1))}", 23, RED)
            tag_bg = RED_L if pri == "高" else AMBER_L
            tag_fg = RED if pri == "高" else AMBER
            _tp = ry + (h - 34) / 2
            pen.rect((CX[4], _tp, CX[4] + 70, _tp + 34), fill=tag_bg, r=8)
            pen.t((CX[4] + 35, _tp + 17), pri, 21, tag_fg, anchor="mm")
            ry += h
    else:
        pen.t((M, ty), "丢分在哪 · 复习优先级", 26, INK, True)
        rows = [m for m in d["模块"] if m["满分"] - m["得分"] > 0]
        if not rows:
            # 一个模块都没丢。既不能画表（只有一个空表头，家长以为数据漏了），
            # 也不能什么都不画（标题下面直接空着）。换成一句绿底结论。
            _ch = 78
            _ty = ty + 44
            pen.rect((M, _ty, W - M, _ty + _ch), fill=GREEN_L, r=14)
            pen.t((M + 36, _ty + 27),
                  "这次每个模块都拿满了，没有需要补的地方。", 23, INK)
            ry = _ty + _ch
        else:
            hy = ty + 48
            pen.rect((M, hy, W - M, hy + 48), fill=SOFT, r=8)
            for x, name in zip(RX[:4], ["模块", "丢分", "得分率", "优先级"]):
                pen.t((x, hy + 13), name, 20, SUB)
            ry = hy + 48
            drew_tbl = True
        rows.sort(key=lambda m: -(m["满分"] - m["得分"]))
        for i, m in enumerate(rows):
            h = RH
            if i % 2 == 0:
                pen.rect((M, ry, W - M, ry + h), fill=(252, 252, 253))
            lost = m["满分"] - m["得分"]
            rate = m["得分"] / m["满分"] if m["满分"] else 0
            pri = "高" if i < 2 else "中"
            _m = ry + h / 2
            pen.tmid(RX[0], _m, fit(m["名"], pen, 190, 23), 23, INK)
            pen.tmid(RX[1], _m, f"-{num(round(lost, 1))}", 23, RED)
            pen.tmid(RX[2], _m, f"{rate * 100:.0f}%", 23,
                     RED if rate < 0.8 else INK, rate < 0.8)
            tag_bg = RED_L if pri == "高" else AMBER_L
            tag_fg = RED if pri == "高" else AMBER
            _tp = ry + (h - 34) / 2
            pen.rect((RX[3], _tp, RX[3] + 70, _tp + 34), fill=tag_bg, r=8)
            pen.t((RX[3] + 35, _tp + 17), pri, 21, tag_fg, anchor="mm")
            ry += h
        drew_tbl = True
    if drew_tbl:                      # 满分时画的是卡片，补线会压在卡片边缘上
        pen.line((M, ry), (W - M, ry), color=(203, 213, 225))

    # ---- 全班普遍失分（不够 2 人同错就不显示这块） ----
    if d["全班薄弱"]:
        ty = ry + 54
        pen.t((M, ty), "本讲全班普遍丢分的地方", 26, INK, True)
        pen.t((M, ty + 38), "不是只有一个人错，这几处全班都栽了", 20, HINT)
        by = ty + 84
        for i, (kp, cnt, tot) in enumerate(d["全班薄弱"]):
            ly = by + i * 62
            pen.t((M + 24, ly), kp, 23, INK)
            bw0, bwid = M + 320, 420
            pen.rect((bw0, ly + 6, bw0 + bwid, ly + 26), fill=(238, 240, 243), r=10)
            pen.rect((bw0, ly + 6, bw0 + bwid * cnt / tot, ly + 26), fill=(133, 183, 235), r=10)
            pen.t((bw0 + bwid + 20, ly + 3), f"{cnt} / {tot} 人", 21, SUB)
        ry = by + len(d["全班薄弱"]) * 62

    # ---- 教师寄语（盒子高度跟着文字走，别留一大块空白） ----
    py = ry + 46
    pen.t((M, py), "老师的话", 26, INK, True)
    words = d["寄语"]
    # 标点不能落在行首。中文排版里「，」跑到下一行第一格是最刺眼的一种错，
    # 一张报告上有三四处，家长不会说「排版不好」，只会觉得「做得糙」。
    NO_HEAD = "，。、；：？！）》」』%,.;:?!)]}"
    line_s, cur, lines = 24, "", []
    for ch in words:
        if pen.width(cur + ch, line_s) > CW - 80:
            if ch in NO_HEAD and len(cur) > 1:
                lines.append(cur[:-1])      # 把上一行最后一个字带下来
                cur = cur[-1] + ch
            else:
                lines.append(cur)
                cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    lines = lines[:6]
    box_h = max(140, 76 + 40 * len(lines))
    pen.rect((M, py + 48, W - M, py + 48 + box_h), fill=CARD, outline=LINE, w=2, r=18)
    for i, s in enumerate(lines):
        pen.t((M + 40, py + 84 + i * 40), s, line_s, (55, 65, 81))

    pen.line((M, 1640), (W - M, 1640))
    pen.t((M, 1660), f"{d['姓名']} · {d['卷名']}", 18, HINT)
    pen.t((W - M, 1660), "第 2 页 / 共 2 页", 18, HINT, anchor="ra")



# ---------------------------------------------------------------- 真实数据
def _read_meta_students(path: Path):
    from roster import parse_roster
    return parse_roster(path)


def _read_answers(path: Path):
    """答案文件 -> ({题号: (考点, 题型)}, meta)"""
    import re
    kp, meta = {}, {}
    if not path.exists():
        return kp, meta
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("@"):
            if ":" in s:
                k, v = s[1:].split(":", 1)
                meta[k.strip()] = v.strip()
            continue
        m = re.match(r"^(\d+)\s*[.、]\s*(.+)$", s)
        if not m:
            continue
        parts = [x.strip() for x in m.group(2).split("|")]
        kp[int(m.group(1))] = (parts[1] if len(parts) > 1 else "",
                               parts[2] if len(parts) > 2 else "")
    return kp, meta


def _split(s):
    import re
    return [x.strip() for x in re.split(r"[,，、]", str(s or "")) if x.strip()]


def load_from_workspace(root: Path, cls: str, student: str, lesson: str = "") -> dict:
    """从工作区算出报告需要的全部数字。缺什么算什么，不编。"""
    import json

    roster = root / "班级名单" / f"{cls}.txt"
    if not roster.exists():
        raise SystemExit(f"[错] 没有名单：{roster}")
    meta, students = _read_meta_students(roster)
    lesson = lesson or meta.get("课次", "")
    full = float(meta.get("满分") or 40)

    # 模块：名称 / 题数 / 分值
    names = _split(meta.get("大题名称"))
    counts = [int(x) for x in _split(meta.get("大题题数")) if x.isdigit()]
    if not names:
        names = ["全卷"]
    if len(counts) != len(names):
        counts = [0] * len(names)
    # 各大题分值：答案文件里写了就用答案的（更权威），其次名单，最后按题数平摊
    _ans = root / "课测答案" / f"{lesson}.txt"
    if not _ans.exists() and (root / "课测答案").is_dir():
        # 课次名和答案文件名不完全一致时（多了日期、少了「五年级」）模糊找一次
        _cands = [q for q in (root / "课测答案").glob("*.txt")
                  if lesson and (lesson in q.stem or q.stem in lesson)]
        if len(_cands) == 1:
            _ans = _cands[0]
            print(f"  [提示] 答案按名字找到：{_cands[0].name}")
        elif not _cands:
            print(f"  [提示] 课测答案/{lesson}.txt 不存在，报告只出到模块级")
    kp_map, ans_meta = _read_answers(_ans)
    v_src = ans_meta.get("大题分值") or meta.get("大题分值") or ""
    vals = [float(x) for x in _split(v_src) if x.replace(".", "").isdigit()]
    _val_est = False
    if len(vals) != len(names):
        # 没给各模块分值（或数量和名字对不上）→ 按题数占比分总分。
        # 这时候「单选 10 分」可能被估成 11.4 分，得分率跟着错，
        # 所以必须吭声：控制台打警告，报告上也要标出来。
        if vals:
            print(f"  [警告] 大题分值给了 {len(vals)} 个、大题名称 {len(names)} 个，"
                  f"对不上 → 模块分值改按题数折算，得分率会略偏。"
                  f"核对名单里的 @大题分值 / @大题名称")
        tot_cnt = sum(counts) or 1
        vals = [full * c / tot_cnt for c in counts] if sum(counts) else [full / len(names)] * len(names)
        _val_est = True
    per_q = [v / c if c else 0 for v, c in zip(vals, counts)]   # 每题分值

    # 扣分表
    ded_path = root / "生成结果" / (f"{cls}_{lesson}" if lesson else cls) / "大题扣分.json"
    ded_all = {}
    if ded_path.exists():
        try:
            raw = json.loads(ded_path.read_text(encoding="utf-8-sig"))
            ded_all = {k: v for k, v in raw.items() if not k.startswith("_")}
        except Exception:
            ded_all = {}
    cn = "一二三四五六七八九十"

    def mod_scores(st):
        """某人各模块得分。优先用扣分表，退回用错题号×每题分值。"""
        name = st["姓名"]
        d = ded_all.get(name)
        out = []
        wrong = set(st.get("错题") or [])
        for i, (v, c) in enumerate(zip(vals, counts)):
            lost = None
            _k = cn[i] if i < len(cn) else str(i + 1)   # 超过 10 个大题也不崩
            if isinstance(d, dict) and _k in d:
                lost = float(d[_k] or 0)
            elif wrong and counts:
                lo = sum(counts[:i]) + 1
                hi = lo + c - 1
                lost = len([q for q in wrong if lo <= q <= hi]) * per_q[i]
            if lost is None:
                out.append(None)
            else:
                out.append(max(0.0, v - lost))
        return out

    scored = [s for s in students if not s["缺考"] and s["分数"] is not None]
    me = next((s for s in students if s["姓名"] == student), None)
    if me is None:
        raise SystemExit(f"[错] 名单里没有这个人：{student}")
    if me["缺考"] or me["分数"] is None:
        raise SystemExit(f"[提示] {student} 这次缺考，没有报告可出")

    ranked = sorted(scored, key=lambda s: -(s["分数"] or 0))
    rank = next(i for i, s in enumerate(ranked, 1) if s["姓名"] == student)
    avg = sum(s["分数"] for s in scored) / len(scored)

    # 模块数据
    mine = mod_scores(me)
    all_mods = [mod_scores(s) for s in scored]
    modules = []
    _est = False
    for i, name in enumerate(names):
        got = mine[i] if i < len(mine) else None
        if got is None:
            # 没有分项数据时按总分等比折算。绝填 0——否则会出现
            # 「满分 40 考了 38，但每个模块都是 0 分」这种自相矛盾的表
            got = (me["分数"] * vals[i] / full) if full else 0.0
            _est = True
        cls_vals = [m[i] for m in all_mods if i < len(m) and m[i] is not None]
        _has_cls = bool(cls_vals and vals[i])
        cls_rate = (sum(cls_vals) / len(cls_vals) / vals[i]) if _has_cls else 0.0
        modules.append({"名": name, "满分": round(vals[i], 1),
                        "得分": round(got, 1), "班级率": round(cls_rate, 3),
                        "班级有数": _has_cls,
                        "题数": counts[i] if i < len(counts) else 0})

    # 分数段
    cuts = [0.60, 0.75, 0.875]
    edges = [0] + [round(full * c) for c in cuts] + [round(full)]
    segs = []
    for i in range(4):
        lo, hi = edges[i], edges[i + 1]
        n = sum(1 for s in scored if (lo <= s["分数"] < hi) or (i == 3 and s["分数"] == hi))
        label = f"{lo:g}-{hi:g}" if i else f"{hi:g} 以下"
        segs.append((label, n))
    # 本人落在哪一段
    my_seg = 0
    for i in range(4):
        lo, hi = edges[i], edges[i + 1]
        if lo <= me["分数"] <= hi and (i == 3 or me["分数"] < hi):
            my_seg = i
            break

    # 考点
    wrong = me.get("错题") or []
    lost_list = []
    for q in wrong:
        kp, _t = kp_map.get(q, ("", ""))
        mi = next((i for i in range(len(names))
                   if sum(counts[:i]) < q <= sum(counts[:i]) + (counts[i] if i < len(counts) else 0)), 0)
        lost_list.append((f"第 {q} 题", kp or "—", names[mi] if mi < len(names) else "—",
                          per_q[mi] if mi < len(per_q) else 1, ""))
    # 优先级：按该题分值（占分多 = 优先级高）
    lost_list.sort(key=lambda x: -x[3])
    lost_list = [(a, b, c, e, "高" if e >= 2 or i < 2 else "中")
                 for i, (a, b, c, e, _) in enumerate(lost_list)]

    # 全班普遍丢分：按考点聚合
    from collections import Counter
    cnt = Counter()
    for s in scored:
        seen = set()
        for q in (s.get("错题") or []):
            kp, _t = kp_map.get(q, ("", ""))
            if kp and kp not in seen:
                seen.add(kp)
                cnt[kp] += 1
    weak = [(k, v, len(scored)) for k, v in cnt.most_common(3) if v >= 2]

    # 结论（规则版；AI 可以覆盖）
    # 按「丢了多少分」排，不按得分率——丢 1 分的模块不该被说成"主要问题"
    worst = sorted(modules, key=lambda m: -(m["满分"] - m["得分"]))[:2]
    # 「没填错题号」不等于「全对」——判据是分数，不是错题列表空不空
    if not wrong and me["分数"] >= full:
        concl = [f"{student}这次全对，{full:g} 分满分，这一讲的内容是真掌握了。",
                 "保持节奏就行，不用额外加练。"]
    else:
        loose = [m for m in modules if m["满分"] - m["得分"] > 0]
        # 口径以总分为准：没有扣分表时，模块失分是按错题号估的，加起来可能
        # 和「满分 − 分数」对不上（卷面写每题 2 分、老师实际扣 1 分就会高估）
        share = full - me["分数"]
        if loose:
            names = "、".join(m["名"] for m in worst
                             if m["满分"] - m["得分"] > 0) or loose[0]["名"]
            desc = f"丢了 {share:g} 分，主要在{names}"
            rest = "其余模块没丢分" if len(loose) < len(modules) else "各模块都有失分"
        else:
            desc, rest = "分项扣分还没对齐", "先把每道错的题补上"
        concl = [f"{student}这次 {me['分数']:g}/{full:g}，{desc}，{rest}。",
                 "先把这几处的规则记准，下次同类题别再错，分就回来了。"]

    return {
        "学校": meta.get("学校", ""),
        "科目": meta.get("科目", ""),
        "卷名": meta.get("考试名称", lesson or "阶段诊断"),
        "日期": meta.get("日期", "") or __import__("datetime").date.today().isoformat(),
        "姓名": student,
        "分数": me["分数"],
        "满分": full,
        "排名": rank,
        "总人数": len(scored),
        "班级平均": round(avg, 1),
        "档位": me.get("档位") or "",
        "模块": modules,
        "分数段": segs, "我段": my_seg,
        "失分": lost_list[:4],
        "错题数": len(wrong),
        "模块估算": _est or _val_est,      # 真实错题数，题量按它算（失分表只显示前 4 行）
        "全班薄弱": weak,
        "结论": concl,
        "寄语": "",
        "未定位": not kp_map,
        "班级": cls,
        "课次": lesson,
    }


def fill_text(d: dict, root: Path, cls: str) -> None:
    """档位、寄语：尽量从现有产物里取，取不到就不写。"""
    from _ws import short_name

    full = float(d.get("满分") or 0)
    score = float(d.get("分数") or 0)
    rate = score / full if full else 0
    short = short_name(d.get("姓名", ""))

    if not d.get("档位"):
        # 满分要单独一档：100% 和 90% 都叫「优秀」，家长看不出孩子考满了
        d["档位"] = ("满分" if (full and score >= full) else
                     "优秀" if rate >= 0.9 else
                     "良好" if rate >= 0.8 else
                     "合格" if rate >= 0.7 else
                     "待提升" if rate >= 0.6 else "需重点关注")

    if not d.get("寄语"):
        # 寄语必须按分数档位分开写。原来只有「有失分 / 没失分」两种，结果：
        #   满分的孩子看到「把错的那几处再做一遍」（他一道没错了）
        #   0 分的孩子看到「基础不错，记准了就能拿回来」（他 40 分全丢了）
        # 都是同一句万能话套到底。家长不会说「文案不好」，只会觉得老师没看卷子。
        kps = []
        for _row in (d.get("失分") or []):
            _kp = (_row[1] if len(_row) > 1 else "") or ""
            _kp = _kp.replace(",", "、").strip()
            if _kp and _kp != "—" and _kp not in kps:
                kps.append(_kp)
        if not kps:
            # 没有答案文件时考点列全是「—」，得退回模块名。
            # 不退回就会写出「主要在 这一讲的基础内容」这种病句。
            kps = [m["名"] for m in (d.get("模块") or [])
                   if float(m.get("满分") or 0) - float(m.get("得分") or 0) > 0]

        # 寄语必须和「档位」一一对齐，六档分开写。
        # 只有四档时会出这种事：55 分和 10 分收到同一句「丢分比较多」，
        # 93 分和 86 分也是同一句 —— 家长拿两份报告一对比就看得出来。
        share = max(0.0, full - score)
        if full and score >= full:
            d["寄语"] = (f"{short}这次一分没丢，整张卷子干干净净。"
                         f"这不是运气——课上跟得紧，回家也确实练了。"
                         f"下一讲我往难一点的地方推，这个节奏保持住就行。")
        elif rate >= 0.9:
            _w = "、".join(kps[:2]) or "个别小题"
            # 一处说「这一点上」，两处得说「这两处」——不然是病句
            _here = f"{_w}这一点上" if len(kps[:2]) == 1 else f"{_w}这两处"
            d["寄语"] = (f"{short}这次只丢了 {share:g} 分，在{_here}。"
                         f"到这个份上，比的不是会不会，是细不细——"
                         f"课上我带着把这几处抠一遍，下回就能拿满。")
        elif rate >= 0.8:
            _w = "、".join(kps[:2]) or "个别小题"
            d["寄语"] = (f"{short}这次基础是扎实的，分主要丢在{_w}上。"
                         f"这几处不是没学会，是还没记牢，"
                         f"课上我带你再过一遍，回家不用加量。")
        elif rate >= 0.7:
            _w = "、".join(kps[:2]) or "几处基础题"
            d["寄语"] = (f"{short}这次还有几处没吃透：{_w}。"
                         f"不是不会，是还没练熟，换个问法就犹豫，"
                         f"课上我把这几块拆开再走一遍，回家照练习过两遍就稳了。")
        elif rate >= 0.6:
            _w = "、".join(kps[:3]) or "这一讲的几处重点"
            d["寄语"] = (f"{short}这次明显的薄弱在{_w}，不止一处。"
                         f"这些是这一讲的地基，后面的内容都压在上面，"
                         f"得先补上——我课上补，回家也请照练习过一遍。")
        else:
            _w = "、".join(kps[:3]) or "这一讲的重点内容"
            d["寄语"] = (f"{short}这次丢分比较多，主要在{_w}。"
                         f"说实话这些内容课上讲过，但孩子还没真正接住。"
                         f"不是一两天能补完的，我们先从最基础的开始，"
                         f"我课上补，也麻烦您在家盯着每天过一遍，"
                         f"有进展我随时跟您同步。")

def build(d: dict, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    files = []
    for idx, fn in ((1, page1), (2, page2)):
        img = Image.new("RGB", (W, H), "white")
        fn(d, Pen(img))
        p = out_dir / f"{d['姓名']}_诊断报告_第{idx}页.png"
        # compress_level=1：比默认的 6 快两三倍，肉眼看不出差别
        img.save(p, "PNG", compress_level=1)
        files.append(p)
    return files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--演示", action="store_true")
    ap.add_argument("--班级", default=None)
    ap.add_argument("--学生", default=None)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--输出", default=None, help="默认放到 生成结果/诊断报告/<班级>/<课次>/")
    args = ap.parse_args()

    if args.演示:
        d = DEMO
        out = Path(args.输出 or ".").expanduser().resolve()
    else:
        if not (args.班级 and args.学生):
            print("[错] 需要 --班级 和 --学生（或加 --演示 看样张）")
            return 2
        from _ws import workspace_root, dirs_for
        root = workspace_root(args.工作区)
        d = load_from_workspace(root, args.班级, args.学生, args.课次 or "")
        fill_text(d, root, args.班级)
        if args.输出:
            out = Path(args.输出).expanduser().resolve()
        else:
            out = dirs_for(root, args.班级, d["课次"])["work"] / "诊断报告"

    files = build(d, out)
    for p in files:
        print(p)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
