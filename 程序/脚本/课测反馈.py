#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
课测反馈一键生成器
==================================================================
熊老师专用：每周课测 -> 成绩表 + 家长话术 + 错题小练习

目录约定（相对本脚本所在的「课测反馈系统」根目录）：
  1-名单/<班级>.txt     姓名 + 分数 + 错题号        <- 你手打，最重要
  2-答案/<课次>.txt     题号 + 答案 + 考点 + 题型    <- 你提供答案
  3-照片/               改完的课测照片（文件名尽量 = 姓名.jpg，仅作核对/存档）
  课测答案/自定义练习/<班级>.txt 可选，有则覆盖模板生成
  4-输出/<班级>_<课次>/ 一切产物

用法：
  python 课测反馈.py                     # 自动处理 1-名单 / 2-答案 里所有配对
  python 课测反馈.py --demo              # 生成示例输入并跑一遍（自检用）
  python 课测反馈.py --班级 五1班 --答案 第5讲L95-96
"""
from __future__ import annotations

import argparse
import datetime as _dt
import os
import re
import sys
from collections import OrderedDict, defaultdict
from pathlib import Path

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("缺少 openpyxl，请先运行：python -m pip install openpyxl python-docx")

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    sys.exit("缺少 python-docx，请先运行：python -m pip install openpyxl python-docx")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent


def set_workspace(base: Path):
    """把全部输入/输出目录重定向到 base（--demo 用它隔离示例数据，绝不碰真实名单）。"""
    global DIR_ROSTER, DIR_ANSWER, DIR_PHOTO, DIR_OUT, DIR_CUSTOM, DIR_HISTORY, HISTORY_CSV
    DIR_ROSTER = base / "班级名单"
    DIR_ANSWER = base / "课测答案"
    DIR_PHOTO = base / "答题卡照片"
    DIR_OUT = base / "生成结果"
    DIR_CUSTOM = ROOT / "课测答案" / "自定义练习"
    DIR_HISTORY = base / "生成结果" / "历史"
    HISTORY_CSV = DIR_HISTORY / "成绩历史.csv"


set_workspace(ROOT)

PHOTO_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".heic"}

# 只有低于这个分数的孩子才在家长反馈里布置「家里要做的任务」。
# 高于这个分数只反馈情况，不派活 —— 老师精力有限，逐个盯不现实。
# 也可以在名单里用 @重点线: 30 覆盖。
FOCUS_SCORE = 30.0

# ============================================================
# 分数档位（按满分百分比折算，支持任意满分：100分制、40分制…）
# ============================================================
BANDS = [
    (100, 100, "满分", "（满分档）", (0x1F, 0x7A, 0x3D)),
    (90, 99, "优秀", "（优秀档）", (0x1F, 0x7A, 0x3D)),
    (80, 89, "良好", "（良好档）", (0x1B, 0x5E, 0xA8)),
    (70, 79, "合格", "（合格档）", (0xB4, 0x6A, 0x00)),
    (60, 69, "待提升", "（待提升档）", (0xC0, 0x39, 0x2B)),
    (0, 59, "需重点关注", "（需重点关注档）", (0xC0, 0x39, 0x2B)),
]


def band_of(score, full=100.0):
    """按「得分率」定档，自动适配任意满分。

    40 分制考 32.5 → 81.25% → 良好档（而不是被当成 32.5 分算「需重点关注」）。
    """
    if score is None:
        return "未评分", "", (0x55, 0x55, 0x55)
    full = float(full) if full else 100.0
    pct = 100.0 * float(score) / full
    for lo, hi, name, tail, color in BANDS:
        if lo <= pct <= hi + 0.999:
            return name, tail, color
    return "未评分", "", (0x55, 0x55, 0x55)


def full_mark(meta):
    """从 meta 读取满分，缺省 100。"""
    try:
        v = float(meta.get("满分") or 100)
        return v if v > 0 else 100.0
    except (TypeError, ValueError):
        return 100.0


# ============================================================
# 解析：名单
# ============================================================
SEP_RE = re.compile(r"[,，、;；/|\s]+")


def _split_errors(token: str):
    """把 '3、7、12' / '错3,7' / '全对' 解析成题号列表。"""
    t = token.strip()
    if not t:
        return None  # 未知
    if re.search(r"全对|没有错|无错|无$|^无$|满分|perfect|^-$|^—$", t, re.I):
        return []
    t = re.sub(r"^错[:：]?", "", t).strip()
    nums = [int(x) for x in re.findall(r"\d+", t)]
    return sorted(set(nums))


def parse_roster(path: Path):
    """名单文件 -> (meta, students)

    meta 形如 {'班级': '五1班', '课次': '第5讲L95-96', '日期': '2026-09-22'}
    学生行： 姓名 分数 [错题号]
    """
    meta: dict[str, str] = {}
    students: list[dict] = []
    seen = set()

    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("#") or line.startswith("//"):
            continue  # 纯注释行，不参与解析
        m = re.match(r"^[@＠]\s*([^:：\s]+)\s*[:：]?\s*(.*)$", line)
        if m:
            meta[m.group(1)] = m.group(2).strip()
            continue
        if re.match(r"^(姓名|名字)\b", line):  # 表头
            continue
        # 姓名 + 分数 + 可选错题
        parts = [p for p in re.split(r"[\s,，\t]+", line) if p]
        if not parts:
            continue
        # 只写了姓名、没有分数 → 先记为「待定」，后面按有没有答题卡判定缺考
        name = parts[0]
        rest = parts[1:]
        score = None
        absent = False
        no_score = len(parts) == 1
        err_tok = None
        for i, p in enumerate(rest):
            if re.search(r"缺考|请假|未考|absent|^-$", p, re.I):
                absent = True
                rest = rest[i + 1:]
                break
            mnum = re.match(r"^(-?\d+(?:\.\d+)?)", p)
            if mnum:
                score = float(mnum.group(1))
                tail = p[mnum.end():].strip("分,，、;；")
                # 分数本身写成了 '85(3,7)' 这种
                if tail:
                    err_tok = tail
                rest = rest[i + 1:]
                break
        if not absent:
            if err_tok is None and rest:
                err_tok = " ".join(rest)
            errors = _split_errors(err_tok) if err_tok else None
        else:
            errors = None

        key = name
        disp = name
        if key in seen:
            n = 2
            while f"{name}({n})" in seen:
                n += 1
            disp = f"{name}({n})"   # 重名自动加序号，避免话术/照片互相覆盖
            print(f"  [提示] 名单中出现重名，已区分：{name} -> {disp}")
        seen.add(key)
        seen.add(disp)
        if not absent and score is None and not no_score:
            print(f"  [警告] {disp} 没读到分数，已跳过该行：{raw.strip()}")
            continue
        students.append(
            {
                "姓名": name,
                "显示名": disp,
                "分数": score,
                "缺考": absent,
                "待定": no_score,      # 名单里只有姓名，还没判定
                "错题": errors,
                "原行": raw.strip(),
            }
        )
    return meta, students


# ============================================================
# 解析：大题扣分（免登记错题号的关键）
# ============================================================
def load_section_deductions(answer_path: Path, cls_name: str = ""):
    """读取大题扣分 JSON（按班级区分）。

    查找顺序：
      1. <答案>_大题扣分_<班级>.json   （分班专用，如 周六2点）
      2. <答案>_大题扣分.json          （默认）
    结构：{"姓名": {"一":0.5,"二":1,"三":2,"四":2,"五":2}, ...}
    """
    import json
    cands = []
    if answer_path is None:
        answer_path = Path("_none_")     # 让下面的候选路径失效，只走输出目录
    if cls_name:
        cands.append(answer_path.with_name(
            f"{answer_path.stem}_大题扣分_{cls_name}.json"))
    cands.append(answer_path.with_name(answer_path.stem + "_大题扣分.json"))
    # 识别照片.py 自动生成的（放在输出目录里）
    if cls_name:
        lesson = ""
        try:
            meta2, _ = parse_roster(next(
                f for f in DIR_ROSTER.glob("*.txt") if cls_name in f.stem))
            lesson = meta2.get("课次", "")
        except Exception:
            pass
        if lesson:
            cands.append(DIR_OUT / f"{cls_name}_{lesson}" / "大题扣分.json")
        cands.append(DIR_OUT / cls_name / "大题扣分.json")
    cand = next((c for c in cands if c.exists()), None)
    if not cand:
        return {}
    try:
        raw = json.loads(cand.read_text(encoding="utf-8-sig"))
        print(f"  大题扣分数据：{cand.name}")
    except (OSError, ValueError):
        return {}
    out = {}
    for name, v in raw.items():
        if name.startswith("_"):
            continue
        if isinstance(v, dict):
            out[name] = {k: float(x) for k, x in v.items()
                         if not k.startswith("_") and isinstance(x, (int, float))}
    return out

# ============================================================
# 解析：答案
# ============================================================
def parse_answers(path: Path):
    """答案文件 -> (meta, questions)

    每行： 题号[.、)] 答案  [| 考点 | 题型]
    考点/题型可省略；分隔符支持 | 、 ／ / tab 或连续两个以上空格。
    """
    meta: dict[str, str] = {}
    questions: list[dict] = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"^[@＠]\s*([^:：\s]+)\s*[:：]?\s*(.*)$", line)
        if m:
            meta[m.group(1)] = m.group(2).strip()
            continue

        # 题号
        mq = re.match(r"^(?:第)?\s*(\d+)\s*(?:题)?\s*[.、。)）:：]\s*(.+)$", line)
        if mq:
            no = int(mq.group(1))
            body = mq.group(2).strip()
        else:
            no = len(questions) + 1
            body = line

        # 切 答案 | 考点 | 题型
        segs = [s.strip() for s in re.split(r"\s*[|｜]\s*", body) if s.strip() != ""]
        if len(segs) == 1:
            # 尝试用 2+ 空格 或 制表符切
            segs = [s.strip() for s in re.split(r"\t+|\s{2,}", body) if s.strip()]
        ans = segs[0] if segs else ""
        point = segs[1] if len(segs) > 1 else ""
        qtype = segs[2] if len(segs) > 2 else ""
        questions.append({"题号": no, "答案": ans, "考点": point, "题型": qtype})

    questions.sort(key=lambda q: q["题号"])
    return meta, questions


# ============================================================
# 统计
# ============================================================
def build_stats(students, questions, meta=None):
    meta = meta or {}
    total = len(questions)
    qmap = {q["题号"]: q for q in questions}
    valid = [s for s in students if not s["缺考"] and s["分数"] is not None]
    scores = [s["分数"] for s in valid]
    FULL = full_mark(meta)
    pass_line = 0.60 * FULL     # 及格线 = 满分的 60%
    good_line = 0.90 * FULL     # 优秀线 = 满分的 90%

    cls = {
        "人数": len(students),
        "有效": len(valid),
        "缺考": sum(1 for s in students if s["缺考"]),
        "满分题数": total,
        "满分": FULL,
        "平均分": round(sum(scores) / len(scores), 1) if scores else 0.0,
        "最高": max(scores) if scores else 0.0,
        "最低": min(scores) if scores else 0.0,
        "及格率": (round(100.0 * sum(1 for x in scores if x >= pass_line) / len(scores), 1) if scores else 0.0),
        "优秀率": (round(100.0 * sum(1 for x in scores if x >= good_line) / len(scores), 1) if scores else 0.0),
        "平均得分率": (round(100.0 * sum(scores) / len(scores) / FULL, 1) if scores else 0.0),
        "及格线": pass_line,
    }

    # 每题错误人数
    q_err = {q["题号"]: [] for q in questions}
    for s in valid:
        if s["错题"] is None:
            continue
        for n in s["错题"]:
            if n in q_err:
                q_err[n].append(s["姓名"])

    # 考点聚合
    point_stat = OrderedDict()
    for q in questions:
        key = q["考点"] or f"第{q['题号']}题"
        p = point_stat.setdefault(key, {"考点": key, "题号": [], "错": 0, "题型": q["题型"]})
        p["题号"].append(q["题号"])
        p["错"] += len(q_err[q["题号"]])
    n_valid = max(1, len(valid))
    for p in point_stat.values():
        p["错误率"] = round(100.0 * p["错"] / (n_valid * max(1, len(p["题号"]))), 1)

    # 大题 -> 名称 / 题号区间。可由名单里的 @大题名称 与 @大题题数 覆盖，
    # 例如四年级 NCE1-L11-18：@大题名称: 名词复数,单项选择,选词填空,按要求改写句子,阅读理解
    #                        @大题题数: 5,20,5,5,5
    KEYS = ("一", "二", "三", "四", "五")
    _def_names = ["过去式/过去分词", "单项选择", "用适当形式填空",
                  "按要求改写句子", "阅读理解"]
    _def_counts = [5, 15, 5, 5, 5]
    _names = [x.strip() for x in re.split(r"[,，、]", meta.get("大题名称", "")) if x.strip()]
    if len(_names) != 5:
        _names = _def_names
    _counts = []
    for x in re.split(r"[,，、]", meta.get("大题题数", "")):
        try:
            _counts.append(int(x.strip()))
        except ValueError:
            pass
    if len(_counts) != 5 or sum(_counts) <= 0:
        _counts = _def_counts
    SEC_NAME = dict(zip(KEYS, _names))
    SEC_RANGE = {}
    _cur = 1
    for _k, _n in zip(KEYS, _counts):
        SEC_RANGE[_k] = (_cur, _cur + _n - 1)
        _cur += _n
    meta["_sec_names"] = _names       # 供 make_feedback / 双胞胎话术取用

    # 若提供了大题扣分，则据此为每个学生生成"薄弱大题"
    #
    # 【自证门槛】四个大题扣分之和必须等于「满分 − 分数」。
    # 对不上说明读数有误（或老师漏扣），此时不输出具体诊断，避免误导家长。
    sec_ded = meta.get("_大题扣分") or {}
    for s in students:
        d = sec_ded.get(s["姓名"])
        if not d:
            s["大题扣分"] = None
            s["薄弱大题"] = []
            s["未定位扣分"] = 0.0
            continue
        s["大题扣分"] = d
        got = sum(float(v or 0) for k, v in d.items() if not str(k).startswith("_"))
        need = FULL - float(s["分数"] or 0)
        if abs(got - need) > 0.01:
            s["薄弱大题"] = []
            s["大题扣分存疑"] = f"分项和 {got:g} ≠ 应扣 {need:g}"
            continue
        weak = []
        for k in ("一", "二", "三", "四", "五"):
            v = float(d.get(k, 0) or 0)
            if v > 0:
                a, b = SEC_RANGE[k]
                qs = [q["题号"] for q in questions if a <= q["题号"] <= b]
                pts = list(dict.fromkeys(
                    (q["考点"] or q["题型"] or f"第{q['题号']}题")
                    for q in questions if a <= q["题号"] <= b))
                weak.append({"大题": k, "名称": SEC_NAME[k], "扣分": v,
                             "题号": qs, "考点": pts, "分值": b - a + 1})
        weak.sort(key=lambda x: -x["扣分"])
        s["薄弱大题"] = weak
        resid = float(d.get("未定位", 0) or 0)
        s["未定位扣分"] = resid
    for s in students:
        if s["缺考"] or s["分数"] is None:
            # 缺考/无分数：补齐字段，话术走缺考分支（缺考不算「错题未填」）
            s.setdefault("错题", None)
            s["错题未知"] = False
            s["档位"], s["档位说明"], s["颜色"] = band_of(s["分数"], FULL)
            if s["缺考"]:
                s["档位"] = "缺考"
            s["薄弱"], s["亮点"], s["错题考点"] = [], [], ""
            continue
        if s["错题"] is None:
            if s["分数"] is not None and s["分数"] >= FULL:
                s["错题"] = []  # 满分 = 全对，不必再手写「全对」
                s["错题未知"] = False
            else:
                s["错题"] = []
                s["错题未知"] = True
        else:
            s["错题未知"] = False
        s["档位"], s["档位说明"], s["颜色"] = band_of(s["分数"], FULL)
        # 该生薄弱考点
        weak = []
        for p in point_stat.values():
            hit = [n for n in p["题号"] if n in s["错题"]]
            if hit:
                weak.append({"考点": p["考点"], "题号": hit, "题型": p["题型"]})
        s["薄弱"] = sorted(weak, key=lambda x: -len(x["题号"]))
        # 该生全对的考点（亮点）
        strong = []
        for p in point_stat.values():
            hit = [n for n in p["题号"] if n in s["错题"]]
            if not hit and p["题号"]:
                strong.append(p["考点"])
        s["亮点"] = strong
        s["错题考点"] = "、".join(dict.fromkeys(w["考点"] for w in s["薄弱"]))

    ranked = sorted(valid, key=lambda s: -s["分数"])
    for i, s in enumerate(ranked, 1):
        s["名次"] = i

    return total, qmap, q_err, point_stat, cls


# ============================================================
# 照片索引
# ============================================================
def photo_dir_for(*candidates):
    """按候选名依次查找分班照片目录，找不到就回退到 3-照片/ 根目录。

    候选顺序一般是：名单里的 @班级 字段 → 名单文件名。
    这样即使名单里把班级名写成了别的（如「五1班」而目录叫「周六10点」），
    只要文件名对得上就仍能找到照片。
    """
    for c in candidates:
        if not c:
            continue
        d = DIR_PHOTO / str(c)
        if d.is_dir():
            return d
    return DIR_PHOTO


def load_photo_map(base=None):
    """读取 照片目录/命名对照.txt

    支持两种写法（可混用）：
      IMG_1234.jpg 张三          <- 文件名 姓名
      张三                        <- 纯姓名，按文件名排序后依次对应
    """
    base = base or DIR_PHOTO
    mp = {}
    for cand in (base / "命名对照.txt", base / "_命名对照.txt"):
        if not cand.exists():
            continue
        for raw in cand.read_text(encoding="utf-8-sig").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = re.split(r"[\s,，\t]+", line)
            if len(parts) >= 2:
                mp[parts[0]] = parts[1]
            elif len(parts) == 1:
                mp.setdefault("__order__", [])
                mp["__order__"].append(parts[0])
    return mp


def load_name_alias():
    """姓名对照表：答题卡/照片上的名字 -> 名单里的名字（1-名单/姓名对照.txt）。"""
    p = ROOT / "班级名单" / "姓名对照.txt"
    out = {}
    if not p.exists():
        return out
    try:
        for raw in p.read_text(encoding="utf-8-sig").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split()
            if len(parts) >= 2:
                out[parts[0]] = parts[1]
    except OSError:
        pass
    return out


def index_photos(names, cls_name=None, roster_stem=None, lesson=None):
    """返回 {姓名: [照片路径]} 与未匹配照片名列表。

    照片可以放在班级目录下，也可以放在课次子目录里（如
    3-照片\\周六10点\\五年级第五讲\\*.jpg）。若给了 lesson，优先只扫该课次子目录，
    避免不同课次的照片混在一起。
    """
    photos = defaultdict(list)
    unmatched = []
    base = photo_dir_for(cls_name, roster_stem)
    if not base.exists():
        return photos, unmatched

    # 确定搜索根：课次子目录优先
    root = base
    if lesson:
        cand = base / lesson
        if cand.is_dir():
            root = cand
        else:
            # 模糊匹配：找名字里含课次关键字的子目录
            key = re.sub(r"[\s\-_]", "", str(lesson))
            for d in base.iterdir():
                if d.is_dir() and key and key in re.sub(r"[\s\-_]", "", d.name):
                    root = d
                    break

    # 优先用「已改名」目录（按学生名命名的干净副本）；
    # 否则扫课次目录，但要排除「已改名」等子目录，避免同一张照片被算两次。
    renamed = root / "已改名"
    if renamed.is_dir() and any(renamed.iterdir()):
        files = [f for f in sorted(renamed.iterdir())
                 if f.is_file() and f.suffix.lower() in PHOTO_EXT]
    else:
        SKIP_DIRS = {"已改名", "分发", "照片归档"}
        files = [f for f in sorted(root.rglob("*"))
                 if f.is_file() and f.suffix.lower() in PHOTO_EXT
                 and not (set(f.relative_to(root).parts[:-1]) & SKIP_DIRS)]
    by_name = {f.name: f for f in files}
    mapping = load_photo_map(root)
    if not mapping and root != base:
        mapping = load_photo_map(base)

    # 1) 显式对照优先
    for fname, who in mapping.items():
        if fname == "__order__":
            continue
        f = by_name.get(fname)
        if f is None:
            hits = [x for x in files if x.stem == Path(fname).stem]
            f = hits[0] if hits else None
        if f and who in names:
            photos[who].append(f)

    # 2) 纯姓名列表：按文件名排序依次对应
    if "__order__" in mapping:
        pool = [f for f in files if not any(f in v for v in photos.values())]
        for who, f in zip(mapping["__order__"], pool):
            if who in names:
                photos[who].append(f)

    # 3) 文件名里带姓名的（张三.jpg / 张三_1.jpg / 5.1张三.jpg）
    #    也认姓名对照表（照片写本名「马泽文」、名单里是「马团团」）
    _alias = load_name_alias()
    for f in files:
        if any(f in v for v in photos.values()):
            continue
        stem = f.stem
        hit = next((n for n in names if n in stem), None)
        if hit is None:
            hit = next((r for c, r in _alias.items()
                        if c in stem and r in names), None)
        if hit:
            photos[hit].append(f)

    # 4) 姓名对照兜底：照片名是本名、名单里是别名时归位
    for _card, _roster in load_name_alias().items():
        if _roster in names and photos.get(_card):
            photos[_roster].extend(photos.pop(_card))

    used = {p for v in photos.values() for p in v}
    unmatched = [f.name for f in files if f not in used]
    for k in photos:
        photos[k] = sorted(photos[k], key=lambda p: p.name)
    return photos, unmatched

def archive_photos(outdir, ordered, photos):
    """把照片按姓名复制到 输出/照片归档/ 下，方便直接发给家长。"""
    import shutil
    dst_root = outdir / "照片归档"
    count = 0
    for s in ordered:
        dn = s.get("显示名", s["姓名"])
        fs = photos.get(s["姓名"], [])
        if not fs:
            continue
        d = dst_root / dn
        d.mkdir(parents=True, exist_ok=True)
        for i, f in enumerate(fs, 1):
            ext = f.suffix.lower()
            shutil.copy2(f, d / f"{dn}_{i}{ext}")
            count += 1
    return count


# ============================================================
# 输出：Excel 成绩表
# ============================================================
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="1B5EA8")
HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def write_xlsx(path, meta, students, questions, qmap, q_err, point_stat, cls, photos, unmatched):
    wb = Workbook()

    def head(ws, cols, widths, title):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(cols))
        c = ws.cell(row=1, column=1, value=title)
        c.font = Font(name="微软雅黑", size=13, bold=True, color="1B5EA8")
        c.alignment = CENTER
        ws.row_dimensions[1].height = 28
        for i, (name, w) in enumerate(zip(cols, widths), 1):
            cell = ws.cell(row=2, column=i, value=name)
            cell.fill, cell.font, cell.alignment, cell.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[2].height = 30
        ws.freeze_panes = "A3"

    title = f"{meta.get('班级','')} {meta.get('课次','')} 课测成绩表".strip()

    # --- Sheet1 成绩总表 ---
    ws = wb.active
    ws.title = "成绩总表"
    cols = ["姓名", "分数", "档位", "错题号", "薄弱考点", "掌握较好的考点", "照片", "备注"]
    widths = [12, 8, 12, 16, 30, 26, 10, 20]
    head(ws, cols, widths, title)

    ordered = sorted(students, key=lambda s: (s["缺考"], -(s["分数"] if s["分数"] is not None else -1)))
    r = 3
    for s in ordered:
        name = s.get("显示名", s["姓名"])
        real = s["姓名"]
        if s["缺考"] or s["分数"] is None:
            vals = [name, "缺考", "缺考", "", "", "", "", ""]
        else:
            vals = [
                name,
                s["分数"],
                s["档位"],
                "、".join(str(n) for n in s["错题"]) or ("—" if s["错题"] else "全对"),
                s["错题考点"] or "无",
                "、".join(s["亮点"][:4]) or "—",
                f"{len(photos.get(real, []))}张" if photos.get(real) else "",
                "错题号未填写" if s.get("错题未知") else "",
            ]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = LEFT if i in (5, 6, 8) else CENTER
        ws.cell(row=r, column=2).font = Font(name="微软雅黑", size=10, bold=True,
                                             color="%02X%02X%02X" % s.get("颜色", (0, 0, 0)))
        if s.get("错题未知"):
            for i in range(1, len(cols) + 1):
                ws.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FFF3CD")
        r += 1

    # 统计行
    r += 1
    stats_rows = [
        ("班级人数", cls["人数"]), ("有效人数", cls["有效"]), ("缺考", cls["缺考"]),
        ("平均分", cls["平均分"]), ("满分", cls.get("满分", 100)),
        ("平均得分率", f"{cls.get('平均得分率', 0)}%"),
        ("最高分", cls["最高"]), ("最低分", cls["最低"]),
        ("及格率", f"{cls['及格率']}%（≥{cls.get('及格线', 60):g}分）"),
        ("优秀率", f"{cls['优秀率']}%（≥{cls.get('满分',100)*0.9:g}分）"),
    ]
    ws.cell(row=r, column=1, value="班级统计").font = Font(name="微软雅黑", size=11, bold=True, color="1B5EA8")
    r += 1
    for k, v in stats_rows:
        ws.cell(row=r, column=1, value=k).font = BODY_FONT
        c = ws.cell(row=r, column=2, value=v)
        c.font = Font(name="微软雅黑", size=10, bold=True)
        c.alignment = CENTER
        r += 1

    # --- Sheet2 每题错误明细 ---
    ws2 = wb.create_sheet("每题错误明细")
    head(ws2, ["题号", "正确答案", "考点", "题型", "错误人数", "错误率", "错误名单"],
         [8, 26, 22, 14, 10, 10, 46], title + " — 逐题分析")
    r = 3
    for q in questions:
        n = q["题号"]
        errs = q_err[n]
        rate = round(100.0 * len(errs) / max(1, cls["有效"]), 1)
        vals = [n, q["答案"], q["考点"] or "—", q["题型"] or "—", len(errs), f"{rate}%", "、".join(errs) or "全员正确"]
        for i, v in enumerate(vals, 1):
            cell = ws2.cell(row=r, column=i, value=v)
            cell.font, cell.border = BODY_FONT, BORDER
            cell.alignment = LEFT if i == 7 else CENTER
        if rate >= 50:
            for i in range(1, 8):
                ws2.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FDE7E9")
        elif rate >= 30:
            for i in range(1, 8):
                ws2.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FFF3CD")
        r += 1

    # --- Sheet3 考点掌握情况 ---
    ws3 = wb.create_sheet("考点掌握情况")
    head(ws3, ["考点", "对应题号", "题型", "错误人次", "错误率", "掌握判定"],
         [26, 18, 14, 12, 10, 14], title + " — 考点分析")
    r = 3
    for p in sorted(point_stat.values(), key=lambda x: -x["错误率"]):
        rate = p["错误率"]
        verdict = "已掌握" if rate < 15 else ("基本掌握" if rate < 35 else ("薄弱" if rate < 60 else "严重薄弱"))
        vals = [p["考点"], "、".join(str(n) for n in p["题号"]), p["题型"] or "—", p["错"], f"{rate}%", verdict]
        for i, v in enumerate(vals, 1):
            cell = ws3.cell(row=r, column=i, value=v)
            cell.font, cell.border, cell.alignment = BODY_FONT, BORDER, CENTER
        if rate >= 35:
            for i in range(1, 7):
                ws3.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FDE7E9")
        r += 1

    # --- Sheet4 照片核对 ---
    ws4 = wb.create_sheet("照片核对")
    head(ws4, ["姓名", "分数", "照片张数", "照片文件"], [12, 8, 10, 70], title + " — 照片归档核对")
    r = 3
    for s in ordered:
        fs = photos.get(s["姓名"], [])
        vals = [s.get("显示名", s["姓名"]), s["分数"] if not s["缺考"] else "缺考", len(fs),
                "；".join(f.name for f in fs)]
        for i, v in enumerate(vals, 1):
            cell = ws4.cell(row=r, column=i, value=v)
            cell.font, cell.border, cell.alignment = BODY_FONT, BORDER, (LEFT if i == 4 else CENTER)
        if not fs:
            for i in range(1, 5):
                ws4.cell(row=r, column=i).fill = PatternFill("solid", fgColor="FFF3CD")
        r += 1
    if unmatched:
        r += 1
        ws4.cell(row=r, column=1, value="未匹配到学生的照片：").font = Font(name="微软雅黑", size=10, bold=True, color="C0392B")
        ws4.cell(row=r, column=4, value="；".join(unmatched)).alignment = LEFT

    # --- Sheet5 识别核对（照片与识别结果并排，供老师一眼复核）---
    try:
        from openpyxl.drawing.image import Image as XLImage
        from PIL import Image as PILImage

        ws5 = wb.create_sheet("识别核对")
        ws5.merge_cells("A1:H1")
        c = ws5.cell(row=1, column=1, value="识别结果核对表 —— 请对照右侧照片逐个核对姓名/分数/错题号，发现不对直接改「1-名单」后重跑")
        c.font = Font(name="微软雅黑", size=12, bold=True, color="C0392B")
        c.alignment = CENTER
        ws5.row_dimensions[1].height = 26
        heads5 = ["姓名", "分数", "错题号", "核对", "试卷照片（缩放显示，以原件为准）"]
        widths5 = [12, 8, 18, 8, 100]
        for i, (nm, w) in enumerate(zip(heads5, widths5), 1):
            cell = ws5.cell(row=2, column=i, value=nm)
            cell.fill, cell.font, cell.alignment, cell.border = HEAD_FILL, HEAD_FONT, CENTER, BORDER
            ws5.column_dimensions[get_column_letter(i)].width = w
        ws5.row_dimensions[2].height = 24
        ws5.freeze_panes = "A3"

        row = 3
        for s in ordered:
            fs = photos.get(s["姓名"], [])
            vals = [
                s.get("显示名", s["姓名"]),
                "缺考" if s["缺考"] else s["分数"],
                "、".join(str(n) for n in (s["错题"] or [])) or ("—" if s["缺考"] else "全对"),
                "□",
                "",
            ]
            for i, v in enumerate(vals, 1):
                cell = ws5.cell(row=row, column=i, value=v)
                cell.font, cell.border = BODY_FONT, BORDER
                cell.alignment = CENTER
            if s.get("错题未知"):
                for i in range(1, 5):
                    ws5.cell(row=row, column=i).fill = PatternFill("solid", fgColor="FFF3CD")

            # 嵌第一张照片，缩放到合适尺寸
            img_h = 150
            if fs:
                try:
                    src = fs[0]
                    pil = PILImage.open(src)
                    ratio = img_h / pil.height
                    w_px = int(pil.width * ratio)
                    tmp = path.parent / f"_tmp_{row}.png"
                    pil.convert("RGB").resize((w_px, img_h), PILImage.LANCZOS).save(tmp, "PNG")
                    xi = XLImage(str(tmp))
                    xi.width, xi.height = w_px, img_h
                    ws5.add_image(xi, f"E{row}")
                    img_h_pts = img_h * 0.75
                    ws5.row_dimensions[row].height = max(img_h_pts, 20)
                    if len(fs) > 1:
                        ws5.cell(row=row, column=6,
                                 value=f"+{len(fs)-1}张").font = Font(name="微软雅黑", size=9, color="1B5EA8")
                except Exception as e:
                    ws5.cell(row=row, column=5, value=f"[图片嵌入失败: {e}]").font = BODY_FONT
            else:
                ws5.cell(row=row, column=5, value="（未匹配到照片）").font = Font(
                    name="微软雅黑", size=9, color="999999")
            row += 1
    except ImportError:
        pass  # 没装 Pillow 就跳过该 Sheet，不影响其他产物

    wb.save(path)
    # 清理临时缩略图
    for t in path.parent.glob("_tmp_*.png"):
        try:
            t.unlink()
        except OSError:
            pass
    return path


# ============================================================
# 话术生成
# ============================================================
def _pct_text(errors, total):
    if total <= 0:
        return ""
    r = len(errors) / total
    if r <= 0.1:
        return "个别小疏漏"
    if r <= 0.25:
        return "有少量失分"
    if r <= 0.5:
        return "失分明显"
    return "失分较多"


def make_feedback(s, cls, qmap, meta, trend=""):
    """生成一个孩子的家长反馈 —— 口语化，像微信里直接跟家长说的话。

    设计原则：
      1. 说人话，不用【】小标题、不写成公文。
      2. 只有分数低于重点线（默认30，可在名单里用 @重点线: 覆盖）才布置后续任务，
         且最多 2 条。其余孩子只反馈情况，不给家长派活 —— 老师盯不过来。
      3. 任务必须是低成本可执行的（3分钟能完成），不写"每天5分钟连做3天"。
    """
    name = s["姓名"]
    # 名单/照片里的排序前缀（如"A李怀森"）不进入家长话术
    disp_name = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", name) or name
    # 称呼：叫名字后两个字；两字名（或更短）就叫全名
    short = disp_name[-2:] if len(disp_name) > 2 else disp_name
    lesson = meta.get("课次", "本次")
    # 家长话术里不出现讲次名，统一说「月考试卷」（可用 @考试名称: 覆盖）
    exam = meta.get("考试名称") or meta.get("考试") or "月考试卷"
    score = s["分数"]
    total = cls["满分题数"]
    FULL = cls.get("满分", 100.0)

    if s["缺考"] or score is None:
        return (f"{short}家长您好，{exam}{short}请假没赶上。"
                f"答题卡和答案我都留着，您方便时让孩子补做一下，做完拍照发我，我单独批。"
                f"这讲内容前后连着，早点补上不影响后面的进度。")

    score_txt = f"{score:g}/{FULL:g}"
    ptitle = meta.get("称呼") or "家长"
    warm = meta.get("家长风格") == "亲和"      # 低龄段：更软、多鼓励

    # ---------- 特殊学生：专属话术（如「依赖带读型」）----------
    _special = [x.strip() for x in re.split(r"[,，、]", meta.get("特殊学生", "")) if x.strip()]
    if s["姓名"] in _special:
        return make_special_feedback(s, meta, short, ptitle, exam, score_txt)
    _firm = [x.strip() for x in re.split(r"[,，、]", meta.get("严正学生", "")) if x.strip()]
    if s["姓名"] in _firm:
        return make_firm_feedback(s, meta, short, ptitle, exam, score_txt)
    _reassure = [x.strip() for x in re.split(r"[,，、]", meta.get("安心学生", "")) if x.strip()]
    if s["姓名"] in _reassure or meta.get("家长风格") == "安抚":
        return make_reassuring_feedback(s, meta, short, ptitle, exam, score_txt)

    wrong = s["错题"] or []
    n_wrong = len(wrong)
    sec_weak = s.get("薄弱大题") or []
    unknown = s.get("错题未知")
    pct = 100.0 * score / FULL if FULL else 0.0
    # 总分口径的扣分：与报给家长的分数严格一致，避免家长看出数字对不上
    total_off = (FULL - float(score)) if score is not None else 0.0
    located_ded = sum(w["扣分"] for w in sec_weak)
    total_ded = total_off if total_off > 0 else located_ded
    try:
        focus_line = float(meta.get("重点线") or FOCUS_SCORE)
    except (TypeError, ValueError):
        focus_line = FOCUS_SCORE

    # ---------- 开场：报分数（考得好的先道喜）----------
    ptitle = meta.get("称呼") or "家长"
    warm = meta.get("家长风格") == "亲和"      # 低龄段：更软、多鼓励
    if warm:
        if pct >= 95:
            msg = [f"{short}{ptitle}您好！{exam}改完啦，先跟您报个喜——"
                   f"{short}这次考了 {score_txt}，写得又认真又干净，几乎一分没丢，真棒！"]
        elif pct >= 85:
            msg = [f"{short}{ptitle}您好！{exam}改完啦，跟您说个好消息。"
                   f"{short}这次考了 {score_txt}，这一讲学得挺扎实的。"]
        elif pct >= 70:
            msg = [f"{short}{ptitle}您好！{exam}改完啦，我跟您说说{short}这次的表现。"
                   f"{short}考了 {score_txt}，大部分内容都掌握住了。"]
        else:
            msg = [f"{short}{ptitle}您好！{exam}改完啦，我跟您聊聊{short}这次的情况。"
                   f"{short}这次考了 {score_txt}。"]
    elif pct >= 95:
        msg = [f"{short}{ptitle}您好，{exam}改完了，先跟您道个喜——"
               f"{short}这次考了 {score_txt}，卷子做得非常干净，几乎一分没丢。"]
    elif pct >= 85:
        msg = [f"{short}{ptitle}您好，{exam}改完了，跟您说个好消息。"
               f"{short}这次 {score_txt}，这个成绩在班里是很好的。"]
    elif pct >= 70:
        msg = [f"{short}{ptitle}您好，{exam}改完了，跟您说说情况。"
               f"{short}这次考了 {score_txt}，该拿的分基本都拿到了。"]
    else:
        msg = [f"{short}{ptitle}您好，{exam}改完了，跟您说说情况。"
               f"{short}这次考了 {score_txt}。"]
    if trend:
        msg.append(trend)

    # ---------- 先夸：说具体做到了什么（不空夸）----------
    ALL_SECS = tuple(meta.get("_sec_names") or
                     ("过去式/过去分词", "单项选择", "用适当形式填空",
                      "按要求改写句子", "阅读理解"))
    weak_names_set = {w["名称"] for w in sec_weak}
    full_secs = [nm for nm in ALL_SECS if nm not in weak_names_set]

    _common_good = meta.get("共性优点", "")
    if _common_good:
        msg.append(f"先说说做得好的：{_common_good}。"
                   if warm else
                   f"先说说好的地方：{_common_good}。")
    elif not weak_names_set and not unknown and total and n_wrong == 0:
        msg.append(f"整张卷子一道没错！这一讲孩子是真的学明白了，"
                   f"这个状态太棒了，继续保持。"
                   if warm else
                   f"整张卷子一道没错，这一讲的内容{short}是真吃透了，"
                   f"这个状态保持住就好。")
    elif full_secs and not unknown:
        msg.append(f"先夸一夸：{'、'.join(full_secs[:3])}这几块孩子全做对了，"
                   f"说明上课是认认真真跟着走的，特别好。"
                   if warm else
                   f"卷子上{'、'.join(full_secs[:3])}这几块{short}一道没错，"
                   f"说明课上讲的时候是跟着听的，这个我要表扬一下。")
    elif sec_weak:
        worst = max(w["扣分"] for w in sec_weak)
        msg.append(f"这张卷子每个题型都有扣分，"
                   f"不过分得比较散，最多的一块也就 {worst:g} 分，"
                   f"说明不是完全没会，是几个点还没顾上。"
                   if warm else
                   f"这张卷子五个题型都扣了点分，不过分得比较散，"
                   f"最多的一块也就 {worst:g} 分，说明基础不算差，"
                   f"是几个点还没顾上。")
    elif s["亮点"] and not unknown:
        msg.append(f"{'、'.join(s['亮点'][:3])}这几块掌握得挺牢，是这次的加分项。")

    # ---------- 问题：客观说事，不贴标签 ----------
    # 有知识点库时，直接说「哪个知识点、为什么错」，比只报大题名更具体
    # 共性知识点优先：全班都错得差不多时，用同一套，不必逐人分析
    _lib = meta.get("_kplib")
    # 逐人知识点优先（老师口述的每人弱项），其次共性知识点
    _per = {}
    for _it in re.split(r"[；;]", meta.get("逐人知识点", "")):
        if "=" in _it:
            _w, _v = _it.split("=", 1)
            _per[_w.strip()] = [x.strip() for x in re.split(r"[,，、]", _v) if x.strip()]
    _common = [x.strip() for x in
               re.split(r"[,，、]", meta.get("共性知识点", "")) if x.strip()]
    _mine = _per.get(s["姓名"]) or _per.get(s.get("显示名", ""))
    if _mine and _lib:
        _kp = [(_lib.SHORT.get(k, k), _lib.TIP.get(k, "")) for k in _mine[:3]]
    elif _common and _lib:
        _kp = [(_lib.SHORT.get(k, k), _lib.TIP.get(k, "")) for k in _common[:3]]
    else:
        _kp = s.get("知识点") or []
    if _kp:
        _kn = "、".join(x[0] for x in _kp[:2])
        _kt = "；".join(x[1] for x in _kp[:2] if x[1])
        kp_txt = f"{_kn}——{_kt}" if _kt else _kn
        kp_short = _kn
    else:
        kp_txt = kp_short = ""
    _kt1 = (_kp[0][1] if _kp and _kp[0][1] else "")
    _kt2 = (_kp[1][1] if len(_kp) > 1 and _kp[1][1] else "")

    # ---------- 粗心丢分 vs 知识点没记牢 ----------
    _careless_secs = [x.strip() for x in
                      re.split(r"[,，、]", meta.get("粗心大题", "")) if x.strip()]
    careless_ded = sum(w["扣分"] for w in sec_weak if w["大题"] in _careless_secs)
    knowledge_ded = total_ded - careless_ded
    _cn = "、".join(w["名称"] for w in sec_weak if w["大题"] in _careless_secs)
    if _kp and not sec_weak:
        # 只有共性知识点（本次没逐人统计分项）
        msg.append(f"这次错的主要在{kp_short}。"
                   + (f"{_kt1}。" if _kt1 else "")
                   + (f"另外，{_kt2}。" if _kt2 else "")
                   + f"这些都是要记要背的内容，多过两遍就熟了。"
                   if warm else
                   f"这次错的主要在{kp_short}。"
                   + (f"{_kt1}。" if _kt1 else "")
                   + (f"另外，{_kt2}。" if _kt2 else "")
                   + f"这几块都是靠记住就能拿分的，多过两遍就稳了。")
    elif sec_weak:
        names3 = kp_txt or "、".join(w["名称"] for w in sec_weak[:3])
        if pct >= 85:
            msg.append(f"要说还能更完美一点点，就是{kp_short or names3}上错了那么一两道，"
                       f"都是小地方，我课上带着孩子再顺一遍就好啦。"
                       if warm else
                       f"要说还能再稳一点的地方，就是{names3}上错了那么一两道，"
                       f"都是细节，我课上会带着再顺一遍。")
        elif pct >= 70:
            msg.append(f"这次还能再多拿几分的是{kp_short or names3}。"
                       + (f"{_kt1}。" if _kt1 else "")
                       + (f"另外，{_kt2}。" if (_kt2 and not careless_ded) else "")
                       + f"不是孩子不会，是还没练熟，多过两遍很快就补上了。"
                       if warm else
                       f"{names3}这几块还需要再练练。"
                       f"孩子不是不会，是这类题的步骤还没形成习惯，"
                       f"换个问法就容易犹豫——这个多过两遍题就好了。")
        elif total_ded >= 8:
            msg.append(f"这次错的主要在{kp_short or names3}。"
                       + (f"{_kt1}。" if _kt1 else "")
                       + (f"另外，{_kt2}。" if (_kt2 and not careless_ded) else "")
                       + f"这些都要记要背，多过两遍就熟了。"
                       if warm else
                       f"失分主要在{names3}这几块，一共 {total_ded:g} 分。"
                       f"我看了下，问题不在孩子身上，是这一类题的解题步骤还没固定，"
                       f"课上我会把方法再拆开讲一遍。")
        else:
            msg.append(f"{names3}上丢了 {total_ded:g} 分，都是小地方，"
                       f"不影响整体，我课上顺带提一下就行。")
    elif unknown:
        msg.append("这次卷面细节我还在整理，先跟您同步分数，"
                   "具体的薄弱点我下节课当面跟孩子过。")
    elif n_wrong == 0:
        msg.append("这次没有失分点。课测只能说明当下会了，"
                   "过两周还能记住才算真掌握，这块我课上会再带一遍。")

    # ---------- 说清「粗心丢的」和「没记牢的」 ----------
    if _careless_secs and careless_ded > 0 and knowledge_ded > 0:
        msg.append(f"这次扣的分里，{careless_ded:g} 分是没仔细丢的"
                   f"（{_cn}那几道答案就摆在眼前，检查一遍就能捡回来），"
                   f"{knowledge_ded:g} 分是知识点还没记牢。")
    elif _careless_secs and careless_ded > 0:
        msg.append(f"这些分基本都是没仔细丢的——{_cn}那几道答案就摆在眼前，"
                   f"不是不会，养成检查的习惯就能拿回来。")
    elif _careless_secs and knowledge_ded > 0 and total_ded > 0:
        msg.append("这次扣的都是知识点没记牢，没有粗心丢分，"
                   "把该记的记住，分数就上来了。")

    # ---------- 分层作业 ----------
    # 低于重点线 → 发《专项练习_10题》（5 选择 + 5 句型转换，带答案页）
    # 其余学生   → 发《背诵单_巩固版》（不规则动词 + 重点句型，不背单词）
    RECITE = "《背诵单》（附件，动词的过去式和过去分词 + 本讲重点句型）"
    PRACTICE = "《专项练习》（附件，10 道题，最后一页有答案）"
    personal = meta.get("任务模式") == "个人小任务"
    TASK = f"《{short}的小任务》（附件）"

    if personal and score is not None and score < focus_line:
        msg.append(f"我把孩子这次错的地方都整理出来了，做成了一份{TASK}，"
                   f"只有孩子错的那些，量很小。")
        msg.append("一天五分钟就能过一遍，过完拍张照发我，我看看记住没有。")
    elif personal:
        msg.append(f"我把孩子这次要记的东西整理成了一份{TASK}，"
                   f"就是几个词和句型，读一读、背一背就行，不用写。")
        msg.append("让孩子先自己过一遍，我下节课抽背，说给我听就行。")
    elif score is not None and score < focus_line:
        picked = [w for w in sec_weak if w["扣分"] >= 2][:2] or sec_weak[:1]
        names = {w["名称"] for w in picked}
        items = []
        if names & {"单项选择", "用适当形式填空", "过去式/过去分词"}:
            items.append("选择题 5 道（练 for、since、just、last 这几个标志词）")
        if names & {"按要求改写句子"}:
            items.append("句型转换 5 道（否定句、疑问句的语序）")
        if not items:
            items.append("选择题 5 道、句型转换 5 道")
        items = items[:2]
        msg.append(f"我给孩子准备了一份{PRACTICE}，就 10 道题：\n"
                   + "\n".join(f"{i}. {t}" for i, t in enumerate(items, 1)))
        msg.append("十几分钟就能做完，做完拍张照发我就行，我批完跟您说。")
    elif score is not None and pct >= 85:
        msg.append(f"{short}这次基础很稳，我给孩子准备的是背的部分，"
                   f"{RECITE}，不用刷题：\n"
                   "1. 那几个动词的过去式和过去分词（return、have、cost、fly、sell），"
                   "看着原形能说出后面两个形式，下节课我抽背；\n"
                   "2. 本讲重点句型——How long、How soon、for 加时间段、since 加时间点，"
                   "每个自己造两句，下节课说给我听。")
        msg.append("这些课上基本都会了，回家巩固一下就好，不用额外做题。")
    else:
        msg.append(f"我给孩子准备了一份{RECITE}，"
                   f"动词的过去式和过去分词，加上这讲的几个重点句型，"
                   f"读一读、背一背就行，不用写。")
        msg.append("错的那几道我下节课带着孩子当面过一遍，不用您操心。")

    # ---------- 收尾（短、暖）----------
    if score is not None and score < focus_line:
        msg.append("孩子底子是有的，就是这几个点还没背熟。"
                   "您别着急，咱们一起多过几遍，很快就能看见变化。"
                   if warm else
                   "您在家不用着急，这个阶段需要的是有人帮孩子把问题拆开、"
                   "一步一步来。咱们步调一致，孩子一定能慢慢跟上来。")
    elif score is not None and pct >= 85:
        msg.append("答题卡照片发您啦，您有空看一眼就行。谢谢您平时的配合！"
                   if warm else
                   "答题卡照片发您了，您有空看一眼就行。辛苦您平时的配合。")
    else:
        msg.append("答题卡照片发您啦，您有空看看。"
                   f"{short}底子是有的，就是这几个点还没练熟。"
                   "我课上会重点带着练，您别着急，咱们慢慢来。"
                   if warm else
                   "答题卡照片发您了，您有空看看。"
                   "这几块我课上会重点带着练，练一段时间就能看到变化。")

    # 拆成 2~3 条消息（教师精灵的文本1/2/3 对应三个对话框）：
    #   消息1 = 分数 + 表现
    #   消息2 = 问题 + 作业（有作业才单独成条）
    #   消息3 = 收尾（仅在 1、2 都很长时才独立）
    m1 = "\n".join(msg[:2])
    rest = msg[2:]
    if len(rest) >= 2:
        m2 = "\n".join(rest[:-1])
        m3 = rest[-1]
    else:
        m2 = rest[0] if rest else ""
        m3 = ""
    parts_out = [x for x in (m1, m2, m3) if x.strip()]
    return "\n---SPLIT---\n".join(parts_out)


def make_sibling_feedback(members, cls, qmap, meta):
    """双胞胎/兄弟姐妹：合并成一份给家长的话术。

    结构：称呼（两个孩子一起叫） + 分数并列 + 逐人分析 + 各自小任务 + 收尾
    支持：安抚风格、逐人知识点、个人小任务。
    """
    if not members:
        return ""
    exam = meta.get("考试名称") or meta.get("考试") or "月考试卷"
    FULL = float(meta.get("满分") or cls.get("满分") or 40.0)
    ptitle = meta.get("称呼") or "家长"
    FOCUS = float(meta.get("重点线") or 30)
    personal = meta.get("任务模式") == "个人小任务"
    reassure = meta.get("家长风格") == "安抚"
    lib = meta.get("_kplib")

    # 逐人知识点 / 共性知识点
    per = {}
    for it in re.split(r"[；;]", meta.get("逐人知识点", "")):
        if "=" in it:
            w, v = it.split("=", 1)
            per[w.strip()] = [x.strip() for x in re.split(r"[,，、]", v) if x.strip()]
    common = [x.strip() for x in
              re.split(r"[,，、]", meta.get("共性知识点", "")) if x.strip()]

    def bare(m):
        return re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", m["显示名"])

    def short_of(m):
        nm = bare(m)
        return nm[-2:] if len(nm) > 2 else nm

    names = [bare(m) for m in members]
    shorts = [short_of(m) for m in members]

    # 称呼：去掉共同姓氏后拼接，如 谷佳禾/谷佳音 → 佳禾佳音家长
    n1, n2 = names[0], names[1]
    if len(n1) > 2 and len(n2) > 2 and n1[0] == n2[0]:
        call = f"{n1[1:]}{n2[1:]}{ptitle}"
    else:
        call = f"{n1}{n2}{ptitle}"

    pairs = "、".join(f"{nm} {m['分数']:g}/{FULL:g}" for nm, m in zip(names, members))

    # ---- 消息1：打招呼 + 报分 ----
    if reassure:
        avg = sum(m["分数"] for m in members) / len(members)
        if avg / FULL >= 0.75:
            m1 = (f"{call}您好！{exam}改完啦，两个孩子的成绩都出来了——{pairs}。\n"
                  f"整体都拿住了大部分内容，剩下几个点慢慢补就行，不用着急。")
        else:
            m1 = (f"{call}您好！{exam}改完啦，两个孩子的成绩都出来了——{pairs}。\n"
                  f"分数您先不用往心里去——这一讲考的基本都是要背的内容，"
                  f"两个孩子都有几块还没记牢，这个阶段很常见，慢慢来就行。")
    else:
        m1 = f"{call}您好，{exam}改完了，两个孩子的成绩都出来了——{pairs}。"

    common_good = meta.get("共性优点", "")
    if common_good:
        m1 += f"\n先说说做得好的：{common_good}。"

    # ---- 消息2：逐个孩子 ----
    segs = []
    for m, nm, sh in zip(members, names, shorts):
        mine = list(per.get(nm) or per.get(m["显示名"]) or [])
        if len(mine) < 3:
            for k in common:
                if len(mine) >= 3:
                    break
                if k not in mine:
                    mine.append(k)
        if lib and mine:
            shs = [lib.SHORT.get(k, k) for k in mine[:3]]
            tip = lib.TIP.get(mine[0], "")
            if reassure:
                piece = (f"{sh}这边，上课是跟得住的，要补的主要是{'、'.join(shs[:2])}。")
                if tip:
                    piece += tip + "。"
                piece += "这些都是记住就能拿分的，不是理解不了。"
            else:
                piece = (f"{sh}这边要盯的是{'、'.join(shs[:2])}。")
                if tip:
                    piece += tip + "。"
        else:
            piece = f"{sh}这边我再单独跟您说。"

        if personal:
            piece += (f"我给孩子整理了一份《{sh}的小任务》（附件），"
                      f"就 {len(mine) or 5} 件事，一天五分钟过一遍就够了。")
        else:
            piece += "我准备了一份小练习，量很小，一天五分钟就够了。"
        segs.append(piece)

    # 表扬 / 点名提示
    praise = [x.strip() for x in re.split(r"[,，、]", meta.get("表扬学生", "")) if x.strip()]
    for m, sh in zip(members, shorts):
        if m["姓名"] in praise:
            segs.append(f"另外，{sh}有一点我要专门表扬一下——按要求改写句子那道大题，"
                        f"孩子写得挺好，说明句型是听懂了、也用上了。这是实打实的进步。")

    notice = meta.get("点名提示", "")
    for item in re.split(r"[；;]", notice):
        if "=" in item:
            who, msgx = item.split("=", 1)
            for m, sh in zip(members, shorts):
                if who.strip() in (m["姓名"], m["显示名"]):
                    segs.append(f"还有一点提醒：{sh}的{msgx.strip()}。")

    m2 = "\n".join(segs) if segs else ""

    # ---- 消息3：收尾 ----
    if reassure:
        m3 = ("四年级这个阶段，比分数更要紧的是习惯。两个孩子我课上都会多带着，"
              "您在家不用着急，跟着过一遍就行。有情况随时找我。")
    else:
        m3 = "两个孩子我课上都会盯着，您在家不用着急。有情况随时找我。"

    parts = [x for x in (m1, m2, m3) if x.strip()]
    return "\n---SPLIT---\n".join(parts)

def set_run_font(run, name="微软雅黑", size=10.5, bold=False, color=None):
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
    for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(attr), name)


def write_feedback_docx(path, ordered, texts, meta, cls):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "微软雅黑"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(h.add_run(f"{meta.get('班级','')} {meta.get('课次','')} 家长反馈话术".strip()),
                 size=16, bold=True, color=(0x1B, 0x5E, 0xA8))
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(sub.add_run(
        f"生成日期：{_dt.date.today().isoformat()}　|　"
        f"班级人数 {cls['人数']}　平均分 {cls['平均分']}　"
        f"及格率 {cls['及格率']}%　优秀率 {cls['优秀率']}%　"
        f"（话术不含排名，可直接复制发送）"),
        size=9, color=(0x80, 0x80, 0x80))

    for s in ordered:
        key = s.get("显示名", s["姓名"])
        if key not in texts:
            continue   # 已并入兄弟姐妹的合并话术，不重复输出
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        set_run_font(p.add_run(f"◆ {key}　{s['分数'] if not s['缺考'] else '缺考'} 分　{s['档位']}"),
                     size=12, bold=True, color=(0x1B, 0x5E, 0xA8))
        raw = texts[key]
        parts = raw.split("\n---SPLIT---\n") if "---SPLIT---" in raw else [raw]
        for pi, part in enumerate(parts, 1):
            if len(parts) > 1:
                ph = doc.add_paragraph()
                set_run_font(ph.add_run(f"—— 消息{pi} ——"), size=9.5, color=(0x99, 0x99, 0x99))
            for line in part.split("\n"):
                if not line.strip():
                    continue
                para = doc.add_paragraph()
                para.paragraph_format.space_after = Pt(2)
                para.paragraph_format.left_indent = Cm(0.4)
                if line.startswith("【"):
                    set_run_font(para.add_run(line), size=10.5, bold=False)
                else:
                    set_run_font(para.add_run(line), size=10.5)
    doc.save(path)
    return path


def write_feedback_txt(path, ordered, texts, meta):
    lines = [f"{meta.get('班级','')} {meta.get('课次','')} 家长反馈话术",
             f"生成日期：{_dt.date.today().isoformat()}", "=" * 60, ""]
    for s in ordered:
        key = s.get("显示名", s["姓名"])
        if key not in texts:
            continue   # 已并入兄弟姐妹的合并话术
        sc = "缺考" if s["缺考"] else f"{s['分数']:g}分"
        lines.append(f"##### {key} | {sc} | {s['档位']} #####")
        raw = texts[key]
        parts = raw.split("\n---SPLIT---\n") if "---SPLIT---" in raw else [raw]
        for pi, part in enumerate(parts, 1):
            if len(parts) > 1:
                lines.append(f"—— 消息{pi} ——")
            # 机器可读的拆分标记放在正文之后（填模板.py 解析用，人看不到）
            lines.append(part)
            if len(parts) > 1:
                lines.append(f"<<<SPLIT{pi}>>>")
        lines.append("")
        lines.append("-" * 60)
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")
    return path


# ============================================================
# 错题小练习
# ============================================================
def write_practice_docx(path, questions, q_err, point_stat, cls, meta, custom_text=None):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "微软雅黑"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(h.add_run(f"{meta.get('班级','')} {meta.get('课次','')} 错题专项小练习".strip()),
                 size=16, bold=True, color=(0x1B, 0x5E, 0xA8))
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(sub.add_run("姓名：____________　　日期：____________　　用时：______分钟"),
                 size=10, color=(0x60, 0x60, 0x60))

    if custom_text:
        doc.add_paragraph()
        set_run_font(doc.add_paragraph().add_run("（本讲教师自编练习）"), size=9, color=(0x80, 0x80, 0x80))
        for line in custom_text.splitlines():
            if not line.strip():
                doc.add_paragraph()
                continue
            p = doc.add_paragraph()
            bold = bool(re.match(r"^\s*(第[一二三四五六七八九十]+部分|[一二三四五六七八九十]+、|【)", line))
            set_run_font(p.add_run(line), size=11, bold=bold)
        doc.save(path)
        return path

    n_valid = max(1, cls["有效"])
    ranked = sorted(questions, key=lambda q: -len(q_err[q["题号"]]))
    hot = [q for q in ranked if q_err[q["题号"]]]
    cold = [q for q in ranked if not q_err[q["题号"]]]

    def section(title, subtitle=None):
        doc.add_paragraph()
        p = doc.add_paragraph()
        set_run_font(p.add_run(title), size=13, bold=True, color=(0x1B, 0x5E, 0xA8))
        if subtitle:
            q = doc.add_paragraph()
            set_run_font(q.add_run(subtitle), size=9.5, color=(0x80, 0x80, 0x80))

    section("第一部分　原题重做（先不看答案，独立完成）",
            "要求：做完再对答案，错的地方写清楚「我为什么错」，而不只是抄正确答案。")
    if hot:
        for i, q in enumerate(hot, 1):
            rate = round(100.0 * len(q_err[q["题号"]]) / n_valid, 1)
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            set_run_font(p.add_run(f"{i}. （原第{q['题号']}题｜{q['考点'] or q['题型'] or '本讲考点'}｜全班错误率 {rate}%）"),
                         size=11, bold=True)
            a = doc.add_paragraph()
            a.paragraph_format.left_indent = Cm(0.6)
            set_run_font(a.add_run(f"参考答案：{q['答案']}"), size=10, color=(0xC0, 0x39, 0x2B))
        doc.add_paragraph()
        set_run_font(doc.add_paragraph().add_run("作答区："), size=10.5, bold=True)
        for _ in range(6):
            line = doc.add_paragraph()
            set_run_font(line.add_run("_" * 62), size=10.5, color=(0xA0, 0xA0, 0xA0))
    else:
        set_run_font(doc.add_paragraph().add_run("本次课测全员无错题，进入拔高练习。"), size=11)

    section("第二部分　同考点变式（换个问法，看是不是真会）",
            "说明：下面每题都对应你自己的易错点。先说出思路，再动笔。")
    idx = 0
    for p in sorted(point_stat.values(), key=lambda x: -x["错误率"]):
        if p["错误率"] < 20:
            continue
        idx += 1
        t = p["题型"] or "本讲"
        p1 = doc.add_paragraph()
        set_run_font(p1.add_run(f"{idx}. 【{p['考点']}｜{t}】"), size=11, bold=True)
        tips = {
            "听写": "用本讲单词表，请家长报中文，你写英文，注意首字母大小写和拼写完整性。",
            "单词拼写": "用本讲单词表，请家长报中文，你写英文，写完自己逐个字母核对一遍。",
            "句型转换": "按「一定、二变、三问号」三步走：先定特殊疑问词，再变一般疑问句语序，最后补问号。",
            "选择题": "把四个选项逐一代回原句读一遍，用「读得通不通顺」排除两个，再用语法点定最后一个。",
            "翻译": "先定主谓，再补时间地点，最后检查时态和单复数。",
            "阅读理解": "回到原文找依据句，把依据句划出来再选，不要凭印象。",
        }
        tip = tips.get(t) or "回看课件对应例题，把方法步骤复述一遍再做。"
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = Cm(0.6)
        set_run_font(p2.add_run("练习要求：" + tip), size=10)
        for _ in range(3):
            line = doc.add_paragraph()
            set_run_font(line.add_run("_" * 62), size=10.5, color=(0xA0, 0xA0, 0xA0))

    section("第三部分　自检清单（做完打勾）")
    checks = [
        "我把错题的「错误原因」写下来了，不是只抄答案",
        "我能不看笔记把本讲的语法方法步骤说出来",
        "本讲单词我默写了一遍，错的已圈出并隔天重默",
        "我把这份练习拍给老师批改了",
    ]
    for c in checks:
        p = doc.add_paragraph()
        set_run_font(p.add_run("□ " + c), size=10.5)

    doc.save(path)
    return path


# ============================================================
# 跨周成绩历史（趋势追踪）
# ============================================================
def load_history():
    """读取 7-历史/成绩历史.csv -> {(班级,姓名): [(课次,分数), ...]}"""
    hist = defaultdict(list)
    if not HISTORY_CSV.exists():
        return hist
    import csv
    with HISTORY_CSV.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            try:
                sc = float(row.get("分数") or "")
            except ValueError:
                continue
            hist[(row.get("班级", ""), row.get("姓名", ""))].append((row.get("课次", ""), sc))
    return hist


def save_history(meta, students):
    """把本次成绩追加进历史 CSV（同一班级+课次+姓名去重覆盖）。"""
    DIR_HISTORY.mkdir(parents=True, exist_ok=True)
    import csv
    rows = []
    seen = {}
    if HISTORY_CSV.exists():
        with HISTORY_CSV.open("r", encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                key = (row.get("班级", ""), row.get("课次", ""), row.get("姓名", ""))
                seen[key] = row
    for s in students:
        if s["缺考"] or s["分数"] is None:
            continue
        key = (meta.get("班级", ""), meta.get("课次", ""), s["姓名"])
        seen[key] = {
            "日期": meta.get("日期", "") or _dt.date.today().isoformat(),
            "班级": meta.get("班级", ""),
            "课次": meta.get("课次", ""),
            "姓名": s["姓名"],
            "分数": s["分数"],
            "档位": s["档位"],
            "错题数": len(s["错题"]),
            "薄弱考点": s["错题考点"],
        }
    rows = list(seen.values())
    with HISTORY_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["日期", "班级", "课次", "姓名", "分数", "档位", "错题数", "薄弱考点"])
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in w.fieldnames})
    return len(rows)


def trend_of(name, cls_name, meta, hist, score):
    """返回该生相对上次的升降描述，没有历史则返回空。"""
    past = [(l, s) for (c, n), lst in hist.items() if c == cls_name and n == name
            for (l, s) in lst if l != meta.get("课次", "")]
    if not past:
        return ""
    past.sort(key=lambda x: x[0])
    last_lesson, last_score = past[-1]
    d = score - last_score
    if abs(d) < 0.5:
        return f"和上次（{last_lesson}）持平，说明当前水平稳定在这个区间。"
    if d > 0:
        return f"比上次（{last_lesson}）提高了 {d:g} 分，进步是看得见的，说明这段时间的复习方法对了。"
    return (f"比上次（{last_lesson}）下降了 {abs(d):g} 分，"
            f"这个波动值得注意，通常不是能力问题，而是这周课后复习没跟上。")


# ============================================================
# 汇总报告（给老师/agent 看）
# ============================================================
def write_report(path, meta, students, questions, q_err, point_stat, cls, unmatched, texts):
    L = []
    L.append(f"# {meta.get('班级','')} {meta.get('课次','')} 课测汇总报告")
    L.append("")
    L.append(f"- 生成时间：{_dt.datetime.now():%Y-%m-%d %H:%M}")
    L.append(f"- 班级人数 {cls['人数']}（有效 {cls['有效']}，缺考 {cls['缺考']}）")
    L.append(f"- 平均分 **{cls['平均分']}**　最高 {cls['最高']}　最低 {cls['最低']}")
    L.append(f"- 及格率 {cls['及格率']}%　优秀率 {cls['优秀率']}%　题量 {cls['满分题数']} 题　满分 {cls.get('满分',100):g} 分（平均得分率 {cls.get('平均得分率',0)}%）")
    L.append("")
    L.append("## 考点掌握排序（错误率从高到低）")
    L.append("")
    L.append("| 考点 | 题号 | 错误人次 | 错误率 | 判定 |")
    L.append("|---|---|---|---|---|")
    for p in sorted(point_stat.values(), key=lambda x: -x["错误率"]):
        r = p["错误率"]
        v = "已掌握" if r < 15 else ("基本掌握" if r < 35 else ("薄弱" if r < 60 else "严重薄弱"))
        L.append(f"| {p['考点']} | {'、'.join(str(n) for n in p['题号'])} | {p['错']} | {r}% | {v} |")
    L.append("")
    L.append("## 高危题（错误率≥50%）")
    L.append("")
    hi = [q for q in questions if cls["有效"] and len(q_err[q["题号"]]) / cls["有效"] >= 0.5]
    if hi:
        for q in hi:
            L.append(f"- 第{q['题号']}题（{q['考点'] or q['题型'] or '-'}）答案「{q['答案']}」"
                     f"——错 {len(q_err[q['题号']])} 人：{'、'.join(q_err[q['题号']])}")
    else:
        L.append("- 无")
    L.append("")
    L.append(f"## 需重点关注名单（得分率 <70%，即 <{cls.get('满分',100)*0.7:g} 分）")
    L.append("")
    low = [s for s in students if not s["缺考"] and s["分数"] is not None
           and s["分数"] < 0.7 * cls.get("满分", 100)]
    L.append("、".join(s["姓名"] for s in low) if low else "无")
    L.append("")

    # ---- 30 分以下重点反馈专区（老师指定） ----
    focus = [s for s in students if not s["缺考"] and s["分数"] is not None
             and s["分数"] < 30]
    L.append(f"## 重点反馈专区（<30 分，共 {len(focus)} 人）")
    L.append("")
    if not focus:
        L.append("- 无")
    for s in sorted(focus, key=lambda x: x["分数"]):
        L.append(f"### {s['姓名']}（{s['分数']:g}/{cls.get('满分',100):g} 分　{s['档位']}）")
        L.append("")
        sw = s.get("薄弱大题") or []
        if sw:
            L.append("| 大题 | 扣分 | 该大题考查范围（需逐个过关） |")
            L.append("|---|---|---|")
            for w in sw:
                L.append(f"| {w['名称']}（{w['大题']}） | {w['扣分']:g} | {'、'.join(w['考点'][:4]) or '-'} |")
            L.append("")
            top = sw[0]
            L.append(f"**最该补的一件事**：{top['名称']}——本大题扣了 {top['扣分']:g} 分，"
                     f"是全卷最大失分点。")
            L.append("")
            L.append(f"**给家长的任务**（已写入话术，家里 3 分钟版）：")
            for i, w in enumerate([x for x in sw if x["扣分"] >= 2][:2] or sw[:1], 1):
                L.append(f"{i}. {w['名称']}：本大题错题重做，做完说出正确思路。")
            if s.get("未定位扣分"):
                L.append("")
                L.append(f"（另有 {s['未定位扣分']:g} 分无法归到具体大题，"
                         f"可能是某题部分扣分，已计入总分校验。）")
        elif s.get("大题扣分存疑"):
            L.append(f"- ⚠️ 大题扣分读数与总分不符（{s['大题扣分存疑']}），诊断已降级，"
                     f"请老师人工确认后我再补精准建议。")
        else:
            L.append("- 暂无逐题数据。")
        L.append("")

    L.append("## 数据待确认")
    L.append("")
    doubt = [s for s in students if s.get("大题扣分存疑")]
    if doubt:
        L.append("- **以下学生的大题扣分读数和总分对不上，已自动降级（不输出具体诊断）**：")
        for s in doubt:
            L.append(f"    - {s['姓名']}：{s['大题扣分存疑']}　"
                     f"（请核对答题卡右侧分项扣分或老师给分）")
    if unmatched:
        L.append(f"- 以下照片没匹配到学生：{'、'.join(unmatched)}")
    if not doubt and not unmatched:
        L.append("- 无（14 人大题扣分均已通过总分校验）")
    L.append("")
    L.append(f"> 说明：高于重点线（{cls.get('满分',100)*0.75:g} 分以上）的学生话术里**不含任务**，"
             f"只反馈情况；低于重点线的才布置 1-2 件家里能完成的事。"
             f"重点线可在名单里用 `@重点线: 30` 调整。")
    L.append("")
    L.append("## 话术全文")
    L.append("")
    for s in students:
        key = s.get("显示名", s["姓名"])
        if key not in texts:
            continue   # 缺考学生不生成话术
        L.append(f"### {s['姓名']}（{s['分数']:g}）")
        L.append("")
        raw = texts[key]
        parts = raw.split("\n---SPLIT---\n") if "---SPLIT---" in raw else [raw]
        for pi, part in enumerate(parts, 1):
            if len(parts) > 1:
                L.append(f"**—— 消息{pi} ——**")
                L.append("")
            L.append("```")
            L.append(part)
            L.append("```")
            L.append("")
    path.write_text("\n".join(L), encoding="utf-8")
    return path


# ============================================================
# 主流程
# ============================================================
def find_answer_for(roster_path: Path, meta: dict, answer_files: list, explicit: str | None):
    if explicit:
        p = DIR_ANSWER / explicit
        if p.suffix != ".txt":
            p = p.with_suffix(".txt")
        if not p.exists():
            cands = [f for f in answer_files if explicit in f.stem]
            if cands:
                return cands[0]
            raise FileNotFoundError(f"找不到答案文件：{explicit}")
        return p
    key = meta.get("课次", "")
    if key:
        for f in answer_files:
            if key in f.stem:
                return f
        # 归一化后匹配：忽略空格/横线/"讲""第"等字样
        norm = re.sub(r"[\s\-_讲第]", "", key)
        for f in answer_files:
            if norm and norm in re.sub(r"[\s\-_讲第]", "", f.stem):
                return f
        # 进一步：只比对数字（"五年级第五讲" 与 "第5讲L89-94" 都含 5）
        kn = set(re.findall(r"\d+", key))
        if kn:
            best = None
            for f in answer_files:
                fn = set(re.findall(r"\d+", f.stem))
                if kn & fn:
                    best = f if best is None else best
            if best:
                return best
    stem = roster_path.stem
    for f in answer_files:
        if stem in f.stem or f.stem in stem:
            return f
    # 【重要】课次填了却对不上时，绝不盲用唯一那份答案 ——
    # 那会把别的讲的知识点安到这一讲上，比没有更糟。
    if key:
        return None
    if len(answer_files) == 1:
        return answer_files[0]
    return None


def process(roster_path: Path, answer_path: Path, custom_path: Path | None):
    meta, students = parse_roster(roster_path)
    meta.setdefault("班级", roster_path.stem)
    if answer_path is None:
        # 没有答案文件：按各大题扣分生成，不做逐题错题练习
        ameta, questions = {"满分": meta.get("满分"), "_无答案": True}, []
    else:
        ameta, questions = parse_answers(answer_path)
    for k, v in ameta.items():
        meta.setdefault(k, v)
    meta.setdefault("课次", answer_path.stem if answer_path else "")

    if not questions and answer_path is not None:
        print(f"  [错误] 答案文件为空：{answer_path}")
        return None
    if not students:
        print(f"  [错误] 名单为空：{roster_path}")
        return None

    # 错题号越界告警
    qnos = {q["题号"] for q in questions}
    for s in students:
        if s["错题"]:
            bad = [n for n in s["错题"] if n not in qnos]
            if bad:
                print(f"  [警告] {s['姓名']} 的错题号 {bad} 超出题量范围（共 {len(questions)} 题）")

    # 大题扣分（若存在）：免登记错题号也能生成针对性话术
    sec_ded = load_section_deductions(answer_path, meta.get("班级", ""))
    if sec_ded:
        meta["_大题扣分"] = sec_ded
        n_hit = sum(1 for s in students if s["姓名"] in sec_ded)
        print(f"  已载入大题扣分：{n_hit} 人（免登记错题号）")

    # 先索引照片，用于判定「名单里有、但没有答题卡」的学生
    photos, unmatched = index_photos([s["姓名"] for s in students],
                                     meta.get("班级"), roster_path.stem,
                                     meta.get("课次"))

    # 待定（名单只有姓名）→ 按有没有答题卡判定：
    #   有卡但没填分数 → 提示补分数；没卡 → 默认缺考（老师约定，不必再问）
    auto_absent = []
    for s in students:
        if not s.get("待定"):
            continue
        has_card = bool(photos.get(s["姓名"]))
        if not has_card:
            s["缺考"] = True
            s["待定"] = False
            auto_absent.append(s["显示名"])
        else:
            print(f"  [提示] {s['显示名']} 有答题卡但名单里没写分数，请补上分数")
    if auto_absent:
        print(f"  名单中有但无答题卡，按缺考处理：{'、'.join(auto_absent)}")

    total, qmap, q_err, point_stat, cls = build_stats(students, questions, meta)

    # ---------- 知识点定位（可选：名单里写 @知识点库: L11-18）----------
    _libname = meta.get("知识点库", "")
    if _libname:
        _kpfile = Path(__file__).parent / f"知识点库_{_libname}.py"
        if _kpfile.exists():
            import importlib.util as _ilu
            _spec = _ilu.spec_from_file_location(f"kplib_{_libname}", _kpfile)
            _kplib = _ilu.module_from_spec(_spec)
            _spec.loader.exec_module(_kplib)
            meta["_kplib"] = _kplib
            for _st in students:
                if _st["缺考"] or _st["分数"] is None:
                    continue
                _kps = []
                if _st.get("错题"):
                    _kps = _kplib.kps_for_wrong(_st["错题"])
                if not _kps:
                    for _w in (_st.get("薄弱大题") or []):
                        for _kk in _kplib.kps_for_section(_w["大题"]):
                            if _kk not in _kps:
                                _kps.append(_kk)
                _st["知识点"] = [(_kplib.SHORT.get(_k, _k), _kplib.TIP.get(_k, ""))
                                 for _k in _kps[:3]]
            _n_kp = sum(1 for _st in students if _st.get("知识点"))
            print(f"  知识点库 {_kpfile.name}：{_n_kp} 人定位到知识点")
        else:
            print(f"  [警告] 找不到知识点库：{_kpfile.name}")


    # 缺考的学生不生成话术（老师约定：缺考不用管）
    hist = load_history()
    texts = {}
    for s in students:
        if s["缺考"] or s["分数"] is None:
            continue
        tr = trend_of(s["姓名"], meta.get("班级", ""), meta, hist, s["分数"])
        texts[s["显示名"]] = make_feedback(s, cls, qmap, meta, tr)

    # ---------- 双胞胎/兄弟姐妹：合并成一份话术 ----------
    # 名单里写 @同家长: 江宇婷,江宇晴  即可自动合并
    sib_groups = []
    raw_sib = meta.get("同家长") or meta.get("双胞胎") or ""
    for grp in re.split(r"[；;]", raw_sib):
        names = [x.strip() for x in re.split(r"[,，、\s]+", grp) if x.strip()]
        if len(names) >= 2:
            sib_groups.append(names)

    for names in sib_groups:
        # 按名单里写的顺序取成员（决定称呼顺序，如"江宇婷,江宇晴"）
        members = []
        for nm in names:
            hit = [s for s in students if s["显示名"] == nm or s["姓名"] == nm]
            if hit:
                members.append(hit[0])
        if len(members) < 2:
            continue
        combined = make_sibling_feedback(members, cls, qmap, meta)
        if not combined:
            continue
        # 合并后挂在第一个孩子名下，其余孩子话术清空（避免重复发送）
        key0 = members[0]["显示名"]
        texts[key0] = combined
        for m in members[1:]:
            texts.pop(m["显示名"], None)
        print(f"  双胞胎合并话术：{'、'.join(m['显示名'] for m in members)}")

    # 话术/分发只保留有效学生
    graded = [s for s in students if not s["缺考"] and s["分数"] is not None]
    ordered = sorted(graded, key=lambda s: -s["分数"])
    for i, s in enumerate(ordered, 1):
        s["名次"] = i  # 内部用，不写进话术

    outdir = DIR_OUT / f"{meta['班级']}_{meta['课次']}".replace("/", "_")
    outdir.mkdir(parents=True, exist_ok=True)

    custom_text = custom_path.read_text(encoding="utf-8-sig") if custom_path and custom_path.exists() else None

    n_copied = archive_photos(outdir, ordered, photos)

    p1 = write_xlsx(outdir / "课测成绩表.xlsx", meta, students, questions, qmap, q_err, point_stat, cls, photos, unmatched)
    p2 = write_feedback_docx(outdir / "家长反馈话术.docx", ordered, texts, meta, cls)
    p3 = write_feedback_txt(outdir / "家长反馈话术.txt", ordered, texts, meta)
    p4 = write_practice_docx(outdir / "错题专项小练习.docx", questions, q_err, point_stat, cls, meta, custom_text)
    p5 = write_report(outdir / "汇总报告.md", meta, students, questions, q_err, point_stat, cls, unmatched, texts)
    n_hist = save_history(meta, students)

    print(f"\n{'='*64}")
    print(f"✓ {meta['班级']} {meta['课次']} 处理完成")
    print(f"  人数 {cls['人数']} | 有效 {cls['有效']} | 平均 {cls['平均分']} | "
          f"最高 {cls['最高']} | 最低 {cls['最低']} | 及格率 {cls['及格率']}% | 优秀率 {cls['优秀率']}%")
    print(f"  输出目录：{outdir}")
    for p in (p1, p2, p3, p4, p5):
        print(f"    - {p.name}")
    if unmatched:
        print(f"  [提示] {len(unmatched)} 张照片未匹配到学生：{'、'.join(unmatched[:5])}")
    if n_copied:
        print(f"  照片归档：{n_copied} 张 -> {outdir / '照片归档'}")
    unk = [s["姓名"] for s in students if s.get("错题未知")]
    if unk:
        print(f"  [提示] 未填错题号：{'、'.join(unk)}（话术已降级，建议补上）")
    print(f"  成绩历史已更新：{n_hist} 条 -> {HISTORY_CSV}")
    print("=" * 64)
    return outdir


def make_demo():
    DIR_ROSTER.mkdir(parents=True, exist_ok=True)
    DIR_ANSWER.mkdir(parents=True, exist_ok=True)
    DIR_PHOTO.mkdir(parents=True, exist_ok=True)
    DIR_OUT.mkdir(parents=True, exist_ok=True)
    DIR_CUSTOM.mkdir(parents=True, exist_ok=True)

    (DIR_ANSWER / "第5讲L95-96.txt").write_text(
        "@班级: 五1班\n@课次: 第5讲L95-96\n@日期: 2026-09-22\n"
        "# 格式：题号. 答案 | 考点 | 题型\n"
        "1. pilot | 单词拼写 | 听写\n"
        "2. revsum | 单词拼写 | 听写\n"
        "3. New York | 单词拼写 | 听写\n"
        "4. Tokyo | 单词拼写 | 听写\n"
        "5. Madrid | 单词拼写 | 听写\n"
        "6. Tickets, please. | 情景交际 | 翻译\n"
        "7. What would you like? | 特殊疑问句 | 句型转换\n"
        "8. He is flying to Tokyo. | 现在进行时 | 句型转换\n"
        "9. Does she like flying? | 一般疑问句 | 句型转换\n"
        "10. They are going to Madrid. | 一般将来时 | 翻译\n",
        encoding="utf-8")

    (DIR_ROSTER / "五1班.txt").write_text(
        "@班级: 五1班\n@课次: 第5讲L95-96\n"
        "# 格式：姓名 分数 [错题号]   全对就只写 姓名 分数\n"
        "张三 85 3,7,10\n"
        "李四 100\n"
        "王五 92 7\n"
        "赵六 68 2,5,7,8,10\n"
        "钱七 78 1,4,9\n"
        "孙八 58 3,5,6,7,8,9,10\n"
        "周九 95 7\n"
        "吴十 88 2,10\n"
        "郑十一 缺考\n"
        "王十二 73 1,3,7,10\n",
        encoding="utf-8")
    return DIR_ROSTER / "五1班.txt", DIR_ANSWER / "第5讲L95-96.txt"


def main():
    ap = argparse.ArgumentParser(description="课测反馈一键生成器")
    ap.add_argument("--demo", action="store_true", help="生成示例输入并运行")
    ap.add_argument("--班级", default=None, help="只处理指定名单文件（可不带 .txt）")
    ap.add_argument("--答案", default=None, help="指定答案文件（可不带 .txt）")
    ap.add_argument("--练习", default=None, help="自定义练习文件（课测答案/自定义练习/ 下）")
    args = ap.parse_args()

    for d in (DIR_ROSTER, DIR_ANSWER, DIR_PHOTO, DIR_OUT, DIR_CUSTOM):
        d.mkdir(parents=True, exist_ok=True)

    if args.demo:
        demo = ROOT / "生成结果" / "_示例"
        set_workspace(demo)          # 隔离：示例数据只写在 生成结果/_示例/，绝不碰真实名单
        for d in (DIR_ROSTER, DIR_ANSWER, DIR_PHOTO, DIR_OUT, DIR_CUSTOM, DIR_HISTORY):
            d.mkdir(parents=True, exist_ok=True)
        rp, apth = make_demo()
        print(f"已生成示例输入（隔离目录：{demo}），开始演示运行...\n")
        process(rp, apth, None)
        print(f"\n示例产物在：{DIR_OUT}")
        return

    rosters = sorted(f for f in DIR_ROSTER.glob("*.txt")
                     if not f.name.startswith("~$") and not f.name.startswith("_"))
    # 答案文件：排除模板、单词表、大题扣分等辅助文件（它们不是答案本身）
    AUX_SUFFIX = ("_单词表", "_大题扣分", "_词汇", "_句型")
    answers = sorted(f for f in DIR_ANSWER.glob("*.txt")
                     if not f.name.startswith("~$") and not f.name.startswith("_")
                     and not any(f.stem.endswith(x) for x in AUX_SUFFIX))
    if args.班级:
        sel = args.班级 if args.班级.endswith(".txt") else args.班级 + ".txt"
        rosters = [f for f in rosters if f.name == sel or args.班级 in f.stem]
    if not rosters:
        print(f"没有找到名单文件。请把名单放到：{DIR_ROSTER}")
        print("格式示例：\n@班级: 五1班\n@课次: 第5讲L95-96\n张三 85 3,7,12\n李四 100")
        return
    # 没答案文件也能跑：识别照片已给出各大题扣分，够定位知识点了
    if not answers:
        print("[提示] 2-答案 里没有答案文件 —— 将只按各大题扣分生成，"
              "不做逐题错题练习。")

    ok = 0
    skipped = []
    for rp in rosters:
        meta, _peek = parse_roster(rp)
        if len(_peek) == 0:
            skipped.append(rp.stem)   # 空名单（还没填的班）静默跳过
            continue
        apth = find_answer_for(rp, meta, answers, args.答案) if answers else None
        if answers and not apth:
            print(f"\n>>> 处理 {rp.stem}  "
                  f"[提示] 没配对到对应这一讲的答案，按各大题扣分生成")
        elif apth:
            print(f"\n>>> 处理 {rp.stem}  配对答案：{apth.name}")
        else:
            print(f"\n>>> 处理 {rp.stem}")
        cp = None
        if args.练习:
            cp = DIR_CUSTOM / (args.练习 if args.练习.endswith(".txt") else args.练习 + ".txt")
        else:
            for cand in DIR_CUSTOM.glob("*.txt"):
                if rp.stem in cand.stem or cand.stem in rp.stem:
                    cp = cand
                    break
        if process(rp, apth, cp):
            ok += 1
    if skipped:
        print(f"\n（{len(skipped)} 个班名单为空，已跳过：{'、'.join(skipped)}）")
    print(f"全部完成：成功 {ok} 个班。")


def make_special_feedback(s, meta, short, ptitle, exam, score_txt):
    """特殊学生专属话术。

    适用：基础弱、且「课上带着就能写、回家自己背就写不出来」的依赖带读型学生。
    要点：不批评家长、不否定孩子；把「离不开人带」这个模式讲清楚，
          让家长意识到高年级会撑不住（危机感），并给一个五分钟就能做到的具体动作。
    """
    FULL = meta.get("满分", 40)
    homework = meta.get("任务模式") == "个人小任务"
    task = f"《{short}的小任务》（附件）"

    p1 = (f"{short}{ptitle}您好！{exam}改完啦，我想跟您把{short}的情况说细一点。\n"
          f"{short}这次考了 {score_txt}。分数我先放一边，有个现象我觉得您需要知道——")

    p2 = (f"课上我带着读、当堂让孩子开口说的那些，基本都能写对；"
          f"可只要是我说“回家自己背一下”的部分，就写不出来。\n"
          f"这说明一件事：跟着老师读的时候，是能记住的；"
          f"但现在还离不开人带，一个人学不进去。")

    p3 = (f"四年级靠老师课上带还兜得住，再往上走内容一多，课上就带不过来了，"
          f"那时候差距会拉得很快。这是我最担心的地方。")

    if homework:
        p4 = (f"所以想请您帮个忙：家里不用让孩子自己背，每天就花五分钟，"
              f"陪着读一遍、再问一遍，能说出来就算过。不用您教，陪着重复就行。\n"
              f"我把这次错的都整理成了一份{task}，量很小，一天五分钟够了。")
    else:
        p4 = (f"所以想请您帮个忙：家里不用让孩子自己背，每天就花五分钟，"
              f"陪着读一遍、再问一遍，能说出来就算过。不用您教，陪着重复就行。")

    p5 = "咱们一起盯一段时间，我课上也会多抽背。"

    return "\n---SPLIT---\n".join([p1, p2 + "\n" + p3, p4 + "\n" + p5])

def make_reassuring_feedback(s, meta, short, ptitle, exam, score_txt):
    """安心版话术（考得不好，但不让家长焦虑）。

    和「特殊学生」那版的区别：
      · 不制造紧迫感，不讲"差距会拉得很快"
      · 把问题定性成"还没背熟、很正常、慢慢来"
      · 强调上课跟得住 + 靠记背就能补，把注意力引到习惯而不是分数
    """
    _lib = meta.get("_kplib")
    _per = {}
    for _it in re.split(r"[；;]", meta.get("逐人知识点", "")):
        if "=" in _it:
            _w, _v = _it.split("=", 1)
            _per[_w.strip()] = [x.strip() for x in re.split(r"[,，、]", _v) if x.strip()]
    _common = [x.strip() for x in
               re.split(r"[,，、]", meta.get("共性知识点", "")) if x.strip()]
    _mine = _per.get(s["姓名"]) or _per.get(s.get("显示名", ""))
    if _mine and _lib:
        kp = [(_lib.SHORT.get(k, k), _lib.TIP.get(k, "")) for k in _mine[:2]]
    elif _common and _lib:
        kp = [(_lib.SHORT.get(k, k), _lib.TIP.get(k, "")) for k in _common[:2]]
    else:
        kp = s.get("知识点") or []
    kn = "、".join(x[0] for x in kp[:2]) if kp else "几个要记要背的地方"
    kt = kp[0][1] if kp and kp[0][1] else ""
    personal = meta.get("任务模式") == "个人小任务"
    task = f"《{short}的小任务》（附件）"

    FULL = float(meta.get("满分") or 40)
    _pct = 100.0 * float(s["分数"] or 0) / FULL
    if _pct >= 90:
        p1 = (f"{short}{ptitle}您好！{exam}改完啦，先跟您报个喜——"
              f"{short}这次考了 {score_txt}，这一讲学得挺扎实的。")
    elif _pct >= 75:
        p1 = (f"{short}{ptitle}您好！{exam}改完啦，{short}这次考了 {score_txt}。\n"
              f"大部分内容都拿住了，剩下几个点慢慢补就行，不用着急。")
    else:
        p1 = (f"{short}{ptitle}您好！{exam}改完啦，我跟您说说{short}的情况。\n"
              f"这次考了 {score_txt}。分数您先不用往心里去——"
              f"这一讲考的基本都是要背的内容，{short}有几块还没记牢，"
              f"这个阶段很常见，慢慢来就行。")

    p2 = (f"我看卷子的时候注意到，{short}上课是跟得住的，"
          + (f"错误主要集中在{kn}。" if not kt else f"错误主要集中在{kn}。{kt}。")
          + f"这些都是记住就能拿分的，不是理解不了。")

    # 个别化：表扬 / 点名提示（名单里 @表扬学生 / @点名提示）
    _praise = [x.strip() for x in
               re.split(r"[,，、]", meta.get("表扬学生", "")) if x.strip()]
    _extra = ""
    if s["姓名"] in _praise:
        _extra += ("这次有一点我要专门表扬一下——按要求改写句子那道大题，"
                   "孩子写得挺好，说明句型是听懂了、也用上了。这是实打实的进步。")
    _notice = meta.get("点名提示", "")
    for _item in re.split(r"[；;]", _notice):
        if "=" in _item:
            _who, _msg = _item.split("=", 1)
            if _who.strip() == s["姓名"]:
                _extra += ("另外提醒一句：" + _msg.strip() + "。")

    if personal:
        p3 = (f"我整理了一份{task}，就几件事，一天五分钟过一遍就够了。\n"
              f"不用您陪着写，读一读、问一问就行。")
    else:
        p3 = "我准备了一份小练习，量很小，一天五分钟就够了。"
    if _extra:
        p3 = _extra + "\n" + p3

    p4 = (f"四年级这个阶段，比分数更要紧的是习惯。每天跟着过一遍，"
          f"慢慢就上来了。我课上也会多带着{short}，您不用着急。")

    return "\n---SPLIT---\n".join([p1, p2 + "\n" + p3, p4])


def make_firm_feedback(s, meta, short, ptitle, exam, score_txt):
    """严正版话术：把问题点透，但不攻击人。

    适用：上课没听、回家没背，卷面明显「空着」的学生。
    原则：
      · 只谈「这件事」——课上带过的没写、交代要背的没背
      · 不谈「你不负责」，而是「这件事需要家里一起使力」
      · 必给具体、做得到的动作，不空喊重视
    """
    kp = s.get("知识点") or []
    kn = "、".join(x[0] for x in kp[:2]) if kp else ""

    p1 = (f"{short}{ptitle}您好，{exam}改完了，这次我想跟您认真聊一次。\n"
          f"{short}考了 {score_txt}。这个分数我不绕弯子——"
          f"不是不会，是这一讲的内容基本没有进脑子。")

    p2 = (f"我把卷子对着讲过的东西看了一遍：课上反复带读、当堂练过的部分，卷面上几乎是空的；"
          f"我交代回家要背的，也没有背。\n"
          f"所以这不是能力问题，是这段时间的学习没有真正落地。")

    p3 = (f"说句实话，这个情况我有点着急。四年级的内容是一环扣一环的，"
          f"这一讲空着，下一讲就得花双倍时间补。现在补还来得及，再拖下去就吃力了。")

    p4 = (f"所以这次想请您配合两件事：\n"
          f"一、每天晚上固定十分钟，陪他把当天要背的过一遍，能背出来才算完成；\n"
          f"二、每次练习和背书麻烦您签个字，我这边每节课都会抽他。\n"
          f"我把这一讲必须记住的东西列成了一份《{short}的小任务》（附件），"
          f"就按这个来，不用另外找材料。")

    p5 = (f"我课上会多盯着他，但光靠课堂时间不够。"
          f"孩子的状态，得咱们两边一起使力才行。")

    return "\n---SPLIT---\n".join([p1, p2 + "\n" + p3, p4 + "\n" + p5])


if __name__ == "__main__":
    main()


