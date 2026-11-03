# -*- coding: utf-8 -*-
r"""roster.py —— 名单解析的唯一实现（原来 6 个脚本各写一套，容易各自出错）。

名单行支持四种写法，老师怎么方便怎么填：

    李小明 30                     姓名 分数
    STU2025001 李小明 30        学号 姓名 分数          ← 教师精灵导出名单常用
    李小明 30 1,4,10              姓名 分数 错题号
    STU2025001 李小明 30 1,4,10 学号 姓名 分数 错题号

学号的判定：纯数字，或「字母前缀 + 至少 3 位数字」（STU2025001、2025001、A12）。
姓名的判定：2-4 个汉字，允许一个字母前缀（A李小明）。
两个条件不会同时成立，所以不会误判。
"""
from __future__ import annotations

import re
from pathlib import Path

NAME_RE = re.compile(r"^[\u4e00-\u9fff]{2,4}$")
ID_RE = re.compile(r"^[A-Za-z]{0,6}\d{3,}$")


def _bare_name(t: str) -> str:
    return re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", t or "")


def looks_name(t: str) -> bool:
    return bool(NAME_RE.fullmatch(_bare_name(t)))


def looks_id(t: str) -> bool:
    return bool(ID_RE.fullmatch((t or "").strip())) and not looks_name(t)


def strip_id(parts: list[str]) -> list[str]:
    """如果第一个 token 是学号，去掉它。各脚本解析名单时统一调这个。"""
    if len(parts) >= 2 and looks_id(parts[0]):
        return parts[1:]
    return parts


def norm_key(x: str) -> str:
    """姓名归一化，用于跨文件匹配（去字母前缀、空格、括号、数字）。"""
    x = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", x or "")
    return re.sub(r"[\s_\-（）()0-9.]", "", x)


def split_errors(token: str):
    """'3、7、12' / '错3,7' / '全对' -> [3,7,12] / []"""
    t = (token or "").strip()
    if not t:
        return None
    if re.search(r"全对|没有错|无错|无$|满分|perfect|^-$|^—$", t, re.I):
        return []
    t = re.sub(r"^错[:：]?", "", t).strip()
    return sorted({int(x) for x in re.findall(r"\d+", t)})


def parse_line(line: str):
    """一行 -> dict(学号, 姓名, 分数, 缺考, 待定, 错题)；不是学生行返回 None。"""
    parts = [p for p in re.split(r"[\s,，\t]+", line.strip()) if p]
    if not parts:
        return None
    sid = ""
    if len(parts) >= 2 and looks_id(parts[0]):
        sid, parts = parts[0], parts[1:]
    name, rest = parts[0], parts[1:]
    score, absent, no_score = None, False, len(parts) == 1
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
            if tail:
                err_tok = tail
            rest = rest[i + 1:]
            break
    if not absent:
        if err_tok is None and rest:
            err_tok = " ".join(rest)
        errors = split_errors(err_tok) if err_tok else None
    else:
        errors = None
    return {"学号": sid, "姓名": name, "分数": score, "缺考": absent,
            "待定": no_score, "错题": errors, "原行": line.strip()}


def parse_roster(path: Path):
    """名单文件 -> (meta, students)。重名自动加 (2) 区分，和原版行为一致。"""
    meta, students, seen = {}, [], set()
    for raw in Path(path).read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith("//"):
            continue
        m = re.match(r"^[@＠]\s*([^:：\s]+)\s*[:：]?\s*(.*)$", line)
        if m:
            meta[m.group(1)] = m.group(2).strip()
            continue
        if re.match(r"^(姓名|名字|学号)\b", line):
            continue
        d = parse_line(line)
        if not d:
            continue
        name = d["姓名"]
        disp = name
        if name in seen:
            n = 2
            while f"{name}({n})" in seen:
                n += 1
            disp = f"{name}({n})"
        seen.update({name, disp})
        if not d["缺考"] and d["分数"] is None and not d["待定"]:
            continue
        students.append({
            "学号": d["学号"], "姓名": name, "显示名": disp,
            "分数": d["分数"], "缺考": d["缺考"], "待定": d["待定"],
            "错题": d["错题"], "原行": d["原行"],
        })
    return meta, students


def name_id_map(path: Path) -> dict[str, str]:
    """姓名 -> 学号（给填群发表时用学号匹配）。"""
    _, students = parse_roster(path)
    return {s["姓名"]: s["学号"] for s in students if s.get("学号")}
