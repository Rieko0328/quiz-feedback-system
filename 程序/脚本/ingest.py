# -*- coding: utf-8 -*-
r"""ingest.py —— 把「AI 看图读到的结果」落成程序认的输入。

这是技能的核心接口：AI 读完答题卡照片后产出一个 readings.json，
本脚本负责校验（扣分之和必须 = 满分 − 分数）、合并进名单、写出分项扣分。

原来 识别照片.py 里调视觉 API 的部分全部删除 —— 现在由 AI 直接看图，
这个脚本只做确定性的事，不联网。

用法：
    python ingest.py --班级 五1班 --批 readings.json
    python ingest.py --班级 五1班 --批 readings.json --dry-run
    python ingest.py --班级 五1班 --批 readings.json --核对表 核对表.xlsx

readings.json 格式：
{
  "课次": "第5讲L95-96",
  "满分": 100,
  "重点线": 60,
  "大题名称": ["听力", "词汇", "语法"],
  "大题题数": [10, 10, 10],
  "students": [
    {"姓名": "张三", "分数": 85, "错题号": "3,7,10",
     "扣分": {"听力": 1, "词汇": 2, "语法": 2},
     "照片": "IMG_0001.jpg", "置信度": 0.97}
  ]
}
只需要「姓名 + 分数」就能跑；错题号、扣分都是可选，有则话术更准。
"""
from __future__ import annotations

import argparse
import json
import re
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

from _ws import workspace_root, ensure_workspace, photo_root, safe_name  # noqa: E402
from roster import parse_line, parse_roster, norm_key  # noqa: E402

META_ORDER = ["班级", "课次", "满分", "重点线", "大题名称", "大题题数",
              "大题分值", "日期", "知识点库"]
CN_SEC = "一二三四五六七八九十"


# ---------------------------------------------------------------- 工具

def _num(x):
    if x is None or x == "":
        return None
    if isinstance(x, (int, float)):
        return float(x)
    m = re.search(r"-?\d+(?:\.\d+)?", str(x))
    return float(m.group()) if m else None


def _split_list(v):
    if v is None:
        return []
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [x.strip() for x in re.split(r"[,，、;；/\s]+", str(v)) if x.strip()]


def split_errors(tok):
    """'3,7,10' / '错3、7' / '全对' -> [3,7,10] / []"""
    t = str(tok or "").strip()
    if not t:
        return None
    if re.search(r"全对|没有错|无错|满分|^无$|^-$|^—$", t, re.I):
        return []
    t = re.sub(r"^错[:：]?", "", t).strip()
    return sorted({int(x) for x in re.findall(r"\d+", t)})


def norm_name(x):
    x = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", x or "")
    return re.sub(r"[\s_\-（）()0-9.]", "", x)


def best_match(raw, names):
    """把读到的名字对到名单上（手写体常认错一个字）。

    规则：长度相同 + 至少一半字相同，且首字或末字相同才认。
    返回 (名单里的写法, 相似度)；认不出返回 ("", 0)。
    """
    r = norm_name(raw)
    if not r:
        return "", 0.0
    best, best_score = "", 0.0
    for n in names:
        nn = norm_name(n)
        if len(nn) != len(r):
            continue
        same = sum(1 for a, b in zip(r, nn) if a == b)
        score = same / len(r)
        if r[0] == nn[0]:
            score += 0.15
        if r[-1] == nn[-1]:
            score += 0.10
        if score > best_score:
            best, best_score = n, score
    if best_score >= 0.55:
        return best, round(best_score, 2)
    return "", 0.0


def read_roster(path: Path):
    """读回已有名单，保留 meta 与学生顺序（合并用）。"""
    meta, students = {}, []
    if not path.exists():
        return meta, students
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"^[@＠]\s*([^:：\s]+)\s*[:：]?\s*(.*)$", line)
        if m:
            meta[m.group(1)] = m.group(2).strip()
            continue
        students.append(line)
    return meta, students


# ---------------------------------------------------------------- 校验
def check_student(s, cols, full):
    """返回 (状态, 提示, 未定位扣分)"""
    problems = []
    name = (s.get("姓名") or "").strip()
    if not name:
        return "name_missing", "没有姓名", None
    if s.get("缺考") or re.search(r"缺考|请假", str(s.get("状态") or "")):
        return "absent", "缺考", None

    score = _num(s.get("分数"))
    if score is None:
        return "score_missing", "没有总分", None

    ded = s.get("扣分") or {}
    got, unlocated = 0.0, None
    if ded and full:
        for c in cols:
            v = _num(ded.get(c))
            if v:
                got += v
        need = round(full - score, 2)
        if abs(got - need) > 0.01:
            unlocated = round(need - got, 2)
            problems.append(f"分项扣分合计 {got:g} ≠ 应扣 {need:g}（差 {unlocated:g}）")
    _conf = _num(s.get("置信度"))
    if _conf is not None and _conf < 0.9:
        problems.append(f"识别置信度偏低 {s['置信度']}")
    return ("review" if problems else "ok"), "；".join(problems), unlocated


# ---------------------------------------------------------------- 主流程
def archive_photos(root: Path, cls: str, lesson: str, rows):
    """把照片按姓名复制一份到 <课次>/已改名/姓名.jpg。

    只复制、不移动原件 —— 原图永远留着，改名失败也不丢东西。
    下游的「群发打包」从「已改名」取照片，所以这步不能少。
    """
    import shutil
    if not lesson:
        return {"copied": 0, "skipped": 0, "missing": []}
    base = photo_root(root) / cls / lesson
    if not base.is_dir():
        return {"copied": 0, "skipped": 0, "missing": [], "note": "没有课次照片目录"}
    dest = base / "已改名"
    dest.mkdir(exist_ok=True)
    copied, skipped, missing = 0, 0, []
    for r in rows:
        ph, nm = (r.get("照片") or "").strip(), (r.get("姓名") or "").strip()
        if not ph or not nm:
            continue
        src = base / ph
        if not src.exists():
            src = base / "原图" / ph
        if not src.exists():
            missing.append(ph)
            continue
        tgt = dest / f"{nm}{src.suffix.lower()}"
        if tgt.exists():
            skipped += 1
            continue
        shutil.copy2(src, tgt)
        copied += 1
    return {"copied": copied, "skipped": skipped, "missing": missing}


def edit_roster(root: Path, args):
    """不带 --批 时只改名单：删人（退费/转班走了）或标缺考。"""
    cls = (args.班级 or "").strip()
    if not cls:
        print(json.dumps({"ok": False, "error": {
            "code": "no_class", "message": "没给班级名"}}, ensure_ascii=False))
        return 2
    path = root / "班级名单" / f"{cls}.txt"
    if not path.exists():
        print(json.dumps({"ok": False, "error": {
            "code": "no_roster", "message": f"没有这个班的名单：{path}"}},
            ensure_ascii=False))
        return 2
    meta, lines = read_roster(path)
    rm = {norm_key(x) for x in _split_list(args.移除)}
    ab = {norm_key(x) for x in _split_list(args.缺考)}
    out, removed, absent, unknown = [], [], [], []
    hit_rm, hit_ab = set(), set()
    for ln in lines:
        d = parse_line(ln)
        if not d:
            continue
        k = norm_key(d["姓名"])
        if k in rm:
            removed.append(d["姓名"])
            hit_rm.add(k)
            continue
        if k in ab:
            absent.append(d["姓名"])
            hit_ab.add(k)
            out.append((f"{d['学号']} " if d.get("学号") else "") + f"{d['姓名']} 缺考")
        else:
            out.append(ln)
    for k, who in ((rm, args.移除), (ab, args.缺考)):
        for x in _split_list(who):
            if norm_key(x) not in hit_rm and norm_key(x) not in hit_ab:
                unknown.append(x)
    text = "\n".join([f"@{k}: {v}" for k, v in meta.items()] + [""] + out) + "\n"
    if not args.dry_run:
        path.write_text(text, encoding="utf-8")
    print(json.dumps({"ok": True, "class": cls, "removed": removed,
                      "absent": absent, "not_found": unknown,
                      "left": len(out), "dry_run": bool(args.dry_run)},
                     ensure_ascii=False, indent=2))
    return 0


def main():
    ap = argparse.ArgumentParser(description="把 AI 读到的成绩落进名单")
    ap.add_argument("--批", default=None, help="AI 产出的 readings.json")
    ap.add_argument("--班级", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--核对表", default=None, help="顺便导出一份 xlsx 核对表")
    ap.add_argument("--dry-run", action="store_true", help="只校验，不写文件")
    ap.add_argument("--不归档", action="store_true",
                    help="不把照片复制到 已改名/（默认会复制）")
    ap.add_argument("--允许新生", action="store_true",
                    help="名单里没有的学生也写进名单（默认只报出来等确认）")
    ap.add_argument("--最后一批", action="store_true",
                    help="这是本次课测的最后一批 → 顺便报出「名单里有、但这次没交卷」的人")
    ap.add_argument("--移除", default=None,
                    help="从名单里删掉这些人（退费/转班走了），逗号分隔")
    ap.add_argument("--缺考", default=None,
                    help="把这些人标成缺考（分数清空），逗号分隔")
    args = ap.parse_args()

    root = ensure_workspace(workspace_root(args.工作区))

    if not args.批:
        if args.移除 or args.缺考:
            return edit_roster(root, args)
        print(json.dumps({"ok": False, "error": {
            "code": "no_batch", "message": "要么给 --批，要么用 --移除/--缺考 改名单"}},
            ensure_ascii=False))
        return 2

    batch_path = Path(args.批).expanduser()
    if not batch_path.is_absolute():          # 相对路径按工作区算，不按 cwd
        batch_path = root / batch_path
    batch_path = batch_path.resolve()
    if not batch_path.exists():
        print(json.dumps({"ok": False, "error": {"code": "no_batch",
              "message": f"找不到 {batch_path}"}}, ensure_ascii=False))
        return 2
    data = json.loads(batch_path.read_text(encoding="utf-8-sig"))

    cls = (args.班级 or data.get("班级") or "").strip()
    if not cls:
        print(json.dumps({"ok": False, "error": {
            "code": "no_class", "message": "没给班级名"}}, ensure_ascii=False))
        return 2
    lesson = (data.get("课次") or "").strip()
    full = _num(data.get("满分"))
    cols = _split_list(data.get("大题名称"))
    counts = [_num(x) or 0 for x in _split_list(data.get("大题题数"))]

    roster_path = root / "班级名单" / f"{cls}.txt"
    old_meta, old_lines = read_roster(roster_path)
    # 换课次了：上一讲的卷子结构和这一讲通常完全不同（题型、题数、分值都变），
    # 一律不继承，必须由本讲重新给。宁可报错问老师，也不能拿上一讲的算这一讲。
    _prev_lesson = (old_meta.get("课次") or "").strip()
    _new_lesson = (lesson or str(data.get("课次") or "")).strip()
    switched = bool(_prev_lesson and _new_lesson and _prev_lesson != _new_lesson)
    if switched:
        _keep = {"班级", "称呼", "学校", "科目", "家长风格", "共性优点"}
        old_meta = {k: v for k, v in old_meta.items() if k in _keep}
    full = full or _num(old_meta.get("满分"))
    lesson = lesson or old_meta.get("课次", "")
    if not cols:
        cols = _split_list(old_meta.get("大题名称"))
    # 课次和满分都不能靠猜：课次错了产物会串课，满分错了所有得分率都错
    if not lesson:
        print(json.dumps({"ok": False, "error": {
            "code": "no_lesson",
            "message": "不知道这是第几讲。readings.json 里加 \"课次\": \"五年级第五讲\""}},
            ensure_ascii=False))
        return 2
    if not full:
        print(json.dumps({"ok": False, "error": {
            "code": "no_full_score",
            "message": "不知道满分。readings.json 里加 \"满分\": 40（卷头就印着）"}},
            ensure_ascii=False))
        return 2

    students = data.get("students") or []
    if not students:
        print(json.dumps({"ok": False, "error": {
            "code": "no_students", "message": "批次里没有 students"}},
            ensure_ascii=False))
        return 2

    # ---- 第一次核对：把读到的名字对到老师给的名单上 ----
    old_recs = [d for d in (parse_line(ln) for ln in old_lines) if d]
    roster_names = [d["姓名"] for d in old_recs]
    sid_map = {norm_key(d["姓名"]): d["学号"] for d in old_recs if d.get("学号")}
    # 名字纠正分三步：原文精确 → 归一后**唯一**才敢改 → 模糊匹配。
    # 「归一后唯一」这个前提很重要：norm_name 会抹掉数字和括号，
    # 名单里「学生01…学生40」全归一成「学生」，如果直接查表，
    # 40 个人会被统一改成名单最后一个名字，分数全串。
    _norm_cnt: dict[str, int] = {}
    for _n0 in roster_names:
        _k0 = norm_name(_n0)
        _norm_cnt[_k0] = _norm_cnt.get(_k0, 0) + 1
    by_norm = {norm_name(n): n for n in roster_names
               if _norm_cnt.get(norm_name(n), 0) == 1}
    _exact_names = set(roster_names)
    corrected, not_in_roster = [], []
    for s in students:
        raw = (s.get("姓名") or "").strip()
        if not raw or not roster_names:
            continue
        if raw in _exact_names:            # 原文就对得上，一个字都不用改
            continue
        if by_norm.get(norm_name(raw)):
            s["姓名"] = by_norm[norm_name(raw)]
            continue
        hit, score = best_match(raw, roster_names)
        if hit:
            s["姓名"] = hit
            corrected.append({"读到的": raw, "按名单纠正为": hit, "相似度": score})
        else:
            s["_不在名单"] = True
            not_in_roster.append(raw)

    # ---- 逐条校验 ----
    rows, review, ded_out = [], [], {}
    # 同名的人分项扣分没法区分（扣分表按姓名做 key），只给唯一名字写；
    # 同名的报告会降级成「按总分折算」，并在核对表里提示老师逐行核对。
    _name_dup: dict[str, int] = {}
    for _nm0 in roster_names:
        _k0 = norm_name(_nm0)
        _name_dup[_k0] = _name_dup.get(_k0, 0) + 1
    if any(v > 1 for v in _name_dup.values()):
        ded_out.setdefault("_说明", []).append(
            "有同名同学，分项扣分无法按人区分，这些人的报告会按总分折算")
    for s in students:
        st, msg, unloc = check_student(s, cols, full)
        name = (s.get("姓名") or "").strip()
        score = _num(s.get("分数"))
        errs = split_errors(s.get("错题号"))
        if s.get("_不在名单") and not args.允许新生:
            st = "review"
            msg = ("；".join(x for x in [msg, "名单里没有这个人：是来试听的、转班进来的，"
                                        "还是新报名的？确认后加 --允许新生 再跑"] if x))
        sid = (s.get("学号") or "").strip() or sid_map.get(norm_key(name), "")
        rows.append({"学号": sid, "姓名": name, "分数": score, "错题": errs,
                     "状态": st, "提示": msg, "照片": s.get("照片", ""),
                     "扣分": s.get("扣分") or {},
                     "_不在名单": bool(s.get("_不在名单"))})
        if st in ("review", "name_missing", "score_missing"):
            review.append({"姓名": name, "状态": st, "提示": msg,
                           "照片": s.get("照片", "")})
        if (name and st not in ("name_missing", "absent") and score is not None
                and _name_dup.get(norm_name(name), 0) <= 1):
            d = {}
            for i, c in enumerate(cols):
                # 超过十个大题用阿拉伯数字，和 课测反馈/学情报告 的 KEYS 保持一致。
                # 原来这里用大题名称当 key，消费端用 "11"，两边对不上，
                # 第 11 个之后大题的扣分会被静默丢掉。
                key = CN_SEC[i] if i < len(CN_SEC) else str(i + 1)
                v = _num((s.get("扣分") or {}).get(c))
                d[key] = v or 0.0
            if unloc is not None:
                d["未定位"] = unloc
            raw_err = s.get("错题号")
            if raw_err is None and errs:
                raw_err = ",".join(str(x) for x in errs)
            if raw_err and str(raw_err).strip():
                ded_out.setdefault("_错题", {})[name] = str(raw_err).strip()
            ded_out[name] = d

    # ---- 合并进名单 ----
    # 重名（王芳 / 王芳(2) / 王芳(3)）必须**按行一一对上**，不能按名字归并。
    # norm_name 会把括号和数字抹掉，三个人撞成同一个 key；用 dict 归并的后果是
    # 三行全写成同一个人的分数和学号（实测：34/28/20 三个人只剩 20 重复三遍，
    # 还全部标了 ok，谁看都发现不了）。改成按名单顺序排队取。
    from collections import defaultdict
    pend: dict[str, list] = defaultdict(list)
    for r in rows:
        if r["姓名"]:
            pend[norm_name(r["姓名"])].append(r)
    # 真正的重名只看**原文完全相同**。
    # 不能用 norm_name 比：它会把数字和括号抹掉，「学生01/学生02…」会全撞成
    # 「学生」，40 人的班被报成「40 个同名」；「王芳/王芳(2)」也会被误报
    # （这两个本来就是同一批人被建班时加了序号，不需要再核）。
    _exact: dict[str, int] = {}
    for r in rows:
        if r["姓名"]:
            _exact[r["姓名"]] = _exact.get(r["姓名"], 0) + 1
    dup_names = {k: n for k, n in _exact.items() if n > 1}

    merged, seen = [], set()
    for line in old_lines:
        _d = parse_line(line)
        nm = _d["姓名"] if _d else ""
        key = norm_name(nm)
        bucket = pend.get(key) or []
        r = bucket.pop(0) if bucket else None
        if r:
            # 名单里已经有学号和写法，就用名单的：
            # 读图那边可能把「王芳(2)」认成「王芳」，用回去会把三个人又搞混
            if _d.get("学号"):
                r["学号"] = _d["学号"]
            r["姓名"] = nm or r["姓名"]
            seen.add(key)
            merged.append(format_line(r))
        elif switched and _d and _d.get("学号"):
            # 换课次了还没交卷：清成「学号 姓名」，分数留给本讲
            merged.append(f"{_d['学号']} {_d['姓名']}")
        else:
            merged.append(line)

    added = []
    for key, bucket in pend.items():
        for r in bucket:                  # 名单里没对应行的（新生）
            if r.get("_不在名单") and not args.允许新生:
                continue                  # 名单外的先不入库，等老师确认
            merged.append(format_line(r))
            added.append(r["姓名"])

    for k, cnt in dup_names.items():
        review.append({"姓名": k, "状态": "重名",
                       "提示": f"名单里有 {cnt} 个同名，分数按名单先后顺序对应，"
                               f"请逐行核对别搞混",
                       "照片": ""})

    meta = dict(old_meta)
    meta["班级"] = cls
    if lesson:
        meta["课次"] = lesson
    if full:
        meta["满分"] = f"{full:g}"
    if data.get("重点线") is not None:
        meta["重点线"] = f"{_num(data['重点线']):g}"
    if cols:
        meta["大题名称"] = ",".join(cols)
    if counts and any(counts):
        meta["大题题数"] = ",".join(f"{int(c)}" for c in counts)
    # 各大题分值：卷面标题下面印着「每题1分，共20分」，读图时抄下来
    result_warn = ""
    marks = [_num(x) or 0 for x in _split_list(data.get("大题分值"))]
    if marks and any(marks) and len(marks) == len(cols or marks):
        meta["大题分值"] = ",".join(f"{m:g}" for m in marks)
        if full and abs(sum(marks) - full) > 0.01:
            result_warn = (f"各大题分值加起来是 {sum(marks):g}，"
                           f"跟满分 {full:g} 对不上，检查一下")

    out_txt = "\n".join(
        [f"@{k}: {meta[k]}" for k in META_ORDER if meta.get(k)] +
        [f"@{k}: {v}" for k, v in meta.items() if k not in META_ORDER] +
        [""] + merged) + "\n"

    result = {
        "ok": not review,
        "class": cls, "lesson": lesson,
        "count": len(rows),
        "ok_count": len(rows) - len(review),
        "added": added,
        "corrected_by_roster": corrected,
        "not_in_roster": not_in_roster,
        "review": review,
        "roster": str(roster_path),
        "dry_run": bool(args.dry_run),
    }
    if result_warn:
        result["warning"] = result_warn

    if args.dry_run:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0

    roster_path.parent.mkdir(parents=True, exist_ok=True)
    if roster_path.exists():
        bak = roster_path.with_suffix(".txt.bak")
        if not bak.exists():
            bak.write_text(roster_path.read_text(encoding="utf-8-sig"),
                           encoding="utf-8")
    roster_path.write_text(out_txt, encoding="utf-8")

    # ---- 照片按姓名归档（下游「分发/群发」靠它取图）----
    if not args.不归档:
        result["photos"] = archive_photos(root, cls, lesson, rows)

    # ---- 最后一批：报出「名单里有、这次没交卷」的人（缺考/退费/转班）----
    if args.最后一批:
        _, allst = parse_roster(roster_path)
        miss = [s["姓名"] for s in allst
                if s["分数"] is None and not s["缺考"]]
        result["not_in_this_exam"] = miss

    # ---- 分项扣分 json（话术 / 诊断报告 / 巩固单都靠它定位考点）----
    if ded_out and (cols or any(1 for _ in ded_out)):
        outdir = root / "生成结果" / (f"{safe_name(cls)}_{safe_name(lesson)}"
                                       if lesson else safe_name(cls))
        outdir.mkdir(parents=True, exist_ok=True)
        ded_out["_说明"] = f"{cls} 各大题扣分（ingest 自动生成）"
        if full:
            ded_out["_校验"] = f"各大题扣分之和必须等于 {full:g}−分数"
        (outdir / "大题扣分.json").write_text(
            json.dumps(ded_out, ensure_ascii=False, indent=2), encoding="utf-8")
        result["deductions"] = str(outdir / "大题扣分.json")

    # ---- 可选：导出核对表给老师改 ----
    if args.核对表:
        try:
            xlsx = Path(args.核对表).expanduser()
            if not xlsx.is_absolute():        # 相对路径按工作区算
                xlsx = root / xlsx
            xlsx = xlsx.resolve()
            write_review_xlsx(xlsx, cls, lesson, full, cols, rows)
            result["review_xlsx"] = str(xlsx)
        except Exception as e:
            result.setdefault("warnings", []).append(f"核对表导出失败：{e}")

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if not review else 1


def format_line(r):
    parts = []
    if r.get("学号"):
        parts.append(r["学号"])
    parts += [r["姓名"], f"{r['分数']:g}" if r["分数"] is not None else "缺考"]
    if r["错题"]:
        parts.append(",".join(str(x) for x in r["错题"]))
    elif r["错题"] == []:
        parts.append("全对")
    # 注意：名单行里绝不能带 # 注释或说明文字，
    # 否则会被 parse_roster 当成「错题号」解析出错误的题号。
    return " ".join(parts)


def write_review_xlsx(path, cls, lesson, full, cols, rows):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "核对表"
    head = ["#", "姓名", "分数"] + cols + ["错题号", "状态", "提示", "照片"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1B5EA8")
        c.alignment = Alignment(horizontal="center")
    red = Font(color="C0392B", bold=True)
    for i, r in enumerate(rows, 1):
        sc = f"{r['分数']:g}" if r["分数"] is not None else ""
        errs = ",".join(str(x) for x in (r["错题"] or [])) if r["错题"] is not None else ""
        _ded = r.get("扣分") or {}
        ws.append([i, r["姓名"], sc]
                  + [(_ded.get(c) if _ded.get(c) is not None else "") for c in cols]
                  + [errs, r["状态"], r["提示"], r["照片"]])
        if r["状态"] != "ok":
            for c in ws[ws.max_row]:
                c.font = red
    widths = [5, 12, 8] + [9] * len(cols) + [12, 10, 30, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


if __name__ == "__main__":
    sys.exit(main() or 0)
