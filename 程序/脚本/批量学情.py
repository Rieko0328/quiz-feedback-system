# -*- coding: utf-8 -*-
"""批量出学情报告 + 巩固单（一个班一次跑完）。

老师要的是「每个孩子一份报告 + 一份练习」，逐个命令敲 40 次不现实。
题目由 AI 预先写成 <题目目录>/<姓名>.json；没有的孩子退回抄写默写版。

用法：
    python 批量学情.py --班级 五年级A班
    python 批量学情.py --班级 五年级A班 --只报告
    python 批量学情.py --班级 五年级A班 --题目 题目/五年级A班_五年级第五讲
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from _ws import workspace_root, dirs_for  # noqa: E402


def _absent_names(roster: Path):
    """名单里标了「缺考」的人——他们的旧产物要清掉，不能跟着发出去。"""
    out = []
    if not roster.exists():
        return out
    for raw in roster.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith(("#", "@")):
            continue
        parts = s.split()
        if any(t in ("缺考", "请假", "免考") for t in parts):
            # 姓名是「缺考」前面那个词
            idx = next(i for i, t in enumerate(parts)
                       if t in ("缺考", "请假", "免考"))
            if idx >= 1:
                out.append(parts[idx - 1])
    return out


def _students_with_score(roster: Path):
    """名单里这次有分数的人（跳过缺考和没登分的）。"""
    out = []
    if not roster.exists():
        return out
    for raw in roster.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith(("#", "@")):
            continue
        parts = s.split()
        if len(parts) < 2:
            continue
        # 可能是 [学号] 姓名 分数 ...，从后往前找一个像分数的
        name, score = None, None
        for i in range(1, len(parts)):
            tok = parts[i]
            if tok in ("缺考", "请假", "免考"):
                break
            try:
                score = float(tok)
            except ValueError:
                _m = re.match(r"[-+]?\d+(?:\.\d+)?", tok)
                if not _m:
                    continue
                score = float(_m.group(0))
            name = parts[0] if i == 1 else parts[i - 1]
            break
        if name and score is not None:
            out.append((name, score))
    return out


def _work_one(job: dict):
    """处理一个学生：出报告 + 出巩固单。多进程时会把这个函数 pickled 给子进程，
    所以里面不能引用 main() 里的局部变量，所有依赖都从 job 里取。
    """
    from 学情报告 import load_from_workspace, fill_text, build as build_report
    from 错题巩固单 import compose_from_bank
    import json as _json

    name = job["name"]
    i = job["i"]
    root = job["root"]
    cls = job["班级"]
    lesson = job["课次"]
    rep_dir = job["rep_dir"]
    ex_dir = job["ex_dir"]
    qdir = job["qdir"]
    bank = job["bank"]
    n_questions = job["n_questions"]
    only_report = job["只报告"]
    only_practice = job["只巩固单"]
    build_student = job["build_student"]
    build_blanks = job["build_blanks"]

    d = load_from_workspace(root, cls, name, lesson)
    fill_text(d, root, cls)
    tag = []
    if not only_practice:
        try:
            build_report(d, rep_dir)
            tag.append("报告")
        except (Exception, SystemExit) as e:
            tag.append(f"报告✗({type(e).__name__})")
    if not only_report:
        qs = {"背诵": [], "题目": []}
        qp = qdir / f"{name}.json"
        if qp.exists():
            try:
                qs = _json.loads(qp.read_text(encoding="utf-8-sig"))
                qs.setdefault("背诵", [])
                qs.setdefault("题目", [])
            except Exception:
                qs = {"背诵": [], "题目": []}
        rate = d["分数"] / d["满分"] if d["满分"] else 0
        n = n_questions(rate, d.get("错题数") or len(d.get("失分") or []))
        src = "专属"
        if not qs["题目"] and bank:
            qs = compose_from_bank(d, bank, n, seed=i)
            src = "题库"
        if len(qs["题目"]) > n:
            qs["题目"] = qs["题目"][:n]
        out = ex_dir / f"{name}_巩固练习.docx"
        try:
            if qs["题目"]:
                build_student(d, qs, out)
                tag.append(f"巩固单{len(qs['题目'])}题·{src}")
            else:
                pts = [x[1] for x in (d.get("失分") or [])
                       if x[1] and x[1] != "—"] or \
                      [m["名"] for m in d.get("模块", [])]
                build_blanks(d, list(dict.fromkeys(pts))[:4], out)
                tag.append("巩固单-抄写默写")
        except Exception as e:
            tag.append(f"巩固单✗({type(e).__name__})")
    return i, name, "　".join(tag)


def main():
    ap = argparse.ArgumentParser(description="批量出诊断报告和巩固单")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--题目", default=None,
                    help="题目目录，里面放 <姓名>.json；默认 题目/<班级>_<课次>/")
    ap.add_argument("--并行", type=int, default=0,
                    help="并行进程数（0=串行）。40 人以上的班开 4 能快 2 倍多，"
                         "但进度提示会晚一点出来")
    ap.add_argument("--只报告", action="store_true")
    ap.add_argument("--只巩固单", action="store_true")
    args = ap.parse_args()
    if args.只报告 and args.只巩固单:
        print("[错] --只报告 和 --只巩固单 不能同时给——那样两个都不做，"
              "什么也不会产出。只要其中一样就去掉另一个；两样都要就都不加。")
        return 2

    root = workspace_root(args.工作区)
    roster = root / "班级名单" / f"{args.班级}.txt"
    students = _students_with_score(roster)
    if not students:
        print(f"[错] {roster.name} 里没有已登分的学生")
        return 2

    from 学情报告 import load_from_workspace, fill_text, build as build_report
    from 错题巩固单 import (n_questions, build_student, build_blanks,
                            load_bank, compose_from_bank)

    # 课次从名单头里读；命令行给了就用命令行的
    lesson = args.课次
    if not lesson:
        for raw in roster.read_text(encoding="utf-8-sig").splitlines():
            if raw.startswith("@课次"):
                lesson = raw.split(":", 1)[-1].strip()
                break
    if not lesson:
        print("[错] 名单里没有 @课次，也没传 --课次")
        return 2

    d0 = dirs_for(root, args.班级, lesson)
    work = d0["work"]
    rep_dir = work / "诊断报告"
    ex_dir = work / "巩固单"

    qdir = Path(args.题目).expanduser() if args.题目 else \
        root / "题目" / f"{args.班级}_{lesson}"

    # 缺考的人的旧巩固单要清掉，否则重跑时还躺在目录里，老师会发错
    absent = _absent_names(roster)
    dropped = []
    for nm in absent:
        stale = ex_dir / f"{nm}_巩固练习.docx"
        if stale.exists():
            stale.unlink()
            dropped.append(nm)
    if dropped:
        print(f"清掉缺考的旧巩固单：{'、'.join(dropped)}")

    bank = load_bank(qdir / "题库.json")
    if bank and not args.只报告:
        print(f"题库：{len(bank)} 个考点　"
              f"{sum(len(v.get('题') or []) for v in bank.values())} 道题")

    total = len(students)
    n_rep = n_ex = n_q = 0
    print(f"{args.班级} · {lesson}　{total} 人")

    # ---- 并行分支（可选）----
    if args.并行 and args.并行 > 1 and total > 3:
        from concurrent.futures import ProcessPoolExecutor
        jobs = []
        for i, (name, _score) in enumerate(students, 1):
            jobs.append({
                "name": name, "i": i, "root": root, "班级": args.班级,
                "课次": lesson, "rep_dir": rep_dir, "ex_dir": ex_dir,
                "qdir": qdir, "bank": bank, "n_questions": n_questions,
                "只报告": args.只报告, "只巩固单": args.只巩固单,
                "build_student": build_student, "build_blanks": build_blanks,
            })
        done = 0
        with ProcessPoolExecutor(max_workers=args.并行) as pool:
            for i, name, tags in pool.map(_work_one, jobs):
                done += 1
                print(f"  [{done}/{total}] {name}　{tags}")
                if "报告" in tags:
                    n_rep += 1
                if "巩固单" in tags:
                    n_ex += 1
                    if "题库" in tags or "专属" in tags:
                        n_q += 1
        print(f"\n报告 {n_rep}/{total}　巩固单 {n_ex}/{total}（其中带题 {n_q} 人）")
        return 0

    for i, (name, _score) in enumerate(students, 1):
        d = load_from_workspace(root, args.班级, name, lesson)
        fill_text(d, root, args.班级)

        tag = []
        if not args.只巩固单:
            try:
                build_report(d, rep_dir)
                n_rep += 1
                tag.append("报告")
            except (Exception, SystemExit) as e:      # noqa: BLE001
                tag.append(f"报告✗({type(e).__name__})")

        if not args.只报告:
            qs = {"背诵": [], "题目": []}
            qp = qdir / f"{name}.json"
            if qp.exists():
                try:
                    qs = json.loads(qp.read_text(encoding="utf-8-sig"))
                    qs.setdefault("背诵", [])
                    qs.setdefault("题目", [])
                except Exception:                      # noqa: BLE001
                    qs = {"背诵": [], "题目": []}
            rate = d["分数"] / d["满分"] if d["满分"] else 0
            n = n_questions(rate, d.get("错题数") or len(d.get("失分") or []))
            src = "专属"
            if not qs["题目"] and bank:
                # 没写专属题目 → 按他的失分考点从题库抽，seed 用序号错开
                qs = compose_from_bank(d, bank, n, seed=i)
                src = "题库"
            if len(qs["题目"]) > n:
                qs["题目"] = qs["题目"][:n]
            out = ex_dir / f"{name}_巩固练习.docx"
            try:
                if qs["题目"]:
                    build_student(d, qs, out)
                    n_q += 1
                    tag.append(f"巩固单{len(qs['题目'])}题·{src}")
                else:
                    pts = [x[1] for x in (d.get("失分") or [])
                           if x[1] and x[1] != "—"] or \
                          [m["名"] for m in d.get("模块", [])]
                    build_blanks(d, list(dict.fromkeys(pts))[:4], out)
                    tag.append("巩固单-抄写默写")
                n_ex += 1
            except Exception as e:                     # noqa: BLE001
                tag.append(f"巩固单✗({type(e).__name__})")

        print(f"  [{i}/{total}] {name}　{'　'.join(tag)}")

    print(f"\n报告 {n_rep}/{total}　巩固单 {n_ex}/{total}"
          f"（其中带题 {n_q} 人）")
    if n_q < n_ex:
        if not bank:
            print(f"没写题目的孩子出了抄写默写版。把本讲课题库放这里就能自动抽题：{qdir / '题库.json'}")
        else:
            print(f"有 {n_ex - n_q} 人出了抄写默写版：题库没覆盖到他的失分考点，补几个考点就行")

    # 有失败就别报「全部完成」——一个都没成功更要说清楚
    bad = (total - n_rep if not args.只巩固单 else 0)
    bad2 = (total - n_ex if not args.只报告 else 0)
    if bad or bad2:
        print(f"\n[注意] 有失败：报告 {bad} 人没出，巩固单 {bad2} 人没出。"
              f"看上面标 ✗ 的那几行，把报错截图给技术同学。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
