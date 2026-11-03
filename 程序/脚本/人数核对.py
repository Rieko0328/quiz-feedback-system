# -*- coding: utf-8 -*-
r"""人数核对.py —— 名单和这次收到的卷子对不上时，一个孩子一个孩子地处置。

为什么单独做这一步：
  名单 27 人，收到 24 份卷子。少的 3 个人可能是三种完全不同的情况 ——
  一个缺考（下次还来）、一个退费走了、一个转到别的班了。
  打包问「这 3 个都怎么了」一定是错的，而错一次就是成绩表和汇总表全错。

用法：
    python 人数核对.py --班级 周五6-8 --课次 NCE1-L9-14月考          # 只看差在哪
    python 人数核对.py --班级 周五6-8 --课次 X --缺考 张三,李四
    python 人数核对.py --班级 周五6-8 --课次 X --移除 王五
    python 人数核对.py --班级 周五6-8 --课次 X --新生 赵六,钱七
    python 人数核对.py --班级 周五6-8 --课次 X --试听 孙八

处置规则（照技能版 SKILL.md 第 5.5 节）：
  缺考    —— 名单不动，只在成绩表里标「缺考」；不出话术/报告/巩固单/群发
  移除    —— 从名单删掉（挪到 班级名单/_已退班.txt 留档），历史和汇总不再有他
  试听    —— 不写进名单，这次也不出任何产物
  新生    —— 正式加进名单，以后每次都统计
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _ws import workspace_root, safe_name  # noqa: E402


def read_roster(p: Path):
    meta, lines = {}, []
    for raw in p.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if s.startswith("@") and ":" in s:
            k, v = s[1:].split(":", 1)
            meta[k.strip()] = v.strip()
        lines.append(raw)
    return meta, lines


def students_of(lines):
    out = []
    for ln in lines:
        s = ln.strip()
        if not s or s.startswith(("@", "#")):
            continue
        out.append(s.split()[0])
    return out


def scored_of(lines):
    """名单里已经有分数的（这次交卷了）"""
    out = []
    for ln in lines:
        s = ln.strip()
        if not s or s.startswith(("@", "#")):
            continue
        parts = s.split()
        if len(parts) >= 2:
            try:
                float(parts[1])
                out.append(parts[0])
            except ValueError:
                pass
    return out


def main():
    ap = argparse.ArgumentParser(description="人数核对与处置")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--缺考", default=None, help="逗号分隔；名单不动，只在成绩表标缺考")
    ap.add_argument("--移除", default=None, help="逗号分隔；退费或转班走了，从名单删掉")
    ap.add_argument("--试听", default=None, help="逗号分隔；只记这一次，不进名单")
    ap.add_argument("--新生", default=None, help="逗号分隔；转班进来或新报名，正式入班")
    args = ap.parse_args()

    root = workspace_root(args.工作区)
    cls = safe_name(args.班级)
    rp = root / "班级名单" / f"{cls}.txt"
    if not rp.exists():
        raise SystemExit(f"[错] 名单不存在：{rp}")

    meta, lines = read_roster(rp)
    names = students_of(lines)
    scored = scored_of(lines)
    lesson = args.课次 or meta.get("课次", "")

    # ---------- 只看差在哪 ----------
    if not any([args.缺考, args.移除, args.试听, args.新生]):
        # 核对表里识别到的名字
        got = []
        tbl = root / "生成结果" / f"{cls}_{lesson}" / "识别核对表.xlsx"
        if tbl.exists():
            try:
                import openpyxl
                ws = openpyxl.load_workbook(tbl, data_only=True).active
                head = [c.value for c in ws[1]]
                i_nm = head.index("识别姓名") if "识别姓名" in head else 1
                i_sc = head.index("分数") if "分数" in head else 3
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if not row or not row[i_nm]:
                        continue
                    if row[i_sc] in (None, ""):
                        continue
                    got.append(str(row[i_nm]).strip())
            except Exception as e:
                print(f"  [注意] 核对表读不出来：{type(e).__name__}")
        else:
            got = scored

        miss = [n for n in names if n not in got]
        extra = [n for n in got if n not in names]
        print(f"班级 {cls}　课次 {lesson or '（未填）'}")
        print(f"  名单 {len(names)} 人　这次交了 {len(got)} 份")
        if not got:
            print("  [提示] 还没识别过（核对表里没有分数），先跑识别再来核对人数")
            return
        if not miss and not extra:
            print("  ✔ 人数对得上，不用处置")
            return
        if miss:
            print(f"\n  少 {len(miss)} 个（名单里有、这次没卷子）：")
            for n in miss:
                print(f"    {n}　→ 缺考？退费/转班走了？")
        if extra:
            print(f"\n  多 {len(extra)} 个（有卷子、名单里没有）：")
            for n in extra:
                print(f"    {n}　→ 来试听？转班进来？新报名？")
        print("\n  一个孩子一个孩子地处置，别打包：")
        print(f'    python 人数核对.py --班级 "{args.班级}" --缺考 {miss[0] if miss else "张三"}')
        print(f'    python 人数核对.py --班级 "{args.班级}" --移除 {miss[-1] if miss else "李四"}')
        print(f'    python 人数核对.py --班级 "{args.班级}" --新生 {extra[0] if extra else "赵六"}')
        return

    # ---------- 处置 ----------
    def split(s):
        return [x.strip() for x in re.split(r"[,，、]", s or "") if x.strip()]

    changed = []

    if args.缺考:
        # 名单不动；把分数那栏写成「缺考」，下游脚本据此跳过专属产物
        targets = split(args.缺考)
        new = []
        for ln in lines:
            s = ln.strip()
            if s and not s.startswith(("@", "#")):
                nm = s.split()[0]
                if nm in targets:
                    new.append(f"{nm} 缺考")
                    changed.append(f"缺考 {nm}")
                    continue
            new.append(ln)
        lines = new

    if args.移除:
        targets = split(args.移除)
        arch = root / "班级名单" / "_已退班.txt"
        keep = []
        with arch.open("a", encoding="utf-8") as f:
            for ln in lines:
                s = ln.strip()
                if s and not s.startswith(("@", "#")):
                    nm = s.split()[0]
                    if nm in targets:
                        f.write(f"{cls}\t{nm}\t{lesson}\n")
                        changed.append(f"退班 {nm}（已留档）")
                        continue
                keep.append(ln)
        lines = keep

    if args.试听:
        for nm in split(args.试听):
            changed.append(f"试听 {nm}（不进名单，不出产物）")

    if args.新生:
        targets = split(args.新生)
        # 加在最后一个学生后面
        last = max((i for i, ln in enumerate(lines)
                    if ln.strip() and not ln.strip().startswith(("@", "#"))),
                   default=len(lines) - 1)
        for k, nm in enumerate(targets):
            if nm in names:
                changed.append(f"新生 {nm}（已在名单里，跳过）")
                continue
            lines.insert(last + 1 + k, nm)
            changed.append(f"新生 {nm}（已入班）")

    rp.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"[完成] {cls}")
    for c in changed:
        print(f"  ✔ {c}")
    if args.缺考:
        print("\n  缺考的孩子：成绩表会标「缺考」，不出话术/诊断报告/巩固单/群发，"
              "也不参与平均分和排名。")


if __name__ == "__main__":
    main()
