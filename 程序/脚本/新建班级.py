# -*- coding: utf-8 -*-
r"""新建班级.py —— 一个班一个文本文件，不用装软件不用注册。

    python 新建班级.py                          # 交互式，一步步问
    python 新建班级.py --班级 周六10点 --课次 五年级第五讲 --满分 40
    python 新建班级.py --列出                    # 看现有班级
    python 新建班级.py --删 --班级 测试班         # 删掉一个班（可恢复）
    python 新建班级.py --删 --班级 测试班 --连产物 # 连产物一起清掉
"""
from __future__ import annotations

import argparse
import re
import shutil
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
from _ws import workspace_root, photo_root  # noqa: E402
from roster import strip_id, parse_line  # noqa: E402

ROOT = workspace_root()
DIR_ROSTER = ROOT / "班级名单"
DIR_PHOTO = photo_root(ROOT)

# 建班不预设卷子结构：每次课测的题型/题数/分值都不一样，写死就等于逼老师重建班。



def ask(prompt, default=""):
    d = f"（默认 {default}）" if default else ""
    v = input(f"{prompt}{d}：").strip()
    return v or default


def build_text(班级, 课次="", 满分="", 考试名称="", 称呼="家长", 风格="",
               大题名称="", 大题题数="", 重点线="", 拍照说明=""):
    """名单头部只写「已经知道的」。

    建班是一次性的，那会儿还不知道这讲考什么、满分多少、有几个大题，
    所以这些一律留空，等第一次 ingest 时按卷面写进来。
    """
    lines = [f"@班级: {班级}"]
    for key, val in (("课次", 课次), ("满分", 满分), ("重点线", 重点线),
                     ("考试名称", 考试名称), ("称呼", 称呼)):
        if str(val or "").strip():
            lines.append(f"@{key}: {val}")
    if 风格 and 风格 != "亲和":
        lines.append(f"@家长风格: {风格}")
    if str(大题名称 or "").strip():
        lines.append(f"@大题名称: {大题名称}")
    if str(大题题数 or "").strip():
        lines.append(f"@大题题数: {大题题数}")
    where = f"{班级}/{课次}" if 课次 else f"{班级}/<课次>"
    lines += [
        "# ============================================================",
        f"# {班级}　学员名单（建好就不用动了）",
        "#",
        "# 每次课测只需要读照片、登记分数，这个文件里的【人名部分】永远不碰；",
        "# 只有出现 缺考 / 退费 / 转班 / 新报名 时，才回头改这一份。",
        "#",
        "# 【照片放哪】",
        f"#     照片/{where}/张三.jpg",
        "#",
        "# 【其他可选设置】写在最上面，需要哪个加哪个：",
        "#     @特殊学生: 张三          危机感版话术",
        "#     @安心学生: 张三          不焦虑版话术",
        "#     @严正学生: 张三          认真谈一次（问题点透但不伤人）",
        "#     @表扬学生: 张三,李四      额外点名表扬",
        "#     @同家长: 张三,张三妹      双胞胎合并成一条话术",
        "#     @家长风格: 安抚           全班都走安抚版",
        "#     @任务模式: 个人小任务      每人一份独立小任务（四年级常用）",
        "#     @知识点库: L11-18         挂知识点库，话术会说具体考点",
        "# ============================================================",
        "",
    ]
    return "\n".join(lines)


def import_from_xlsx(args):
    """从老师的官方学员列表建班（一个班一张表，老师只提供材料）。

    能认的表：表头里有「姓名/学生姓名」，学号列叫「听课证号/学员号/学号」均可。
    表头行不固定，前 15 行里找。
    """
    import openpyxl

    src = Path(args.从表).expanduser()
    if not src.exists():
        print(f"[错] 找不到表：{src}")
        return 1
    cls = (args.班级 or src.stem).strip()
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb["Sheet0"] if "Sheet0" in wb.sheetnames else wb.worksheets[0]

    NAME_KEYS = ("姓名", "学生姓名", "学员姓名", "名字")
    ID_KEYS = ("听课证号", "听课证", "学员号", "学号", "学生编号", "证号")

    hdr, col_name, col_id = None, None, None
    for r in range(1, min(ws.max_row, 15) + 1):
        row = [(c, str(ws.cell(row=r, column=c).value or "").strip())
               for c in range(1, ws.max_column + 1)]
        nm = [c for c, v in row if v in NAME_KEYS]
        if nm:
            hdr, col_name = r, nm[0]
            for c, v in row:
                if any(k in v for k in ID_KEYS):
                    col_id = c
                    break
            break
    if hdr is None:
        print("[错] 这张表里找不到「姓名」表头。")
        print("     支持的列名：姓名 / 学生姓名 / 学员姓名")
        return 1

    students, seen, dup = [], set(), []
    for r in range(hdr + 1, ws.max_row + 1):
        nm = str(ws.cell(row=r, column=col_name).value or "").strip()
        if not nm:
            continue
        if nm in seen:
            # 重名不能丢人：加 (2)(3) 区分。丢掉一个孩子，他以后永远收不到东西。
            dup.append(nm)
            k = 2
            while f"{nm}({k})" in seen:
                k += 1
            nm = f"{nm}({k})"
        sid = ""
        if col_id:
            sid = str(ws.cell(row=r, column=col_id).value or "").strip()
        seen.add(nm)
        students.append((sid, nm))
    if dup:
        print(f"  [注意] 班里有重名：{'、'.join(sorted(set(dup)))}")
        for _nm in sorted(set(dup)):
            _rows = [f"{s} {n0}" if s else n0
                     for s, n0 in students if n0.startswith(_nm)]
            print(f"      {_nm} → 名单里写成：{'、'.join(_rows)}")
        print("      ⚠ 重名的孩子，答题卡/照片上也要写清是哪一位（或靠学号区分）。")
        print("        只写「王芳」的话，读图时认不出是哪个，分数会串到一个人身上。")
    if not students:
        print("[错] 表头下面一个学生都没读到。")
        return 1

    rp = DIR_ROSTER / f"{cls}.txt"
    head = build_text(cls, args.课次 or "", args.满分 or "", args.考试名称,
                      args.称呼, args.风格, args.大题名称,
                      args.大题题数, args.重点线 or "")
    # build_text 里带了示例学生，去掉它们，换成真名单
    lines = [ln for ln in head.splitlines()
             if not (ln and not ln.startswith(("#", "@")) and ln.strip())]

    # 重建班（转来新生、重新导出名单）时，把已经登过的分数接回来，不能丢
    old_scores, old_head = {}, {}
    if rp.exists():
        for raw in rp.read_text(encoding="utf-8-sig").splitlines():
            s = raw.strip()
            if not s or s.startswith("#"):
                continue
            if s.startswith("@"):
                if ":" in s:
                    k, v = s[1:].split(":", 1)
                    old_head[k.strip()] = v.strip()
                continue
            parts = strip_id(s.split())
            if parts:
                old_scores[parts[0]] = " ".join(parts[1:]).strip()
        # 课次/满分别被覆盖成空
        for k in ("课次", "满分", "重点线", "大题名称", "大题题数"):
            if not old_head.get(k):
                continue
            if f"@{k}: " not in head:
                lines.insert(len([x for x in lines if x.startswith("@")]),
                             f"@{k}: {old_head[k]}")

    body = []
    for sid, nm in students:
        got = old_scores.get(nm, "")
        body.append(" ".join(x for x in (sid, nm, got) if x))
    rp.parent.mkdir(parents=True, exist_ok=True)
    if rp.exists():
        bak = rp.with_suffix(".txt.bak")
        if not bak.exists():
            bak.write_text(rp.read_text(encoding="utf-8-sig"), encoding="utf-8")
    rp.write_text("\n".join(lines + [""] + body) + "\n", encoding="utf-8")
    (DIR_PHOTO / cls).mkdir(parents=True, exist_ok=True)

    kept = [nm for _, nm in students if old_scores.get(nm)]
    no_id = [nm for sid, nm in students if not sid]
    print(f"✔ 建好了：{rp}")
    print(f"  班级：{cls}　学生：{len(students)} 人"
          f"（学号 {len(students) - len(no_id)} 个）")
    if kept:
        print(f"  已接回 {len(kept)} 人的历史分数，没覆盖掉")
    print("  课次/满分/大题结构不写死，第一次课测时按卷面自动补进去。")
    if no_id:
        print(f"  [注意] {len(no_id)} 人没有学号：{'、'.join(no_id[:6])}"
              "　—— 填群发表时只能按姓名匹配")
    return 0


def main():
    ap = argparse.ArgumentParser(description="新建/管理班级")
    ap.add_argument("--班级")
    ap.add_argument("--课次", default=None, help="一般不用给，第一次课测时自动写入")
    ap.add_argument("--满分", default=None, help="一般不用给")
    ap.add_argument("--考试名称", default=None)
    ap.add_argument("--称呼", default="家长")
    ap.add_argument("--风格", default=None)
    ap.add_argument("--大题名称", default=None)
    ap.add_argument("--大题题数", default=None)
    ap.add_argument("--重点线", default=None)
    ap.add_argument("--从表", default=None,
                    help="老师的学员列表 xlsx（一个班一张表），直接建班")
    ap.add_argument("--列出", action="store_true")
    ap.add_argument("--删", action="store_true", help="删掉一个班")
    ap.add_argument("--连产物", action="store_true",
                    help="删班时连这个班的照片/成绩表/群发/报告一起移走")
    args = ap.parse_args()

    if args.从表:
        return import_from_xlsx(args)

    if args.列出:
        print(f"{'班级':14s} {'课次':24s} {'人数':>4s} {'已填分':>6s}")
        print("-" * 56)
        for f in sorted(DIR_ROSTER.glob("*.txt")):
            if f.name.endswith(".bak") or f.stem == "姓名对照":
                continue
            meta, n, sc = {}, 0, 0
            for raw in f.read_text(encoding="utf-8-sig").splitlines():
                s = raw.strip()
                if not s or s.startswith("#"):
                    continue
                if s.startswith("@"):
                    if ":" in s:
                        k, v = s[1:].split(":", 1)
                        meta[k.strip()] = v.strip()
                    continue
                n += 1
                # 「已填分」只数真的有分数的行。
                # 不能按「这一行有两块」判断——名单建好时每行都是「学号 姓名」，
                # 两块但没分数，会把新建的 27 人班显示成 27/27，
                # 老师以为已经登过分了。交给 roster.parse_line 判，
                # 它认得「张三 85」「张三 85(3,7)」「张三 缺考」这几种写法。
                _d = parse_line(s)
                if _d and _d["分数"] is not None:
                    sc += 1
            if not meta.get("班级"):
                continue
            print(f"{meta['班级']:14s} {(meta.get('课次') or '（未开课）'):24s} {n:4d} {sc:6d}")
        return 0

    if args.删:
        if not args.班级:
            print("请用 --班级 指定要删的班（先 --列出 看有哪些）")
            return 2
        cls = args.班级
        rp = DIR_ROSTER / f"{cls}.txt"
        if not rp.exists():
            print(f"[错] 没有叫「{cls}」的班。先 --列出 看全部班级。")
            return 2

        import datetime as _dt
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        trash = ROOT / "_已删除" / f"{cls}_{stamp}"
        trash.mkdir(parents=True, exist_ok=True)

        # ① 名单 → _已删除/（可恢复，不是真删）
        shutil.move(str(rp), str(trash / f"{cls}.txt"))
        print(f"✓ 名单已移到：_已删除/{trash.name}/{cls}.txt")

        # ② 成绩历史里这个班的行要删干净。
        # 不删的话，以后重建一个同名班，第一讲就会显示「比上次进步 12 分」——
        # 那个「上次」是上一个班的成绩，纯粹是串的。
        hist = ROOT / "生成结果" / "历史" / "成绩历史.csv"
        if hist.exists():
            import csv as _csv
            rows = list(_csv.reader(hist.read_text(encoding="utf-8-sig").splitlines()))
            if rows:
                head, body = rows[0], rows[1:]
                ci = head.index("班级") if "班级" in head else -1
                kept = [r for r in body if not (ci >= 0 and len(r) > ci
                                                and r[ci].strip() == cls)]
                gone = len(body) - len(kept)
                if gone:
                    with hist.open("w", encoding="utf-8-sig", newline="") as fh:
                        w = _csv.writer(fh)
                        w.writerow(head)
                        w.writerows(kept)
                    print(f"✓ 成绩历史里删掉 {gone} 行（这个班的记录）")
                else:
                    print("· 成绩历史里没有这个班的记录")

        # ③ 产物：列出来，问老师要不要一起清
        left = []
        for pat in (ROOT / "照片" / cls,
                    ROOT / "成绩表" / cls,
                    ROOT / "群发" / cls):
            if pat.exists():
                left.append(pat)
        for sub in ("诊断报告", "巩固单", "话术"):
            d = ROOT / "生成结果" / sub / cls
            if d.exists():
                left.append(d)
        for d in (ROOT / "生成结果").glob(f"{cls}_*"):
            if d.is_dir():
                left.append(d)

        if left:
            if args.连产物:
                for d in left:
                    # 按原来的相对路径搬，别只用目录名：
                    # 照片/示例班 和 生成结果/诊断报告/示例班 同名，撞在一起会报错
                    rel = d.relative_to(ROOT)
                    tgt = trash / rel
                    tgt.parent.mkdir(parents=True, exist_ok=True)
                    if tgt.exists():
                        shutil.rmtree(tgt)
                    shutil.move(str(d), str(tgt))
                print(f"✓ 产物 {len(left)} 个目录已一起移到 _已删除/{trash.name}/")
            else:
                print(f"\n· 这些目录还留着（里面有这个班的历史产物）：")
                for d in left:
                    print(f"    {d.relative_to(ROOT)}")
                print("  不想留就重跑一次加 --连产物")

        print(f"\n班级「{cls}」已删除。想恢复：把 _已删除/{trash.name}/ 里的东西搬回去。")
        return 0

    if not args.班级:
        print("[用法] 两种建班方式：")
        print('  从企微导出的学员列表建：')
        print('    python 新建班级.py --从表 "学员列表.xlsx" --班级 "周五慢班"')
        print('  手写一个空名单：')
        print('    python 新建班级.py --班级 "周五慢班" --称呼 家长')
        print('  看已有哪些班：python 新建班级.py --列出')
        return 2

    text = build_text(args.班级, args.课次 or "", args.满分 or "",
                      args.考试名称, args.称呼, args.风格, args.大题名称,
                      args.大题题数, args.重点线 or "")
    rp = DIR_ROSTER / f"{args.班级}.txt"
    rp.write_text(text, encoding="utf-8")
    (DIR_PHOTO / args.班级).mkdir(parents=True, exist_ok=True)
    print(f"已建：{rp}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
