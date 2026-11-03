# -*- coding: utf-8 -*-

r"""归档.py —— 把散落在「生成结果」里的产物，整理成老师要的 5 个主目录。



老师看到的最终结构：



    <工作区>/

    ├── 照片/<班级>/<课次>/            这堂课测的照片

    ├── 成绩表/<班级>/<课次>_成绩表.xlsx

    ├── 小练习/<课次>/背诵单.docx 专项练习.docx     （通用版，全班共用）

    ├── 群发/<班级>/<课次>/<姓名>.jpg + <姓名>.docx

    └── 成绩表/<班级>/<班级>_成绩总表.xlsx          （随课次累加，看进步）



过程文件（大题扣分、小任务、话术 txt…）留在 生成结果/<班级>_<课次>/；

其中两样要长期留档的，归档时搬到：



    生成结果/话术/<班级>/<课次>.docx

    生成结果/学情反馈单/<班级>/<课次>/<姓名>.png



用法：

    python 归档.py --班级 五年级A班          # 整理这个班的全部课次

    python 归档.py                          # 整理所有班

    python 归档.py --班级 五年级A班 --课次 五年级第五讲

    python 归档.py --试运行                 # 只看会怎么放，不动文件

"""

from __future__ import annotations



import argparse

import csv

import os

import shutil

import sys

from collections import OrderedDict, defaultdict

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



from _ws import fast_remove, workspace_root, ensure_workspace, dirs_for, photo_root, safe_name  # noqa: E402

from roster import parse_line  # noqa: E402



IMG_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".heic"}







# ---------------------------------------------------------------- 工具



def _safe_copy(src, dst):

    """复制一个文件；被别的程序占用时给一句人话，不抛 traceback。"""

    try:

        shutil.copy2(src, dst)

        return True

    except PermissionError:

        print(f"  [跳过] {Path(dst).name} 正被别的程序打开（Excel/Word？），"

              f"关掉它再重跑就不会漏")

        return False





def _take(src, dst) -> bool:

    """把文件**搬**到 dst（不是拷一份再删）。



    为什么不用 copy2 + unlink：

    1) 同一块盘上 rename 是瞬时的，拷一份 120KB 的 PNG 再删要慢两个数量级；

    2) 某些环境（含 WorkBuddy 的沙箱、装了「删除进回收站」的机器）会把

       每次 unlink 变成一次子进程调用去移回收站，单个文件 60ms，

       40 个学生 × 3 个文件 = 十几秒，全耗在这上面。

    """

    dst = Path(dst)

    dst.parent.mkdir(parents=True, exist_ok=True)

    try:

        os.replace(str(src), str(dst))     # 同盘：原子覆盖，瞬时

        return True

    except OSError:

        pass

    try:

        shutil.copy2(src, dst)             # 跨盘：退回拷贝

        fast_remove(src)                   # 别用 os.unlink：可能被拦成「移回收站」

        return True

    except PermissionError:

        print(f"  [跳过] {dst.name} 正被别的程序打开（Excel/Word？），"

              f"关掉它再重跑就不会漏")

        return False





def copy_files(src: Path, dst: Path, pattern="*", rename=None, dry=False):

    """把 src 下匹配的文件复制到 dst，返回复制个数。"""

    if not src.is_dir():

        return 0

    dst.mkdir(parents=True, exist_ok=True)

    n = 0

    for f in sorted(src.glob(pattern)):

        if not f.is_file():

            continue

        name = rename(f) if rename else f.name

        if not name:

            continue

        tgt = dst / name

        if dry:

            print(f"      {f.name}  ->  {tgt.relative_to(dst.parent.parent)}")

            n += 1

            continue

        if f.resolve() == tgt.resolve():

            continue

        _safe_copy(f, tgt)

        n += 1

    return n





def read_roster(path: Path):

    """名单 -> (meta, 姓名列表)。注意名单可能带学号，不能取第一列当姓名。"""

    meta, students = {}, []

    if not path.exists():

        return meta, students

    for raw in path.read_text(encoding="utf-8-sig").splitlines():

        s = raw.strip()

        if not s or s.startswith("#"):

            continue

        if s[:1] in "@＠":                 # 全角 ＠ 也认，和 roster.py 一致

            if ":" in s or "：" in s:

                k, v = s[1:].replace("：", ":").split(":", 1)

                meta[k.strip()] = v.strip()

            continue

        d = parse_line(s)

        if d and d["姓名"]:

            students.append(d["姓名"])

    return meta, students





# ---------------------------------------------------------------- 汇总表

def load_history(root: Path):

    """读成绩历史 -> {(班级,姓名): {课次: 分数 or None(缺考)}}，并返回课次顺序。"""

    p = root / "生成结果" / "历史" / "成绩历史.csv"

    hist = defaultdict(dict)

    order = OrderedDict()

    if not p.exists():

        return hist, []

    with p.open("r", encoding="utf-8-sig", newline="") as f:

        for row in csv.DictReader(f):

            cls, lesson, name = row.get("班级", ""), row.get("课次", ""), row.get("姓名", "")

            if not (cls and name and lesson):

                continue

            order.setdefault(cls, [])

            if lesson not in order[cls]:

                order[cls].append(lesson)

            if str(row.get("状态") or "").strip() == "缺考":

                hist[(cls, name)][lesson] = None

                continue

            try:

                sc = float(row.get("分数") or "")

            except ValueError:

                continue

            hist[(cls, name)][lesson] = sc

    return hist, order





def write_summary(root, cls, students, hist, lessons, dry=False):

    """成绩表/<班级>/<班级>_成绩总表.xlsx —— 一行一个学生，一列一次课。"""

    from openpyxl import Workbook

    from openpyxl.styles import Alignment, Font, PatternFill

    from openpyxl.utils import get_column_letter



    out = dirs_for(root, cls)["summ"] / f"{safe_name(cls)}_成绩总表.xlsx"

    # 只列名单里的人（没考的留空）。

    # 历史里有、名单里没有的（退费/转走的）不再补回来——

    # 之前补回来会让他继续进平均分，老师明明已经把他移除了。

    names = list(students)

    gone = [n for (c, n) in hist if c == cls and n not in names]

    if gone:

        print(f"  [提示] {len(gone)} 人已不在名单（退费/转走），汇总表不再统计："

              f"{'、'.join(gone[:5])}")

    if not names:

        return None



    if dry:

        print(f"  汇总表 -> {out}（{len(names)} 人 × {len(lessons)} 次课）")

        return out



    wb = Workbook()

    ws = wb.active

    ws.title = "成绩总表"

    head = ["姓名"] + lessons + ["最近一次", "上一次", "变化", "平均分", "趋势"]

    ws.append([f"{cls} 成绩总表（随课次自动累加）"] + [""] * (len(head) - 1))

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(head))

    ws.cell(row=1, column=1).font = Font(bold=True, size=13, color="1B5EA8")

    ws.append(head)

    for c in ws[2]:

        c.font = Font(bold=True, color="FFFFFF")

        c.fill = PatternFill("solid", fgColor="1B5EA8")

        c.alignment = Alignment(horizontal="center", vertical="center")



    up, down = "C0392B", "1F7A3D"      # 涨红跌绿

    for name in names:

        row = [name]

        got = hist.get((cls, name), {})

        cells = []

        for l in lessons:

            if l not in got:

                cells.append("")

            elif got[l] is None:

                cells.append("缺考")          # 缺考留痕，不是空白

            else:

                cells.append(round(got[l], 1))

        row += cells

        nums = [(l, got[l]) for l in lessons

                if l in got and got[l] is not None]

        vals = [v for _, v in nums]

        last = vals[-1] if vals else None

        prev = vals[-2] if len(vals) > 1 else None

        avg = round(sum(vals) / len(vals), 1) if vals else None

        delta = None if (last is None or prev is None) else round(last - prev, 1)

        if delta is None:

            trend = "首次" if vals else ""

        elif delta > 0.5:

            trend = f"↑ 进步 {delta:g}"

        elif delta < -0.5:

            trend = f"↓ 退步 {abs(delta):g}"

        else:

            trend = "→ 持平"

        row += [last, prev, delta, avg, trend]

        ws.append(row)

        r = ws.max_row

        if delta is not None and abs(delta) > 0.5:

            ws.cell(row=r, column=len(head) - 1).font = Font(

                bold=True, color=up if delta > 0 else down)

            ws.cell(row=r, column=len(head)).font = Font(

                color=up if delta > 0 else down)

        if len(vals) == 1:

            ws.cell(row=r, column=len(head)).font = Font(color="5F5E5A")

        for ci, v in enumerate(cells, start=2):

            if v == "缺考":

                ws.cell(row=r, column=ci).font = Font(color="888780")



    # 班级平均

    ws.append([])

    avg_row = ["班级平均"]

    for l in lessons:

        xs = [hist.get((cls, n), {}).get(l) for n in names]

        xs = [v for v in xs if v is not None]

        avg_row.append(round(sum(xs) / len(xs), 1) if xs else "")

    ws.append(avg_row)

    for c in ws[ws.max_row]:

        c.font = Font(bold=True)



    ws.column_dimensions["A"].width = 12

    for i in range(2, len(head) + 1):

        ws.column_dimensions[get_column_letter(i)].width = 12

    ws.freeze_panes = "B3"

    out.parent.mkdir(parents=True, exist_ok=True)

    wb.save(out)

    return out





# ---------------------------------------------------------------- 主流程

def archive_one(root: Path, cls: str, lesson: str, dry=False, keep_copy=False):

    d = dirs_for(root, cls, lesson)

    work = d["work"]

    if not work.exists():

        return None

    done = {}



    # ① 照片：老工作区的「答题卡照片」并进「照片」

    old_p = root / "答题卡照片" / cls / lesson

    new_p = d["photos"]

    n = 0

    if old_p.is_dir() and old_p.resolve() != new_p.resolve():

        n += copy_files(old_p, new_p, "*", dry=dry)

        n += copy_files(old_p / "已改名", new_p / "已改名", "*", dry=dry)

    if new_p.is_dir():

        n = max(n, len([f for f in new_p.iterdir()

                        if f.is_file() and f.suffix.lower() in IMG_EXT]))

    done["照片"] = (new_p, n)



    # ② 成绩表：成绩表/<班级>/<课次>_成绩表.xlsx

    # Excel 开着时课测反馈会另存成 课测成绩表_新.xlsx（原名存不进去）。

    # 只认原名的话，归档进目录的是上一版的旧表，日志还照样说「成绩表：1 个」。

    src = work / "课测成绩表.xlsx"

    _alts = sorted(work.glob("课测成绩表*.xlsx"),

                   key=lambda f: f.stat().st_mtime, reverse=True)

    if _alts:

        src = _alts[0]

    if src.exists():

        tgt = d["scores"] / f"{safe_name(lesson)}_成绩表.xlsx"

        if not dry:

            tgt.parent.mkdir(parents=True, exist_ok=True)

            _safe_copy(src, tgt)

        done["成绩表"] = (tgt, 1)



    # ③ 通用练习：生成结果/通用练习/<课次>/ 两份，名字简化

    pairs = [("背诵单_巩固版.docx", "背诵单.docx"),

             ("专项练习_10题.docx", "专项练习.docx")]

    cnt = 0

    for s, t in pairs:

        f = work / s

        if not f.exists():

            continue

        if not dry:

            d["pract"].mkdir(parents=True, exist_ok=True)

            _safe_copy(f, d["pract"] / t)

        cnt += 1

    if cnt:

        done["通用练习"] = (d["pract"], cnt)



    # ④ 群发：群发/<班级>/<课次>/<姓名>.jpg + <姓名>.docx

    srcp = work / "群发" / f"{safe_name(cls)}_群发"

    if srcp.is_dir():

        # 先拷、拷成功了再清旧文件。

        # 反过来的话，一旦某个文件被看图软件占着拷失败，目录里就只剩清空后的残局，

        # 老师要发的时候才发现群发里没图。

        cnt = copy_files(srcp, d["pack"], "*", dry=dry)

        if not dry and d["pack"].is_dir():

            # 以本次为准清掉上一轮留下的（被标缺考/转走的孩子），

            # 但本轮想拷的一律不动——万一刚才拷失败了，旧的那份还在。

            _want = {x.name for x in srcp.iterdir() if x.is_file()}

            stale = [f for f in d["pack"].iterdir()

                     if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg",

                                                             ".png", ".docx")

                     and f.name not in _want]

            if stale:

                fast_remove(*stale)

            # 一张图都没有 —— 多半是照片没对上名字（见 命名对照.txt），

            # 这时候发出去家长只看得到文字。

            if cnt and not any(f.suffix.lower() in (".jpg", ".jpeg", ".png")

                               for f in d["pack"].iterdir() if f.is_file()):

                print(f"  [注意] {d['pack'].name} 里一张照片都没有——"

                      f"检查 照片/{cls}/{lesson}/ 下的照片名字，"

                      f"或写一份 命名对照.txt")

        done["群发"] = (d["pack"], cnt)

        if not dry and not keep_copy and cnt:

            fast_remove(srcp)                            # 搬走了，别留重复

            if not any(srcp.parent.iterdir()):

                fast_remove(srcp.parent)



    # ④b 照片归档/（课测反馈按人建的那份）内容已在「照片」和「群发」里，删掉避免重复

    src_arch = work / "照片归档"

    if src_arch.is_dir() and not keep_copy and not dry:

        fast_remove(src_arch)



    # ⑤ 家长话术：生成结果/话术/<班级>/<课次>.docx

    src_docx = work / "家长反馈话术.docx"

    if src_docx.exists():

        tgt = (root / "生成结果" / "话术" / safe_name(cls)

               / f"{safe_name(lesson)}.docx")

        if not dry:

            tgt.parent.mkdir(parents=True, exist_ok=True)

            _safe_copy(src_docx, tgt)

            if not keep_copy:

                fast_remove(src_docx)

        done["话术"] = (tgt, 1)



    # ⑥ 学情反馈单：生成结果/学情反馈单/<班级>/<课次>/<姓名>.png

    src_png = work / "学情反馈单"

    if src_png.is_dir():

        imgs = [f for f in src_png.iterdir()

                if f.is_file() and f.suffix.lower() == ".png"]

        if imgs:

            tgt_dir = (root / "生成结果" / "学情反馈单"

                       / safe_name(cls) / safe_name(lesson))

            if not dry:

                tgt_dir.mkdir(parents=True, exist_ok=True)

                moved = []

                for f in imgs:

                    if _safe_copy(f, tgt_dir / f.name):

                        moved.append(f)

                if not keep_copy:

                    fast_remove(*moved)

            done["学情反馈单"] = (tgt_dir, len(imgs))

            if not dry and not keep_copy and not any(src_png.iterdir()):

                fast_remove(src_png)         # 搬空了，删掉空目录



    # ⑦⑧ 诊断报告 / 巩固单：同样按 类别/<班级>/<课次>/ 归位

    for folder, pat, cat, label in (

            ("诊断报告", ".png", "诊断报告", "诊断报告"),

            ("巩固单", ".docx", "巩固单", "巩固单")):

        src = work / folder

        if not src.is_dir():

            continue

        files = [f for f in src.iterdir()

                 if f.is_file() and f.suffix.lower() == pat]

        if not files:

            continue

        tgt_dir = (root / "生成结果" / cat / safe_name(cls) / safe_name(lesson))

        if not dry:

            tgt_dir.mkdir(parents=True, exist_ok=True)

            # 顺序很关键：先复制，成功之后才删源文件、才清上一轮的旧文件。

            # 原来「先清后拷」，一旦某个文件被 Word/看图软件占着就 PermissionError，

            # 旧文件已经删了、新文件没进来 —— 老师那边的报告直接少一页。

            ok_names = set()

            for f in files:

                if keep_copy:

                    ok = _safe_copy(f, tgt_dir / f.name)

                else:

                    ok = _take(f, tgt_dir / f.name)   # 搬：瞬时，且不触发回收站

                if ok:

                    ok_names.add(f.name)

            # 以本次为准：清掉上一轮留下的（被标缺考/转走的孩子不该还留着）。

            # 但「本轮想拷却拷失败」的那个不能删——源文件还在，删了这边就真缺件了。

            _want = {f.name for f in files}

            stale = [old for old in tgt_dir.iterdir()

                     if (old.is_file() and old.suffix.lower() == pat

                         and old.name not in ok_names

                         and old.name not in _want)]

            if stale:

                fast_remove(*stale)          # 可能几十个，一次删完

            if not keep_copy and src.is_dir() and not any(src.iterdir()):

                fast_remove(src)

        done[label] = (tgt_dir, len(files))



    return done





def main():

    ap = argparse.ArgumentParser(description="把产物整理成标准目录")

    ap.add_argument("--班级", default=None)

    ap.add_argument("--课次", default=None)

    ap.add_argument("--工作区", default=None)

    ap.add_argument("--试运行", action="store_true")

    ap.add_argument("--保留副本", action="store_true",

                    help="话术/反馈单复制一份，过程目录里的原件不删（默认搬走）")

    args = ap.parse_args()



    root = ensure_workspace(workspace_root(args.工作区))

    hist, order = load_history(root)



    rosters = [f for f in sorted((root / "班级名单").glob("*.txt"))

               if not f.name.startswith(("_", "~$"))]

    if args.班级:

        rosters = [f for f in rosters if args.班级 in f.stem]

    if not rosters:

        print("[提示] 没有找到班级名单，先在 班级名单/ 里建一个")

        return 1



    total = 0

    for rp in rosters:

        meta, names = read_roster(rp)

        cls = meta.get("班级", rp.stem)

        lesson = args.课次 or meta.get("课次", "")

        if not lesson:

            print(f"[跳过] {cls}：名单里没有 @课次")

            continue

        print(f"\n{cls}　{lesson}")

        done = archive_one(root, cls, lesson, dry=args.试运行,

                           keep_copy=args.保留副本)

        if done is None:

            print(f"  [跳过] {cls}_{lesson}：还没生成产物（先跑 一键生成）")

            continue

        for k, (path, n) in done.items():

            try:

                rel = path.relative_to(root)

            except ValueError:

                rel = path

            print(f"  {k}：{n} 个 -> {rel}")



        s = write_summary(root, cls, names, hist, order.get(cls, []), dry=args.试运行)

        if s:

            print(f"  汇总：-> {s.relative_to(root)}（{len(order.get(cls, []))} 次课）")

        total += 1



    print(f"\n{'[试运行] ' if args.试运行 else ''}整理完成 {total} 个课次")

    return 0





if __name__ == "__main__":

    sys.exit(main() or 0)

