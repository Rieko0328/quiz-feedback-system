# -*- coding: utf-8 -*-

r"""

群发打包 —— 把所有孩子的「课测照片 + 应发资料」放进同一个文件夹，

每个文件都用学生名字命名，方便公司群发系统按名字匹配。



输出：

  生成结果\<班级>_<课次>\群发\<班级>_群发\

      郭小刚.jpg      ← 课测照片

      郭小刚.docx     ← 该生要发的资料（背诵单 / 专项练习）

      丁小明.jpg

      丁小明.docx

      ...



规则（与话术里承诺的一致）：

  · 分数 < 重点线（默认30）→ 该生资料 = 专项练习

  · 其余学生              → 该生资料 = 背诵单



用法：

  python 群发打包.py --班级 周六2点 --试运行

  python 群发打包.py --班级 周六2点

  python 群发打包.py                # 处理所有有名单的班

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

CP_FAIL = []          # 复制失败的文件名，main 结尾汇总并返回非 0





def _cp(src, dst):

    """复制文件；被占用/权限不足时给一句人话，不中断整批。"""

    try:

        shutil.copy2(src, dst)

        return True

    except (PermissionError, OSError) as e:

        CP_FAIL.append(Path(dst).name)

        print(f"    [跳过] {Path(dst).name}：{type(e).__name__}"

              f"（文件被别的程序占用？关掉再重跑）")

        return False



sys.path.insert(0, str(Path(__file__).resolve().parent))

from _ws import fast_remove, workspace_root, photo_root, safe_name  # noqa: E402

from roster import strip_id  # noqa: E402



ROOT = workspace_root()

DIR_ROSTER = ROOT / "班级名单"

DIR_OUT = ROOT / "生成结果"

DIR_PHOTO = photo_root(ROOT)



PHOTO_EXT = (".jpg", ".jpeg", ".png", ".webp")



# 名字写法映射（名单里的写法 -> 群发系统里的写法）

# 一般不用填。只有「名单和群发系统写法确实不同」时才加，例如：

#   NAME_AS_IS = {"张小　明": "张小明"}      # 名单里有全角空格

# 别把某个班的真名写死在这里——换老师换班就变成别人的错误数据。

NAME_AS_IS: dict[str, str] = {}





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

        p = strip_id(s.split())

        if len(p) >= 2:

            try:

                students.append((p[0], float(p[1].rstrip("分"))))

            except ValueError:

                pass

        elif len(p) == 1:

            students.append((p[0], None))   # 缺考

    return meta, students





def load_name_alias():

    """姓名对照表：照片/答题卡上的名字 -> 名单里的名字。"""

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





def find_photos(cls: str, lesson: str):

    """返回 {姓名: [照片路径]}"""

    base = DIR_PHOTO / cls

    if not base.exists() and DIR_PHOTO.is_dir():

        # 名单里的 @班级 和文件夹名不完全一样时（多了空格、改了叫法）兜一下

        _hit = [d for d in DIR_PHOTO.iterdir()

                if d.is_dir() and (d.name == cls or cls in d.name

                                   or d.name in cls)]

        if len(_hit) == 1:

            base = _hit[0]

    if not base.exists():

        return {}

    cand = [d for d in base.iterdir() if d.is_dir() and lesson and

            re.sub(r"[\s\-_]", "", lesson) in re.sub(r"[\s\-_]", "", d.name)]

    root = cand[0] if cand else base



    photos = {}



    def scan(d: Path):

        for f in sorted(d.iterdir()):

            if f.is_file() and f.suffix.lower() in PHOTO_EXT:

                same = photos.setdefault(f.stem, [])

                # 同名文件只留一份（优先 已改名/，因为它已经是对的名字）

                if any(x.name == f.name for x in same):

                    continue

                same.append(f)



    def mapping(mp: Path):

        for raw in mp.read_text(encoding="utf-8-sig").splitlines():

            s = raw.strip()

            if not s or s.startswith("#"):

                continue

            p = re.split(r"[\s,，\t]+", s)

            if len(p) >= 2:

                f = mp.parent / p[0]

                if f.exists():

                    photos.setdefault(p[1], []).append(f)



    renamed = root / "已改名"

    if renamed.is_dir() and any(renamed.iterdir()):

        scan(renamed)

    mp = root / "命名对照.txt"

    if mp.exists():

        mapping(mp)

    # 兜底：文件名本身就是学生姓名的（后补的照片、老师手工改的名）。

    # 原来这里是 if/else——有对照表就只查表，直接放在目录里的 张小明.jpg

    # 反而被忽略；没对照表就只扫目录，IMG_001.jpg 跟姓名对不上，

    # 于是「每人都缺照片」但一声不吭。两头都要兜。

    scan(root)



    # 姓名对照兜底：照片名是本名（李明），名单里是别名（李小明）时归位

    for _card, _roster in load_name_alias().items():

        if photos.get(_card):

            photos.setdefault(_roster, []).extend(photos.pop(_card))

    return photos





def main():

    ap = argparse.ArgumentParser(description="群发资料打包（扁平化，按姓名命名）")

    ap.add_argument("--班级", default=None)

    ap.add_argument("--试运行", action="store_true")

    args = ap.parse_args()



    rosters = [f for f in sorted(DIR_ROSTER.glob("*.txt"))

               if not f.name.startswith(("_", "~$"))]

    if args.班级:

        rosters = [f for f in rosters if args.班级 in f.stem]

    if not rosters:

        print("没有找到名单。")

        return



    for rp in rosters:

        meta, students = read_roster(rp)

        graded = [(n, s) for n, s in students if s is not None]

        if not graded:

            continue

        cls = meta.get("班级", rp.stem)

        lesson = meta.get("课次", "")

        try:

            focus = float(meta.get("重点线") or 30)

        except ValueError:

            focus = 30.0



        outdir = DIR_OUT / f"{safe_name(cls)}_{safe_name(lesson)}"

        recite = outdir / "背诵单_巩固版.docx"

        practice = outdir / "专项练习_10题.docx"

        # 每人一份的产物现在叫「巩固单」，不再需要名单里的 @任务模式

        personal = any((outdir / "巩固单").glob("*.docx"))

        # 没有练习也照发：很多课测只把卷子和照片发给家长，不需要附练习



        pack = outdir / "群发" / f"{safe_name(cls)}_群发"

        photos = find_photos(cls, lesson)



        print(f"\n{cls} {lesson}　重点线 {focus:g}"

              f"{'　每人一份巩固单' if personal else ''}　→ {pack}")

        if not args.试运行:

            pack.mkdir(parents=True, exist_ok=True)



        n = 0

        no_photo = []            # 找不到照片的人，结尾统一提示

        for name, score in graded:

            disp = NAME_AS_IS.get(name, name)

            is_low = score < focus

            # 优先用这孩子的专属巩固单（要背的+要做的+答案解析）。

            # 找不到时要去归档后的位置再找一次：归档会把 outdir/巩固单/ 整个搬到

            # 生成结果/巩固单/<班级>/<课次>/，先归档后群发就会扑空。

            # 不再退回「小任务」——它和巩固单内容重复，一个孩子会收到两份。

            ex = outdir / "巩固单" / f"{name}_巩固练习.docx"

            if not ex.exists():

                ex = (DIR_OUT / "巩固单" / safe_name(cls) / safe_name(lesson)

                      / f"{name}_巩固练习.docx")

            if ex.exists():

                doc, tag = ex, "巩固单"

            elif practice.exists() or recite.exists():

                doc = practice if is_low else recite

                tag = "练习" if is_low else "背诵"

            else:

                doc, tag = None, "仅照片"

            got = photos.get(name) or photos.get(disp) or []



            if not args.试运行:

                # 照片：<姓名>.jpg

                for i, p in enumerate(got, 1):

                    suffix = p.suffix.lower()

                    fname = f"{disp}{suffix}" if len(got) == 1 else f"{disp}_{i}{suffix}"

                    _cp(p, pack / fname)

                # 资料：<姓名>.docx（没有练习就只发照片）

                if doc is not None and doc.exists():

                    _cp(doc, pack / f"{disp}{doc.suffix}")



            flags = ""

            if not got:

                flags += "  [缺照片]"

                no_photo.append(disp)

            if doc is not None and not doc.exists():

                flags += "  [缺资料]"

            print(f"  ✓ {disp:<8}{score:>6g}  → {tag}　照片{len(got)}张{flags}")

            n += 1



        print(f"  {'[试运行] ' if args.试运行 else ''}共 {n} 人")



        # 重跑时清掉不属于本次的人（比如刚被标缺考的）的旧文件，

        # 否则老师照着这个目录发，会把材料发给不该发的人

        if not args.试运行 and pack.is_dir():

            keep = {NAME_AS_IS.get(nm, nm) for nm, _ in graded}

            gone, _stale = [], []

            for f in pack.iterdir():

                if not f.is_file():

                    continue

                base = f.stem

                tail = base.rsplit("_", 1)

                if len(tail) == 2 and tail[1].isdigit():

                    base = tail[0]

                if base not in keep:

                    _stale.append(f)

                    gone.append(f.name)

            if _stale:

                fast_remove(*_stale)       # 一次系统命令删完

            if gone:

                print(f"  清掉 {len(gone)} 个不该发的旧文件："

                      f"{'、'.join(gone[:4])}{'…' if len(gone) > 4 else ''}")



        # 一张照片都没配上 —— 这是「发出去才发现没图」的经典场景，

        # 逐行的 [缺照片] 老师根本不会看，必须在结尾说清楚。

        if no_photo and not args.试运行:

            print(f"\n[注意] {len(no_photo)} 个人没找到照片："

                  f"{'、'.join(no_photo[:5])}{'…' if len(no_photo) > 5 else ''}")

            print("       照片放 照片/<班级>/<课次>/ 下，文件名直接用学生姓名；")

            print("       或者每行写「文件名 姓名」存成 命名对照.txt。")

            print("       只发文字不发照片也行，但群里那份要补上。")



    if CP_FAIL:

        print(f"\n[注意] 有 {len(CP_FAIL)} 个文件没拷进群发目录："

              f"{'、'.join(CP_FAIL[:5])}{'…' if len(CP_FAIL) > 5 else ''}")

        print("       关掉正在打开这些文件的 Word/看图软件，重跑一次就好。")

        return 1

    return 0





if __name__ == "__main__":

    sys.exit(main() or 0)

