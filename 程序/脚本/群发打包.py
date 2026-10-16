# -*- coding: utf-8 -*-
r"""
群发打包 —— 把所有孩子的「课测照片 + 应发资料」放进同一个文件夹，
每个文件都用学生名字命名，方便公司群发系统按名字匹配。

输出：
  4-输出\<班级>_<课次>\群发\<班级>_群发\
      郭凌硕.jpg      ← 课测照片
      郭凌硕.docx     ← 该生要发的资料（背诵单 / 专项练习）
      丁禾谕.jpg
      丁禾谕.docx
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

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_OUT = ROOT / "生成结果"
DIR_PHOTO = ROOT / "答题卡照片"

PHOTO_EXT = (".jpg", ".jpeg", ".png", ".webp")

# 名字写法映射（名单里的写法 -> 群发系统里的写法）
# 默认一致；如模板用了别的写法，在这里加一行
NAME_AS_IS = {
    "谢西文": "谢西文",
}


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
        p = s.split()
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
    if not base.exists():
        return {}
    cand = [d for d in base.iterdir() if d.is_dir() and lesson and
            re.sub(r"[\s\-_]", "", lesson) in re.sub(r"[\s\-_]", "", d.name)]
    root = cand[0] if cand else base

    photos = {}

    def scan(d: Path):
        for f in sorted(d.iterdir()):
            if f.is_file() and f.suffix.lower() in PHOTO_EXT:
                photos.setdefault(f.stem, []).append(f)

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
    else:
        mp = root / "命名对照.txt"
        if mp.exists():
            mapping(mp)
        else:
            scan(root)

    # 姓名对照兜底：照片名是本名（马泽文），名单里是别名（马团团）时归位
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

        outdir = DIR_OUT / f"{cls}_{lesson}".replace("/", "_")
        recite = outdir / "背诵单_巩固版.docx"
        practice = outdir / "专项练习_10题.docx"
        personal = meta.get("任务模式") == "个人小任务"   # 每人一份专属小任务
        taskdir = outdir / "小任务"
        if personal:
            if not taskdir.is_dir():
                print(f"[跳过] {cls}：任务模式=个人小任务，但 {taskdir} 还没生成"
                      f"（先跑 小任务.py）")
                continue
        elif not recite.exists() and not practice.exists():
            print(f"[跳过] {cls}：还没生成背诵单/练习单")
            continue

        pack = outdir / "群发" / f"{cls}_群发"
        photos = find_photos(cls, lesson)

        print(f"\n{cls} {lesson}　重点线 {focus:g}"
              f"{'　每人一份小任务' if personal else ''}　→ {pack}")
        if not args.试运行:
            pack.mkdir(parents=True, exist_ok=True)

        n = 0
        for name, score in graded:
            disp = NAME_AS_IS.get(name, name)
            is_low = score < focus
            if personal:
                doc = taskdir / f"{name}.docx"
                tag = "小任务"
            else:
                doc = practice if is_low else recite
                tag = "练习" if is_low else "背诵"
            got = photos.get(name) or photos.get(disp) or []

            if not args.试运行:
                # 照片：<姓名>.jpg
                for i, p in enumerate(got, 1):
                    suffix = p.suffix.lower()
                    fname = f"{disp}{suffix}" if len(got) == 1 else f"{disp}_{i}{suffix}"
                    shutil.copy2(p, pack / fname)
                # 资料：<姓名>.docx
                if doc.exists():
                    shutil.copy2(doc, pack / f"{disp}{doc.suffix}")

            flags = ""
            if not got:
                flags += "  [缺照片]"
            if not doc.exists():
                flags += "  [缺资料]"
            print(f"  ✓ {disp:<8}{score:>6g}  → {tag}　照片{len(got)}张{flags}")
            n += 1

        print(f"  {'[试运行] ' if args.试运行 else ''}共 {n} 人")


if __name__ == "__main__":
    main()
