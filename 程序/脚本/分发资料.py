# -*- coding: utf-8 -*-
r"""
按学生建文件夹，把「课测照片 + 该生要发的资料」放进去。

输出结构：
  4-输出\周六10点_五年级第五讲\分发\<姓名>\
      课测照片.jpg          （原名不变）
      背诵单_巩固版.docx    （好生 / 中间生）
      专项练习_10题.docx    （低于重点线的学生）

规则（与话术里承诺的一致）：
  · 分数 < 重点线（默认30）  → 发《专项练习_10题》
  · 其余学生                → 发《背诵单_巩固版》

用法：
  python 分发资料.py --班级 周六10点 --试运行
  python 分发资料.py --班级 周六10点
"""
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

# 名字以老师模板为准：我的写法 -> 模板写法
NAME_FIX = {
    "谢西雯": "谢西文",
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
        parts = s.split()
        if len(parts) >= 2:
            try:
                students.append((parts[0], float(parts[1].rstrip("分"))))
            except ValueError:
                pass
    return meta, students


def find_photos(cls: str, lesson: str):
    """返回 {姓名: [照片路径]}

    查找顺序：
      1. 课次子目录/已改名/姓名.jpg   （照片改名的产物）
      2. 课次子目录/命名对照.txt      （IMG_xxx -> 姓名）
      3. 班级目录根下的命名对照.txt
    """
    base = DIR_PHOTO / cls
    if not base.exists():
        return {}
    cand = [d for d in base.iterdir() if d.is_dir() and lesson and
            re.sub(r"[\s\-_]", "", lesson) in re.sub(r"[\s\-_]", "", d.name)]
    root = cand[0] if cand else base

    photos = {}

    def scan_named_dir(d: Path):
        for f in sorted(d.iterdir()):
            if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg", ".png", ".webp"):
                photos.setdefault(f.stem, []).append(f)

    def apply_mapping(mp: Path):
        n0 = len(photos)
        for raw in mp.read_text(encoding="utf-8-sig").splitlines():
            s = raw.strip()
            if not s or s.startswith("#"):
                continue
            p = re.split(r"[\s,，\t]+", s)
            if len(p) >= 2:
                f = mp.parent / p[0]
                if f.exists():
                    photos.setdefault(p[1], []).append(f)
        return len(photos) > n0

    # 1) 已改名目录
    renamed = root / "已改名"
    if renamed.is_dir() and any(renamed.iterdir()):
        scan_named_dir(renamed)
    # 2) 课次目录的命名对照
    if not photos:
        mp = root / "命名对照.txt"
        if mp.exists():
            apply_mapping(mp)
    # 3) 班级根目录的命名对照
    if not photos:
        mp = base / "命名对照.txt"
        if mp.exists():
            apply_mapping(mp)
    # 4) 课次目录里文件名即姓名的
    if not photos:
        scan_named_dir(root)
    return photos


def main():
    ap = argparse.ArgumentParser(description="按学生分发资料")
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
        if not students:
            continue
        cls = meta.get("班级", rp.stem)
        lesson = meta.get("课次", "")
        try:
            focus = float(meta.get("重点线") or 30)
        except ValueError:
            focus = 30.0

        outdir = DIR_OUT / f"{cls}_{lesson}".replace("/", "_")
        if not outdir.exists():
            print(f"[跳过] {cls}：还没生成产物（先跑 课测反馈.py）")
            continue

        recite = outdir / "背诵单_巩固版.docx"
        practice = outdir / "专项练习_10题.docx"
        dist = outdir / "分发"

        photos = find_photos(cls, lesson)
        print(f"\n{cls} {lesson}　重点线 {focus:g}")

        n_ok = 0
        for name, score in students:
            disp = NAME_FIX.get(name, name)
            d = dist / disp
            got_photo = photos.get(name) or photos.get(disp) or []
            is_low = score < focus
            doc = practice if is_low else recite
            tag = "练习单" if is_low else "背诵单"

            if args.试运行:
                print(f"  {disp:<7}{score:>6}  → {tag}"
                      f"　照片{len(got_photo)}张"
                      + ("  [缺资料]" if not doc.exists() else "")
                      + ("  [缺照片]" if not got_photo else ""))
                n_ok += 1
                continue

            d.mkdir(parents=True, exist_ok=True)
            # 照片（保留原扩展名）
            for i, p in enumerate(got_photo, 1):
                ext = p.suffix.lower()
                dst = d / (f"{disp}{ext}" if len(got_photo) == 1
                           else f"{disp}_{i}{ext}")
                shutil.copy2(p, dst)
            # 资料
            if doc.exists():
                shutil.copy2(doc, d / doc.name)
            print(f"  ✓ {disp:<7}{score:>6}  → {tag}"
                  f"　照片{len(got_photo)}张"
                  + ("  [缺资料]" if not doc.exists() else "")
                  + ("  [缺照片]" if not got_photo else ""))
            n_ok += 1

        print(f"  {'[试运行] ' if args.试运行 else ''}共 {n_ok} 人，输出：{dist}")


if __name__ == "__main__":
    main()
