# -*- coding: utf-8 -*-
r"""转jpg.py —— 把 HEIC/HEIF/PNG 之类统一转成 jpg（答题卡照片偶尔是 iPhone 原图）。

为什么需要：Pillow 默认读不了 HEIC，而 macOS 上 AirDrop 过来的照片经常就是 .HEIC。
不转的话后面所有环节都会说「找不到照片」。

用法：
    python 转jpg.py --目录 "照片/周五慢班/五年级第五讲"
    python 转jpg.py --目录 "..." --删原图      # 转完删掉 HEIC 原件

优先级：macOS 自带 sips（零依赖）→ Pillow + pillow-heif（需额外装）
转换后原文件保留在 原图_HEIC/ 里，不会丢。
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

HEIC = {".heic", ".heif"}
OTHER = {".png", ".webp", ".bmp", ".tif", ".tiff"}


def convert_one(src: Path, dst: Path) -> bool:
    """先试 sips（Mac），再试 pillow-heif。"""
    if shutil.which("sips"):
        r = subprocess.run(["sips", "-s", "format", "jpeg", str(src),
                            "--out", str(dst)],
                           capture_output=True)
        if r.returncode == 0 and dst.exists():
            return True
    try:
        from pillow_heif import register_heif_opener
        register_heif_opener()
        from PIL import Image, ImageOps
        im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
        im.save(dst, "JPEG", quality=92)
        return True
    except Exception:
        pass
    try:
        from PIL import Image, ImageOps
        im = ImageOps.exif_transpose(Image.open(src)).convert("RGB")
        im.save(dst, "JPEG", quality=92)
        return True
    except Exception as e:
        print(f"  ✗ {src.name}: {type(e).__name__}")
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--目录", required=True)
    ap.add_argument("--删原图", action="store_true", help="转换成功后删掉原文件")
    args = ap.parse_args()

    base = Path(args.目录).expanduser()
    if not base.is_absolute() and not base.is_dir():
        # 相对路径默认按工作区算，这样「照片/五1班/第五讲」不管从哪儿执行都对
        try:
            from _ws import workspace_root
            _ws = workspace_root()
            if (_ws / base).is_dir():
                base = _ws / base
        except SystemExit:
            pass
    base = base.resolve()
    if not base.is_dir():
        print(f"[错] 没有这个目录：{base}")
        print("     相对路径是相对**工作区**算的（含「班级名单」那一层）。"
              "也可以直接给绝对路径。")
        return 2
    files = [f for f in base.rglob("*")
             if f.is_file() and f.suffix.lower() in (HEIC | OTHER)]
    if not files:
        print("没有需要转换的文件（已经是 jpg 了）")
        return 0

    ok, bad = 0, []
    for f in files:
        dst = f.with_suffix(".jpg")
        if dst.exists() and not args.删原图:
            print(f"  · {f.name} 已存在同名 jpg，跳过")
            continue
        if convert_one(f, dst):
            ok += 1
            print(f"  ✓ {f.name} -> {dst.name}")
            if args.删原图:
                trash = base / "原图_HEIC"
                trash.mkdir(exist_ok=True)
                shutil.move(str(f), str(trash / f.name))
        else:
            bad.append(f.name)

    print(f"\n完成 {ok} 个" + (f"，失败 {len(bad)} 个：{'、'.join(bad[:5])}" if bad else ""))
    if bad and not shutil.which("sips"):
        print("提示：Windows 读不了 HEIC，需要装一个解码包。"
              "用技能自带的 Python 装，别用系统的：")
        print(f'     "{sys.executable}" -m pip install pillow-heif')
        print("     装不上也不卡流程 —— 让老师把照片导出成 jpg 再发一次就行。")
    return 0 if not bad else 1


if __name__ == "__main__":
    sys.exit(main() or 0)
