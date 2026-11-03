# -*- coding: utf-8 -*-
r"""fonts.py —— 跨平台中文字体解析。

原来写死 C:\Windows\Fonts，Mac 上一张反馈单都出不来。
现在按「技能自带 → 系统」的顺序找，三个平台都覆盖：

    技能自带  <SKILL>/fonts/*.ttf|otf|ttc      （想完全不依赖系统字体就丢一个进来）
    Windows   msyh.ttc / simhei.ttf
    macOS     PingFang.ttc / Hiragino Sans GB / STHeiti
    Linux     Noto Sans CJK / 文泉驿

找不到就用 Pillow 默认字体（中文会变方块），并打印一次警告。
"""
from __future__ import annotations

import sys
from functools import lru_cache
from pathlib import Path

_HERE = Path(__file__).resolve().parent
_OWN_DIR = _HERE.parent / "fonts"

_WIN = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]
_MAC = [
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
]
_LINUX = [
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
]

_warned = False


@lru_cache(maxsize=4)
def _candidates(bold: bool = False) -> tuple[str, ...]:
    """按优先级列出候选字体路径（结果缓存，别每次都去 glob 磁盘）。"""
    own = []
    if _OWN_DIR.is_dir():
        pats = ["*.ttf", "*.otf", "*.ttc"]
        for pat in pats:
            own += [str(p) for p in sorted(_OWN_DIR.glob(pat))]
    sys_list = _WIN + _MAC + _LINUX
    if bold:
        # 粗体优先找带 bold 字样的，其次退回普通体（合成粗体效果差，宁可用普通体）
        boldish = [p for p in own + sys_list
                   if "bold" in p.lower() or "bd" in Path(p).stem.lower()]
        return tuple(boldish + own + sys_list)
    return tuple(own + sys_list)


def candidates(bold: bool = False) -> list[str]:
    """按优先级列出候选字体路径（对外还是 list，缓存的是 tuple）。"""
    return list(_candidates(bool(bold)))


@lru_cache(maxsize=256)
def pil_font(size: int, bold: bool = False):
    """返回 PIL 字体对象。

    一定要缓存：一份诊断报告要画七八百个字符串，每画一次都去
    ImageFont.truetype() 重新读一遍 TTF 的话，一个学生就多花 60ms，
    40 个学生白等两秒多。字体对象本身是只读的，缓存没有副作用。
    """
    from PIL import ImageFont

    global _warned
    for p in candidates(bold):
        if not Path(p).exists():
            continue
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    if not _warned:
        _warned = True
        print("[警告] 没找到中文字体，报告里的中文会显示成方块。", file=sys.stderr)
        print(f"       临时办法：把任意中文字体（ttf/otf）放一份到 {_OWN_DIR}", file=sys.stderr)
    return ImageFont.load_default()


def docx_font_name() -> str:
    """Word 里用的中文字体名（只影响显示，不影响内容）。"""
    if sys.platform == "darwin":
        return "PingFang SC"
    if sys.platform.startswith("linux"):
        return "Noto Sans CJK SC"
    return "微软雅黑"
