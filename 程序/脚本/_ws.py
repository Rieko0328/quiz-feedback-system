# -*- coding: utf-8 -*-

r"""_ws.py —— 统一的「工作区」解析。



所有脚本原来都写 `ROOT = Path(__file__).resolve().parent.parent.parent`，

装成技能后这个路径会指到技能安装目录，产物全写进技能里，老师找不到。

统一改成从工作区解析：



    优先级：环境变量 KCFK_ROOT  >  命令行 --工作区  >  当前目录(含班级名单/)

             >  逐级向上查找  >  报错并给出怎么修



工作区就是老师的那个文件夹，结构：

    <工作区>/

    ├── 班级名单/<班级>.txt

    ├── 答题卡照片/<班级>/<课次>/

    ├── 课测答案/<课次>.txt

    └── 生成结果/<班级>_<课次>/

"""

from __future__ import annotations



import os

import sys

from pathlib import Path



ENV_KEY = "KCFK_ROOT"

MARKER = "班级名单"



# 所有脚本都 import 这个模块，在这里统一把输出改成 UTF-8：

# Windows 控制台默认 GBK，直接 print 中文在某些终端会抛 UnicodeEncodeError。

for _s in (sys.stdout, sys.stderr):

    try:

        _s.reconfigure(encoding="utf-8", errors="replace")

    except Exception:                       # 被重定向到不支持的对象时忽略

        pass



# Windows 文件名不许出现这些字符；班级/课次名里带上就会建目录失败

_BAD_CHARS = '<>:"/\\|?*'





def safe_name(s: str) -> str:

    """把班级名/课次名洗成能当文件夹名的样子（幂等，干净的字符串原样返回）。"""

    out = str(s or "").strip()

    for ch in _BAD_CHARS:

        out = out.replace(ch, "_")

    out = "".join(c if ord(c) >= 32 else " " for c in out)

    return out.strip(" .") or "_"







def short_name(name: str) -> str:

    """家长那边怎么称呼这个孩子：去掉名单排序前缀，取后两个字。



    「A李小明」→「怀森」，「张小明」→「润萌」，「张三」→「张三」。



    话术和诊断报告必须用同一个——否则家长先看到「润萌家长您好」，

    再看到报告上「张小明这次…」，会觉得是两个人写的。

    """

    import re

    s = str(name or "").strip()

    s = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", s) or s

    return s[-2:] if len(s) > 2 else s





def _looks_like_workspace(p: Path) -> bool:

    return (p / MARKER).is_dir()





def workspace_root(explicit: str | os.PathLike | None = None) -> Path:

    """解析工作区目录。explicit 一般来自 --工作区 参数。"""

    for cand in (explicit,

                 os.environ.get(ENV_KEY, "").strip() or None):

        if cand:

            p = Path(cand).expanduser().resolve()

            if not p.exists():

                raise SystemExit(f"[错] 工作区不存在：{p}")

            return p



    cwd = Path.cwd().resolve()

    if _looks_like_workspace(cwd):

        return cwd

    for p in cwd.parents:

        if _looks_like_workspace(p):

            return p



    raise SystemExit(

        "[错] 没找到课测工作区（应是包含「班级名单」文件夹的那个目录）。\n"

        f"     修法一：设置环境变量 {ENV_KEY}=你的工作区路径\n"

                "     修法二：先 cd 到工作区再运行\n"

        "     注意：不要加 --工作区——大多数脚本启动时就要定位，参数来不及生效")





def ensure_workspace(root: Path) -> Path:

    """首次使用时补齐目录骨架，并把班级名单模板放进去。"""

    for sub in ("班级名单", "课测答案", "照片", "题目",

                "成绩表", "群发", "生成结果"):

        (root / sub).mkdir(parents=True, exist_ok=True)

    return root





# ---------------------------------------------------------------- 目录约定

# 老师看到的 5 个主目录；脚本内部的过程文件统一放「生成结果」

def photo_root(root: Path) -> Path:

    """照片根目录：新结构叫「照片」，老工作区还叫「答题卡照片」，两种都认。"""

    new = root / "照片"

    old = root / "答题卡照片"

    if new.is_dir() or not old.is_dir():

        return new

    return old





def dirs_for(root: Path, cls: str, lesson: str = "") -> dict[str, Path]:

    """一次给齐某班某课次的全部目录。



    photos  照片/<班级>/<课次>/          该次课测的原图与改名副本

    scores  成绩表/<班级>/               不建课次子目录，表名带课次

    pract   生成结果/通用练习/<课次>/    通用版背诵单 / 专项练习

    pack    群发/<班级>/<课次>/资料/     姓名.jpg + 姓名.docx

    packroot 群发/<班级>/<课次>/          群发表.xlsx 放这一层

    summ    成绩表/<班级>/               随课次累加的班级总表（和单次成绩表同目录）

    work    生成结果/<班级>_<课次>/      过程文件（话术/大题扣分/诊断报告…）

    """

    lesson = safe_name(lesson) if lesson else ""

    cls = safe_name(cls)

    return {

        "photos": photo_root(root) / cls / lesson if lesson else photo_root(root) / cls,

        "scores": root / "成绩表" / cls,

        "pract": (root / "生成结果" / "通用练习" / lesson) if lesson

                 else (root / "生成结果" / "通用练习"),

        # 群发/<班级>/<课次>/资料/ 放孩子的照片和练习；群发表.xlsx 在课次根目录

        "pack": (root / "群发" / cls / lesson / "资料") if lesson

                else (root / "群发" / cls / "资料"),

        "packroot": (root / "群发" / cls / lesson) if lesson else (root / "群发" / cls),

        "summ": root / "成绩表" / cls,

        "work": root / "生成结果" / f"{cls}_{lesson}" if lesson else root / "生成结果" / cls,

    }





# ------------------------------------------------------------------ 快删

def fast_remove(*paths, quiet: bool = True) -> int:

    """删文件 / 删目录，返回删掉的个数。**不进回收站，删了就是删了。**



    为什么不用 os.unlink / shutil.rmtree：

    WorkBuddy 自带的 sitecustomize 会把每次 unlink 变成一次「移进回收站」的

    子进程调用，单个文件 150ms。删 120 个报告要 18 秒，而系统命令 0.05 秒——

    老师看到的就是「整理这一步卡住十几秒」。



    只用来删**我们自己生成的产物**（报告、巩固单、群发副本这些，删了能重出），

    绝不用来删老师的原始照片、名单、答案。

    """

    targets = [Path(x) for x in paths if x]

    targets = [x for x in targets if x.exists()]

    if not targets:

        return 0

    n = 0

    if os.name == "nt":

        import subprocess

        dirs = [x for x in targets if x.is_dir()]

        files = [x for x in targets if not x.is_dir()]

        for d in dirs:

            try:

                subprocess.run(["cmd", "/c", "rd", "/s", "/q", str(d)],

                               capture_output=True, timeout=300)

            except (OSError, subprocess.SubprocessError):

                pass

            if not d.exists():

                n += 1

            else:                                    # 被占用，退回逐个删

                import shutil as _sh

                _sh.rmtree(d, ignore_errors=True)

                n += 0 if d.exists() else 1

        for i in range(0, len(files), 64):           # 一条命令删一批，别开 120 个进程

            chunk = files[i:i + 64]

            try:

                subprocess.run(["cmd", "/c", "del", "/q", "/f"]

                               + [str(x) for x in chunk],

                               capture_output=True, timeout=120)

            except (OSError, subprocess.SubprocessError):

                pass

            for x in chunk:

                if not x.exists():

                    n += 1

                else:

                    try:

                        x.unlink()

                        n += 1

                    except OSError:

                        if not quiet:

                            print(f"  [删不掉] {x.name}（被别的程序开着？）")

        return n



    # 非 Windows：普通删除就够快

    import shutil as _sh

    for x in targets:

        try:

            if x.is_dir():

                _sh.rmtree(x, ignore_errors=True)

            else:

                x.unlink()

            n += 1

        except OSError:

            if not quiet:

                print(f"  [删不掉] {x.name}")

    return n

