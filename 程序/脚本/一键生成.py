# -*- coding: utf-8 -*-

r"""一键生成.py —— 出成绩之后的所有事。



**固定做**（100% 都要，不给选项）：

    成绩表 + 家长话术 + 汇总报告  →  群发打包  →  整理目录



**可选**（老师这讲要不要，先问再传）：

    2  通用练习（背诵单 / 专项练习，缺考的孩子补）

    3  诊断报告（每人 2 页图，看这次考得怎么样）

    4  巩固练习（每人一份 docx，按他的错题出的题）



    python 一键生成.py --班级 五年级A班                 # 默认连 2,3,4 一起做

    python 一键生成.py --班级 五年级A班 --要 3,4        # 只要报告和巩固单

    python 一键生成.py --班级 五年级A班 --要 ""         # 什么都不要，只出成绩表+群发

    python 一键生成.py --班级 五年级A班 --要 2,3,4 --模板 x.xlsx

"""

from __future__ import annotations



import argparse

import os

import subprocess

import time

import sys

from pathlib import Path



# line_buffering：子进程的输出会先到，父进程的标题要立刻刷出去，

# 否则日志顺序会乱（标题全排到子进程输出后面），AI 读日志会误判阶段

try:

    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

except (AttributeError, ValueError, OSError):

    pass

sys.path.insert(0, str(Path(__file__).resolve().parent))

from _ws import workspace_root  # noqa: E402



ROOT = workspace_root()

os.environ["KCFK_ROOT"] = str(ROOT)   # 子脚本继承同一个工作区

SCRIPTS = Path(__file__).resolve().parent



# 可选的三样（编号 -> 标题 / 脚本 / 额外参数）

OPTIONS = {

    2: ("通用练习（背诵单 / 专项练习，缺考的孩子补）", "分层作业.py", []),

    3: ("诊断报告（每人 2 页图，看这次考得怎么样）", "批量学情.py", ["--只报告"]),

    4: ("巩固练习（每人一份 docx，按他的错题出的题）", "批量学情.py", ["--只巩固单"]),

}

DEFAULT_OPT = "2,3,4"





def parse_want(s: str) -> list[int]:

    out = []

    for tok in str(s or "").replace("，", ",").replace(" ", ",").split(","):

        tok = tok.strip()

        if tok.isdigit() and int(tok) in OPTIONS:

            n = int(tok)

            if n not in out:

                out.append(n)

    return sorted(out)





# 脚本名 -> 模块名。能内联就内联，见 run() 里的说明

_INPROC = {

    "课测反馈.py": "课测反馈",

    "分层作业.py": "分层作业",

    "批量学情.py": "批量学情",

    "群发打包.py": "群发打包",

    "归档.py": "归档",

}





def _run_subprocess(script: str, extra: list[str]) -> int:

    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")

    return subprocess.run([sys.executable, str(SCRIPTS / script)] + extra,

                          cwd=str(ROOT), env=env).returncode





def run(script: str, extra: list[str]) -> int:

    """跑一个步骤。默认在**当前进程里直接调它的 main()**。



    为什么要这样：每起一个 python 子进程，光解释器启动就 0.35 秒，

    再加上 import openpyxl / python-docx / Pillow 又是 0.6-1 秒——

    5 个步骤就是 5 秒的纯开销，老师看到的是「点了之后卡十几秒」。

    同一个进程里这些库只导一次，后面几步直接用，整体能快 2-3 秒。



    真遇到导不进来、或者某一步把状态弄脏了，退回子进程跑，功能不受影响。

    排查问题时可以设 KCFK_FORCE_SUBPROCESS=1 强制走子进程。

    """

    mod_name = _INPROC.get(script)

    if mod_name and not os.environ.get("KCFK_FORCE_SUBPROCESS"):

        try:

            import importlib

            mod = importlib.import_module(mod_name)

            fn = getattr(mod, "main", None)

            if fn is not None:

                old_argv = sys.argv

                sys.argv = [script] + extra

                try:

                    rc = fn()

                    return 0 if rc is None else int(rc)

                except SystemExit as e:               # 脚本自己 sys.exit 的

                    if e.code is None:

                        return 0

                    return e.code if isinstance(e.code, int) else 1

                finally:

                    sys.argv = old_argv

        except (ImportError, AttributeError) as e:

            print(f"      [提示] {script} 改走子进程（{type(e).__name__}）")

    return _run_subprocess(script, extra)





def main():

    ap = argparse.ArgumentParser(description="出成绩之后的所有事")

    ap.add_argument("--班级", required=True)

    ap.add_argument("--要", default=DEFAULT_OPT,

                    help=f"可选的那三样做哪些：2 通用练习 / 3 诊断报告 / 4 巩固练习，"

                         f"默认 {DEFAULT_OPT}；传空字符串就是都不要")

    ap.add_argument("--模板", default=None, help="教师精灵 xlsx，给了就在最后填好")

    ap.add_argument("--不群发", action="store_true",

                    help="这次只出成绩表和留档，不打包群发文件夹")

    args = ap.parse_args()



    want = parse_want(args.要)



    # 组装执行计划：成绩表(必做) → 可选 → 群发(必做) → 整理(必做)

    plan = [("成绩表 + 家长话术 + 汇总报告", "课测反馈.py", [])]

    if 2 in want:

        plan.append(OPTIONS[2])

    if 3 in want and 4 in want:

        # 合成一次跑，省一半时间

        plan.append(("诊断报告 + 巩固练习（每人两份）", "批量学情.py", []))

    else:

        for n in (3, 4):

            if n in want:

                plan.append(OPTIONS[n])

    if not args.不群发:

        plan.append(("群发打包（照片和练习改成孩子名字）", "群发打包.py", []))

    plan.append(("整理成标准目录", "归档.py", []))



    extras = [OPTIONS[n][0].split("（")[0] for n in want]

    print("=" * 64)

    print(f"  一键生成：{args.班级}")

    print(f"  必做：成绩表 / 话术 / 汇总"

          f"{'' if args.不群发 else '　群发打包'}　整理目录")

    print(f"  另加：{'、'.join(extras) if extras else '无（只出成绩表）'}")

    print("=" * 64)



    total = len(plan)

    warn = []

    for i, (label, script, extra) in enumerate(plan, 1):

        print(f"\n[{i}/{total}] {label}")

        print("-" * 64)

        _t0 = time.time()

        rc = run(script, ["--班级", args.班级] + extra)

        _dt = time.time() - _t0

        if rc == 0 and _dt >= 1:

            print(f"      （{_dt:.1f}s）")

        if rc == 0:

            continue

        if rc == 3:

            # 分层作业：没题库，跳过。不是错误，但老师要知道「这次没有通用练习」

            print(f"\n[跳过] {label}：没有本讲题库，这一样没生成。"

                  f"缺考的孩子这次没有通用练习可补——想补就先把题库写进 "

                  f"题目/{args.班级}_<课次>/题库.json")

            warn.append(f"{label}（没题库，跳过）")

            continue

        if rc == 2:

            # 前置条件没满足（还没登分、参数冲突、名单为空）—— 不是程序坏了

            print(f"\n[停] 第 {i} 步没法做：上面那条 [错] 就是原因。")

            print("     这不是程序坏了，把上面说的补上再重跑就行。")

            return 2

        if script == "批量学情.py" and rc == 1:

            # 部分人失败：群发和归档照做，别因为一个人没出报告就整批不出货

            print(f"\n[注意] {label} 有部分失败——群发和归档照常往下做，"

                  f"但家长那边会缺这几个人。要不要先修再看，你定。")

            warn.append(f"{label} 有失败")

            continue

        print(f"\n[中断] {script} 出错了（返回码 {rc}）。")

        print("       上面有报错信息，把它截图给技术同学。")

        return rc



    if args.模板:

        print("\n[附] 填教师精灵群发表")

        print("-" * 64)

        rc = run("填模板.py", ["--模板", args.模板, "--班级", args.班级])

        if rc != 0:

            print(f"\n[注意] 填模板出错（返回码 {rc}），但前面的都已生成。")

    else:

        print("\n[跳过] 填教师精灵群发表（没给 --模板 路径）")



    outdirs = [d for d in (ROOT / "生成结果").glob(f"{args.班级}_*") if d.is_dir()]

    print("\n" + "=" * 64)

    if warn:

        print("  跑完了，但有这几样不完整：")

        for w in warn:

            print(f"    · {w}")

    else:

        print("  全部完成 ✔")

    if outdirs:

        d = outdirs[0]

        print(f"  结果目录：{d}")

        for x in sorted(d.iterdir()):

            mark = "[目录]" if x.is_dir() else "      "

            print(f"    {mark} {x.name}")

    print("=" * 64)

    return 1 if warn else 0





if __name__ == "__main__":

    sys.exit(main() or 0)

