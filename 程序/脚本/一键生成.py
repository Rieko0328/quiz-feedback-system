# -*- coding: utf-8 -*-
r"""一键生成.py —— 把「分数已经填好」之后的所有事一次做完。

    课测反馈（成绩表 + 话术 + 汇总报告）
    → 分层作业（背诵单 / 练习单）
    → 每人小任务
    → 重命名归档（照片和资料统一改成学生姓名）
    → 群发打包（放同一个文件夹，直接拖进企业微信）
    → 填教师精灵群发表（可选，给了模板路径就填）

    python 一键生成.py --班级 周六10点
    python 一键生成.py --班级 周六10点 --模板 "C:\\...\\教师精灵.xlsx"
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent.parent
SCRIPTS = ROOT / "程序" / "脚本"

STEPS = [
    ("生成成绩表 / 家长话术 / 汇总报告", "课测反馈.py"),
    ("生成分层作业（背诵单 / 练习单）", "分层作业.py"),
    ("生成每人一份小任务", "小任务.py"),
    ("重命名归档（照片 + 资料统一用学生姓名）", "分发资料.py"),
    ("群发打包（放同一个文件夹）", "群发打包.py"),
]


def run(script: str, extra: list[str]):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run([sys.executable, str(SCRIPTS / script)] + extra,
                          cwd=str(ROOT), env=env).returncode


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--班级", required=True)
    ap.add_argument("--模板", default=None, help="教师精灵 xlsx，填了就在最后自动填好")
    ap.add_argument("--跳过归档", action="store_true")
    args = ap.parse_args()

    print("=" * 64)
    print(f"  一键生成：{args.班级}")
    print("=" * 64)

    steps = STEPS if not args.跳过归档 else STEPS[:3]
    for i, (label, script) in enumerate(steps, 1):
        print(f"\n[{i}/{len(steps)}] {label}")
        print("-" * 64)
        rc = run(script, ["--班级", args.班级])
        if rc != 0:
            print(f"\n[中断] {script} 出错了（返回码 {rc}）。")
            print("       上面有报错信息，把它截图给技术同学。")
            return rc

    if args.模板:
        print(f"\n[{len(steps)+1}/{len(steps)+1}] 填教师精灵群发表")
        print("-" * 64)
        rc = run("填模板.py", ["--模板", args.模板])
        if rc != 0:
            print(f"\n[注意] 填模板出错（返回码 {rc}），但前面的都已生成。")
    else:
        print("\n[跳过] 填教师精灵群发表（没给 --模板 路径）")

    from pathlib import Path as _P
    outdirs = [d for d in (ROOT / "生成结果").glob(f"{args.班级}_*") if d.is_dir()]
    print("\n" + "=" * 64)
    print("  全部完成 ✔")
    if outdirs:
        d = outdirs[0]
        print(f"  结果目录：{d}")
        for x in sorted(d.iterdir()):
            mark = "📁" if x.is_dir() else "  "
            print(f"    {mark} {x.name}")
    print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
