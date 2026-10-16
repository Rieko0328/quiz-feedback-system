# -*- coding: utf-8 -*-
r"""课测反馈系统 · 命令行版。由 启动.bat 调用。"""
import os
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
PY = ROOT / "python" / "python.exe"
SCRIPTS = ROOT / "程序" / "脚本"

MENU = """
============================================================
            课测反馈系统
============================================================

  每周就四步：

    [1] 新建班级
    [2] 识别答题卡 · 出核对表       ★ 自动认分数
    [3] 核对完，并入名单（登分）
    [4] 一键生成（话术 / 小任务 / 归档 / 群发）
    [5] 填教师精灵群发表

  ── 其他 ──
    [6] 看我的班
    [7] 打开输出文件夹
    [8] 改识别配置（api_key）
    [0] 退出

  提示：网页版更好用 —— 双击「网页版.bat」
"""


def _env():
    return dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")


def run(args, pause=True):
    print()
    print("-" * 62)
    r = subprocess.run([str(PY)] + args, cwd=str(ROOT), env=_env())
    print("-" * 62)
    print("[完成]" if r.returncode == 0 else f"[出错] 返回码 {r.returncode}")
    if pause:
        input("\n按回车返回菜单...")


def classes():
    return sorted(f.stem for f in (ROOT / "班级名单").glob("*.txt")
                  if not f.name.endswith(".bak") and f.stem != "姓名对照")


def ask_class():
    cs = classes()
    if cs:
        print("  已有班级：" + "、".join(cs))
    return input("请输入班级名：").strip()


def main():
    if not PY.exists():
        print("找不到 python\\python.exe，包可能不完整。")
        input("按回车退出...")
        return
    while True:
        os.system("cls")
        print(MENU)
        ch = input("请输入数字回车：").strip()
        if ch == "0":
            return
        elif ch == "1":
            run([str(SCRIPTS / "新建班级.py")])
        elif ch == "2":
            cls = ask_class()
            if not cls:
                continue
            lesson = input("课次（回车=用名单里的）：").strip()
            args = [str(SCRIPTS / "识别照片.py"), "--班级", cls]
            if lesson:
                args += ["--课次", lesson]
            run(args)
        elif ch == "3":
            cls = ask_class()
            if not cls:
                continue
            tab = input("核对表路径（回车=自动找）：").strip().strip('"')
            args = [str(SCRIPTS / "识别照片.py"), "--班级", cls, "--并入名单"]
            if tab:
                args += ["--核对表", tab]
            run(args)
        elif ch == "4":
            cls = ask_class()
            if not cls:
                continue
            tpl = input("教师精灵 xlsx 路径（回车=跳过）：").strip().strip('"')
            args = [str(SCRIPTS / "一键生成.py"), "--班级", cls]
            if tpl:
                args += ["--模板", tpl]
            run(args)
        elif ch == "5":
            tpl = input("教师精灵 xlsx 路径：").strip().strip('"')
            if tpl:
                run([str(SCRIPTS / "填模板.py"), "--模板", tpl])
        elif ch == "6":
            run([str(SCRIPTS / "新建班级.py"), "--列出"])
        elif ch == "7":
            os.startfile(str(ROOT / "生成结果"))
            input("已打开，按回车返回...")
        elif ch == "8":
            os.startfile(str(ROOT / "识别配置.txt"))
            input("已打开，改完保存，按回车返回...")
        else:
            input("没这个选项，按回车重来...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
