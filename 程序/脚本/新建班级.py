# -*- coding: utf-8 -*-
r"""新建班级.py —— 一个班一个文本文件，不用装软件不用注册。

    python 新建班级.py                          # 交互式，一步步问
    python 新建班级.py --班级 周六10点 --课次 五年级第五讲 --满分 40
    python 新建班级.py --列出                    # 看现有班级
    python 新建班级.py --删 --班级 测试班         # 删掉一个班
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
DIR_PHOTO = ROOT / "答题卡照片"

# 常见卷子的大题结构模板
TEMPLATES = {
    "1": ("五年级第五讲（过去式/分词 单选 填空 改写 阅读）",
          "过去式/过去分词,单项选择,用适当形式填空,按要求改写句子,阅读理解",
          "5,15,5,5,5"),
    "2": ("四年级 NCE1（名词复数 单选 选词 改写 阅读）",
          "名词复数,单项选择,选词填空,按要求改写句子,阅读理解",
          "5,20,5,5,5"),
    "3": ("四年级 L9-14（单选 选词 改写 阅读）",
          "单项选择,选词填空,按要求改写句子,阅读理解",
          "20,5,5,5"),
    "4": ("自定义", "", ""),
}


def ask(prompt, default=""):
    d = f"（默认 {default}）" if default else ""
    v = input(f"{prompt}{d}：").strip()
    return v or default


def build_text(班级, 课次, 满分, 考试名称, 称呼, 风格, 大题名称, 大题题数,
               重点线, 拍照说明=""):
    lines = [
        f"@班级: {班级}",
        f"@课次: {课次}",
        f"@满分: {满分}",
        f"@重点线: {重点线}",
        f"@考试名称: {考试名称}",
        f"@称呼: {称呼}",
    ]
    if 风格 and 风格 != "亲和":
        lines.append(f"@家长风格: {风格}")
    lines.append(f"@大题名称: {大题名称}")
    lines.append(f"@大题题数: {大题题数}")
    lines += [
        "# ============================================================",
        f"# {班级}　{课次}",
        "#",
        "# 【怎么用】在每个人名字后面写上他的分数，例如：",
        "#     张三 32",
        "#     李四 27",
        "#     王五          ← 只写名字不写分 = 缺考，自动跳过",
        "#",
        "# 【照片放哪】",
        f"#     3-照片\\{班级}\\{课次}\\张三.jpg",
        "#     照片名用学生姓名最省事（系统就不必认名字了）",
        "#",
        "# 【其他可选设置】写在最上面，需要哪个加哪个：",
        "#     @特殊学生: 张三          危机感版话术",
        "#     @安心学生: 张三          不焦虑版话术",
        "#     @严正学生: 张三          认真谈一次（问题点透但不伤人）",
        "#     @表扬学生: 张三,李四      额外点名表扬",
        "#     @同家长: 张三,张三妹      双胞胎合并成一条话术",
        "#     @家长风格: 安抚           全班都走安抚版",
        "#     @任务模式: 个人小任务      每人一份独立小任务（四年级常用）",
        "#     @知识点库: L11-18         挂知识点库，话术会说具体考点",
        "# ============================================================",
        "",
    ]
    return "\n".join(lines)


def interactive():
    print("=" * 60)
    print("  新建班级")
    print("=" * 60)
    print()
    print("  这个班叫什么？（就是你在机构里怎么称呼这个班）")
    print("  例如：周六10点 / 周五慢班 / 四年级")
    print()
    班级 = ask("班级名")
    if not 班级:
        print("没填班级名，退出。")
        return 1
    rp = DIR_ROSTER / f"{班级}.txt"
    if rp.exists():
        print(f"\n[注意] {班级}.txt 已经存在了。")
        if ask("要覆盖吗？输入 y 覆盖，其他任意键取消", "n").lower() != "y":
            return 0

    课次 = ask("这次考的是什么", "五年级第五讲")
    print()
    print("  这份卷子有几道大题？选一个最接近的：")
    for k, (name, _, _) in TEMPLATES.items():
        print(f"    [{k}] {name}")
    print()
    pick = ask("选", "1")
    if pick not in TEMPLATES:
        pick = "1"
    tname, 大题名称, 大题题数 = TEMPLATES[pick]
    if not 大题名称:
        print("  请按顺序填大题名称，用逗号隔开（例如：单项选择,选词填空,按要求改写句子）")
        大题名称 = ask("大题名称", "单项选择,选词填空,按要求改写句子,阅读理解")
        print("  每个大题有几道题？按同样顺序填数字")
        大题题数 = ask("大题题数", "20,5,5,5")

    满分 = ask("满分", "40")
    try:
        总分 = float(满分)
    except ValueError:
        总分 = 40.0
    重点线 = ask("重点线（低于这条线出练习单，高于的出背诵单）",
                 f"{round(总分 * 0.75)}")
    考试名称 = ask("卷子叫什么（会写进话术里）", "月考试卷")
    称呼 = ask("话术里怎么称呼家长", "家长")
    print()
    print("  话术风格：[1] 亲和（默认）  [2] 安抚（全班不制造焦虑，适合基础弱的班）")
    风格 = "安抚" if ask("选", "1") == "2" else "亲和"

    text = build_text(班级, 课次, 满分, 考试名称, 称呼, 风格, 大题名称, 大题题数, 重点线)
    DIR_ROSTER.mkdir(parents=True, exist_ok=True)
    rp.write_text(text, encoding="utf-8")
    (DIR_PHOTO / 班级 / 课次).mkdir(parents=True, exist_ok=True)

    print()
    print("=" * 60)
    print(f"  ✔ 建好了：{rp}")
    print(f"  ✔ 照片目录：{DIR_PHOTO / 班级 / 课次}")
    print()
    print("  下一步：")
    print(f"   1) 把答题卡照片放进  3-照片\\{班级}\\{课次}\\")
    print(f"      照片名改成学生姓名，例如 张三.jpg")
    print(f"   2) 打开 1-名单\\{班级}.txt，在每人名字后面写分数")
    print(f"      或者在网页版里点「识别答题卡」自动读分数")
    print("=" * 60)
    return 0


def main():
    ap = argparse.ArgumentParser(description="新建/管理班级")
    ap.add_argument("--班级")
    ap.add_argument("--课次")
    ap.add_argument("--满分", default="40")
    ap.add_argument("--考试名称", default="月考试卷")
    ap.add_argument("--称呼", default="家长")
    ap.add_argument("--风格", default="亲和")
    ap.add_argument("--大题名称", default="单项选择,选词填空,按要求改写句子,阅读理解")
    ap.add_argument("--大题题数", default="20,5,5,5")
    ap.add_argument("--重点线", default=None)
    ap.add_argument("--列出", action="store_true")
    ap.add_argument("--删", action="store_true")
    args = ap.parse_args()

    if args.列出:
        print(f"{'班级':14s} {'课次':24s} {'人数':>4s} {'已填分':>6s}")
        print("-" * 56)
        for f in sorted(DIR_ROSTER.glob("*.txt")):
            if f.name.endswith(".bak") or f.stem == "姓名对照":
                continue
            meta, n, sc = {}, 0, 0
            for raw in f.read_text(encoding="utf-8-sig").splitlines():
                s = raw.strip()
                if not s or s.startswith("#"):
                    continue
                if s.startswith("@"):
                    if ":" in s:
                        k, v = s[1:].split(":", 1)
                        meta[k.strip()] = v.strip()
                    continue
                n += 1
                if len(s.split()) > 1:
                    sc += 1
            if not meta.get("班级"):
                continue
            print(f"{meta['班级']:14s} {meta.get('课次',''):24s} {n:4d} {sc:6d}")
        return 0

    if args.删:
        if not args.班级:
            print("请用 --班级 指定要删的班")
            return 1
        rp = DIR_ROSTER / f"{args.班级}.txt"
        if rp.exists():
            rp.rename(rp.with_suffix(".txt.bak"))
            print(f"已停用（改名为 {rp.name}.bak）")
        else:
            print("没找到这个班")
        return 0

    if not args.班级:
        return interactive()

    try:
        total = float(args.满分)
    except ValueError:
        total = 40.0
    focus = args.重点线 or str(round(total * 0.75))
    text = build_text(args.班级, args.课次 or "未命名课次", args.满分,
                      args.考试名称, args.称呼, args.风格, args.大题名称,
                      args.大题题数, focus)
    rp = DIR_ROSTER / f"{args.班级}.txt"
    rp.write_text(text, encoding="utf-8")
    if args.课次:
        (DIR_PHOTO / args.班级 / args.课次).mkdir(parents=True, exist_ok=True)
    print(f"已建：{rp}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
