# -*- coding: utf-8 -*-

"""扫描「写死在代码里的学科内容」，防止换个讲次就把上一讲的内容说给家长。



为什么要有：同一类 bug 出过五次 —— for/since、return/have、How long、

「动词的过去式和过去分词」、「过去式/过去分词」默认大题名。每次都是

只改一处漏另一处，人眼靠不住。



用法：

    python 检查硬编码.py          # 列出可疑项

    python 检查硬编码.py --严     # 有可疑项就返回 1（CI / 提交前用）



两个白名单标记（写在源码行上）：

    # hc-ok                     这一行确认过，不是问题

    # hc-demo-start / hc-demo-end   中间是一段演示假数据，不参与扫描

"""

from __future__ import annotations



import argparse

import re

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



SCRIPTS = Path(__file__).resolve().parent



# 出现这些词就认为「这是某个讲次才成立的内容」

PATTERNS = [

    r"过去式", r"过去分词", r"现在完成", r"现在进行", r"一般将来", r"一般过去",

    r"比较级", r"最高级", r"名词复数", r"句型转换", r"what\s*color", r"how\s*long",

    r"how\s*soon", r"since", r"不规则动词",

]



# 这些是通用说法或文件名，不算写死

SAFE = [

    "本讲", "这次", "卷面", "该讲", "上一讲", "本课",

]



STR_RE = re.compile(r"""(?:"([^"\\]*(?:\\.[^"\\]*)*)"|'([^'\\]*(?:\\.[^'\\]*)*)')""")





def scan_file(p: Path):

    src = p.read_text(encoding="utf-8", errors="replace")

    src_lines = src.splitlines()



    # 先抹掉注释和 docstring，剩下的字符串才算数。

    # 抹的时候要**保留换行数**，否则报出来的行号和真实文件对不上

    # （多行 docstring 被整块删掉，后面的行号全往前错位）。

    text = re.sub(r"#.*", "", src)



    def _blank(m):

        return "\n" * m.group(0).count("\n")



    text = re.sub(r'"""[\s\S]*?"""', _blank, text)

    text = re.sub(r"'''[\s\S]*?'''", _blank, text)



    # 白名单：单行 hc-ok、成段 hc-demo

    skip = [False] * (len(src_lines) + 1)

    in_demo = False

    for i, raw in enumerate(src_lines, 1):

        if "hc-demo-start" in raw:

            in_demo = True

        if in_demo or "hc-ok" in raw:

            skip[i] = True

        if "hc-demo-end" in raw:

            in_demo = False



    out = []

    for ln, line in enumerate(text.splitlines(), 1):

        if not line.strip() or (ln < len(skip) and skip[ln]):

            continue

        for m in STR_RE.finditer(line):

            s = next((g for g in m.groups() if g), "")

            if not s or not re.search(r"[\u4e00-\u9fa5]", s):

                continue

            if any(k in s for k in SAFE):

                continue

            hit = [pat for pat in PATTERNS if re.search(pat, s, re.I)]

            if hit:

                out.append((ln, s.strip()[:70], hit[0]))

    return out





def main():

    ap = argparse.ArgumentParser(description="扫写死在代码里的学科内容")

    ap.add_argument("--严", action="store_true", help="有可疑项返回 1")

    args = ap.parse_args()



    files = sorted(p for p in SCRIPTS.glob("*.py")

                   if not p.name.startswith("知识点库")

                   and p.name != Path(__file__).name)

    total = 0

    for p in files:

        items = scan_file(p)

        if not items:

            continue

        print(f"\n{p.name}")

        for ln, s, pat in items:

            print(f"  {ln:>5}  [{pat}]  {s}")

        total += len(items)



    print()

    if total:

        print(f"共 {total} 处可疑。逐个确认：是不是只有某个讲次才成立？")

        print("是的话改成从 名单 @ / 题库.json / 答案文件 里取，别写在代码里。")

        print("确认没问题的那一行，行尾加  # hc-ok；成段的演示假数据用 "

              "# hc-demo-start / # hc-demo-end 包起来。")

        return 1 if args.严 else 0

    print("没发现写死的学科内容 ✔")

    return 0





if __name__ == "__main__":

    sys.exit(main() or 0)

