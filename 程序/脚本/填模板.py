# -*- coding: utf-8 -*-
r"""
按老师给的模板名单，把话术填进 Excel，并把照片按学生名字命名。

模板：C:\Users\77290\Downloads\教师精灵_个性化文本_学员名单_*.xlsx
  A列=学员号  B列=学生姓名  C列=文本1（要填的话术）

用法：
  python 填模板.py --试运行     # 只打印，不写文件
  python 填模板.py              # 写出新文件
"""
import argparse
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_PHOTO = ROOT / "答题卡照片" / "周六10点" / "五年级第五讲"
DIR_OUT = ROOT / "生成结果"

TPL = Path(r"C:\Users\77290\Downloads\教师精灵_个性化文本_学员名单_ca2013f8-d2e3-4389-926d-965a31bbb490.xlsx")

# 模板写法 -> 我的名单写法（以模板为准，做映射）
# 模板写法 -> 话术里的写法（仅当两者不同时才需要）
# 名字已在名单里统一成模板写法，这里通常为空。
ALIAS = {
    "谢西文": "谢西文",
}


def load_feedback():
    """从 4-输出 读话术，返回 {姓名: 话术}

    文件格式（##### 把标题夹在中间，正文紧随其后）：
        ##### 王韵茜 | 36.5分 | 优秀 #####
        王韵茜家长好，跟您说下孩子这次课测的情况。……
        ------------------------------------------------------------
        ##### 下一位 … #####
    """
    out = {}
    for d in sorted(DIR_OUT.glob("*")):
        f = d / "家长反馈话术.txt"
        if not f.exists():
            continue
        txt = f.read_text(encoding="utf-8")
        parts = txt.split("#####")
        # parts = [文件头, 标题1, 正文1, 标题2, 正文2, ...]
        for i in range(1, len(parts) - 1, 2):
            head = parts[i].strip()
            body_raw = parts[i + 1]
            m = re.match(r"^(\S+?)\s*\|", head)
            if not m:
                continue
            name = m.group(1).strip()
            body_lines = []
            for ln in body_raw.split("\n"):
                st = ln.strip()
                if st and set(st) <= {"-"}:
                    break
                # 人工阅读标记行跳过；<<<SPLITn>>> 转成统一分隔符保留
                if re.match(r"^—— 消息\d+ ——$", st):
                    continue
                if re.match(r"^<<<SPLIT\d+>>>$", st):
                    body_lines.append("---SPLIT---")
                    continue
                body_lines.append(ln)
            body = "\n".join(body_lines).strip()
            if name and body:
                out[name] = body
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--试运行", action="store_true")
    ap.add_argument("--模板", default=str(TPL))
    args = ap.parse_args()

    tpl = Path(args.模板)
    if not tpl.exists():
        print(f"找不到模板：{tpl}")
        return

    feedback = load_feedback()
    print(f"已读取话术：{len(feedback)} 人\n")

    wb = openpyxl.load_workbook(tpl)
    ws = wb["Sheet1"]

    # 定位表头行
    hdr = None
    for r in range(1, 6):
        if ws.cell(row=r, column=1).value == "学员号":
            hdr = r
            break
    if hdr is None:
        print("找不到表头（学员号）")
        return
    col_name, col_text = 2, 3

    filled, missing = 0, []
    for r in range(hdr + 1, ws.max_row + 1):
        tpl_name = ws.cell(row=r, column=col_name).value
        if not tpl_name:
            continue
        tpl_name = str(tpl_name).strip()
        key = ALIAS.get(tpl_name, tpl_name)
        text = feedback.get(key)
        if not text:
            # 兼容名单里带排序前缀的学生（如"A李怀森"），话术的 key 无前缀
            stripped = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", key)
            if stripped != key:
                text = feedback.get(stripped)
        if text:
            # 正文里若出现旧写法，替换成模板里的写法（家长看到的必须是模板名）
            if key != tpl_name:
                text = text.replace(key, tpl_name)
            # 话术正文不该出现排序前缀（家长看到 A李怀森 会奇怪）
            text = re.sub(r"^([A-Za-z])(?=[\u4e00-\u9fff])", "", text)
            # 拆成多个对话框：
            #   优先用 txt 里的 ---SPLIT---（由 <<<SPLITn>>> 标记转换而来）
            parts = [p.strip() for p in text.split("---SPLIT---") if p.strip()]
            for pi, part in enumerate(parts):
                col = col_text + pi          # 文本1 / 文本2 / 文本3
                cell = ws.cell(row=r, column=col, value=part)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = Font(name="微软雅黑", size=10)
            filled += 1
            print(f"  ✓ 行{r} {tpl_name}  （{len(parts)} 条消息）"
                  + (f"（映射自 {key}）" if key != tpl_name else ""))
        else:
            missing.append((r, tpl_name))
            print(f"  ✗ 行{r} {tpl_name}  —— 没有对应话术")

    # 列宽
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 12
    for col in ("C", "D", "E"):
        ws.column_dimensions[col].width = 62

    out = tpl.with_name(tpl.stem + "_已填话术.xlsx")
    if args.试运行:
        print(f"\n[试运行] 未写文件。将保存为：{out}")
    else:
        try:
            wb.save(out)
        except PermissionError:
            # 文件被占用（多半是正开着 Excel）→ 换个带时间戳的名字
            import datetime as _d
            stamp = _d.datetime.now().strftime("%H%M%S")
            out = tpl.with_name(f"{tpl.stem}_已填话术_{stamp}.xlsx")
            wb.save(out)
            print("（原文件被占用，已另存为新文件）")
        print(f"\n已保存：{out}")

    print(f"\n填写 {filled} 人" + (f"，缺 {len(missing)} 人" if missing else ""))
    if missing:
        print("以下学员在模板里有、但我没有话术（多半是没有答题卡）：")
        for r, n in missing:
            print(f"  行{r} {n}")


if __name__ == "__main__":
    main()
