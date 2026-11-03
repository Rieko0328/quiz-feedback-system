# -*- coding: utf-8 -*-
r"""
按老师给的模板名单，把话术填进 Excel，并把照片按学生名字命名。

模板：教师精灵导出的 xlsx（用 --模板 指定路径）
  A列=学员号  B列=学生姓名  C列=文本1（要填的话术）

用法：
  python 填模板.py --试运行     # 只打印，不写文件
  python 填模板.py              # 写出新文件
"""
import argparse
import os
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

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
# 工作区：默认是脚本往上三级；也可以 --工作区 指定，或设环境变量 KCFK_ROOT
sys.path.insert(0, str(Path(__file__).resolve().parent))
from _ws import workspace_root  # noqa: E402
from roster import name_id_map  # noqa: E402
from fonts import docx_font_name  # noqa: E402

ROOT = workspace_root()

DIR_ROSTER = ROOT / "班级名单"
DIR_OUT = ROOT / "生成结果"

# 模板路径：必须由 --模板 传入，不再写死某个人的下载目录
TPL = None

# 模板写法 -> 我的名单写法（以模板为准，做映射）
# 模板写法 -> 话术里的写法（仅当两者不同时才需要）
# 名字已在名单里统一成模板写法，这里通常为空，别写死某个班的真名。
ALIAS: dict[str, str] = {}


def load_feedback(cls: str | None = None):
    """从 生成结果 读话术，返回 {姓名: 话术}

    文件格式（##### 把标题夹在中间，正文紧随其后）：
        ##### 王小红 | 36.5分 | 优秀 #####
        王小红家长好，跟您说下孩子这次课测的情况。……
        ------------------------------------------------------------
        ##### 下一位 … #####

    cls：只读这个班的话术（多个班都跑过的时候必须指定，否则会混）
    """
    out = {}
    dirs = [d for d in sorted(DIR_OUT.glob("*")) if d.is_dir()]
    if cls:
        dirs = [d for d in dirs if cls in d.name]
    # 兜底：指定了班级却一个都没匹配上，就别乱读别的班
    for d in dirs:
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


def slim(text: str, n: int) -> str:
    """把一段话压到 n 字以内，但**不许把「附件/小任务」那句裁掉**。

    做法：先按句切，保留第一句 + 第一句提到附件/任务的 + 最后一句；
    还是超长才硬裁。盲目从头截断会把「我整理了一份小任务（附件）」裁没，
    家长就不知道有附件了。
    """
    t = text.strip()
    if n <= 0 or len(t) <= n:
        return t
    sents = [s.strip() for s in re.split(r"(?<=[。！？\n])", t) if s.strip()]
    if len(sents) <= 1:
        return t[:n].rstrip() + "…"
    keep = {0, len(sents) - 1}
    for i, s in enumerate(sents[1:-1], 1):
        if re.search(r"附件|小任务|任务|练习单|背诵单|发我|发您", s):
            keep.add(i)
            break
    out = "".join(sents[i] for i in sorted(keep))
    if len(out) > n:
        out = out[:n].rstrip() + "…"
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--模板", default=None, help="教师精灵导出的 xlsx 路径")
    ap.add_argument("--班级", default=None, help="班级名（用来定位照片目录）")
    ap.add_argument("--工作区", default=None, help="数据目录，默认脚本往上三级")
    ap.add_argument("--课次", default=None, help="课次名（定位照片目录 / 决定归档子目录）")
    ap.add_argument("--输出", default=None, help="输出文件路径，默认放 群发/<班级>/<课次>/群发表.xlsx")
    ap.add_argument("--精简", type=int, default=0,
                    help="每段压到 N 字以内（0=不压）。群发表里太长家长不读，建议 100-120")
    ap.add_argument("--试运行", action="store_true")
    args = ap.parse_args()

    # --工作区 / 环境变量 KCFK_ROOT 可覆盖数据目录
    global ROOT, DIR_ROSTER, DIR_OUT
    if args.工作区:
        ROOT = Path(args.工作区).resolve()
        DIR_ROSTER = ROOT / "班级名单"
        DIR_OUT = ROOT / "生成结果"
    if not args.模板:
        print("[提示] 请用 --模板 指定教师精灵导出的 xlsx 路径")
        print("       例如：python 填模板.py --模板 \"D:\教师精灵名单.xlsx\"")
        return 1
    tpl = Path(args.模板)
    if not tpl.exists():
        print(f"找不到模板：{tpl}")
        return 1

    feedback = load_feedback(args.班级)
    print(f"已读取话术：{len(feedback)} 人"
          + (f"（班级：{args.班级}）" if args.班级 else "") + "\n")
    if not feedback:
        print("[提示] 没找到话术文件。先跑 一键生成.py 出话术，再来填模板。")
        return 1

    wb = openpyxl.load_workbook(tpl)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]

    # 定位表头：同一行里既有姓名列，又有「文本N / 文案N」列
    NAME_KEYS = ("学生姓名", "学员姓名", "姓名")
    CN_NUM = "一二三四五六七八九十"

    def _order(v):
        m = re.search(r"(\d+)\s*$", v)
        if m:
            return int(m.group(1))
        m = re.search(r"([%s])\s*$" % CN_NUM, v)
        if m:
            return CN_NUM.index(m.group(1)) + 1
        return 99

    hdr, col_name, text_cols, col_id = None, None, [], None
    for r in range(1, min(ws.max_row, 12) + 1):
        vals = [(c, str(ws.cell(row=r, column=c).value or "").strip())
                for c in range(1, ws.max_column + 1)]
        nm = [c for c, v in vals if v in NAME_KEYS]
        tx = [(c, v) for c, v in vals
              if re.fullmatch(r"\*?\s*(文本|文案)\s*[0-9%s]" % CN_NUM, v)]
        if nm and tx:
            hdr, col_name = r, nm[0]
            text_cols = [c for c, _ in sorted(tx, key=lambda x: _order(x[1]))]
            ids = [c for c, v in vals if v in ("学员号", "学号", "学生编号")]
            col_id = ids[0] if ids else None
            break
    if hdr is None:
        print("找不到表头：需要同一行里有「学生姓名」和「文本1/文案一」这样的列")
        return 1
    print(f"表头在第 {hdr} 行，姓名列 {col_name}，文本列 {text_cols}")
    if not text_cols:
        return 1

    id2name = {}
    try:
        rp = next((f for f in DIR_ROSTER.glob("*.txt")
                   if args.班级 and args.班级 in f.stem), None)
        if rp:
            id2name = {v: k for k, v in name_id_map(rp).items() if v}
    except Exception as e:
        print(f"  [提示] 学号映射读取失败（改用姓名匹配）：{e}")

    filled, missing, matched = 0, [], set()
    for r in range(hdr + 1, ws.max_row + 1):
        tpl_name = ws.cell(row=r, column=col_name).value
        if not tpl_name:
            continue
        tpl_name = str(tpl_name).strip()
        sid = str(ws.cell(row=r, column=col_id).value or "").strip() if col_id else ""
        key = ALIAS.get(tpl_name, tpl_name)
        # 学号优先：模板写得再乱，只要学员号对得上就能填对
        if sid and sid in id2name:
            key = id2name[sid]
            if not tpl_name:
                tpl_name = key
        text = feedback.get(key)
        if not text:
            # 兼容名单里带排序前缀的学生（如"A李小明"），话术的 key 无前缀
            stripped = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", key)
            if stripped != key:
                text = feedback.get(stripped)
        if text:
            matched.add(key)
            matched.add(re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", key))
            # 正文里若出现旧写法，替换成模板里的写法（家长看到的必须是模板名）
            if key != tpl_name:
                text = text.replace(key, tpl_name)
            # 话术正文不该出现排序前缀（家长看到 A李小明 会奇怪）
            text = re.sub(r"^([A-Za-z])(?=[\u4e00-\u9fff])", "", text)
            # 拆成多个对话框：
            #   优先用 txt 里的 ---SPLIT---（由 <<<SPLITn>>> 标记转换而来）
            parts = [p.strip() for p in text.split("---SPLIT---") if p.strip()]
            if args.精简:
                parts = [slim(p, args.精简) for p in parts]
            for pi, part in enumerate(parts):
                if pi >= len(text_cols):
                    print(f"    [注意] {tpl_name} 有 {len(parts)} 段，表格只有 "
                          f"{len(text_cols)} 个文本列，多出的没地方放")
                    break
                cell = ws.cell(row=r, column=text_cols[pi], value=part)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.font = Font(name=docx_font_name(), size=10)
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

    if args.输出:
        out = Path(args.输出).expanduser()
    else:
        _cls = args.班级 or ""
        _lesson = args.课次 or ""
        if not _lesson:
            try:
                rp2 = next((f for f in DIR_ROSTER.glob("*.txt")
                            if _cls and _cls in f.stem), None)
                if rp2:
                    for _ln in rp2.read_text(encoding="utf-8-sig").splitlines():
                        if _ln.strip().startswith(("@课次", "@課次")):
                            _lesson = _ln.split(":", 1)[-1].strip()
                            break
            except Exception:
                pass
        if _cls and _lesson:
            out = ROOT / "群发" / _cls / _lesson / "群发表.xlsx"
        else:
            out = tpl.with_name(tpl.stem + "_已填话术.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    if args.试运行:
        print(f"\n[试运行] 未写文件。将保存为：{out}")
    else:
        try:
            wb.save(out)
        except PermissionError:
            # 文件被占用（多半是正开着 Excel）→ 换个带时间戳的名字
            import datetime as _d
            stamp = _d.datetime.now().strftime("%H%M%S")
            out = out.parent / f"群发表_已填话术_{stamp}.xlsx"
            out.parent.mkdir(parents=True, exist_ok=True)
            wb.save(out)
            print("（原文件被占用，已另存到群发目录，文件名带时间戳）")
        print(f"\n已保存：{out}")

    print(f"\n填写 {filled} 人" + (f"，缺 {len(missing)} 人" if missing else ""))
    if missing:
        print("以下学员在模板里有、但我没有话术（多半是没有答题卡）：")
        for r, n in missing:
            print(f"  行{r} {n}")

    # 反向也要报：有话术、模板里却没有这个人 —— 他一个字都收不到，
    # 而日志如果只说「缺 N 人」，看着像正常。这种情况最常见的原因是
    # 模板里姓名写法不同、或者模板是旧的没加这个新生。
    orphan = [k for k in feedback if k not in matched]
    if orphan:
        print(f"\n[注意] 这 {len(orphan)} 个孩子有话术，但模板里找不到对应行，"
              f"他们收不到任何消息：")
        for k in orphan:
            print(f"  {k}")
        print("      多半是模板姓名写法不同（全角空格、多一个学号），"
              "或者模板还没加这个学生。核对一下模板再重跑。")


if __name__ == "__main__":
    main()
