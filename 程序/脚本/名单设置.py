# -*- coding: utf-8 -*-
r"""名单设置.py —— 改话术版本、双胞胎、称呼这些「班级设置」，不用手改 txt。

老师能选的话术版本（其余按分数和进步自动走，不用管）：

    进步      「进步学生」进步版：强调这段时间做对了什么，别继续念缺点
    退步      「退步学生」退步版：承认落差、一起找原因，不追责
    标准      默认语气，按档位自动说
    亲和      「家长风格: 亲和」低年级/家长年轻，话更软、多鼓励
    安抚      「家长风格: 安抚」全班都走不焦虑版
    特殊      「特殊学生」危机感版：家长不太管、孩子有点飘
    安心      「安心学生」不焦虑版：家长容易紧张
    严正      「严正学生」认真谈一次：问题点透但不伤人

用法：
    python 名单设置.py --班级 周五慢班 --查看
    python 名单设置.py --班级 周五慢班 --风格 安抚
    python 名单设置.py --班级 周五慢班 --风格 标准
    python 名单设置.py --班级 周五慢班 --特殊 张三,李四
    python 名单设置.py --班级 周五慢班 --安心 王五
    python 名单设置.py --班级 周五慢班 --严正 赵六
    python 名单设置.py --班级 周五慢班 --表扬 陈七
    python 名单设置.py --班级 周五慢班 --同家长 "江小红,江小晴"
    python 名单设置.py --班级 周五慢班 --清空 特殊
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _ws import workspace_root  # noqa: E402

# 命令行参数 -> 名单里的 @key
PERSON_KEYS = {
    "进步": "进步学生",
    "退步": "退步学生",
    "特殊": "特殊学生",
    "安心": "安心学生",
    "严正": "严正学生",
    "表扬": "表扬学生",
    "同家长": "同家长",
}
STYLE_KEY = "家长风格"
STYLE_OK = {"标准": "", "亲和": "亲和", "安抚": "安抚"}

# 写回时保持这个顺序，和 ingest 一致
ORDER = ["班级", "课次", "满分", "重点线", "大题名称", "大题题数", "大题分值",
         "称呼", "考试名称", "家长风格", "进步学生", "退步学生", "特殊学生",
         "安心学生", "严正学生",
         "表扬学生", "同家长", "附言", "知识点库", "日期", "学校", "科目"]


def read(path: Path):
    meta, body = {}, []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.rstrip()
        if s[:1] in "@＠" and (":" in s or "：" in s):
            k, v = s[1:].replace("：", ":").split(":", 1)
            meta[k.strip()] = v.strip()
        elif s.strip() and not s.startswith("#"):
            body.append(s)
    return meta, body


def write(path: Path, meta: dict, body: list[str]):
    lines = [f"@{k}: {meta[k]}" for k in ORDER if meta.get(k)]
    lines += [f"@{k}: {v}" for k, v in meta.items()
              if v and k not in ORDER]
    path.write_text("\n".join(lines + [""] + body) + "\n", encoding="utf-8")


def names_of(body: list[str]) -> list[str]:
    out = []
    for ln in body:
        parts = ln.split()
        if len(parts) >= 2:
            out.append(parts[1])
    return out


def main():
    ap = argparse.ArgumentParser(description="改班级的话术设置")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--查看", action="store_true", help="看当前设置")
    ap.add_argument("--附言", default=None,
                    help="给这个班每条话术末尾统一加一句（下次课时间、考试范围等）；"
                         "传空字符串表示清掉")
    ap.add_argument("--风格", choices=list(STYLE_OK), default=None,
                    help="全班语气：标准 / 亲和（低龄）/ 安抚（家长焦虑）")
    ap.add_argument("--进步", default=None,
                    help="进步版：强调「这段时间做对了什么」，逗号分隔")
    ap.add_argument("--退步", default=None,
                    help="退步版：承认落差、一起找原因，逗号分隔")
    ap.add_argument("--特殊", default=None, help="危机感版，逗号分隔")
    ap.add_argument("--安心", default=None, help="不焦虑版，逗号分隔")
    ap.add_argument("--严正", default=None, help="认真谈一次，逗号分隔")
    ap.add_argument("--表扬", default=None, help="额外点名表扬，逗号分隔")
    ap.add_argument("--同家长", default=None,
                    help="双胞胎合并话术，如 \"江小红,江小晴\"（多组用分号隔开）")
    ap.add_argument("--清空", default=None,
                    choices=["进步", "退步", "特殊", "安心", "严正",
                             "表扬", "同家长", "风格"],
                    help="把某一项清掉")
    ap.add_argument("--工作区", default=None)
    args = ap.parse_args()

    root = workspace_root(args.工作区)
    path = root / "班级名单" / f"{args.班级}.txt"
    if not path.exists():
        print(f"[错] 没有这个班的名单：{path}")
        print("     建班用：python 新建班级.py --从表 学员列表.xlsx --班级 " + args.班级)
        return 2

    meta, body = read(path)
    roster = names_of(body)

    # ---- 先把现状摆出来 ----
    print(f"{args.班级}　{meta.get('课次') or '（还没开始）'}　共 {len(roster)} 人")
    st = meta.get(STYLE_KEY) or ""
    print(f"  全班语气：{st or '标准'}{'' if st else '（默认）'}")
    for label, key in (("进步版", "进步学生"), ("退步版", "退步学生"),
                       ("危机感版", "特殊学生"), ("不焦虑版", "安心学生"),
                       ("认真谈", "严正学生"), ("额外表扬", "表扬学生"),
                       ("双胞胎", "同家长")):
        if meta.get(key):
            print(f"  {label}：{meta[key]}")

    if args.查看 or not any([args.风格, args.进步, args.退步, args.特殊,
                            args.安心, args.严正, args.表扬, args.同家长,
                            args.附言 is not None, args.清空]):
        return 0

    changed = []

    def check(raw: str) -> tuple[str, list[str]]:
        """把「张三,李四」拆出来，并核对在不在名单里。"""
        got = [x.strip() for x in re.split(r"[,，、\s]+", raw) if x.strip()]
        bad = [x for x in got if x not in roster]
        return ",".join(got), bad

    if args.清空:
        if args.清空 == "风格":
            meta.pop(STYLE_KEY, None)
            changed.append("全班语气 → 标准")
        else:
            meta.pop(PERSON_KEYS[args.清空], None)
            changed.append(f"{PERSON_KEYS[args.清空]} → 清空")

    if args.附言 is not None:
        v = args.附言.strip()
        if v:
            meta["附言"] = v
            changed.append(f"统一附言 → {v}")
        else:
            meta.pop("附言", None)
            changed.append("统一附言 → 清空")

    if args.风格:
        v = STYLE_OK[args.风格]
        if v:
            meta[STYLE_KEY] = v
            changed.append(f"全班语气 → {args.风格}")
        else:
            meta.pop(STYLE_KEY, None)
            changed.append("全班语气 → 标准")

    for arg_name, label in (("进步", "--进步"), ("退步", "--退步"),
                            ("特殊", "--特殊"), ("安心", "--安心"),
                            ("严正", "--严正"), ("表扬", "--表扬")):
        raw = getattr(args, arg_name)
        if not raw:
            continue
        joined, bad = check(raw)
        if bad:
            print(f"[错] 名单里没有：{'、'.join(bad)}")
            print(f"     这个班的学生是：{'、'.join(roster[:8])}"
                  + ("…" if len(roster) > 8 else ""))
            return 2
        meta[PERSON_KEYS[arg_name]] = joined
        changed.append(f"{PERSON_KEYS[arg_name]} → {joined}")

    if args.同家长:
        groups, bad = [], []
        for grp in re.split(r"[；;]", args.同家长):
            joined, b = check(grp)
            if b:
                bad += b
            elif joined:
                groups.append(joined)
        if bad:
            print(f"[错] 名单里没有：{'、'.join(bad)}")
            return 2
        if groups:
            meta["同家长"] = ";".join(groups)
            changed.append(f"同家长 → {meta['同家长']}")

    if not changed:
        return 0

    write(path, meta, body)
    print("\n改好了：")
    for c in changed:
        print(f"  · {c}")
    print(f"\n名单：{path}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
