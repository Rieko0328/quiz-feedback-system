# -*- coding: utf-8 -*-
r"""识别照片.py —— 用视觉大模型读答题卡，生成「核对表」给老师过目。

设计原则：
  · AI 只负责"看"，不负责"写" —— 读出来的是原始数据，不直接生成话术
  · 照片文件名如果已经是学生姓名，就**只读分数**，不读姓名（避免手写名字认错）
  · 全部结果先落到「核对表」，老师扫一眼改掉错的，再并入名单
  · 支持任意 OpenAI 兼容的视觉端点（DeepSeek 官方 flash / 通义 / 智谱 / Kimi / 方舟）

用法：
    python 识别照片.py --班级 周五慢班
    python 识别照片.py --班级 周五慢班 --限 3          # 只读前 3 张，先看效果
    python 识别照片.py --班级 周五慢班 --并入名单       # 把核对表写进名单
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent.parent
DIR_ROSTER = ROOT / "班级名单"
DIR_PHOTO = ROOT / "答题卡照片"
DIR_OUT = ROOT / "生成结果"
CFG = ROOT / "识别配置.txt"

DEFAULT_CFG = """# ============================================================
#  识别配置 —— 填好这里才能用「识别照片」功能
# ============================================================
# 只要是用 OpenAI 兼容格式的视觉大模型，填这三行就行。
#
# 【DeepSeek 官方】—— 实测可用，用 flash（pro 不支持读图！）
api_base = https://api.deepseek.com/chat/completions
api_key  = 在这里填你的 key
api_model = deepseek-flash
#
# 【通义千问】api_base = https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions
#             api_model = qwen-vl-max
# 【智谱】    api_base = https://open.bigmodel.cn/api/paas/v4/chat/completions
#             api_model = glm-4v-plus
# 【Kimi】    api_base = https://api.moonshot.cn/v1/chat/completions
#             api_model = moonshot-v1-8k-vision-preview
# ============================================================
并发数 = 4
图片最长边 = 1400
"""


def load_cfg() -> dict:
    if not CFG.exists():
        CFG.write_text(DEFAULT_CFG, encoding="utf-8")
        print(f"[首次运行] 已生成配置文件：{CFG}")
        print("          请打开它，填上 api_key，然后再运行。")
        sys.exit(0)
    cfg = {}
    for line in CFG.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        cfg[k.strip()] = v.strip()
    if not cfg.get("api_key") or "在这里填" in cfg.get("api_key", ""):
        print(f"[提示] 请先打开 {CFG} 填上 api_key。")
        sys.exit(0)
    return cfg


def read_roster(path: Path):
    meta, students = {}, []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s:
            continue
        if s.startswith("#"):
            continue
        if s.startswith("@"):
            if ":" in s:
                k, v = s[1:].split(":", 1)
                meta[k.strip()] = v.strip()
            continue
        parts = s.split()
        if len(parts) >= 2:
            try:
                students.append({"姓名": parts[0], "分数": float(parts[1])})
                continue
            except ValueError:
                pass
        students.append({"姓名": parts[0], "分数": None})
    return meta, students


def best_match(raw: str, names: list[str]) -> str:
    """把 AI 读出来的名字对到名单上。

    先精确匹配，再模糊匹配：长度相同 + 至少一个字相同（且不重叠率低）就认。
    手写体常把一个字认错（李璞心→李昊心），这种要能兜住。
    """
    if not raw:
        return ""
    r = norm_name(raw)
    if not r:
        return ""
    exact = {norm_name(n): n for n in names}
    if r in exact:
        return exact[r]
    best, best_score = "", 0.0
    for n in names:
        nn = norm_name(n)
        if len(nn) != len(r):
            continue
        same = sum(1 for a, b in zip(r, nn) if a == b)
        score = same / len(r)
        # 首字或末字相同，再加一点权重（姓氏/末字通常更稳）
        if r[0] == nn[0]:
            score += 0.15
        if r[-1] == nn[-1]:
            score += 0.10
        if score > best_score:
            best, best_score = n, score
    return best if best_score >= 0.55 else ""


def to_global_wrong(wrong_txt: str, sec_names: list, sec_counts: list) -> list:
    """把「一:1,3; 二:5」转成全局题号。
    大题顺序：一=第一个大题，偏移按前面各大题题数累加。"""
    if not wrong_txt:
        return []
    # 每个大题在全局里的起始题号
    base, offset = {}, 1
    for i, cnt in enumerate(sec_counts):
        base["一二三四五六七八九十"[i]] = offset
        offset += (cnt or 0)
    out = []
    # 支持「一:1,3; 二:5」和「1,3,5」两种写法
    has_sec = any(k in wrong_txt for k in base)
    if not has_sec:
        for x in re.split(r"[,，、\s]+", wrong_txt):
            if x.strip().isdigit():
                out.append(int(x))
        return sorted(set(out))
    for seg in re.split(r"[;；]", wrong_txt):
        seg = seg.strip()
        if not seg:
            continue
        m = re.match(r"^([一二三四五六七八九十])\s*[:：]?\s*(.*)$", seg)
        if not m:
            continue
        key, rest = m.group(1), m.group(2)
        off = base.get(key)
        if off is None:
            continue
        for x in re.split(r"[,，、\s]+", rest):
            x = x.strip()
            if x.isdigit():
                out.append(off + int(x) - 1)
    return sorted(set(out))


def wrong_ok(txt: str) -> bool:
    """判断第二遍的结果是不是有效（不是抄示例、不是全空、不是全 0）。"""
    t = (txt or "").strip()
    if not t:
        return False
    # 抄了旧示例 / 新示例
    if t in ("一:1,3,7; 二:2; 三:4", "一:0; 二:0; 三:0"):
        return False
    nums = re.findall(r"\d+", t)
    if not nums:
        return False                      # 一个数字都没有 = 无效
    if all(int(x) == 0 for x in nums):
        return False                      # 全是占位 0 = 无效
    return True


def looks_like_name(stem: str) -> bool:
    """文件名像不像中文姓名：2-4 个汉字，允许 A 前缀（如 A李怀森）。"""
    t = re.sub(r"^[A-Za-z]", "", stem)
    return bool(re.fullmatch(r"[\u4e00-\u9fff]{2,4}", t))


def norm_name(x: str) -> str:
    x = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", x or "")
    return re.sub(r"[\s_\-（）()0-9.]", "", x)


def collect_photos(cls: str, lesson: str | None):
    base = DIR_PHOTO / cls
    if not base.exists():
        print(f"[错] 没有照片目录：{base}")
        sys.exit(1)
    dirs = [base / lesson] if lesson and (base / lesson).exists() else []
    if not dirs:
        dirs = [d for d in base.iterdir() if d.is_dir() and d.name not in
                {"已改名", "分发", "照片归档"}]
    photos = []
    for d in dirs:
        ys = d / "已改名"
        target = ys if ys.exists() and any(ys.glob("*.jpg")) else d
        photos += sorted([p for p in target.glob("*.jpg")])
        photos += sorted([p for p in target.glob("*.jpeg")])
        photos += sorted([p for p in target.glob("*.png")])
    # 去重
    seen, out = set(), []
    for p in photos:
        if p.name not in seen:
            seen.add(p.name); out.append(p)
    return out, [d.name for d in dirs]


def encode(p: Path, max_side: int) -> str:
    from PIL import Image, ImageOps
    im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
    im.thumbnail((max_side, max_side * 2), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


def parse_json(txt: str) -> dict | None:
    if not txt:
        return None
    t = re.sub(r"```(?:json)?|```", "", txt).strip()
    m = re.search(r"\{.*\}", t, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(0))
    except Exception:
        try:
            return json.loads(re.sub(r",\s*}", "}", m.group(0)))
        except Exception:
            return None


def call_api(cfg, prompt, b64, timeout=180):
    body = {"model": cfg["api_model"], "max_tokens": 8000, "stream": False,
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/jpeg;base64,{b64}"}}]}]}
    req = urllib.request.Request(
        cfg["api_base"], data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json",
                 "Authorization": f"Bearer {cfg['api_key']}"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        d = json.loads(r.read().decode())
    m = d["choices"][0]["message"]
    return m.get("content", ""), (m.get("reasoning_content") or "")


def build_prompt(sec_names, want_name: bool):
    keys = ['"分数": 总分数字']
    if want_name:
        keys.insert(0, '"姓名": "卷面上的名字"')
    sec = "、".join(sec_names)
    p = [
        "这是一张小学英语考试答题卡的照片，上面有老师手写的批改。",
        f"卷面大题依次是：{sec}。右栏通常写有每个大题的扣分（如 -3 -6 -1）。",
        "请读出以下内容：",
        "1) 卷面总分（分数栏里红笔写的数字）",
        "2) 每个大题的扣分（如果右栏没写，就按你看到的批改自己数出来；对不上的留 0）",
    ]
    if want_name:
        p.append("3) 卷面手写的学生姓名")
    p.append("")
    p.append("只输出一行 JSON，不要任何解释和 markdown 代码块。格式：")
    p.append("{" + ", ".join(keys + [f'"{x}": 扣分数字' for x in sec_names]) + "}")
    p.append("分数不确定时，估计一个最可能的数字，不要省略。")
    p.append("请直接给出最终 JSON，不要写长篇分析过程。")
    return "\n".join(p)


def build_wrong_prompt(sec_names, sec_counts, allow, expect=None):
    """第二遍专用：只问错题号，别的不问（问多了会拖垮第一遍的准确率）。

    expect: 每个大题【应该错几道】（由第一遍的扣分推出）—— 这是最有力的约束。
    """
    lines = ["这是同一张答题卡。这次只做一件事：找出老师【圈出或打叉】的题号。"]
    for i, (name, cnt) in enumerate(zip(sec_names, sec_counts)):
        hint = ""
        if expect and i < len(expect) and expect[i] is not None:
            e = expect[i]
            hint = f"　（这个大题应该正好错 {e} 道，请核对数量）"
        lines.append(f"　{name}：第 1-{cnt} 题{hint}")
    lines.append("按上面的大题顺序，用「一」「二」「三」…代表第几个大题。")
    lines.append("")
    lines.append('只输出一行 JSON，格式：{"错题": "一:0; 二:0; 三:0"}')
    lines.append("上面 0 只是占位符，务必换成你实际看到的题号；"
                 "没有错题就留空，写成 一: 这样。")
    lines.append("每个大题的错题号用逗号隔开，不同大题用分号隔开。")
    lines.append(f"题号只能在你确认看清楚了才写，范围是 {allow}。看不准就留空，不要猜。")
    lines.append("不要解释，不要 markdown，只给 JSON。")
    return "\n".join(lines)


def do_rename(args, cfg):
    """照片自动改名：AI 认卷面姓名 → 文件改名成「姓名.jpg」。"""
    import shutil
    rp = next((f for f in DIR_ROSTER.glob("*.txt") if args.班级 in f.stem), None)
    if not rp:
        print(f"[错] 找不到名单：1-名单\\{args.班级}.txt")
        return 1
    meta, students = read_roster(rp)
    lesson = args.课次 or meta.get("课次", "")
    names = [s["姓名"] for s in students]
    norm2real = {norm_name(n): n for n in names}

    base = DIR_PHOTO / args.班级
    if not base.exists():
        print(f"[错] 没有照片目录：{base}")
        return 1
    src = (base / lesson) if (lesson and (base / lesson).exists()) else base
    photos = sorted([p for p in src.glob("*.jpg")] +
                    [p for p in src.glob("*.jpeg")] +
                    [p for p in src.glob("*.png")])
    # 已经改过名的就不动
    photos = [p for p in photos if not looks_like_name(p.stem)]
    if args.限:
        photos = photos[:args.限]
    if not photos:
        print("[提示] 这个文件夹里没有需要改名的照片")
        print("       （已经叫「学生姓名.jpg」的照片不用改）")
        return 0

    print(f"班级：{args.班级}　课次：{lesson or '（自动）'}")
    print(f"待改名：{len(photos)} 张　模型：{cfg['api_model']}")
    print("=" * 64)

    max_side = int(cfg.get("图片最长边") or 1400)
    workers = max(1, int(cfg.get("并发数") or 4))
    prompt = ("这是一张小学英语考试答题卡的照片。\n"
              "请读出卷面上【老师或学生手写的学生姓名】。\n"
              "只输出一行 JSON，不要任何解释和 markdown。格式：{\"姓名\": \"张三\"}\n"
              "看不出姓名就写 {\"姓名\": \"\"}")

    def one(idx_p):
        idx, p = idx_p
        last = ""
        raw_best = ""
        for attempt in range(3):
            pr = prompt
            if attempt > 0:
                # 第一次没对上名单，换个说法让它再仔细看一遍
                pr = ("再仔细看一遍这张答题卡。手写的学生姓名在「姓名：」那一行的横线上。\n"
                      f"名单里可能是这几个之一：{'、'.join(names)}\n"
                      "只输出一行 JSON：{\"姓名\": \"\"}")
            try:
                b64 = encode(p, max_side)
                txt, reasoning = call_api(cfg, pr, b64)
                data = parse_json(txt) or parse_json(reasoning)
                if data is not None:
                    raw = str(data.get("姓名", "")).strip()
                    if raw:
                        raw_best = raw
                        if norm2real.get(norm_name(raw)) or best_match(raw, names):
                            return idx, raw, None      # 对上了，直接用
                    last = f"读了「{raw}」但名单里对不上，重试中"
                else:
                    last = "没解析出 JSON"
            except urllib.error.HTTPError as e:
                last = f"HTTP {e.code}"
            except Exception as e:
                last = f"{type(e).__name__}: {e}"
            time.sleep(1.2 * (attempt + 1))
        # 3 次都没对上，把最好的一次结果返回（交给核对表人工确认）
        return idx, (raw_best or None), last

    results, errs = {}, []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for idx, raw, err in ex.map(one, list(enumerate(photos, 1))):
            p = photos[idx - 1]
            if raw is None:
                errs.append((p.name, err))
                print(f"  [{idx:2d}/{len(photos)}] {p.name:28s} 读取失败：{err}")
                continue
            real = norm2real.get(norm_name(raw)) or best_match(raw, names)
            results[p.name] = (raw, real)
            tag = f"→ {real}" if real else f"（{raw or '没读出'}）✘ 名单里没有"
            print(f"  [{idx:2d}/{len(photos)}] {p.name:28s} {tag}")

    # ---- 真正改名 ----
    out = src / "已改名"
    old = src / "原图"
    if not args.试运行:
        out.mkdir(exist_ok=True)
        old.mkdir(exist_ok=True)
    n_ok = 0
    for p in photos:
        raw, real = results.get(p.name, ("", ""))
        if not real:
            continue
        dst = out / (real + p.suffix.lower())
        if args.试运行:
            print(f"    [试运行] {p.name} -> 已改名/{dst.name}")
            n_ok += 1
            continue
        if dst.exists():
            dst = out / f"{real}_2{p.suffix.lower()}"
        shutil.copy2(p, dst)
        shutil.move(str(p), str(old / p.name))
        n_ok += 1

    # ---- 核对表 ----
    outdir = DIR_OUT / f"{args.班级}_{lesson}" if lesson else DIR_OUT / args.班级
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "改名核对表"
        ws.append(["原文件名", "AI 读到的名字", "改成", "名单里有吗"])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1B5EA8")
        for p in photos:
            raw, real = results.get(p.name, ("", ""))
            ws.append([p.name, raw, real + p.suffix.lower() if real else "",
                       "✔" if real else "✘ 待确认"])
        for i, w in enumerate([30, 16, 20, 12], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        wb.save(outdir / "改名核对表.xlsx")
    except Exception as e:
        print("[警告] 写核对表失败：", e)

    print("=" * 64)
    print(f"改名成功 {n_ok} 张" + ("（试运行，没真的改）" if args.试运行 else ""))
    print(f"  新名字：{out}")
    print(f"  原图  ：{old}")
    print(f"  核对表：{outdir / '改名核对表.xlsx'}")
    print()
    print("下一步：直接跑「识别答题卡」，照片名已经是学生姓名了。")
    return 0


def main():
    ap = argparse.ArgumentParser(description="用视觉大模型读答题卡，生成核对表")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--限", type=int, default=0, help="只处理前 N 张（调试用）")
    ap.add_argument("--改名", action="store_true",
                    help="照片自动改名：AI 认卷面姓名，把 IMG_xxx.jpg 改成 姓名.jpg")
    ap.add_argument("--试运行", action="store_true", help="只演示，不真的改文件")
    ap.add_argument("--并入名单", action="store_true", help="把核对表写回名单")
    ap.add_argument("--核对表", default=None, help="并入名单时指定核对表路径")
    args = ap.parse_args()

    # ---------- 改名模式 ----------
    if args.改名:
        return do_rename(args, load_cfg())

    # ---------- 并入名单模式 ----------
    if args.并入名单:
        return merge_into_roster(args)

    cfg = load_cfg()
    rp = next((f for f in DIR_ROSTER.glob("*.txt") if args.班级 in f.stem), None)
    if not rp:
        print(f"[错] 找不到名单：1-名单\\{args.班级}.txt")
        sys.exit(1)
    meta, students = read_roster(rp)
    lesson = args.课次 or meta.get("课次", "")
    names = [s["姓名"] for s in students]
    norm2real = {}
    for n in names:
        norm2real[norm_name(n)] = n
    # 姓名对照
    alias = {}
    ap_ = ROOT / "班级名单" / "姓名对照.txt"
    if ap_.exists():
        for line in ap_.read_text(encoding="utf-8-sig").splitlines():
            t = line.strip()
            if not t or t.startswith("#"):
                continue
            q = t.split()
            if len(q) >= 2:
                alias[norm_name(q[0])] = q[1]
    sec_names = [x.strip() for x in
                 re.split(r"[,，、]", meta.get("大题名称", "")) if x.strip()]

    photos, dirs = collect_photos(args.班级, args.课次)
    if args.限:
        photos = photos[:args.限]
    print(f"班级：{args.班级}　课次：{lesson or '（自动）'}")
    print(f"照片目录：{', '.join(dirs)}　共 {len(photos)} 张")
    print(f"模型：{cfg['api_model']}　并发：{cfg.get('并发数','4')}")
    print("=" * 64)

    max_side = int(cfg.get("图片最长边") or 1400)
    workers = max(1, int(cfg.get("并发数") or 4))

    # 只有老师上传了答案文件（说明想要逐题定位）才跑第二遍
    import glob as _glob
    ans_dir = ROOT / "课测答案"
    has_answer = False
    if ans_dir.exists():
        key = (lesson or "").strip()
        for f in ans_dir.glob("*.txt"):
            if f.stem.startswith("_") or f.name.endswith(".bak"):
                continue
            if key and (key in f.stem or f.stem in key):
                has_answer = True
                break
    sec_counts = []
    try:
        sec_counts = [int(x) for x in
                      re.split(r"[,，、]", str(meta.get("大题题数", ""))) if x.strip()]
    except Exception:
        sec_counts = []
    if len(sec_counts) != len(sec_names):
        sec_counts = [0] * len(sec_names)
    allow = "、".join(f"{n} 1-{c}" for n, c in zip(sec_names, sec_counts) if c)
    print(f"  第二遍读错题号：{'开' if has_answer else '关（没上传答案文件）'}")

    def one(idx_p):
        idx, p = idx_p
        stem = p.stem
        # ① 文件名能在名单里对上 → 用名单里的写法
        # ② 对不上、但文件名本身就像中文姓名 → 直接用文件名
        # ③ 都不行（比如 IMG_1234.jpg）→ 才让 AI 读手写名字
        hit = norm2real.get(norm_name(stem)) or alias.get(norm_name(stem))
        if hit is None and looks_like_name(stem):
            hit = re.sub(r"^[A-Za-z](?=[\u4e00-\u9fff])", "", stem)
            src_from = "文件名"
        else:
            src_from = "文件名" if hit else "AI读的"
        want_name = hit is None
        prompt = build_prompt(sec_names, want_name)
        last = ""
        for attempt in range(3):
            pr = prompt
            if attempt > 0:
                keys = ", ".join('"%s": 0' % x for x in sec_names)
                pr = ("不要分析，直接输出一行 JSON。格式："
                      '{"分数": 27, ' + keys + "}")
            try:
                b64 = encode(p, max_side)
                txt, reasoning = call_api(cfg, pr, b64)
                data = parse_json(txt) or parse_json(reasoning)
                if data:
                    if hit:
                        data["姓名"] = hit
                    else:
                        raw_nm = str(data.get("姓名", "")).strip()
                        data["姓名"] = norm2real.get(norm_name(raw_nm), raw_nm)
                    data["_来自文件名"] = (src_from == "文件名")
                    data["_文件"] = p.name
                    # ---- 第二遍：只读错题号（偶尔返回空，重试 2 次）----
                    if has_answer and allow:
                        # 由第一遍的扣分推出每个大题应该错几道
                        expect = []
                        for _sn in sec_names:
                            _v = data.get(_sn, "")
                            try:
                                expect.append(int(round(float(_v))))
                            except (TypeError, ValueError):
                                expect.append(None)
                        wp = build_wrong_prompt(sec_names, sec_counts, allow,
                                                expect=expect)
                        for _try in range(3):
                            try:
                                t2, r2 = call_api(cfg, wp, b64)
                                d2 = parse_json(t2) or parse_json(r2)
                                w = str((d2 or {}).get("错题", "")).strip()
                                if wrong_ok(w):
                                    data["错题"] = w
                                    break
                                time.sleep(0.8)
                            except Exception:
                                time.sleep(0.8)
                    return idx, data, None
                last = f"没解析出 JSON（原文：{(txt or reasoning)[:60]!r}）"
            except urllib.error.HTTPError as e:
                last = f"HTTP {e.code}: " + e.read().decode(errors='replace')[:120]
            except Exception as e:
                last = f"{type(e).__name__}: {e}"
            time.sleep(1.5 * (attempt + 1))
        return idx, None, last

    results, errors = {}, []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        for idx, data, err in ex.map(one, list(enumerate(photos, 1))):
            p = photos[idx - 1]
            if data:
                results[p.name] = data
                sc = data.get("分数")
                flag = "" if data.get("_来自文件名") else "  ← 姓名是AI读的，务必核对"
                print(f"  [{idx:2d}/{len(photos)}] {p.name:26s} "
                      f"{str(data.get('姓名')):8s} {sc}分{flag}")
            else:
                errors.append((p.name, err))
                print(f"  [{idx:2d}/{len(photos)}] {p.name:26s} 读取失败：{err}")

    # ---------- 写核对表 ----------
    outdir = DIR_OUT / f"{args.班级}_{lesson}" if lesson else DIR_OUT / args.班级
    outdir.mkdir(parents=True, exist_ok=True)
    xlsx = outdir / "识别核对表.xlsx"
    txtf = outdir / "识别核对表.txt"
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "核对表"
        head = (["照片文件", "识别姓名", "名单里有吗", "分数"] + sec_names
                + ["错题号", "姓名来源", "备注"])
        ws.append(head)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1B5EA8")
        red = Font(color="C0392B", bold=True)
        for p in photos:
            d = results.get(p.name)
            if not d:
                ws.append([p.name, "", "读取失败", "", *[""] * len(sec_names), "", ""])
                ws.cell(row=ws.max_row, column=1).font = red
                continue
            nm = d.get("姓名", "")
            known = "✔" if nm in names else "✘ 待确认"
            row = [p.name, nm, known, d.get("分数")]
            for s in sec_names:
                row.append(d.get(s, ""))
            row.append(str(d.get("错题", "") or ""))
            row.append("文件名" if d.get("_来自文件名") else "AI读的")
            row.append("")
            ws.append(row)
            if known != "✔":
                ws.cell(row=ws.max_row, column=3).font = red
        widths = [26, 12, 12, 8] + [9] * len(sec_names) + [10, 16]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        wb.save(xlsx)
    except Exception as e:
        print("[警告] 写 xlsx 失败：", e)

    lines = ["照片文件\t识别姓名\t名单里有吗\t分数\t" + "\t".join(sec_names) + "\t姓名来源"]
    for p in photos:
        d = results.get(p.name)
        if not d:
            lines.append(f"{p.name}\t\t读取失败\t\t" + "\t" * len(sec_names))
            continue
        nm = d.get("姓名", "")
        lines.append("\t".join([p.name, str(nm), "✔" if nm in names else "✘待确认",
                                str(d.get("分数", ""))] +
                               [str(d.get(s, "")) for s in sec_names] +
                               [str(d.get("错题", "") or "")] +
                               ["文件名" if d.get("_来自文件名") else "AI读的"]))
    txtf.write_text("\n".join(lines), encoding="utf-8")

    print("=" * 64)
    print(f"成功 {len(results)} 张，失败 {len(errors)} 张")
    print(f"核对表：{xlsx}")
    print(f"        {txtf}")
    print()
    print("下一步：")
    print("  1) 打开核对表，把「识别姓名」和「分数」错的地方改掉")
    print("  2) 跑：python 识别照片.py --班级 "
          f"{args.班级} --并入名单 --核对表 \"{xlsx}\"")
    return 0


def merge_into_roster(args):
    """把核对表里的分数写进名单。"""
    rp = next((f for f in DIR_ROSTER.glob("*.txt") if args.班级 in f.stem), None)
    if not rp:
        print(f"[错] 找不到名单：{rp}")
        sys.exit(1)
    path = Path(args.核对表) if args.核对表 else None
    if not path or not path.exists():
        print("[错] 请用 --核对表 指定 xlsx 路径")
        sys.exit(1)
    import openpyxl
    ws = openpyxl.load_workbook(path).active
    head = [c.value for c in ws[1]]
    i_nm, i_sc = head.index("识别姓名"), head.index("分数")
    got = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        nm, sc = row[i_nm], row[i_sc]
        if nm and sc not in (None, ""):
            try:
                got[str(nm).strip()] = float(sc)
            except (TypeError, ValueError):
                pass

    # 先把「错题号」算成全局题号
    wrong_global = {}
    try:
        import openpyxl
        ws0 = openpyxl.load_workbook(path).active
        h0 = [c.value for c in ws0[1]]
        if "错题号" in h0:
            i_nm0, i_w0 = h0.index("识别姓名"), h0.index("错题号")
            meta0, _ = read_roster(rp)
            sec0 = [x.strip() for x in
                    re.split(r"[,，、]", meta0.get("大题名称", "")) if x.strip()]
            cnt0 = []
            for x in re.split(r"[,，、]", meta0.get("大题题数", "")):
                try:
                    cnt0.append(int(x))
                except ValueError:
                    cnt0.append(0)
            for r0 in ws0.iter_rows(min_row=2, values_only=True):
                nm0 = str(r0[i_nm0] or "").strip()
                w0 = str(r0[i_w0] or "").strip()
                if nm0 and w0:
                    g = to_global_wrong(w0, sec0, cnt0)
                    if g:
                        wrong_global[nm0] = g
    except Exception as e:
        print(f"[注意] 错题号转换失败：{type(e).__name__}: {e}")

    txt = rp.read_text(encoding="utf-8-sig")
    out, n, have = [], 0, set()
    for raw in txt.splitlines():
        s = raw.strip()
        if not s or s.startswith(("@", "#")):
            out.append(raw); continue
        parts = s.split()
        nm = parts[0]
        have.add(nm)
        if nm in got:
            extra = wrong_global.get(nm)
            line2 = f"{nm} {got[nm]:g}"
            if extra:
                line2 += " " + ",".join(str(x) for x in extra)
            out.append(line2); n += 1
        else:
            out.append(nm)

    # 名单里没有、但核对表里识别到的，自动补进去
    added = []
    for nm, sc in got.items():
        if nm not in have:
            extra = wrong_global.get(nm)
            line3 = f"{nm} {sc:g}"
            if extra:
                line3 += " " + ",".join(str(x) for x in extra)
            out.append(line3)
            added.append(nm)
            n += 1
    if added:
        # 把新学生接在最后一条学生后面，保住原来的排版
        out = [x for x in out]
        print(f"[新增] 名单里原来没有，已自动加入 {len(added)} 人："
              + "、".join(added[:8]) + ("…" if len(added) > 8 else ""))

    backup = rp.with_suffix(".txt.bak")
    if not backup.exists():
        backup.write_text(txt, encoding="utf-8")
    rp.write_text("\n".join(out).rstrip() + "\n", encoding="utf-8")
    print(f"[完成] 共写入 {n} 人（原文件备份为 {backup.name}）")

    # ---- 顺便把「各大题扣分」写给话术/小任务用 ----
    try:
        import json
        import openpyxl
        ws = openpyxl.load_workbook(path).active
        head = [c.value for c in ws[1]]
        i_nm = head.index("识别姓名")
        i_sc = head.index("分数")
        # 大题列 = 表头里除固定列之外的
        fixed = {"照片文件", "识别姓名", "名单里有吗", "分数", "姓名来源", "备注"}
        sec_names = [c for c in head if c not in fixed]
        sec_key = {name: k for name, k in zip(sec_names, "一二三四五六七八九")}

        # 名单里的满分 / 课次
        meta2, _ = read_roster(rp)
        full = float(meta2.get("满分") or 40)
        lesson = args.课次 or meta2.get("课次", "")

        data = {"_说明": f"{args.班级} 各大题扣分（识别照片自动生成）",
                "_校验": "各大题扣分之和必须等于 满分−分数"}
        bad = 0
        for row in ws.iter_rows(min_row=2, values_only=True):
            nm = row[i_nm]
            if not nm:
                continue
            nm = str(nm).strip()
            try:
                score = float(row[i_sc])
            except (TypeError, ValueError):
                continue
            d, got = {}, 0.0
            for name in sec_names:
                v = row[head.index(name)]
                try:
                    v = float(v or 0)
                except (TypeError, ValueError):
                    v = 0.0
                d[sec_key[name]] = v
                got += v
            need = full - score
            if abs(got - need) > 0.01:
                d["未定位"] = round(need - got, 2)
                bad += 1
            data[nm] = d

        # 错题号也存一份，供话术/小任务定位考点
        try:
            who_txt = ""
            if "错题号" in head:
                i_w = head.index("错题号")
                for row in ws.iter_rows(min_row=2, values_only=True):
                    nm2 = row[i_nm]
                    if not nm2:
                        continue
                    wrong_txt = str(row[i_w] or "").strip()
                    if wrong_txt:
                        data.setdefault("_错题", {})[str(nm2).strip()] = wrong_txt
        except Exception:
            pass

        if data and len(data) > 1:
            outdir = DIR_OUT / f"{args.班级}_{lesson}" if lesson else DIR_OUT / args.班级
            outdir.mkdir(parents=True, exist_ok=True)
            jp = outdir / "大题扣分.json"
            jp.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                          encoding="utf-8")
            print(f"[顺带] 已写出各大题扣分：{jp.name}"
                  + (f"（{bad} 人分项和对不上，已记『未定位』）" if bad else ""))
            print("       话术和小任务会用它来定位知识点")
    except Exception as e:
        print(f"[注意] 写大题扣分失败（不影响登分）：{type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
