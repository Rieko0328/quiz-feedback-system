# -*- coding: utf-8 -*-
r"""出题库.py —— 统计本讲全班失分考点，导出清单给 AI 写题库。

分工（重要）：
  本脚本 = 只做**统计**：谁错了哪道 → 哪些考点失分最多 → 写出清单
  出题   = 交给对话里的 AI（DeepSeek）。脚本不再自己调 API 出题。

  原因：调免费模型出题会出现错答案（实测「What's this? → 选 Is it a book?」），
  巩固单是直接发家长的，错答案比没题更糟。

用法：
    # ① 导出清单（第一步，看会出哪些考点）
    python 出题库.py --班级 周五慢班

    # ② 把清单里的考点告诉 AI，让它按格式写 题目/<班级>_<课次>/题库.json
    # ③ 出巩固单
    python 批量学情.py --班级 周五慢班 --只巩固单

  老路径（自己调 API 出题）还留着，用 --调API 才走，默认不调。

原说明：按本讲全班实际失分的考点，让 AI 生成一份题库。

巩固单的价值在「针对这次错的考点出题」。没有题库时只能出抄写默写版，
所以这个脚本补的就是那一环。

输入：课测答案（含考点）+ 班级名单（含每人错题号）
输出：题目/<班级>_<课次>/题库.json
      {"说明": "...", "考点": {"<考点>": {"背诵": [...3条...],
                                          "题": [{"题干","选项","答案","解析"}...]}}}

用法：
    python 出题库.py --班级 周五慢班
    python 出题库.py --班级 周五慢班 --课次 NCE1-L9-14月考 --每题 6
    python 出题库.py --班级 周五慢班 --只列考点      # 只看会出哪些考点，不调 API
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _ws import workspace_root, safe_name  # noqa: E402


# ------------------------------------------------------------------ 读输入
def read_answer(path: Path) -> dict[int, dict]:
    """课测答案 → {题号: {答案, 考点, 题型}}"""
    out = {}
    if not path.exists():
        return out
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith(("#", "@")):
            continue
        m = re.match(r"^(\d+)\s*[.、．:：]?\s*(.+)$", s)
        if not m:
            continue
        no = int(m.group(1))
        body = m.group(2)
        parts = [x.strip() for x in re.split(r"[|｜]", body)]
        if len(parts) < 2:
            parts = [x.strip() for x in re.split(r"\s{2,}", body)]
        ans = parts[0] if parts else ""
        kp = parts[1] if len(parts) > 1 else ""
        qtype = parts[2] if len(parts) > 2 else ""
        out[no] = {"答案": ans, "考点": kp, "题型": qtype}
    return out


def read_roster(path: Path):
    """名单 → (meta, [(姓名, 分数, [错题号])])"""
    meta, rows = {}, []
    if not path.exists():
        return meta, rows
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("@"):
            if ":" in s:
                k, v = s[1:].split(":", 1)
                meta[k.strip()] = v.strip()
            continue
        parts = s.split()
        if not parts:
            continue
        name = parts[0]
        score = parts[1] if len(parts) > 1 else ""
        wrong = []
        if len(parts) > 2:
            for tok in re.split(r"[,，、]", parts[2]):
                tok = tok.strip()
                if tok.isdigit():
                    wrong.append(int(tok))
        rows.append((name, score, wrong))
    return meta, rows


def aligned_sections(meta: dict):
    """把 大题名称 / 大题题数 对齐，并按大题把题号编成全局号"""
    names = [x.strip() for x in re.split(r"[,，、]", meta.get("大题名称", "")) if x.strip()]
    counts = []
    for x in re.split(r"[,，、]", meta.get("大题题数", "")):
        try:
            counts.append(int(x))
        except ValueError:
            counts.append(0)
    if names and len(counts) != len(names):
        counts = [0] * len(names)
    # 全局题号 → 大题名
    q2sec, no = {}, 1
    for nm, cnt in zip(names, counts):
        for _ in range(cnt):
            q2sec[no] = nm
            no += 1
    return names, counts, q2sec


# ------------------------------------------------------------------ 挑考点
def pick_points(meta, rows, answers, q2sec, limit=12):
    """统计每个考点被多少人错过，返回 [(考点, 错过人数, [题型...])]"""
    cnt = Counter()
    types = defaultdict(set)
    for _name, _score, wrong in rows:
        # 一个孩子在一个考点上错 3 道，也只算「1 人」，不能算 3 人
        seen_atoms = set()
        for q in wrong:
            a = answers.get(q)
            if not a:
                continue
            kp = (a.get("考点") or "").strip()
            if not kp:
                continue
            for atom in re.split(r"[,，、]", kp):
                atom = atom.strip()
                if not atom:
                    continue
                if a.get("题型"):
                    types[atom].add(a["题型"].strip())
                if atom not in seen_atoms:
                    seen_atoms.add(atom)
                    cnt[atom] += 1
    out = []
    for kp, n in cnt.most_common(limit):
        out.append((kp, n, sorted(types[kp])))
    return out


# ------------------------------------------------------------------ 调 API
def build_prompt(lesson: str, points, full: float):
    lines = [
        f"你在给一位小学英语老师出教辅。这是「{lesson}」这次课测的题库。",
        "这次只要下面这几个考点（不要多写别的考点）：",
    ]
    for kp, n, ts in points:
        t = "、".join(ts) if ts else "选择题"
        lines.append(f"  · {kp}（{n} 人错，题型：{t}）")
    lines += [
        "",
        "请为上面每一个考点各出：",
        "1) 「背诵」3 条要点：一句话讲透规则或固定搭配，要具体，能直接照着背。",
        "2) 「题」4 道单选题：题干用小学四五年级能懂的英语，选项 4 个（A. B. C. D.），",
        "   只有一个正确答案，答案写字母，解析一句话说清为什么。",
        "",
        "要求：",
        "- 考点名必须和上面一字不差地对应（用作 key）",
        "- 难度贴近公校四五年级和 KET 入门",
        "- 不要出现中文句子当题干（可以有中文提示词，比如「填合适的词」）",
        "- 同一个考点下的 4 道题不要雷同",
        "",
        "只输出一行 JSON，不要 markdown 代码块，不要任何解释。格式：",
        '{"说明": "一句话", "考点": {"<考点名>": {"背诵": ["...", "...", "..."], '
        '"题": [{"题干": "...", "选项": ["A. ...", "B. ...", "C. ...", "D. ..."], '
        '"答案": "A", "解析": "..."}]}}}',
    ]
    return "\n".join(lines)


def parse_json(txt: str):
    if not txt:
        return None
    t = txt.strip()
    if "```" in t:
        t = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", t.strip(), flags=re.M)
    i, j = t.find("{"), t.rfind("}")
    if i < 0 or j <= i:
        return None
    try:
        return json.loads(t[i:j + 1])
    except Exception:
        return None


def call_text(cfg, prompt: str, timeout=300):
    """纯文本调用（不带图）。

    识别照片.py 的 call_api 是视觉专用，content 里强制塞 image_url，
    出题库这种纯文字任务传进去会被 400 拒掉，所以单独写一个。
    """
    import os
    import urllib.error

    spec = importlib.util.spec_from_file_location(
        "rec", Path(__file__).resolve().parent / "识别照片.py")
    rec = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(rec)

    key = cfg["api_base"]
    try:
        mt = int(str(cfg.get("max_tokens") or 8000).strip())
    except Exception:
        mt = 8000
    # 智谱 glm-4v-flash 的 max_tokens 硬上限是 1024，写大了直接 400。
    # _MAX_TOKENS 是「撞过 400 之后才记住」的，第一次跑还没记 → 这里按域名先限死。
    if "bigmodel" in key or "zhipu" in key:
        mt = 1000
    mt = min(int(rec._MAX_TOKENS.get(key, mt)), mt) if key in rec._MAX_TOKENS else mt
    body = {"model": cfg["api_model"], "max_tokens": mt, "stream": False,
            "messages": [{"role": "user", "content": prompt}]}
    try:
        d = rec._post(cfg, body, timeout)
    except urllib.error.HTTPError as e:
        msg = ""
        try:
            msg = e.read().decode("utf-8", "replace")
        except Exception:
            pass
        raise RuntimeError(f"HTTP {e.code}: {msg[:300]}")
    m = d["choices"][0]["message"]
    return m.get("content", ""), (m.get("reasoning_content") or "")


def main():
    ap = argparse.ArgumentParser(description="按失分考点生成题库")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--每题", type=int, default=4, help="每个考点出几道（默认4）")
    ap.add_argument("--最多考点", type=int, default=12)
    ap.add_argument("--导出清单", action="store_true",
                    help="（默认就做这个）把考点清单写成 txt，交给 AI 写题库")
    ap.add_argument("--调API", action="store_true",
                    help="老路径：自己调 API 出题（质量一般，错答案风险高）")
    ap.add_argument("--只列考点", action="store_true", help="只看会出哪些考点")
    ap.add_argument("--输出", default=None)
    args = ap.parse_args()

    root = workspace_root(args.工作区)
    cls = safe_name(args.班级)

    meta, rows = read_roster(root / "班级名单" / f"{cls}.txt")
    if not rows:
        print(f"[错] 名单不存在或没有学生：班级名单/{cls}.txt")
        sys.exit(1)
    lesson = args.课次 or meta.get("课次", "")
    if not lesson or lesson.startswith("请填写"):
        print(f"[错] 名单里还没有课次（@课次），请用 --课次 指定")
        sys.exit(1)
    full = float(meta.get("满分") or 40)

    # 答案：先精确找，再模糊找
    ans_dir = root / "课测答案"
    ans_path = ans_dir / f"{lesson}.txt"
    if not ans_path.exists() and ans_dir.is_dir():
        cands = [p for p in ans_dir.glob("*.txt")
                 if p.stem != "_模板" and (lesson in p.stem or p.stem in lesson)]
        if len(cands) == 1:
            ans_path = cands[0]
            print(f"  [提示] 答案按名字找到：{ans_path.name}")
    answers = read_answer(ans_path)
    if not answers:
        print(f"[错] 没有可用的课测答案（{ans_dir / (lesson + '.txt')}）")
        print("      题库要靠答案里的「考点」列来出题，没有它没法针对性出。")
        sys.exit(1)
    print(f"  答案：{ans_path.name}　{len(answers)} 道题")

    names, counts, q2sec = aligned_sections(meta)
    print(f"  班级：{cls}　{len(rows)} 人　课次：{lesson}　满分 {full:g}")

    points = pick_points(meta, rows, answers, q2sec, args.最多考点)
    if not points:
        print("[错] 名单里没有错题号，或者错题号在答案里找不到 → 不知道谁错了哪道")
        print("      先确认名单每行是「姓名 分数 错题号」，且题号与答案一致")
        sys.exit(1)

    print(f"\n  会出的考点（按错过人数排）：")
    for kp, n, ts in points:
        print(f"    {kp:22s} {n:2d} 人错　{'/'.join(ts) if ts else ''}")

    if args.只列考点:
        print("\n（--只列考点：没调 API）")
        return

    # ---------- 默认：导出清单，交给 AI 写题库 ----------
    if not args.调API:
        out_dir = Path(args.输出) if args.输出 else root / "题目" / f"{cls}_{lesson}"
        out_dir.mkdir(parents=True, exist_ok=True)
        brief = out_dir / "出题清单.txt"

        # 每个考点：错的人 + 题型 + 该考点涉及的原题（题干不知道，但答案和题型能给）
        qs_by_kp = {}
        for q, a in answers.items():
            for atom in re.split(r"[,，、]", (a.get("考点") or "")):
                atom = atom.strip()
                if atom:
                    qs_by_kp.setdefault(atom, []).append(
                        f"第{q}题（答案 {a.get('答案')}，{a.get('题型') or '选择题'}）")

        lines = [
            f"# 出题清单：{cls} / {lesson}",
            f"# 满分 {full:g} 分　学生 {len(rows)} 人",
            "#",
            "# 交给 AI 的指令（连同本文件一起发过去）：",
            "#   请按下面每个考点写一份题库，格式为：",
            '#   {"说明": "一句话", "考点": {"<考点名>": {"背诵": ["...3条..."],',
            '#     "题": [{"题干": "...", "选项": ["A. ...","B. ...","C. ...","D. ..."],',
            '#             "答案": "A", "解析": "..."}]}}}',
            f"#   每个考点：背诵 3 条 + 单选题 {args.每题} 道，答案必须正确。",
            "#   写完存成同目录下的 题库.json。",
            "",
        ]
        for kp, n, ts in points:
            lines.append(f"## {kp}")
            lines.append(f"   错的人数：{n} / {len(rows)}　题型：{'、'.join(ts) if ts else '选择题'}")
            sample = qs_by_kp.get(kp) or []
            if sample:
                lines.append(f"   相关原题：{'; '.join(sample[:5])}")
            lines.append("")

        brief.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n[完成] 出题清单 → {brief.relative_to(root)}")
        print(f"\n  把这个文件发给 AI（或直接说「按出题清单写题库」），")
        print(f"  它会写 {out_dir.name}/题库.json，然后跑：")
        print(f"    python 批量学情.py --班级 {cls} --只巩固单")
        return

    # ---- 调 API ----
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location(
        "rec", Path(__file__).resolve().parent / "识别照片.py")
    rec = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(rec)

    import os
    old = os.getcwd()
    os.chdir(str(root))
    try:
        cfg = rec.load_cfg()
    finally:
        os.chdir(old)
    if not cfg:
        print("[错] 没能读到识别配置.txt 里的 API 配置")
        sys.exit(1)

    # ---- 分批调用：批大小按 token 上限自适应 ----
    # 一个考点（3 条背诵 + 4 道题）大约吃 400 token。
    # glm-4v-flash 上限 1024 → 一次只放 1 个考点，放 2 个就会把 JSON 截断。
    _cap = 1000 if ("bigmodel" in (cfg.get("api_base") or "")) else 8000
    BATCH = max(1, min(args.每题 and 3 or 3, (_cap - 200) // 420))
    BATCH = min(BATCH, 3)
    batches = [points[i:i + BATCH] for i in range(0, len(points), BATCH)]
    print(f"\n  正在让 AI 出题（{len(points)} 个考点 × {args.每题} 题，"
          f"分 {len(batches)} 批）…")
    merged = {}
    for bi, bp_ in enumerate(batches, 1):
        names_b = "、".join(x[0] for x in bp_)
        print(f"    第 {bi}/{len(batches)} 批：{names_b}")
        prompt = build_prompt(lesson, bp_, full)
        try:
            txt, reasoning = call_text(cfg, prompt, timeout=300)
        except Exception as e:
            print(f"      [跳过] 这批失败：{type(e).__name__}: {e}")
            continue
        data = parse_json(txt) or parse_json(reasoning or "")
        if not data or not isinstance(data.get("考点"), dict):
            # 多半是输出被 token 上限截断了 —— 拆成一个考点一个考点重试
            if len(bp_) > 1:
                print(f"      [拆分重试] 一次装不下 {len(bp_)} 个考点，改成逐个出")
                for one in bp_:
                    try:
                        t1, r1 = call_text(cfg, build_prompt(lesson, [one], full),
                                           timeout=300)
                        d1 = parse_json(t1) or parse_json(r1 or "")
                        if d1 and isinstance(d1.get("考点"), dict):
                            for k, v in d1["考点"].items():
                                if isinstance(v, dict):
                                    merged[k] = v
                            print(f"        ✔ {one[0]}")
                        else:
                            print(f"        ✘ {one[0]}（还是解析不出）")
                    except Exception as e2:
                        print(f"        ✘ {one[0]}：{type(e2).__name__}")
            else:
                print(f"      [跳过] 解析不出 JSON（原文前 120 字："
                      f"{(txt or reasoning or '')[:120]!r}）")
            continue
        for k, v in data["考点"].items():
            if isinstance(v, dict):
                merged[k] = v
    if not merged:
        print("[错] 所有批次都没出成功，检查 API key / 网络")
        sys.exit(1)
    data = {"考点": merged}

    bank = {}
    for kp, v in data["考点"].items():
        if not isinstance(v, dict):
            continue
        recite = [str(x).strip() for x in (v.get("背诵") or []) if str(x).strip()][:3]
        qs = []
        for q in (v.get("题") or [])[:args.每题]:
            if not isinstance(q, dict):
                continue
            stem = str(q.get("题干") or "").strip()
            opts = [str(x).strip() for x in (q.get("选项") or []) if str(x).strip()]
            ans = str(q.get("答案") or "").strip().upper()[:1]
            if stem and len(opts) >= 2 and ans:
                qs.append({"题干": stem, "选项": opts, "答案": ans,
                           "解析": str(q.get("解析") or "").strip()})
        if recite or qs:
            bank[kp] = {"背诵": recite, "题": qs}

    out_dir = Path(args.输出) if args.输出 else root / "题目" / f"{cls}_{lesson}"
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "题库.json"
    payload = {
        "说明": data.get("说明") or f"{cls} {lesson} 按失分考点生成的题库",
        "班级": cls, "课次": lesson, "满分": full,
        "考点数": len(bank), "总题数": sum(len(v["题"]) for v in bank.values()),
        "考点": bank,
    }
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n[完成] 题库 {len(bank)} 个考点、"
          f"{payload['总题数']} 道题 → {out.relative_to(root)}")
    miss = [kp for kp, _, _ in points if kp not in bank]
    if miss:
        print(f"  [注意] 这几个考点 AI 没出到：{miss}")


if __name__ == "__main__":
    main()
