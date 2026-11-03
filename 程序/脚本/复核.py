# -*- coding: utf-8 -*-
r"""复核.py —— 双模型交叉核对。

为什么需要：单个模型看手写答题卡都会错，而且错的地方不一样。
  我（对话里的 DeepSeek）↓  GLM 免费接口 ↓
两个人各自读一遍，**只把不一致的挑出来给老师看**。
一致的部分直接过，老师的工作量从「40 张全看」降到「可能 3-5 张」。

用法：
    # ① 我读完照片后，把结果写成 JSON（格式见下），然后对比
    python 复核.py --班级 周五慢班 --课次 NCE1-L9-14月考 --我的 我的读取.json

    # ② 只想看 GLM 读成什么（不对比）
    python 复核.py --班级 周五慢班 --课次 NCE1-L9-14月考 --只跑GLM

我的读取.json 格式（只写分数，姓名可省略）：
    {"霍小红": 27, "马小晴": 31, "万小明": 29}
    或带姓名容错：{"IMG_001.jpg": {"姓名": "霍小红", "分数": 27}}

输出：生成结果/<班级>_<课次>/复核差异表.xlsx
     一致的行绿底、不一致的红底，老师只看红的。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _ws import workspace_root, safe_name, photo_root  # noqa: E402

IMPORT = importlib = None  # 占位，下面按需导入


def _load_rec():
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "rec", Path(__file__).resolve().parent / "识别照片.py")
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def sections(meta: dict):
    names = [x.strip() for x in re.split(r"[,，、]", meta.get("大题名称", "")) if x.strip()]
    counts = []
    for x in re.split(r"[,，、]", meta.get("大题题数", "")):
        try:
            counts.append(int(x))
        except ValueError:
            counts.append(0)
    return names or ["全卷"], counts


def read_meta(root: Path, cls: str):
    p = root / "班级名单" / f"{cls}.txt"
    meta = {}
    if not p.exists():
        return meta
    for raw in p.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if s.startswith("@") and ":" in s:
            k, v = s[1:].split(":", 1)
            meta[k.strip()] = v.strip()
    return meta


def pick_lesson(root: Path, cls: str, lesson: str | None):
    meta = read_meta(root, cls)
    if lesson:
        return meta, lesson
    lesson = meta.get("课次", "")
    if not lesson or lesson.startswith("请填写"):
        raise SystemExit(f"[错] 名单里没有 @课次，请用 --课次 指定")
    return meta, lesson


EXT = (".jpg", ".jpeg", ".png", ".webp", ".bmp")


def find_photos(root: Path, cls: str, lesson: str):
    """照片可能在 <课次>/ 下，也可能在 <课次>/原图/ 下，两种都认。"""
    base = photo_root(root) / safe_name(cls) / safe_name(lesson)
    if not base.is_dir():
        raise SystemExit(f"[错] 照片目录不存在：{base}")
    for d in (base, base / "原图", base / "已改名"):
        if not d.is_dir():
            continue
        out = [p for p in sorted(d.glob("*")) if p.suffix.lower() in EXT]
        if out:
            if d != base:
                print(f"  [提示] 照片在子目录：{d.name}/")
            return out
    raise SystemExit(f"[错] {base} 里没有照片（找过 / 原图/ 已改名/）")


def run_glm(root: Path, photos, sec_names, workers=4):
    """调 GLM 读一遍，返回 {文件名: {姓名, 分数}}"""
    rec = _load_rec()
    import os
    old = os.getcwd()
    os.chdir(str(root))
    try:
        cfg = rec.load_cfg()
    finally:
        os.chdir(old)
    if not cfg:
        raise SystemExit("[错] 读不到 识别配置.txt 里的 API 配置")

    print(f"  模型：{cfg.get('api_model')}　并发 {workers}")
    out = {}

    def one(item):
        idx, p = item
        try:
            b64 = rec.encode(p, 1600)
            pr = rec.build_prompt(sec_names, True)
            for attempt in range(2):
                try:
                    txt, reasoning = rec.call_api(cfg, pr, b64, timeout=180)
                    d = rec.parse_json(txt) or rec.parse_json(reasoning)
                    if d:
                        return p.name, d
                except Exception as e:
                    if attempt == 1:
                        return p.name, {"_错误": f"{type(e).__name__}"}
                    time.sleep(1.2)
            return p.name, {"_错误": "没解析出 JSON"}
        except Exception as e:
            return p.name, {"_错误": f"{type(e).__name__}: {e}"}

    with ThreadPoolExecutor(max_workers=workers) as ex:
        for name, d in ex.map(one, list(enumerate(photos, 1))):
            out[name] = d
            sc = d.get("分数", "")
            print(f"    GLM {name[:26]:26s} {str(d.get('姓名', ''))[:8]:8s} {sc}")
    return out


def norm_name(x):
    return re.sub(r"\s+", "", str(x or "")).replace("A", "A").strip()


def main():
    ap = argparse.ArgumentParser(description="双模型交叉核对")
    ap.add_argument("--班级", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--我的", default=None, help="我（对话 AI）读出来的 JSON")
    ap.add_argument("--只跑GLM", action="store_true")
    ap.add_argument("--复用", action="store_true", help="复用上次 GLM 结果，不重新调")
    args = ap.parse_args()

    root = workspace_root(args.工作区)
    cls = safe_name(args.班级)
    meta, lesson = pick_lesson(root, cls, args.课次)
    photos = find_photos(root, cls, lesson)
    sec_names, _ = sections(meta)
    outdir = root / "生成结果" / f"{cls}_{lesson}"
    outdir.mkdir(parents=True, exist_ok=True)
    cache = outdir / "GLM读取.json"

    print(f"班级 {cls}　课次 {lesson}　照片 {len(photos)} 张")
    print(f"大题：{'、'.join(sec_names)}")

    if args.复用 and cache.exists():
        glm = json.loads(cache.read_text(encoding="utf-8-sig"))
        print(f"  复用上次结果：{cache.name}")
    else:
        glm = run_glm(root, photos, sec_names)
        cache.write_text(json.dumps(glm, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"  GLM 结果已存：{cache.relative_to(root)}")

    if args.只跑GLM or not args.我的:
        print("\n（没给 --我的，只出 GLM 结果）")
        if not args.我的:
            print("  下一步：我把读到的分数写成 JSON 再跑一次，就能出差异表")
        return

    mine = json.loads(Path(args.我的).read_text(encoding="utf-8-sig"))
    print(f"\n对比：我 {len(mine)} 条 vs GLM {len(glm)} 条")

    # 统一成「键 → (姓名, 分数)」
    def norm_glm(d):
        if "_错误" in d:
            return None, None, d["_错误"]
        try:
            sc = float(d.get("分数"))
        except (TypeError, ValueError):
            sc = None
        return norm_name(d.get("姓名")), sc, ""

    def norm_mine(v):
        if isinstance(v, dict):
            try:
                sc = float(v.get("分数"))
            except (TypeError, ValueError):
                sc = None
            return norm_name(v.get("姓名")), sc
        try:
            return None, float(v)
        except (TypeError, ValueError):
            return None, None

    rows = []
    all_keys = list(glm.keys())
    # 我的结果可能按姓名给，也可能按文件名给
    used = set()
    for fname in all_keys:
        g_name, g_score, g_err = norm_glm(glm[fname])
        m_name, m_score = None, None
        # ① 先按文件名匹配
        if fname in mine:
            m_name, m_score = norm_mine(mine[fname])
            used.add(fname)
        # ② 再按姓名匹配（我用姓名写的）
        if m_score is None and g_name:
            for k, v in mine.items():
                if k in used:
                    continue
                kn, ks = norm_mine(v)
                if kn and kn == g_name or k == g_name:
                    m_name, m_score = (kn or k), ks
                    used.add(k)
                    break
        same = (m_score is not None and g_score is not None
                and abs(m_score - g_score) < 0.01)
        note = ""
        if g_err:
            note = f"GLM 失败：{g_err}"
        elif m_score is None:
            note = "我这边没读到"
        elif not same:
            note = f"差 {abs((m_score or 0) - (g_score or 0)):g} 分"
        rows.append({
            "照片": fname, "我的姓名": m_name or "", "GLM姓名": g_name or "",
            "我的分数": m_score, "GLM分数": g_score,
            "一致": "✔" if same else "✘", "说明": note,
        })

    # 我的结果里有、GLM 没有的
    for k, v in mine.items():
        if k in used:
            continue
        mn, ms = norm_mine(v)
        rows.append({"照片": "（GLM 没读到）", "我的姓名": mn or k, "GLM姓名": "",
                     "我的分数": ms, "GLM分数": None, "一致": "✘",
                     "说明": "只有我读到了"})

    diff = [r for r in rows if r["一致"] != "✔"]
    print(f"  一致 {len(rows)-len(diff)} 条　不一致 {len(diff)} 条")

    # ---- 出表 ----
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "复核差异"
    ws.sheet_view.showGridLines = False
    FB = "微软雅黑"
    THIN = Side(style="thin", color="D0D7E2")
    BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = (f"复核差异表　{cls} / {lesson}　"
               f"一致 {len(rows)-len(diff)} 张 · 不一致 {len(diff)} 张")
    c.font = Font(name=FB, size=13, bold=True, color="1B5EA8")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:G2")
    c = ws["A2"]
    c.value = "只看红底的。绿底的是我和 GLM 读出来一样的，可以直接过。"
    c.font = Font(name=FB, size=10, color="6B7684")
    c.alignment = Alignment(horizontal="center")

    HEAD = ["照片", "我的姓名", "GLM姓名", "我的分数", "GLM分数", "一致", "说明"]
    for i, h in enumerate(HEAD, 1):
        cc = ws.cell(row=4, column=i, value=h)
        cc.font = Font(name=FB, size=10, bold=True, color="FFFFFF")
        cc.fill = PatternFill("solid", fgColor="1B5EA8")
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc.border = BD

    for i, r in enumerate(rows, 5):
        for j, k in enumerate(["照片", "我的姓名", "GLM姓名", "我的分数",
                               "GLM分数", "一致", "说明"], 1):
            cc = ws.cell(row=i, column=j, value=r.get(k))
            cc.font = Font(name=FB, size=10)
            cc.border = BD
            cc.alignment = Alignment(horizontal="center" if j != 7 else "left")
            if r["一致"] != "✔":
                cc.fill = PatternFill("solid", fgColor="FDE8E8")
            else:
                cc.fill = PatternFill("solid", fgColor="E8F5EC")
        ws.row_dimensions[i].height = 22

    for col, w in zip("ABCDEFG", [26, 12, 12, 10, 10, 7, 22]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A5"

    xlsx = outdir / "复核差异表.xlsx"
    wb.save(xlsx)
    print(f"\n[完成] {xlsx.relative_to(root)}")
    if diff:
        print(f"\n  要你核对的 {len(diff)} 张：")
        for r in diff[:12]:
            print(f"    {r['照片'][:24]:24s} 我 {r['我的分数']} vs GLM {r['GLM分数']}　{r['说明']}")
    else:
        print("\n  两个模型完全一致，可以直接用。")


if __name__ == "__main__":
    main()
