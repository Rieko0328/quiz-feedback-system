# -*- coding: utf-8 -*-
r"""错题巩固单.py —— 一页半到三页的练习题（docx）。

结构固定四块：
    一、要背的          （按失分考点，2-3 条）
    二、要做的          （选择题，题量按得分率自适应 3-12 道）
    三、答案与解析       （统一放卷末，题下只留「我的答案」）
    老师寄语

题目内容由 AI 生成后放进 questions.json（脚本只管排版）：
    python 错题巩固单.py --班级 五年级A班 --学生 谷小明 --题目 questions.json

questions.json：
{
  "背诵": [{"考点": "形容词比较级", "要点": ["tall→taller", "big→bigger（双写）"]}],
  "题目": [{"题干": "Lucy is ___ than her sister.", "选项": ["A. tall", "B. taller", "C. tallest", "D. more tall"],
            "答案": "B", "解析": "than 提示比较级，tall→taller。"}]
}
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from _ws import workspace_root, dirs_for  # noqa: E402
from fonts import docx_font_name  # noqa: E402

from docx import Document  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt, RGBColor  # noqa: E402

INK = RGBColor(0x1F, 0x29, 0x37)
SUB = RGBColor(0x6B, 0x72, 0x80)
HINT = RGBColor(0x9C, 0xA3, 0xAF)
BLUE = RGBColor(0x1B, 0x5E, 0xA8)
RED = RGBColor(0xC0, 0x39, 0x2B)


def n_questions(rate: float, wrong: int) -> int:
    """题量：按得分率定档，再按错题数微调，夹在 3~12。"""
    if rate >= 0.95:
        base = 3
    elif rate >= 0.85:
        base = 6
    elif rate >= 0.70:
        base = 8
    else:
        base = 12
    # 错得少就少出（每个错点配一道变式），不反过来把题量顶上去
    if wrong:
        base = min(base, wrong * 2)
    return max(3, min(12, base))


def set_run(run, size=11, bold=False, color=INK):
    run.font.name = docx_font_name()
    run.font.size = Pt(size)
    run.bold = bold
    run.font.color.rgb = color
    run._element.rPr.rFonts.set(qn("w:eastAsia"), docx_font_name())


def para(doc, text="", size=11, bold=False, color=INK, space_before=0,
         space_after=4, line=1.45, align=None, indent=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    if indent:
        pf.left_indent = Cm(indent)
    if align is not None:
        p.alignment = align
    if text:
        set_run(p.add_run(text), size, bold, color)
    return p


def shade(paragraph, hex_color="F5F7FA"):
    """给段落加浅底纹。"""
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def rule(doc, color="E5E7EB"):
    """一条细分隔线（用段落下边框）。"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def section_title(doc, text):
    p = para(doc, text, size=13, bold=True, color=BLUE, space_before=14, space_after=6)
    return p


# ---------------------------------------------------------------- 题库抽题
def load_bank(path) -> dict:
    """读本讲课题库：{原子考点: {"背诵": [...], "题": [...]}}"""
    p = Path(path)
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8-sig"))
    except Exception:                                  # noqa: BLE001
        return {}
    bank = data.get("考点") or data
    if not isinstance(bank, dict):
        return {}
    return {str(k).strip(): v for k, v in bank.items() if isinstance(v, dict)}


def _atoms(one: str):
    """'疑问词,划线提问' -> ['疑问词', '划线提问']"""
    out = []
    for a in str(one).replace("，", ",").split(","):
        a = a.strip()
        if a and a not in out:
            out.append(a)
    return out


def _resolve_kp(bank: dict, atom: str):
    """把答案文件里的考点名，对到题库里的 key 上。

    同一个考点，答案文件和题库的写法经常不一样：
        had better not be   ↔ had better not
        had better·否定      ↔ had better do
        一般过去时            ↔ 一般过去时·疑问句
        疑问词·How soon      ↔ 对划线部分提问·How soon

    三步：精确 → 包含 → 最长公共子串（占较短那个 60% 以上才算）。
    都找不到返回 None —— 宁可少出题，也不能出不相干的题。
    """
    if not atom:
        return None
    if atom in bank:
        return atom

    def _n(x: str) -> str:
        return (x.replace("·", "").replace(" ", "")
                 .replace("，", "").replace(",", "").strip())

    key = _n(atom)
    if not key:
        return None

    # ② 包含关系（取重叠最长的）
    best, best_len = None, 0
    for k in bank:
        kk = _n(k)
        if key in kk or kk in key:
            n = min(len(key), len(kk))
            if n > best_len:
                best, best_len = k, n
    if best:
        return best

    # ③ 最长公共子串要够长（「疑问词·How soon」vs「对划线部分提问·How soon」）
    import difflib
    cand, cand_ratio = None, 0.0
    for k in bank:
        kk = _n(k)
        sm = difflib.SequenceMatcher(None, key, kk)
        blk = sm.find_longest_match(0, len(key), 0, len(kk))
        common = blk.size
        shorter = min(len(key), len(kk)) or 1
        ratio = common / shorter
        if ratio > cand_ratio:
            cand, cand_ratio = k, ratio
    return cand if cand_ratio >= 0.6 else None


def compose_from_bank(d, bank: dict, n: int, seed: int = 0) -> dict:
    """按这个孩子的失分考点，从题库里凑出一份专属练习。

    - 背诵：每个失分考点一条，要点取它拆出来的原子考点（合并去重）
    - 题目：各考点轮转抽，`seed` 按学生序号错开，同班的两个孩子不会一模一样
    """
    if not bank:
        return {"背诵": [], "题目": []}

    lost = [x[1] for x in (d.get("失分") or []) if x[1] and x[1] != "—"]
    atoms = []
    for s in lost:
        for a in _atoms(s):
            if a not in atoms:
                atoms.append(a)

    recite, seen = [], set()
    for s in lost:
        if s in seen:
            continue
        seen.add(s)
        pts = []
        for a in _atoms(s):
            _rk = _resolve_kp(bank, a)
            for p in (bank.get(_rk) or {}).get("背诵") or []:
                if p not in pts:
                    pts.append(p)
        if pts:
            recite.append({"考点": s.replace(",", "、"), "要点": pts[:3]})

    # 原子考点 → 题库里真实的 key（可能不是同一个写法），去重后再建池子
    _mapped = {}
    for a in atoms:
        rk = _resolve_kp(bank, a)
        if rk and (bank.get(rk) or {}).get("题"):
            _mapped[rk] = list(bank[rk]["题"])
    pools = _mapped
    picked, used = [], set()
    keys = list(pools)
    rnd = 0
    while len(picked) < n and keys and rnd < 30:
        added = False
        for j, k in enumerate(keys):
            if len(picked) >= n:
                break
            pool = pools[k]
            start = (seed + rnd + j) % len(pool)
            for t in range(len(pool)):
                idx = (start + t) % len(pool)
                if (k, idx) in used:
                    continue
                used.add((k, idx))
                picked.append(pool[idx])
                added = True
                break
        if not added:
            break
        rnd += 1

    return {"背诵": recite, "题目": picked[:n]}


def build_student(d, qs, out_path: Path) -> Path:
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    style = doc.styles["Normal"]
    style.font.name = docx_font_name()
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), docx_font_name())

    # ---- 抬头 ----
    para(doc, "巩固练习", 19, True, INK, space_after=2)
    sub = f"{d.get('科目') or '英语'} · {d['卷名']} · {d['姓名']} · {d['分数']:g}/{d['满分']:g}"
    para(doc, sub, 10.5, False, SUB, space_after=2)
    rule(doc)

    # ---- 一、要背的 ----
    if qs.get("背诵"):
        section_title(doc, "一、要背的")
        for i, item in enumerate(qs["背诵"], 1):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.2
            set_run(p.add_run(f"{i}. {item['考点']}"), 11.5, True, INK)
            for pt in item.get("要点", []):
                pp = para(doc, f"· {pt}", 11, color=RGBColor(0x37, 0x41, 0x51),
                          space_after=0, indent=0.6)

    # ---- 二、要做的 ----
    # 背诵占了半页以上时另起一页，免得首页只排进一个标题
    if len(qs.get("背诵") or []) >= 3:
        doc.add_page_break()
    section_title(doc, "二、要做的")
    para(doc, f"共 {len(qs.get('题目', []))} 题　把答案写在横线上，做完再翻到最后一页对答案。",
         10.5, False, SUB, space_after=8)

    for i, q in enumerate(qs.get("题目", []), 1):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.keep_with_next = True
        set_run(p.add_run(f"{i}. "), 11.5, True, BLUE)
        set_run(p.add_run(q["题干"]), 11.5, False, INK)
        opts = q.get("选项") or []
        if opts:
            # 一行两个；两个选项都太长就一行一个，免得挤在一起
            for i in range(0, len(opts), 2):
                pair = opts[i:i + 2]
                if len(pair) > 1 and len(pair[0]) + len(pair[1]) <= 34:
                    pp = para(doc, "　　" + pair[0] + "\t" + pair[1], 11,
                              color=RGBColor(0x37, 0x41, 0x51), space_after=0,
                              indent=0.4, line=1.0)
                    pp.paragraph_format.tab_stops.add_tab_stop(Cm(7.9))
                    pp.paragraph_format.keep_with_next = True
                else:
                    for one in pair:
                        po = para(doc, "　　" + one, 11,
                                  color=RGBColor(0x37, 0x41, 0x51), space_after=0,
                                  indent=0.4, line=1.0)
                        po.paragraph_format.keep_with_next = True
        para(doc, "我的答案：____________", 11, color=SUB, space_before=2, indent=0.4)

    # ---- 三、答案与解析 ----
    doc.add_page_break()
    section_title(doc, "三、答案与解析")
    para(doc, "做完再看，标出自己错的那几道。", 10.5, False, SUB, space_after=8)
    for i, q in enumerate(qs.get("题目", []), 1):
        ans = q.get("答案", "")
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.4
        set_run(p.add_run(f"{i}. "), 11, True, INK)
        set_run(p.add_run(f"{ans}"), 11, True, RED)
        if q.get("解析"):
            set_run(p.add_run(f"　{q['解析']}"), 10.5, False, SUB)

    # ---- 老师寄语 ----
    if d.get("寄语"):
        rule(doc)
        p = para(doc, "老师的话", 11.5, True, BLUE, space_before=8, space_after=4)
        q = para(doc, d["寄语"], 10.5, color=RGBColor(0x37, 0x41, 0x51),
                 space_after=6, line=1.5)
        shade(q)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def build_blanks(d, points, out_path: Path) -> Path:
    """没有题目数据时的兜底：只出背诵 + 抄写/默写，不出选择题。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    style = doc.styles["Normal"]
    style.font.name = docx_font_name()
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), docx_font_name())

    para(doc, "巩固练习", 19, True, INK, space_after=2)
    para(doc, f"{d.get('科目') or '英语'} · {d['卷名']} · {d['姓名']} · {d['分数']:g}/{d['满分']:g}",
         10.5, False, SUB, space_after=2)
    rule(doc)
    # 「抄 3 遍：疑问词」这种任务孩子做不了——考点名不是一个能抄的词。
    # 改成两件真能执行的事：重做错题 + 把规则说一遍。
    _qs = []
    for row in (d.get("失分") or []):
        _no, _kp = (row[0], row[1]) if len(row) >= 2 else ("", "")
        if _no and str(_no) not in ("—", ""):
            _qs.append((str(_no), _kp or ""))
    # 章节编号要动态算：满分的孩子没有「错题重做」这一段，
    # 不能还硬写「二、」，不然文档从「二、」开始，家长以为缺了一页。
    _CN = "一二三四"
    _idx = 0
    if _qs:
        section_title(doc, f"{_CN[_idx]}、这次错的题，重做一遍")
        _idx += 1
        para(doc, "　对着答案自己做，做完再翻答案页对。不用抄题，写答案就行。",
             10, color=SUB, space_after=6, indent=0.4)
        for _no, _kp in _qs[:6]:
            p = para(doc, f"第 {_no} 题" + (f"（{_kp}）" if _kp else ""),
                     11.5, True, INK, space_before=6, space_after=2)
            para(doc, "　　我的答案：__________________________________________",
                 10.5, color=HINT, space_after=2, indent=0.4)

    if points:
        section_title(doc, f"{_CN[_idx]}、这几处要记准")
        _idx += 1
        para(doc, "　不用抄，读一遍，然后合上用自己的话讲给家长听。讲不通的就是没记住。",
             10, color=SUB, space_after=6, indent=0.4)
        for i, kp in enumerate(points, 1):
            para(doc, f"{i}. {kp}", 11.5, True, INK, space_before=4, space_after=2)
            para(doc, "　　讲一遍：__________________________________________",
                 10.5, color=HINT, space_after=2, indent=0.4)

    if d.get("寄语"):
        rule(doc)
        para(doc, "老师的话", 11.5, True, BLUE, space_before=8, space_after=4)
        q = para(doc, d["寄语"], 10.5, color=RGBColor(0x37, 0x41, 0x51), space_after=6, line=1.5)
        shade(q)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--班级", required=True)
    ap.add_argument("--学生", required=True)
    ap.add_argument("--课次", default=None)
    ap.add_argument("--题目", default=None, help="这个孩子专属的 questions.json（可选）")
    ap.add_argument("--题库", default=None,
                    help="本讲课题库 题库.json；没给专属题目时按他的失分考点抽题")
    ap.add_argument("--工作区", default=None)
    ap.add_argument("--输出", default=None)
    args = ap.parse_args()

    root = workspace_root(args.工作区)
    from 学情报告 import load_from_workspace, fill_text  # noqa: E402

    d = load_from_workspace(root, args.班级, args.学生, args.课次 or "")
    fill_text(d, root, args.班级)

    rate = d["分数"] / d["满分"] if d["满分"] else 0
    wrong = len(d.get("失分") or [])
    n = n_questions(rate, wrong)

    qs = {}
    if args.题目:
        qp = Path(args.题目).expanduser()
        if not qp.is_absolute():
            qp = root / qp
        if qp.exists():
            qs = json.loads(qp.read_text(encoding="utf-8-sig"))
    if not qs.get("题目") and args.题库:
        bp = Path(args.题库).expanduser()
        if not bp.is_absolute():
            bp = root / bp
        qs = compose_from_bank(d, load_bank(bp), n)
    qs.setdefault("背诵", [])
    qs.setdefault("题目", [])
    if len(qs["题目"]) > n:
        qs["题目"] = qs["题目"][:n]

    out_dir = Path(args.输出).expanduser() if args.输出 else \
        dirs_for(root, args.班级, d.get("课次", ""))["work"] / "巩固单"
    name = f"{args.学生}_巩固练习.docx"

    if qs["题目"]:
        p = build_student(d, qs, out_dir / name)
    else:
        pts = [x[1] for x in (d.get("失分") or []) if x[1] and x[1] != "—"] or \
              [m["名"] for m in d.get("模块", [])]
        p = build_blanks(d, list(dict.fromkeys(pts))[:4], out_dir / name)
    print(p)
    print(f"题量：{len(qs['题目'])} 道（得分率 {rate * 100:.0f}%，建议 {n} 道）")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
