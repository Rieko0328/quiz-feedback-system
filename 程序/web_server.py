# -*- coding: utf-8 -*-
r"""网页版.py —— 课测反馈系统的本地网页界面。

纯标准库实现，不依赖 Flask。启动后浏览器打开 http://127.0.0.1:8777

    python 网页版.py            （默认端口 8777）
    python 网页版.py --端口 8888
"""
from __future__ import annotations

import argparse
import functools
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
print = functools.partial(print, flush=True)

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "程序" / "脚本"
DIR_ROSTER = ROOT / "班级名单"
DIR_ANSWER = ROOT / "课测答案"
DIR_PHOTO = ROOT / "答题卡照片"
DIR_OUT = ROOT / "生成结果"
CFG = ROOT / "识别配置.txt"
WEB = Path(__file__).resolve().parent / "web"

JOBS: dict[str, dict] = {}
JOB_SEQ = [0]
JOB_LOCK = threading.Lock()


# ---------------------------------------------------------------- 工具
def roster_path(cls: str) -> Path | None:
    for f in DIR_ROSTER.glob("*.txt"):
        if f.name.endswith(".bak"):
            continue
        if f.stem == cls or cls in f.stem:
            return f
    return None


def read_roster_raw(p: Path) -> str:
    return p.read_text(encoding="utf-8-sig")


def parse_roster(p: Path):
    meta, students = {}, []
    for raw in read_roster_raw(p).splitlines():
        s = raw.strip()
        if not s or s.startswith("#"):
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


def count_photos(cls: str) -> tuple[int, list[str]]:
    base = DIR_PHOTO / cls
    if not base.exists():
        return 0, []
    lessons, n = [], 0
    for d in sorted(base.iterdir()):
        if not d.is_dir() or d.name in {"已改名", "分发", "照片归档"}:
            continue
        sub = d / "已改名"
        target = sub if sub.exists() and any(sub.glob("*.jpg")) else d
        c = len(list(target.glob("*.jpg"))) + len(list(target.glob("*.png"))) \
            + len(list(target.glob("*.jpeg")))
        if c:
            lessons.append(d.name)
            n += c
    return n, lessons


def load_config() -> dict:
    cfg = {}
    if CFG.exists():
        for line in CFG.read_text(encoding="utf-8").splitlines():
            t = line.strip()
            if not t or t.startswith("#") or "=" not in t:
                continue
            k, v = t.split("=", 1)
            cfg[k.strip()] = v.strip()
    return cfg


def save_config(cfg: dict):
    lines = ["# 识别配置（网页版可直接在这里改，也可以编辑这个文件）"]
    for k, v in cfg.items():
        lines.append(f"{k} = {v}")
    CFG.write_text("\n".join(lines) + "\n", encoding="utf-8")


def new_job(title: str) -> str:
    with JOB_LOCK:
        JOB_SEQ[0] += 1
        jid = f"j{JOB_SEQ[0]}"
        JOBS[jid] = {"id": jid, "title": title, "status": "running",
                     "log": [], "started": time.time(), "ended": None}
    return jid


def job_log(jid: str, line: str):
    with JOB_LOCK:
        if jid in JOBS:
            JOBS[jid]["log"].append(line)


def job_done(jid: str, status: str = "done"):
    with JOB_LOCK:
        if jid in JOBS:
            JOBS[jid]["status"] = status
            JOBS[jid]["ended"] = time.time()


def run_stream(jid: str, args: list[str], cwd: Path):
    """跑子进程，逐行推到 job 日志。"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    try:
        p = subprocess.Popen([sys.executable] + args, cwd=str(cwd), env=env,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                             text=True, encoding="utf-8", errors="replace", bufsize=1)
        for line in p.stdout:
            line = line.rstrip()
            if line:
                job_log(jid, line)
        p.wait()
        job_log(jid, f"[退出码 {p.returncode}] {Path(args[0]).name}")
        job_done(jid, "done" if p.returncode == 0 else "error")
    except Exception as e:
        job_log(jid, f"[异常] {type(e).__name__}: {e}")
        job_done(jid, "error")


def subprocess_json(args: list[str]) -> dict:
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    r = subprocess.run([sys.executable] + args, cwd=str(ROOT), env=env,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return {"code": r.returncode, "out": r.stdout, "err": r.stderr}


# ---------------------------------------------------------------- 路由
class API:
    @staticmethod
    def state():
        classes = []
        for f in sorted(DIR_ROSTER.glob("*.txt")):
            if f.name.endswith(".bak") or f.name.startswith("_"):
                continue
            if f.stem in ("姓名对照", "班级清单"):
                continue
            cls = f.stem
            meta, students = parse_roster(f)
            if not meta.get("班级"):
                continue  # 不是名单文件
            scored = [s for s in students if s["分数"] is not None]
            nph, lessons = count_photos(cls)
            outdirs = [d for d in DIR_OUT.glob(f"{cls}_*") if d.is_dir()] if DIR_OUT.exists() else []
            classes.append({
                "name": cls,
                "lesson": meta.get("课次", ""),
                "满分": meta.get("满分", "40"),
                "人数": len(students),
                "已填分": len(scored),
                "缺考": len(students) - len(scored),
                "平均": round(sum(s["分数"] for s in scored) / len(scored), 1) if scored else None,
                "照片数": nph,
                "课次目录": lessons,
                "输出": outdirs[0].name if outdirs else "",
            })
        return {"classes": classes, "config": load_config(), "root": str(ROOT)}

    @staticmethod
    def klass(name: str):
        p = roster_path(name)
        if not p:
            return {"error": f"找不到名单 {name}"}
        meta, students = parse_roster(p)
        nph, lessons = count_photos(name)
        photos = []
        base = DIR_PHOTO / name
        for d in sorted(base.iterdir()) if base.exists() else []:
            if d.is_dir() and d.name not in {"已改名", "分发", "照片归档"}:
                sub = d / "已改名"
                target = sub if sub.exists() and any(sub.glob("*.jpg")) else d
                photos += [x.name for x in sorted(target.glob("*.jpg"))]
        outdir = DIR_OUT / f"{name}_{meta.get('课次', '')}"
        files = []
        if outdir.exists():
            for x in sorted(outdir.iterdir(), key=lambda y: (y.is_file(), y.name)):
                if x.is_dir():
                    n = len(list(x.rglob("*")))
                    files.append({"name": x.name + "/", "size": 0, "dir": True,
                                  "n": n})
                else:
                    files.append({"name": x.name, "size": x.stat().st_size,
                                  "dir": False})
        return {"name": name, "meta": meta, "roster_raw": read_roster_raw(p),
                "students": students, "photos": photos, "lessons": lessons,
                "outdir": str(outdir), "files": files,
                "table": API.read_table(name, meta.get("课次", ""))}

    @staticmethod
    def read_table(cls: str, lesson: str):
        import openpyxl
        p = DIR_OUT / f"{cls}_{lesson}" / "识别核对表.xlsx"
        if not p.exists():
            return {"path": str(p), "exists": False, "head": [], "rows": []}
        ws = openpyxl.load_workbook(p).active
        rows = [[("" if c is None else c) for c in r]
                for r in ws.iter_rows(values_only=True)]
        return {"path": str(p), "exists": True,
                "head": rows[0] if rows else [], "rows": rows[1:]}

    @staticmethod
    def save_table(cls: str, lesson: str, head: list, rows: list):
        import openpyxl
        from openpyxl.styles import Font, PatternFill
        outdir = DIR_OUT / f"{cls}_{lesson}"
        outdir.mkdir(parents=True, exist_ok=True)
        p = outdir / "识别核对表.xlsx"
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "核对表"
        ws.append(head)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1B5EA8")
        for r in rows:
            ws.append(r)
        wb.save(p)
        return {"ok": True, "path": str(p)}


def route_get(path: str, q: dict):
    if path in ("/", "/index.html"):
        return None  # 静态文件
    if path == "/api/state":
        return API.state()
    if path == "/api/class":
        return API.klass(q.get("name", ""))
    if path == "/api/job":
        jid = q.get("id", "")
        with JOB_LOCK:
            j = JOBS.get(jid)
            if not j:
                return {"error": "no such job"}
            from_ = int(q.get("from", 0))
            return {"id": jid, "status": j["status"], "title": j["title"],
                    "log": j["log"][from_:], "next": len(j["log"])}
    if path == "/api/output":
        cls = q.get("name", "")
        outdir = Path(q.get("dir", ""))
        if not outdir.exists():
            return {"files": []}
        return {"files": [{"name": x.name, "dir": x.is_dir(),
                           "size": x.stat().st_size if x.is_file() else 0}
                          for x in sorted(outdir.iterdir())]}
    return {"error": "unknown api"}


def route_post(path: str, body: dict):
    if path == "/api/class/save":
        p = roster_path(body.get("name", ""))
        if not p:
            return {"error": "找不到名单"}
        p.write_text(body.get("roster_raw", ""), encoding="utf-8")
        return {"ok": True}
    if path == "/api/table/save":
        return API.save_table(body["name"], body.get("lesson", ""),
                              body["head"], body["rows"])
    if path == "/api/config":
        save_config(body.get("config", {}))
        return {"ok": True}

    if path == "/api/newclass":
        name = (body.get("name") or "").strip()
        if not name:
            return {"error": "班级名不能为空"}
        jid = new_job(f"新建班级：{name}")
        jid_box = jid
        args = [str(SCRIPTS / "新建班级.py"), "--班级", name,
                "--课次", body.get("lesson", "未命名课次"),
                "--满分", str(body.get("full", "40")),
                "--考试名称", body.get("exam", "月考试卷"),
                "--称呼", body.get("call", "家长"),
                "--风格", body.get("style", "亲和"),
                "--大题名称", body.get("secs", "单项选择,选词填空,按要求改写句子,阅读理解"),
                "--大题题数", body.get("counts", "20,5,5,5")]
        if body.get("focus"):
            args += ["--重点线", str(body["focus"])]
        threading.Thread(target=lambda: run_stream(jid_box, args, ROOT),
                         daemon=True).start()
        return {"job": jid}

    if path == "/api/oneclick":
        name = body.get("name", "")
        tpl = body.get("template", "")
        jid = new_job(f"一键生成：{name}")
        jid_box = jid
        args = [str(SCRIPTS / "一键生成.py"), "--班级", name]
        if tpl:
            args += ["--模板", tpl]
        threading.Thread(target=lambda: run_stream(jid_box, args, ROOT),
                         daemon=True).start()
        return {"job": jid}

    if path == "/api/answer":
        lesson = (body.get("lesson") or "").strip()
        if not lesson:
            return {"error": "先填「这次考什么」"}
        f = DIR_ANSWER / f"{lesson}.txt"
        if body.get("get"):
            return {"exists": f.exists(),
                    "text": f.read_text(encoding="utf-8") if f.exists() else "",
                    "path": str(f)}
        text = body.get("text")
        if text is None:
            return {"error": "没有内容"}
        DIR_ANSWER.mkdir(exist_ok=True)
        if f.exists():
            f.with_suffix(".txt.bak").write_text(
                f.read_text(encoding="utf-8"), encoding="utf-8")
        f.write_text(text, encoding="utf-8")
        n = sum(1 for l in text.splitlines()
                if l.strip() and not l.strip().startswith(("#", "@")))
        return {"ok": True, "path": str(f), "题数": n}

    if path == "/api/upload":
        # 浏览器原生文件对话框选的文件，前端读成 base64 传过来
        import base64
        import tempfile
        nm = (body.get("filename") or "上传文件").strip()
        b64 = body.get("b64") or ""
        kind = body.get("kind") or "temp"
        if not b64:
            return {"error": "文件是空的"}
        try:
            raw = base64.b64decode(b64)
        except Exception as e:
            return {"error": f"文件解析失败：{e}"}
        if len(raw) > 30 * 1024 * 1024:
            return {"error": "文件太大了（超过 30MB）"}
        if kind == "answer":
            lesson = (body.get("lesson") or Path(nm).stem).strip()
            DIR_ANSWER.mkdir(exist_ok=True)
            dst = DIR_ANSWER / f"{lesson}.txt"
            try:
                text = raw.decode("utf-8")
            except UnicodeDecodeError:
                try:
                    text = raw.decode("gbk")
                except UnicodeDecodeError:
                    return {"error": "答案文件读不出来，"
                                     "请存成 UTF-8 或 GBK 编码的 txt"}
            dst.write_text(text, encoding="utf-8")
            n = sum(1 for l in text.splitlines()
                    if l.strip() and not l.strip().startswith(("#", "@")))
            return {"ok": True, "path": str(dst), "题数": n,
                    "text": text}
        tmp = Path(tempfile.gettempdir()) / "课测反馈_上传"
        tmp.mkdir(exist_ok=True)
        dst = tmp / nm
        dst.write_bytes(raw)
        return {"ok": True, "path": str(dst)}

    if path == "/api/progress":
        import csv
        hv = ROOT / "生成结果" / "历史"
        if not hv.exists():
            return {"error": "还没有成绩记录。生成一次成绩就会自动存档。"}
        recs = []
        for f in sorted(hv.glob("*.csv")):
            try:
                with open(f, encoding="utf-8-sig", newline="") as fh:
                    for r in csv.DictReader(fh):
                        recs.append(r)
            except Exception:
                continue
        if not recs:
            return {"error": "历史文件是空的"}

        want = (body.get("name") or "").strip()
        if want:
            recs = [r for r in recs if (r.get("班级") or "") == want]

        def num(x):
            try:
                return float(str(x).strip())
            except (TypeError, ValueError):
                return None

        # 考试顺序：按日期 + 课次去重保序
        exams, seen = [], set()
        for r in recs:
            key = (r.get("日期", ""), r.get("课次", ""))
            if key not in seen and r.get("课次"):
                seen.add(key)
                exams.append({"日期": r.get("日期", ""), "课次": r.get("课次", "")})

        # 按学生归组
        stu = {}
        for r in recs:
            nm = (r.get("姓名") or "").strip()
            sc = num(r.get("分数"))
            if not nm or sc is None:
                continue
            stu.setdefault(nm, {})[(r.get("日期", ""), r.get("课次", ""))] = sc

        out = []
        for nm, m in stu.items():
            seq, prev = [], None
            for e in exams:
                k = (e["日期"], e["课次"])
                v = m.get(k)
                delta = None
                if v is not None and prev is not None:
                    delta = round(v - prev, 1)
                seq.append({"日期": e["日期"], "课次": e["课次"],
                            "分数": v, "变化": delta})
                if v is not None:
                    prev = v
            vals = [x["分数"] for x in seq if x["分数"] is not None]
            last = vals[-1] if vals else None
            first = vals[0] if vals else None
            out.append({
                "姓名": nm, "历次": seq,
                "最近": last, "首次": first,
                "总变化": (round(last - first, 1)
                          if last is not None and first is not None
                          and len(vals) > 1 else None),
                "次数": len(vals),
                "最高": max(vals) if vals else None,
                "最低": min(vals) if vals else None,
            })
        out.sort(key=lambda x: (x["总变化"] is None, -(x["总变化"] or 0)))
        classes = sorted({(r.get("班级") or "") for r in recs if r.get("班级")})
        return {"ok": True, "exams": exams, "students": out,
                "classes": classes, "班级": want,
                "条数": len(recs)}

    if path == "/api/rename":
        name = body.get("name", "")
        lesson = body.get("lesson", "")
        trial = bool(body.get("trial"))
        jid = new_job(f"照片自动改名：{name}")
        jid_box = jid
        args = [str(SCRIPTS / "识别照片.py"), "--班级", name, "--改名"]
        if lesson:
            args += ["--课次", lesson]
        if body.get("limit"):
            args += ["--限", str(body["limit"])]
        if trial:
            args += ["--试运行"]
        threading.Thread(target=lambda: run_stream(jid_box, args, ROOT),
                         daemon=True).start()
        return {"job": jid}

    if path == "/api/recognize":
        name = body.get("name", "")
        lesson = body.get("lesson", "")
        limit = int(body.get("limit") or 0)
        jid = new_job(f"识别照片：{name}")
        jid_box = jid

        def work():
            args = [str(SCRIPTS / "识别照片.py"), "--班级", name]
            if lesson:
                args += ["--课次", lesson]
            if limit:
                args += ["--限", str(limit)]
            run_stream(jid_box, args, ROOT)
        threading.Thread(target=work, daemon=True).start()
        return {"job": jid}

    if path == "/api/merge":
        name = body.get("name", "")
        lesson = body.get("lesson", "")
        jid = new_job(f"并入名单：{name}")
        jid_box = jid

        def work():
            tab = DIR_OUT / f"{name}_{lesson}" / "识别核对表.xlsx"
            run_stream(jid_box, [str(SCRIPTS / "识别照片.py"), "--班级", name,
                                 "--并入名单", "--核对表", str(tab)], ROOT)
        threading.Thread(target=work, daemon=True).start()
        return {"job": jid}

    if path == "/api/run":
        name = body.get("name", "")
        scripts = body.get("scripts") or []
        jid = new_job(f"生成：{name}")
        jid_box = jid

        def work():
            for s in scripts:
                job_log(jid_box, f"──── {s} ────")
                run_stream(jid_box, [str(SCRIPTS / s), "--班级", name], ROOT)
        threading.Thread(target=work, daemon=True).start()
        return {"job": jid}

    if path == "/api/fill":
        name = body.get("name", "")
        tpl = body.get("template", "")
        jid = new_job(f"填模板：{name}")
        jid_box = jid
        threading.Thread(target=lambda: run_stream(
            jid_box, [str(SCRIPTS / "填模板.py"), "--模板", tpl], ROOT),
            daemon=True).start()
        return {"job": jid}

    if path == "/api/open":
        target = Path(body.get("path", ""))
        if not target.exists():
            return {"error": f"这个文件/文件夹不在了：{target}"}
        errs = []
        for name, fn in (
            ("os.startfile", lambda: os.startfile(str(target))),
            ("explorer", lambda: subprocess.Popen(["explorer",
                                                   str(target)])),
            ("cmd start", lambda: subprocess.Popen(
                ["cmd", "/c", "start", "", str(target)])),
        ):
            try:
                fn()
                return {"ok": True}
            except Exception as e:
                errs.append(f"{name}: {type(e).__name__}")
        return {"error": "打不开（可能是权限问题）。请手动到这个位置：\n"
                         + str(target) + "\n" + "；".join(errs)}
    return {"error": "unknown api"}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, data, ctype="application/json; charset=utf-8"):
        if isinstance(data, (dict, list)):
            data = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        elif isinstance(data, str):
            data = data.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
        if u.path.startswith("/api/"):
            try:
                return self._send(200, route_get(u.path, q))
            except Exception:
                return self._send(500, {"error": traceback.format_exc()})
        # 静态
        rel = u.path.lstrip("/") or "index.html"
        f = (WEB / rel).resolve()
        if not str(f).startswith(str(WEB.resolve())) or not f.is_file():
            return self._send(404, "not found", "text/plain; charset=utf-8")
        ctype = {".html": "text/html; charset=utf-8",
                 ".js": "application/javascript; charset=utf-8",
                 ".css": "text/css; charset=utf-8"}.get(f.suffix, "application/octet-stream")
        return self._send(200, f.read_bytes(), ctype)

    def do_POST(self):
        u = urllib.parse.urlparse(self.path)
        n = int(self.headers.get("Content-Length") or 0)
        try:
            body = json.loads(self.rfile.read(n).decode("utf-8")) if n else {}
        except Exception:
            body = {}
        if not u.path.startswith("/api/"):
            return self._send(404, {"error": "not found"})
        try:
            return self._send(200, route_post(u.path, body))
        except Exception:
            return self._send(500, {"error": traceback.format_exc()})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--端口", type=int, default=8777)
    ap.add_argument("--不打开浏览器", action="store_true")
    args = ap.parse_args()
    if not (WEB / "index.html").exists():
        print("[错误] 找不到 web\\index.html，这个包可能没解压完整。")
        print(f"       当前目录：{ROOT}")
        input("按回车退出...")
        return 1

    # 端口被占用就往后试
    srv, port = None, args.端口
    for p in range(args.端口, args.端口 + 20):
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            port = p
            break
        except OSError:
            print(f"[提示] 端口 {p} 被占用，换一个…")
    if srv is None:
        print("[错误] 20 个端口都被占用了，请关掉别的程序再试。")
        input("按回车退出...")
        return 1

    url = f"http://127.0.0.1:{port}/"
    print()
    print("=" * 60)
    print("   课测反馈系统 · 网页版  已启动")
    print()
    print(f"   浏览器地址：  {url}")
    print()
    print("   如果浏览器没有自动打开，把上面这行地址")
    print("   复制到浏览器（Edge / Chrome 都行）地址栏回车。")
    print()
    print("   ★ 这个黑窗口别关，关了服务就停了。")
    print("=" * 60)
    print()
    if not args.不打开浏览器:
        def _open():
            # 依次尝试几种方式，Windows 上 os.startfile 最稳
            tries = [
                ("os.startfile", lambda: os.startfile(url)),
                ("webbrowser", lambda: webbrowser.open(url)),
                ("cmd start", lambda: subprocess.Popen(
                    ["cmd", "/c", "start", "", url],
                    shell=False, stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL)),
            ]
            for name, fn in tries:
                try:
                    fn()
                    return
                except Exception as e:
                    print(f"[尝试] {name} 打开浏览器失败：{e}")
            print(f"[提示] 自动打开浏览器失败，请手动访问：{url}")
        threading.Timer(0.6, _open).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    return 0


if __name__ == "__main__":
    main()
