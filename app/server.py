# Local window for PIAS Export. Binds 127.0.0.1 only, serves app\ui and runs the export queue:
# one report at a time, each as a child process (worker.py) so a run can be stopped and a crash never
# takes the window down. Started by "PIAS Export.bat" with pyembed\pythonw.exe (no console window).
#   python.exe app\server.py --diag   -> same server with the log in a console (for troubleshooting)
import os, re, sys, json, time, glob, codecs, shutil, socket, zipfile, ctypes, tempfile, threading, subprocess, webbrowser
import urllib.request, urllib.error
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C
from common import APP, ROOT, DUMP, OUTPUT, RUNS
import pias_engine as E

PORT    = 8716
HOST    = "127.0.0.1"
NAME    = "Pias Export - Nam Doo"
UI      = os.path.join(APP, "ui")
VERFILE = os.path.join(ROOT, "VERSION")
TOKFILE = os.path.join(DUMP, "ui_token.txt")
LOGFILE = os.path.join(DUMP, "server.log")
BAT     = os.path.join(ROOT, "PIAS Export.bat")
DIAG    = "--diag" in sys.argv

PY = [sys.executable] if sys.executable else ["py", "-3.12"]
if PY[0].lower().endswith("pythonw.exe"):          # children need a real stdout to stream from
    pc = os.path.join(os.path.dirname(PY[0]), "python.exe")
    if os.path.exists(pc): PY = [pc]
NO_WINDOW = 0x08000000 if os.name == "nt" else 0
ANSI = re.compile(r"\x1b\[[0-9;]*[A-Za-z]")


def log(msg):
    line = time.strftime("%H:%M:%S ") + str(msg)
    if DIAG: print(line, flush=True)
    try:
        os.makedirs(DUMP, exist_ok=True)
        with open(LOGFILE, "a", encoding="utf-8") as f: f.write(line + "\n")
    except Exception: pass


def box(msg, title=NAME):
    """Fatal error popup: pythonw has no console, so a message box is the only way to be seen."""
    if DIAG: print(msg, flush=True); return
    try: ctypes.windll.user32.MessageBoxW(0, str(msg), title, 0x10)
    except Exception: pass


def version():
    try: return open(VERFILE, encoding="utf-8").read().strip() or "0.0.0"
    except Exception: return "0.0.0"


def policy_volume():
    """How many policies a line writes in a month, measured from the exports already made."""
    try:
        return json.load(open(C.POLICY_VOLUME, encoding="utf-8")).get("rows_per_month", {})
    except Exception:
        return {}


def pias_catalog():
    """The report list of each entry of the PIAS 'Báo cáo' menu, as read off the screen."""
    try:
        c = json.load(open(C.PIAS_CATALOG, encoding="utf-8"))
    except Exception:
        return {"entries": []}
    for e in c.get("entries", []):
        e["ready"] = bool(e.get("access")) and bool(e.get("reports"))
    return c


def chunk_months(y1, m1, y2, m2, how="quarter"):
    """Cut a period into the pieces PIAS is asked for: one month, one quarter or one year at a time."""
    out, y, m = [], y1, m1
    while (y, m) <= (y2, m2) and len(out) < 120:
        if how == "month": end = m
        elif how == "year": end = 12
        else: end = ((m - 1) // 3 + 1) * 3
        if y == y2: end = min(end, m2)
        out.append((y, m, end))
        m = end + 1
        if m > 12: y, m = y + 1, 1
    return out


def session_year():
    """The year PIAS is logged in with, as the login script last wrote it."""
    try:
        return int(open(os.path.join(DUMP, "pias_year.txt"), encoding="utf-8").read().strip())
    except Exception:
        return 0


# ------------------------------------------------------------------ export queue
class Queue:
    """Reports wait in order; one worker thread runs them one at a time."""
    def __init__(self):
        self.items, self.lines = [], []
        self.lock = threading.Lock()
        self.thread, self.stop_after, self.proc, self.cur = None, False, None, None
        self.seq = 0

    def say(self, text):
        with self.lock:
            self.lines.append(time.strftime("%H:%M:%S  ") + text)
            del self.lines[:-500]

    def add(self, year, month, nums, folder_tpl, revenue_only, entry=None):
        entry = entry or E.HQKD
        allowed = {r["num"] for r in E.catalog().get(entry, [])} or set(E.REPORTS)
        added = 0
        with self.lock:
            for n in nums:
                n = int(n)
                if n not in allowed: continue
                if any(i["status"] == "waiting" and (i["entry"], i["year"], i["month"], i["num"])
                       == (entry, year, month, n) for i in self.items):
                    continue
                self.seq += 1
                self.items.append({"id": self.seq, "entry": entry, "year": year, "month": month, "num": n,
                                   "name": E.report_name(entry, n),
                                   "folder": C.folder_for(folder_tpl, year, month), "revenue_only": bool(revenue_only),
                                   "status": "waiting", "seconds": None, "file": "", "error": "", "revenue": None, "progress": None})
                added += 1
        return added

    def add_policies(self, chunks, lines, folder):
        """chunks: [(year, m1, m2)], lines: ['01', '05', ...]; one job per chunk and line."""
        added = 0
        with self.lock:
            for (y, m1, m2) in chunks:
                for ln in lines:
                    ln = str(ln).strip()
                    if ln not in E.LINES: continue
                    if any(i["status"] == "waiting" and i.get("kind") == "policy" and
                           (i["year"], i.get("m1"), i.get("m2"), i.get("line")) == (y, m1, m2, ln)
                           for i in self.items):
                        continue
                    self.seq += 1
                    span = "T%d" % m1 if m1 == m2 else ("cả năm" if (m1, m2) == (1, 12) else "T%d-T%d" % (m1, m2))
                    self.items.append({"id": self.seq, "kind": "policy", "entry": "Tìm kiếm đơn Bảo hiểm",
                                       "year": y, "month": m1, "m1": m1, "m2": m2, "num": 0, "line": ln,
                                       "name": "%s %s - %s" % (ln, span, E.LINES[ln]), "folder": folder,
                                       "revenue_only": False, "status": "waiting", "seconds": None,
                                       "file": "", "error": "", "revenue": None, "progress": None})
                    added += 1
        return added

    def remove(self, year, month, nums, entry=None):
        nums = {int(n) for n in nums}
        with self.lock:
            self.items = [i for i in self.items
                          if not (i["status"] == "waiting" and (i["year"], i["month"]) == (year, month)
                                  and i["num"] in nums and (entry is None or i.get("entry") == entry))]

    def clear(self, which):
        with self.lock:
            keep = ("running",) if which == "all" else ("running", "waiting")
            self.items = [i for i in self.items if i["status"] in keep]

    def remove_one(self, iid):
        """Drop one row from the queue by id, as long as it is not the one running."""
        with self.lock:
            self.items = [i for i in self.items if not (i["id"] == iid and i["status"] != "running")]

    def retry(self, iid):
        with self.lock:
            for i in self.items:
                if i["id"] == iid and i["status"] in ("error", "stopped"):
                    i.update(status="waiting", error="", seconds=None)

    def running(self):
        return bool(self.thread and self.thread.is_alive())

    def start(self):
        if self.running(): return False
        self.stop_after = False
        self.thread = threading.Thread(target=self.loop, daemon=True)
        self.thread.start()
        return True

    def stop_now(self):
        """Stop after the current report, and tell the current one to stop waiting for PIAS."""
        self.stop_after = True
        if self.cur:
            try: open(os.path.join(self.cur, "STOP"), "w").close()
            except Exception: pass

    def loop(self):
        self.say("Queue started")
        year = session_year()                 # PIAS has to be restarted to change the year, so group by it
        while not self.stop_after:
            with self.lock:
                waiting = [i for i in self.items if i["status"] == "waiting"]
                it = next((i for i in waiting if i["year"] == year), None) or (waiting[0] if waiting else None)
                if it:
                    it["status"] = "running"
                    year = it["year"]
            if not it: break
            self.run_one(it)
        with self.lock:
            done = sum(i["status"] == "done" for i in self.items)
            bad = ["T%d.%d %s" % (i["month"], i["year"], i["name"]) for i in self.items if i["status"] == "error"]
        self.say("Queue finished: %d done" % done + (", to re-run: " + "; ".join(bad) if bad else ""))
        self.cur = None

    def run_one(self, it):
        label = "T%d.%d %s" % (it["month"], it["year"], it["name"])
        d = os.path.join(RUNS, time.strftime("%Y%m%d_%H%M%S") + "_%d" % it["id"])
        os.makedirs(d, exist_ok=True)
        json.dump({k: it.get(k) for k in ("year", "month", "num", "folder", "revenue_only", "entry",
                                          "kind", "line", "m1", "m2")},
                  open(os.path.join(d, "args.json"), "w", encoding="utf-8"), ensure_ascii=False)
        self.cur = d
        self.say(label + ": start")
        t0 = time.time()
        err = ""
        try:
            e = dict(os.environ); e["PYTHONIOENCODING"] = "utf-8"; e["PYTHONUTF8"] = "1"
            p = subprocess.Popen(PY + ["-u", os.path.join(APP, "worker.py"), d], cwd=APP, env=e, stdin=subprocess.DEVNULL,
                                 stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=0, creationflags=NO_WINDOW)
            self.proc = p
            dec = codecs.getincrementaldecoder("utf-8")(errors="replace")
            buf = ""
            while True:
                chunk = p.stdout.read(4096)
                if not chunk: break
                for ch in dec.decode(chunk):
                    if ch in "\r\n":
                        t = ANSI.sub("", buf).rstrip(); buf = ""
                        if t.startswith("[[ERROR]]"): err = t[9:].strip()
                        elif t.startswith("[[P]]"):
                            try:
                                dn, tot, label = t[5:].strip().split("|", 2)
                                with self.lock:
                                    it["progress"] = {"done": int(dn), "total": int(tot), "label": label}
                            except Exception: pass
                        elif t: self.say("  " + t)
                    else:
                        buf += ch
            p.wait()
            res = json.load(open(os.path.join(d, "result.json"), encoding="utf-8"))
        except Exception as ex:
            res = {"ok": False, "error": err or str(ex)}
        self.proc = None
        with self.lock:
            it["seconds"] = round(time.time() - t0)
            it["progress"] = None
            if res.get("ok"):
                it.update(status="done", file=res.get("file", ""), revenue=res.get("revenue"))
            else:
                stopped = os.path.exists(os.path.join(d, "STOP"))
                it.update(status="stopped" if stopped else "error", error=res.get("error") or err or "failed")
        if it["status"] == "error" and it.get("kind") == "policy" and "fewer months" in (it.get("error") or ""):
            m1, m2 = int(it.get("m1") or it["month"]), int(it.get("m2") or it["month"])
            if m2 > m1:
                added = self.add_policies([(it["year"], m, m) for m in range(m1, m2 + 1)],
                                          [it["line"]], it["folder"])
                it["error"] += " -> split into %d months" % added
                self.say("%s: too big for PIAS, split into %d monthly jobs" % (label, added))
        if it["status"] == "done" and not it["file"]:
            it["file"] = "(no data)"
        self.say(label + ": " + (("saved " + os.path.basename(it["file"])) if it["status"] == "done" else it["status"] + " - " + it["error"]))

    def view(self, since):
        with self.lock:
            return {"items": [dict(i) for i in self.items], "running": self.running(), "stopping": self.stop_after and self.running(),
                    "lines": self.lines[since:] if since <= len(self.lines) else self.lines[:], "total": len(self.lines)}


Q = Queue()


def cleanup():
    """Run folders older than 30 days."""
    now = time.time()
    for d in glob.glob(os.path.join(RUNS, "*")):
        try:
            if os.path.isdir(d) and now - os.path.getmtime(d) > 30 * 86400: shutil.rmtree(d, ignore_errors=True)
        except Exception: pass


def browse_folder(start):
    """Explorer folder picker (engines/pick_folder.ps1, brought to the front); returns '' when cancelled."""
    if BROWSE_LOCK.locked(): return ""            # a picker is already open, do not stack a second one
    with BROWSE_LOCK:
        script = os.path.join(APP, "engines", "pick_folder.ps1")
        r = subprocess.run(["powershell", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", script,
                            "-Start", start or ""], capture_output=True, creationflags=NO_WINDOW, timeout=600)
        out = r.stdout.decode("utf-8", "replace").strip().lstrip("﻿")
        if r.returncode and not out:
            raise RuntimeError(r.stderr.decode("utf-8", "replace").strip()[:300] or "folder picker failed")
        return out.splitlines()[-1] if out else ""

BROWSE_LOCK = threading.Lock()


# ------------------------------------------------------------------ update from GitHub
UPD = {"cache": None, "checked": 0, "job": None}
SKIP_DIRS  = {"pyembed", "dump", "output", "dist", ".git", "__pycache__", ".vscode"}
ALLOW_EXT  = {".py", ".js", ".html", ".css", ".bat", ".txt", ".json", ".md", ".ico", ".ps1"}
ALLOW_NAME = {"VERSION"}
KEEP_FILES = {"update.json"}                   # each install keeps its own repo + token


def wanted(rel):
    name = rel.split("/")[-1]
    ext = os.path.splitext(name)[1].lower()
    if name in KEEP_FILES or name.startswith("~$"): return False
    return name in ALLOW_NAME or ext in ALLOW_EXT


http_get = C.http_get


def cmp_ver(a, b):
    na = [int(x) for x in str(a).replace("v", "").split(".") if x.isdigit()]
    nb = [int(x) for x in str(b).replace("v", "").split(".") if x.isdigit()]
    for i in range(max(len(na), len(nb))):
        x, y = (na[i] if i < len(na) else 0), (nb[i] if i < len(nb) else 0)
        if x != y: return 1 if x > y else -1
    return 0


def read_remote(c, name):
    """API first (raw.githubusercontent keeps serving the old file for a few minutes)."""
    repo, branch, token = c["repo"], c["branch"] or "main", c["token"]
    try:
        api = "https://api.github.com/repos/%s/contents/%s?ref=%s" % (repo, name, branch)
        j = json.loads(http_get(api, token).decode("utf-8", "replace"))
        import base64
        return base64.b64decode(j.get("content", "")).decode("utf-8", "replace")
    except urllib.error.HTTPError as ex:
        if ex.code == 404: raise
    except Exception:
        pass
    return http_get("https://raw.githubusercontent.com/%s/%s/%s" % (repo, branch, name), token).decode("utf-8", "replace")


def update_check():
    c = C.update_cfg(); cur = version()
    if not c["repo"] or "/" not in c["repo"]:
        out = {"configured": False, "current": cur}
    else:
        try:
            latest = read_remote(c, "VERSION").strip()
            try: notes = read_remote(c, "RELEASE-NOTES.txt")
            except Exception: notes = ""
            out = {"configured": True, "current": cur, "latest": latest,
                   "available": cmp_ver(latest, cur) > 0, "notes": notes[:4000], "repo": c["repo"]}
        except urllib.error.HTTPError as ex:
            msg = ("repo %s not found on GitHub" % c["repo"]) if ex.code in (401, 403, 404) else "GitHub returned %s" % ex.code
            out = {"configured": True, "current": cur, "error": msg, "repo": c["repo"]}
        except Exception as ex:
            out = {"configured": True, "current": cur, "error": "cannot reach GitHub (%s)" % str(ex)[:120], "repo": c["repo"]}
    UPD["cache"] = out; UPD["checked"] = time.time()
    return out


class UpdJob:
    def __init__(self):
        self.lines, self.state, self.error = [], "running", ""
        self.started, self.ended = time.time(), None

    def put(self, t): self.lines.append(t)

    def view(self):
        return {"state": self.state, "lines": self.lines, "error": self.error,
                "elapsed": round((self.ended or time.time()) - self.started)}


def install_tree(src, j):
    backup = os.path.join(APP, "_backup_" + time.strftime("%Y%m%d_%H%M%S"))
    n = 0
    for base, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith("_backup_")]
        for f in files:
            full = os.path.join(base, f)
            rel = os.path.relpath(full, src).replace("\\", "/")
            if not wanted(rel): continue
            dst = os.path.join(ROOT, rel.replace("/", os.sep))
            if os.path.exists(dst):
                if open(dst, "rb").read() == open(full, "rb").read(): continue
                b = os.path.join(backup, rel.replace("/", os.sep))
                os.makedirs(os.path.dirname(b), exist_ok=True); shutil.copy2(dst, b)
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            shutil.copy2(full, dst); n += 1
            j.put("  " + rel)
    j.put("")
    j.put("%d file(s) updated." % n + (" Old copies kept in app\\%s" % os.path.basename(backup) if n else ""))
    j.put("Version on disk is now %s. Restart the tool to use it." % version())
    UPD["cache"] = None


def update_job(mode, zip_path=""):
    if Q.running(): return None, "Wait for the export queue to finish."
    if UPD["job"] and UPD["job"].state == "running": return None, "An update is already running."
    j = UpdJob(); UPD["job"] = j

    def body():
        tmp = tempfile.mkdtemp(prefix="pias_upd_")
        try:
            if mode == "local":
                if not os.path.exists(zip_path): raise ValueError("File not found: %s" % zip_path)
                j.put("Installing from %s" % zip_path)
                zp = zip_path
            else:
                c = C.update_cfg()
                if not c["repo"]: raise ValueError("app\\update.json has no repo yet.")
                url = "https://codeload.github.com/%s/zip/refs/heads/%s" % (c["repo"], c["branch"] or "main")
                j.put("Downloading %s ..." % url)
                zp = os.path.join(tmp, "src.zip")
                open(zp, "wb").write(http_get(url, c["token"], timeout=180))
            ex = os.path.join(tmp, "x")
            with zipfile.ZipFile(zp) as z: z.extractall(ex)
            roots = [os.path.join(ex, d) for d in os.listdir(ex) if os.path.isdir(os.path.join(ex, d))]
            src = roots[0] if len(roots) == 1 and not os.path.exists(os.path.join(ex, "VERSION")) else ex
            install_tree(src, j)
            j.state = "done"
        except urllib.error.HTTPError as ex:
            j.error = "GitHub returned %s. Check the repo and branch in app\\update.json." % ex.code; j.state = "failed"
        except Exception as ex:
            j.error = "Update failed: %s" % ex; j.state = "failed"
        finally:
            j.ended = time.time()
            shutil.rmtree(tmp, ignore_errors=True)
    threading.Thread(target=body, daemon=True).start()
    return j, ""


# ------------------------------------------------------------------ http
TOKEN = ""
LAST_PING = [time.time()]
IDLE_CLOSE = 8 * 3600      # a hidden window stops pinging (the browser freezes its timers), so give
                           # the tool a whole working day before it closes itself
SHUTDOWN = []


class Handler(BaseHTTPRequestHandler):
    server_version = "PIASExport"
    protocol_version = "HTTP/1.1"

    def log_message(self, *a): pass

    def _send(self, code, body=b"", ctype="application/json; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        try: self.wfile.write(body)
        except Exception: pass

    def json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8"))

    def ok_token(self, q):
        ok = (self.headers.get("X-Token") or q.get("t", [""])[0]) == TOKEN
        if ok: LAST_PING[0] = time.time()        # any call from the page means a window is open
        return ok

    def body(self):
        try: n = int(self.headers.get("Content-Length") or 0)
        except ValueError: n = 0
        if not n: return {}
        try: return json.loads(self.rfile.read(n).decode("utf-8"))
        except Exception: return {}

    def do_GET(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        g = lambda k: q.get(k, [""])[0]
        if p in ("/", "/index.html"):
            LAST_PING[0] = time.time()
            f = os.path.join(UI, "index.html")
            if not os.path.exists(f): return self._send(404, b"ui missing", "text/plain")
            return self._send(200, open(f, "rb").read(), "text/html; charset=utf-8")
        if not self.ok_token(q): return self.json({"error": "token"}, 403)
        if p == "/api/state":
            return self.json({"version": version(), "update": UPD.get("cache"), "reports": E.REPORTS, "usual": E.DEFAULT,
                              "settings": C.load_settings(), "pias": E.pias_exe(), "pias_found": os.path.exists(E.pias_exe()),
                              "catalog": pias_catalog(), "ready_entry": E.HQKD, "lines": E.LINES,
                              "volume": policy_volume(),
                              "policy_folder": C.load_settings().get("policy_folder", "")})
        if p == "/api/queue": return self.json(Q.view(int(g("since") or 0)))
        if p == "/api/update/check": return self.json(update_check())
        if p == "/api/update/job":
            j = UPD["job"]; return self.json({"job": j.view() if j else None})
        return self._send(404, b"not found", "text/plain")

    def do_POST(self):
        u = urlparse(self.path); q = parse_qs(u.query); p = u.path
        if not self.ok_token(q): return self.json({"error": "token"}, 403)
        d = self.body()
        if p == "/api/ping":
            LAST_PING[0] = time.time(); return self.json({"ok": True})
        if p == "/api/queue/add":
            s = C.load_settings()
            try: n = Q.add(int(d["year"]), int(d["month"]), d.get("nums") or [], s["folder"],
                           d.get("revenue_only"), d.get("entry"))
            except Exception as ex: return self.json({"error": str(ex)}, 400)
            return self.json({"ok": True, "added": n})
        if p == "/api/queue/policies":
            s = C.load_settings()
            folder = (d.get("folder") or s.get("policy_folder") or "").strip().strip('"')
            if not folder: return self.json({"error": "Choose the folder to save the lists in."}, 400)
            s["policy_folder"] = folder; C.save_settings(s)
            try:
                y1, m1, y2, m2 = int(d["year_from"]), int(d["month_from"]), int(d["year_to"]), int(d["month_to"])
            except Exception:
                return self.json({"error": "bad period"}, 400)
            n = Q.add_policies(chunk_months(y1, m1, y2, m2, d.get("chunk") or "quarter"),
                               d.get("lines") or [], folder)
            return self.json({"ok": True, "added": n})
        if p == "/api/queue/remove":
            Q.remove(int(d["year"]), int(d["month"]), d.get("nums") or [], d.get("entry"))
            return self.json({"ok": True})
        if p == "/api/queue/clear":
            Q.clear(d.get("which") or "finished"); return self.json({"ok": True})
        if p == "/api/queue/remove_one":
            Q.remove_one(int(d.get("id") or 0)); return self.json({"ok": True})
        if p == "/api/queue/retry":
            Q.retry(int(d.get("id") or 0)); return self.json({"ok": True})
        if p == "/api/queue/start":
            Q.start(); return self.json({"ok": True})
        if p == "/api/queue/stop":
            Q.stop_now(); return self.json({"ok": True})
        if p == "/api/settings":
            s = C.load_settings()
            f = (d.get("folder") or "").strip().strip('"')
            if f: s["folder"] = f
            C.save_settings(s); return self.json(s)
        if p == "/api/browse":
            try: picked = browse_folder(d.get("start") or "")
            except Exception as ex: return self.json({"error": str(ex)}, 500)
            return self.json({"path": picked})
        if p == "/api/open":
            path = (d.get("path") or "").strip()
            if not path: return self.json({"error": "Not found"}, 404)
            try:
                if d.get("select") and os.path.isfile(path):
                    subprocess.Popen(["explorer", "/select,", path])
                elif os.path.isfile(path) and path.lower().endswith((".xlsx", ".xls")):
                    os.startfile(path)
                else:
                    if not os.path.isdir(path): os.makedirs(path, exist_ok=True)
                    os.startfile(path)
            except Exception as ex: return self.json({"error": str(ex)}, 500)
            return self.json({"ok": True})
        if p == "/api/update":
            j, err = update_job(d.get("mode") or "github", d.get("zip") or "")
            if not j: return self.json({"error": err}, 409)
            return self.json({"ok": True})
        if p == "/api/restart":
            restart(); return self.json({"ok": True})
        if p == "/api/quit":
            Q.stop_now()
            if Q.proc:
                try: subprocess.call(["taskkill", "/F", "/T", "/PID", str(Q.proc.pid)], creationflags=NO_WINDOW,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception: pass
            threading.Timer(0.4, lambda: SHUTDOWN and SHUTDOWN[0]()).start()
            return self.json({"ok": True})
        return self.json({"error": "not found"}, 404)


def restart():
    """Detached helper waits for this process to free the port, then runs the .bat."""
    helper = "import os,time\ntime.sleep(2.0)\nos.startfile(r'%s')\n" % BAT
    exe = sys.executable or "pythonw.exe"
    pw = os.path.join(os.path.dirname(exe), "pythonw.exe")
    subprocess.Popen([pw if os.path.exists(pw) else exe, "-c", helper], creationflags=NO_WINDOW | 0x00000008)
    threading.Timer(0.5, lambda: SHUTDOWN and SHUTDOWN[0]()).start()


def app_windows(title=NAME):
    if os.name != "nt": return set()
    u32 = ctypes.windll.user32
    proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    hits = set()
    def cb(hwnd, _):
        n = u32.GetWindowTextLengthW(hwnd)
        if n and u32.IsWindowVisible(hwnd):
            buf = ctypes.create_unicode_buffer(n + 1); u32.GetWindowTextW(hwnd, buf, n + 1)
            if buf.value == title: hits.add(hwnd)
        return True
    u32.EnumWindows(proto(cb), 0)
    return hits


def maximize_window(before, wait=20):
    """Edge ignores --start-maximized when it is already running, so maximize the new window ourselves."""
    if os.name != "nt": return
    u32 = ctypes.windll.user32
    end = time.time() + wait
    while time.time() < end:
        new = app_windows() - before
        if new:
            for h in new:
                if not u32.IsZoomed(h): u32.ShowWindow(h, 3)
            return
        time.sleep(0.3)


def open_ui():
    url = "http://%s:%d/?t=%s" % (HOST, PORT, TOKEN)
    if os.environ.get("CT_NO_OPEN"): log("CT_NO_OPEN set, window not opened: " + url); return
    for exe in (r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"):
        if os.path.exists(exe):
            try:
                before = app_windows()
                subprocess.Popen([exe, "--app=" + url, "--start-maximized", "--disable-features=Translate"],
                                 creationflags=NO_WINDOW)
                t = threading.Thread(target=maximize_window, args=(before,), daemon=True)
                t.start()
                return t
            except Exception: pass
    webbrowser.open(url)


def already_running():
    s = socket.socket()
    try:
        s.connect((HOST, PORT)); return True
    except Exception:
        return False
    finally:
        s.close()


def shortcuts():
    """On start: make sure the Desktop + Start-menu shortcut exists and its icon matches app.ico.
    The icon is copied to a content-hashed name so Windows cannot keep showing a stale cached icon after
    it changes; the shortcut is pointed at that copy and the icon cache is nudged, so the Desktop icon
    updates after an app update without a reboot."""
    if os.name != "nt" or os.environ.get("CT_NO_OPEN"): return
    target = os.path.join(APP, "pyembed", "pythonw.exe")
    ico = os.path.join(UI, "app.ico")
    if not os.path.exists(target) or not os.path.exists(ico): return
    import hashlib
    cache_ico = ico
    try:
        tag = hashlib.md5(open(ico, "rb").read()).hexdigest()[:8]
        cache_ico = os.path.join(DUMP, "icon_%s.ico" % tag)
        if not os.path.exists(cache_ico): shutil.copy2(ico, cache_ico)
        for stale in glob.glob(os.path.join(DUMP, "icon_*.ico")):
            if os.path.abspath(stale) != os.path.abspath(cache_ico):
                try: os.remove(stale)
                except OSError: pass
    except Exception as ex:
        log("icon cache: %r" % ex); cache_ico = ico
    srv = os.path.join(APP, "server.py")
    q = lambda s: s.replace("'", "''")
    dirs = "@([Environment]::GetFolderPath('Programs'), [Environment]::GetFolderPath('Desktop'))"
    ps = "$s = New-Object -ComObject WScript.Shell;"
    ps += ("foreach ($d in %s) {"
           " $p = Join-Path $d '%s.lnk'; $new = -not (Test-Path -LiteralPath $p); $l = $s.CreateShortcut($p);"
           " if ($new) { $l.TargetPath = '%s'; $l.Arguments = '\"%s\"'; $l.WorkingDirectory = '%s'; $l.Description = '%s - PVI Digital' };"
           " if ($new -or $l.IconLocation -ne '%s,0') { $l.IconLocation = '%s,0'; $l.Save() } }"
           % (dirs, q(NAME), q(target), q(srv), q(ROOT), q(NAME), q(cache_ico), q(cache_ico)))
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], timeout=60,
                       creationflags=NO_WINDOW, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["ie4uinit.exe", "-show"], timeout=20, creationflags=NO_WINDOW,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as ex:
        log("shortcuts: %r" % ex)

def watchdog():
    """No page talking to us for 2 minutes and nothing running -> close the server."""
    while True:
        time.sleep(10)
        if Q.running() or (UPD["job"] and UPD["job"].state == "running"):
            LAST_PING[0] = time.time(); continue
        if time.time() - LAST_PING[0] > IDLE_CLOSE:
            log("no page for %d hours, closing" % (IDLE_CLOSE // 3600)); SHUTDOWN and SHUTDOWN[0](); return


def main():
    global TOKEN
    for d in (DUMP, RUNS): os.makedirs(d, exist_ok=True)
    if already_running():
        try: TOKEN = open(TOKFILE, encoding="utf-8").read().strip()
        except Exception: TOKEN = ""
        log("already running, opening the window")
        t = open_ui()
        if t: t.join()
        return
    TOKEN = os.urandom(16).hex()
    open(TOKFILE, "w", encoding="utf-8").write(TOKEN)
    try:
        srv = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as ex:
        box("Cannot open port %d (%s).\nClose the running tool, then open it again." % (PORT, ex)); return
    SHUTDOWN.append(lambda: threading.Thread(target=srv.shutdown, daemon=True).start())
    log("listening on %s:%d  v%s" % (HOST, PORT, version()))
    threading.Thread(target=watchdog, daemon=True).start()
    threading.Thread(target=cleanup, daemon=True).start()
    threading.Thread(target=lambda: (time.sleep(1.5), update_check()), daemon=True).start()
    threading.Thread(target=shortcuts, daemon=True).start()
    open_ui()
    try: srv.serve_forever()
    except KeyboardInterrupt: pass
    log("closed")


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        import traceback
        log(traceback.format_exc())
        box("The tool failed to start:\n\n%s\n\nLog: %s" % (ex, LOGFILE))
