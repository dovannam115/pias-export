# Paths and setup shared by server.py (the window) and worker.py (one report export).
import os, sys, json

APP     = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP)
ENGINES = os.path.join(APP, "engines")
DUMP    = os.path.join(APP, "dump")           # token, log, settings, run folders
OUTPUT  = os.path.join(ROOT, "output")        # default save folder: output\<year>\<MM>\
RUNS    = os.path.join(DUMP, "runs")          # one folder per export: args.json, result.json
SETTINGS = os.path.join(DUMP, "settings.json")

if ENGINES not in sys.path: sys.path.insert(0, ENGINES)

# Nam's revenue project keeps the monthly exports here; used as the default save folder when it exists.
REVENUE_RAW = os.path.join(os.path.expanduser("~"), "Documents", "Revenue-Report", "data", "raw")


def default_folder():
    if os.path.isdir(REVENUE_RAW): return os.path.join(REVENUE_RAW, "{year}", "{MM}")
    return os.path.join(OUTPUT, "{year}", "{MM}")


def load_settings():
    s = {}
    try: s = json.load(open(SETTINGS, encoding="utf-8"))
    except Exception: pass
    if not s.get("folder"): s["folder"] = default_folder()
    return s


def save_settings(s):
    os.makedirs(DUMP, exist_ok=True)
    json.dump(s, open(SETTINGS, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


def folder_for(tpl, year, month):
    return os.path.normpath(tpl.replace("{year}", str(year)).replace("{MM}", "%02d" % month).replace("{M}", str(month)))


# Where "Update" looks for new versions; app\update.json wins when it names a repo.
REPO = {"repo": "dovannam115/pias-export", "branch": "main", "token": ""}


def update_cfg():
    out = dict(REPO)
    try:
        j = json.load(open(os.path.join(APP, "update.json"), encoding="utf-8"))
        if j.get("repo"): out.update({k: v for k, v in j.items() if v or k == "token"})
    except Exception:
        pass
    return out


def http_get(url, token="", timeout=25):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "PIASExport", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()
