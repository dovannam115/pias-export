# One report export, started by server.py as a child process:
#   python -u worker.py <run dir>
# Reads <run dir>\args.json {year, month, num, folder, revenue_only}, ends with <run dir>\result.json.
# Lines it prints go to the log in the window. A file named STOP in the run dir asks it to give up waiting.
import os, sys, json, time, traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C  # noqa: E402,F401
import pias_engine as E  # noqa: E402

try: sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception: pass


def main(run_dir):
    a = json.load(open(os.path.join(run_dir, "args.json"), encoding="utf-8"))
    res = {"ok": False}
    t0 = time.time()
    try:
        dst = E.export(int(a["year"]), int(a["month"]), int(a["num"]), a["folder"], bool(a.get("revenue_only")),
                       log=lambda s: print(s, flush=True),
                       stop=lambda: os.path.exists(os.path.join(run_dir, "STOP")))
        res = {"ok": True, "file": dst}
        try:
            tot = E.revenue_total(dst)
            if tot is not None:
                res["revenue"] = tot
                print("Revenue in the file: {:,.0f} VND".format(tot).replace(",", "."), flush=True)
        except Exception:
            pass
        print("Saved " + dst, flush=True)
    except E.UserError as ex:
        res["error"] = str(ex)
        print("[[ERROR]] " + str(ex), flush=True)
    except Exception as ex:
        res["error"] = "Unexpected error: %s" % ex
        res["trace"] = traceback.format_exc()
        print("[[ERROR]] " + res["error"], flush=True)
    res["seconds"] = round(time.time() - t0)
    json.dump(res, open(os.path.join(run_dir, "result.json"), "w", encoding="utf-8"), ensure_ascii=False)
    return 0 if res["ok"] else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
