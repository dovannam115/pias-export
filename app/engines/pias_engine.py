"""Drive the PIAS desktop app (Pias2022.exe) to export one 'Báo cáo hiệu quả kinh doanh' report for one month.

Per report: make sure PIAS is open and logged in (pias_login.ps1, saved account), open the report list through
the menu (MSAA, no mouse), run the report for one calendar month (pias_export.ps1), wait for PIAS to write
BCHQKD_*.xlsx into its 'Excel Files' folder, check the period printed in the file, copy it to the save folder as
'T10.2026 - BH Kỹ thuật.xlsx' (a re-run gets the run date: 'T10.2026 - BH Kỹ thuật 12-11.xlsx'), and close the
copy PIAS opened in Excel. Clicks are posted to window handles: the user's mouse and keyboard are not touched.
"""
import calendar
import ctypes
import threading
import datetime as dt
import glob
import os
import re
import shutil
import subprocess
import time
import zipfile
from ctypes import POINTER, byref, wintypes

HERE = os.path.dirname(os.path.abspath(__file__))
TIMEOUT_MIN = 60

# number in the PIAS list -> name used in the saved file (chị Linh's style: 'T10.2026 - BH Kỹ thuật.xlsx')
REPORTS = {
    1: "BH Cháy - Tài sản", 2: "BH Kỹ thuật", 3: "BH P&I", 4: "BH Thân tàu viễn dương", 5: "BH Thân tàu nội thủy",
    6: "BH Hàng hóa", 7: "BH Trách nhiệm", 8: "BH Năng lượng", 9: "BH Hàng không", 10: "BH Con người",
    11: "BH Xe cơ giới", 12: "BH Nông nghiệp", 13: "BH Khác", 14: "Nhận tái bảo hiểm", 15: "KQKD theo khách hàng",
    16: "Chi phí theo đơn vị", 17: "Chi tiết HQKD theo đơn vị", 18: "HQ nhận tái bảo hiểm",
    19: "BH Cháy - Tài sản CAT4, CAT6", 20: "KH sử dụng nguồn vốn", 21: "BH Phi hàng hải theo rủi ro",
    22: "BH Thiệt hại khác", 23: "Thống kê số lượng đơn",
}
DEFAULT = [1, 2, 4, 5, 6, 7, 10, 11, 13, 22]       # the lines PVI Digital sells

HQKD = "Báo cáo hiệu quả kinh doanh"              # the entry REPORTS above belongs to
CATALOG = os.path.join(os.path.dirname(HERE), "data", "pias_catalog.json")
YEAR_FILE = os.path.join(os.path.dirname(HERE), "dump", "pias_year.txt")

EXE_CANDIDATES = [r"C:\PVI\Pias2022new\Pias2022.exe", r"C:\PVI\Pias2022\Pias2022.exe"]


def catalog():
    """{entry: [{num, name}]} for the entries this PIAS account may open."""
    try:
        import json
        c = json.load(open(CATALOG, encoding="utf-8"))
    except Exception:
        return {}
    return {e["entry"]: e["reports"] for e in c.get("entries", []) if e.get("access") and e.get("reports")}


def report_name(entry, num):
    if entry == HQKD and num in REPORTS: return REPORTS[num]
    for r in catalog().get(entry, []):
        if r["num"] == num: return r["name"]
    return "Báo cáo %d" % num


def pias_exe():
    for e in EXE_CANDIDATES:
        if os.path.exists(e): return e
    hits = glob.glob(r"C:\PVI\*\Pias2022.exe")
    return hits[0] if hits else EXE_CANDIDATES[0]


def pias_out():
    return os.path.join(os.path.dirname(pias_exe()), "Excel Files")


class UserError(Exception):
    pass


# ---------------------------------------------------------------- PIAS driving
def _decode(b):
    s = b.decode("utf-8", "replace")
    if "\ufffd" in s: s = b.decode("mbcs", "replace")
    return s


def ps(script, *args, timeout=None, log=None):
    """Run one of the PowerShell steps. With log=, its lines are passed on as they are printed, so a long
    search does not look like the tool has stopped. The error comes from stderr, not from the first line."""
    p = subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                          os.path.join(HERE, script), *map(str, args)],
                         stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL,
                         creationflags=0x08000000)
    lines, err = [], []

    def pump(stream, keep, cb):
        for raw in iter(stream.readline, b""):
            s = _decode(raw).rstrip("\r\n")
            keep.append(s)
            if cb and s.strip(): cb(s)
        stream.close()

    t1 = threading.Thread(target=pump, args=(p.stdout, lines, log), daemon=True)
    t2 = threading.Thread(target=pump, args=(p.stderr, err, None), daemon=True)
    t1.start(); t2.start()
    try:
        p.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        p.kill(); raise UserError("%s: no answer after %s s" % (script, timeout))
    t1.join(5); t2.join(5)
    out = "\n".join(x for x in lines if x.strip()).strip()
    if p.returncode:
        bad = [x for x in err if x.strip() and not x.startswith(("At ", "+ ", "    +", "ParserError"))]
        msg = bad[0] if bad else (lines[-1] if lines else "failed")
        raise UserError("%s: %s" % (script, msg.strip()))
    return out


user32 = ctypes.windll.user32
EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def pias_main():
    found = []

    def cb(h, _):
        n = ctypes.create_unicode_buffer(256)
        user32.GetWindowTextW(h, n, 256)
        if n.value.startswith("Insurances Solution") and user32.IsWindowVisible(h):
            found.append((h, n.value))
        return True
    user32.EnumWindows(EnumProc(cb), 0)
    return found[0] if found else (None, "")


def logged_in(year):
    """Is PIAS already up on this year? The login script takes ten seconds even when there is nothing
    to do, which is ten seconds on every job of a long queue."""
    try:
        if not _pias_pid(): return False
        return int(open(YEAR_FILE, encoding="utf-8").read().strip()) == int(year)
    except Exception:
        return False


def login(year, log=print):
    """Make sure PIAS is up on this year, cheaply when it already is."""
    unminimize()
    if logged_in(year):
        log("PIAS already on %s" % year)
        return
    t0 = time.time()
    ps("pias_login.ps1", "-Exe", pias_exe(), "-Year", year, timeout=300, log=lambda x: log("  " + x))
    log("PIAS ready in %d s" % (time.time() - t0))


def unminimize():
    """A minimised PIAS hides its title from Windows and its controls have no size: restore it without
    activating it (SW_SHOWNOACTIVATE), so the user's current window keeps the keyboard."""
    h, _ = pias_main()
    if h and user32.IsIconic(h):
        user32.ShowWindow(h, 4)
        time.sleep(1.5)


def report_list_open(entry):
    """Is the report screen already showing this entry? Reopening it costs ten seconds, so it pays to ask."""
    try:
        header, rows = list_on_screen()
    except Exception:
        return False
    return header == entry and bool(rows)


def open_report_list(entry=None):
    """Menu Báo cáo > <entry>, invoked through MSAA (DevExpress menus ignore posted clicks).
    The screen remembers the entry it last showed, so it is closed first and opened again on the one we want."""
    entry = entry or HQKD
    for _ in range(3):
        close_report_screen()
        menu_click(["Báo cáo", entry])
        for _ in range(25):
            time.sleep(1)
            header, rows = list_on_screen()
            if header == entry and rows:
                time.sleep(1.5)
                return
    raise UserError("Could not open '%s' in PIAS (the account may not have it)." % entry)


# ---------------------------------------------------------------- the report screen, read through MSAA
def _acc(hwnd):
    import comtypes.client
    from comtypes import automation  # noqa: F401
    comtypes.client.GetModule("oleacc.dll")
    from comtypes.gen.Accessibility import IAccessible
    p = POINTER(IAccessible)()
    ctypes.oledll.oleacc.AccessibleObjectFromWindow(hwnd, -4, byref(IAccessible._iid_), byref(p))
    return p


def _kids(a):
    import comtypes.client
    from comtypes import automation
    comtypes.client.GetModule("oleacc.dll")
    from comtypes.gen.Accessibility import IAccessible
    try: n = a.accChildCount
    except Exception: return []
    if not n: return []
    arr = (automation.VARIANT * max(n, 1))(); got = ctypes.c_long()
    try: ctypes.oledll.oleacc.AccessibleChildren(a, 0, n, arr, byref(got))
    except Exception: return []
    out = []
    for i in range(got.value):
        v = arr[i].value
        if not isinstance(v, int):
            try: out.append(v.QueryInterface(IAccessible))
            except Exception: pass
    return out


def _name(a):
    try: return (a.accName(0) or "").strip()
    except Exception: return ""


def _children_of_main():
    h, _ = pias_main()
    if not h: raise UserError("PIAS main window not found.")
    out = []
    user32.EnumChildWindows(h, EnumProc(lambda c, _: out.append(c) or True), 0)
    return out


def menu_click(path):
    """Click a menu item by its path, e.g. ['Báo cáo', 'Báo cáo nghiệp vụ bồi thường']."""
    for c in _children_of_main():
        try:
            a = _acc(c)
            if _name(a) != "MenuBar": continue
        except Exception:
            continue
        node = a
        for name in path:
            node = next((k for k in _kids(node) if _name(k) == name), None)
            if node is None: raise UserError("Menu item '%s' not found in PIAS." % name)
        node.accDoDefaultAction(0)
        return
    raise UserError("PIAS menu bar not found.")


def close_report_screen():
    """Press the screen's own 'Kết thúc' button (buttons do take posted clicks)."""
    for c in _children_of_main():
        try:
            a = _acc(c)
            if _name(a) != "Kết thúc": continue
        except Exception:
            continue
        if not user32.IsWindowVisible(c): continue
        lp = (11 << 16) | 30
        user32.PostMessageW(c, 0x0200, 0, lp)
        user32.PostMessageW(c, 0x0201, 1, lp)
        user32.PostMessageW(c, 0x0202, 0, lp)
        time.sleep(0.8)


def _report_navbar():
    """(hwnd, accessible) of the wide list on the right of the report screen."""
    for c in _children_of_main():
        try:
            a = _acc(c)
            if "NavBar" not in _name(a): continue
        except Exception:
            continue
        if not user32.IsWindowVisible(c): continue
        r = wintypes.RECT(); user32.GetWindowRect(c, byref(r))
        if (r.right - r.left) > 1000: return c, a, r
    return None, None, None


def list_on_screen():
    """(entry name, [(num, name, x, y)]) of the report list now shown."""
    c, a, r = _report_navbar()
    if not a: return "", []
    for g in _kids(a):
        rows = []
        for i in _kids(g):
            nm = _name(i)
            m = re.match(r"^\s*(\d+)\.\s*(.*)$", nm)
            if not m: continue
            try: l, t, w, h = i.accLocation(0)
            except Exception: continue
            rows.append((int(m.group(1)), m.group(2).strip(), int(l) + 80, int(t) + int(h) // 2))
        if rows: return _name(g), rows
    return "", []


def scroll_report_list(lines):
    """Wheel the report list up or down (lines > 0 scrolls down)."""
    c, a, r = _report_navbar()
    if not c: return
    x, y = (r.left + r.right) // 2, (r.top + r.bottom) // 2
    for _ in range(abs(lines)):
        delta = -120 if lines > 0 else 120
        user32.PostMessageW(c, 0x020A, (delta << 16) & 0xFFFFFFFF, (y << 16) | (x & 0xFFFF))
        time.sleep(0.15)
    time.sleep(0.5)


def close_report_dialogs():
    """A report's settings box (its title starts with the number) stays open when the report has no data
    or stops with an error, and it blocks every report after it."""
    import re as _re
    import pias_policy as P
    pid = _pias_pid()
    closed = 0
    for _ in range(8):
        left = [h for h in P.windows_of(pid) if _re.match(r"^\d+\.\s", P.text_of(h))]
        if not left: break
        h = left[0]
        btn = next((c for c in P.children(h)
                    if P.acc_name(c) in ("Hủy bỏ", "Hủy", "Thoát", "Đóng")), 0)
        if btn: P.click(btn, 30, 10)
        else: user32.PostMessageW(h, 0x0010, 0, 0)
        gone = False
        for _ in range(12):
            time.sleep(0.4)
            if h not in P.windows_of(pid): gone = True; break
        if not gone: break
        closed += 1
    return closed


def row_position(num):
    """Where to click for report <num>, in client coordinates of the list window. The list scrolls, so a
    row further down has no position until it is brought into view."""
    for attempt in range(8):
        c, a, r = _report_navbar()
        _, rows = list_on_screen()
        here = {n: (x, y) for n, name, x, y in rows}
        if num in here:
            x, y = here[num]
            if r.top + 10 < y < r.bottom - 10:
                return c, x - r.left, y - r.top
        seen = sorted(here)
        if not seen: break
        if num > seen[-1]: scroll_report_list(3)
        elif num < seen[0]: scroll_report_list(-3)
        else: scroll_report_list(2 if num > seen[len(seen) // 2] else -2)
    c, a, r = _report_navbar()
    return c, 192, 38 + 22 * (num - 1)          # fall back to the old guess


def close_in_excel(path):
    """PIAS opens its file in Excel read-only; close that copy (other workbooks untouched)."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        wb = win32com.client.GetObject(path)      # finds the workbook in whichever Excel instance holds it
        if os.path.normcase(wb.FullName) != os.path.normcase(path): return
        xl = wb.Application
        wb.Close(False)
        if xl.Workbooks.Count == 0:
            xl.Quit()
    except Exception:
        pass


class NoData(Exception):
    """PIAS answered 'Không có dữ liệu để xuất báo cáo!': the month has nothing to export (counts as done)."""


REPORT_STEPS = ["PIAS", "Report list", "Running", "File", "Check", "Saved"]


def wait_file(since, stop=lambda: False, log=None, timeout_min=None):
    """New BCHQKD_* file in PIAS's Excel Files folder, written after `since`, size settled."""
    deadline = time.time() + (timeout_min or TIMEOUT_MIN) * 60
    folder = pias_out()
    said = 0
    while time.time() < deadline:
        if stop(): raise UserError("Stopped.")
        new = [f for f in glob.glob(os.path.join(folder, "*.xls*"))
               if os.path.getmtime(f) >= since and not os.path.basename(f).startswith("~$")]
        if new:
            f = max(new, key=os.path.getmtime)
            size = -1
            while size != os.path.getsize(f):
                size = os.path.getsize(f)
                mb = size / 1048576.0
                if log and mb >= said + 1:
                    said = mb
                    log("[[P]] 4|%d|PIAS is writing: %.1f MB so far" % (len(REPORT_STEPS), mb))
                time.sleep(2)
            return f
        tick = getattr(wait_file, "tick", 0) + 1
        wait_file.tick = tick
        if tick % 2 == 0:              # every ~6 s: PIAS answers 'no data' with a box instead of a file
            import pias_policy as P
            pid = _pias_pid()
            if P.answer_boxes(pid, rounds=1):      # "mẫu mới / mẫu cũ" and the like
                continue
            boxes = P.message_boxes(pid)
            if boxes:
                h = boxes[0][0]
                width = P.rect(h)[2]
                P.dismiss_boxes(pid)
                if width <= 330:       # the short box is 'Không có dữ liệu để xuất báo cáo!'
                    raise NoData("PIAS: no data for this period")
                raise UserError("PIAS answered with an error box for this period (width %d px)." % width)
            why = P.crash_box(pid)
            if why:
                raise UserError("PIAS stopped while making this report: %s. Try a shorter period." % why)
        time.sleep(3)
    raise UserError("No file from PIAS after %d min (it may have ignored this report)."
                    % (timeout_min or TIMEOUT_MIN))


# ---------------------------------------------------------------- reading the file back
def _rows_openpyxl(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Doanh thu"] if "Doanh thu" in wb.sheetnames else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(max_row=6, values_only=True)]
    wb.close()
    return rows


def _rows_com(path):
    """Files with a sensitivity label are not plain zip; Excel opens them."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(path), 0, True)
        ws = wb.Worksheets(1)
        vals = ws.Range("A1:Z6").Value
        wb.Close(False)
    finally:
        xl.Quit()
    return [list(r) for r in (vals or [])]


def file_period(path):
    """Report period printed in the first rows ('Từ ngày 01/10/2026 đến ngày 31/10/2026')."""
    rows = _rows_openpyxl(path) if zipfile.is_zipfile(path) else _rows_com(path)
    for r in rows[:4]:
        for v in r:
            m = re.search(r"(\d{2})/(\d{2})/(\d{4}).*?(\d{2})/(\d{2})/(\d{4})", str(v or ""))
            if m:
                return dt.date(int(m[3]), int(m[2]), int(m[1])), dt.date(int(m[6]), int(m[5]), int(m[4]))
    raise UserError("Cannot read the period in " + os.path.basename(path))


# words the report title must contain, so a list numbered differently (other PIAS year) is caught
TITLE_KEY = {1: "CHÁY - TÀI SẢN", 2: "KỸ THUẬT", 3: "P&I", 4: "VIỄN DƯƠNG", 5: "NỘI THỦY", 6: "HÀNG HÓA",
             7: "TRÁCH NHIỆM", 8: "NĂNG LƯỢNG", 9: "HÀNG KHÔNG", 10: "CON NGƯỜI", 11: "XE CƠ GIỚI",
             12: "NÔNG NGHIỆP", 13: "BẢO HIỂM KHÁC", 22: "THIỆT HẠI KHÁC"}


def file_title(path):
    import unicodedata
    rows = _rows_openpyxl(path) if zipfile.is_zipfile(path) else _rows_com(path)
    t = next((str(v) for r in rows[:3] for v in r if v and "BÁO CÁO" in str(v).upper()), "")
    return unicodedata.normalize("NFC", t.upper())


def revenue_total(path):
    """Sum of the 'Doanh thu' VNĐ column (sheet Doanh thu), for the summary line."""
    import openpyxl
    if not zipfile.is_zipfile(path): return None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Doanh thu"] if "Doanh thu" in wb.sheetnames else wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    try:
        h = next(i for i, r in enumerate(rows[:15]) if any(str(v or "").strip() == "Số đơn BH" for v in r))
    except StopIteration:
        return None
    head = [str(v or "").strip() for v in rows[h]]
    dcol = next((i for i, v in enumerate(head) if v.startswith("Doanh thu")), None)
    if dcol is None: return None
    sub = [str(rows[h + k][j] or "") for k in (1, 2) for j in (dcol, dcol + 1)]
    col = dcol + 1 if "VNĐ" in (sub[1] + sub[3]) else dcol
    pol = head.index("Số đơn BH")
    total = 0.0
    for r in rows[h + 1:]:
        v = r[col] if col < len(r) else None
        p = r[pol] if pol < len(r) else None
        if isinstance(v, (int, float)) and isinstance(p, str) and "/" in p:
            total += v
    return total


# ---------------------------------------------------------------- one export
BAD = '<>:"/\\|?*'


def safe_name(s, limit=70):
    s = "".join(" " if c in BAD else c for c in str(s)).strip()
    s = " ".join(s.split())
    return s[:limit].strip(" .")


def target(folder, year, month, num, entry=None):
    folder = os.path.normpath(folder)    # the Save As box rejects '/'; a normalised path is all backslashes
    base = "T%d.%d - %s" % (month, year, safe_name(report_name(entry or HQKD, num)))
    today = dt.date.today()
    path = os.path.join(folder, base + ".xlsx")
    if os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d.xlsx" % (base, today.day, today.month))
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d (%d).xlsx" % (base, today.day, today.month, i))
        i += 1
    return path


def export(year, month, num, folder, revenue_only=False, log=print, stop=lambda: False, tries=3, entry=None):
    """Run report <num> of <entry> for one calendar month and save the file under its standard name."""
    entry = entry or HQKD
    if entry == HQKD and num not in REPORTS: raise UserError("Unknown report %s." % num)
    log("[[P]] 1|%d|Opening PIAS" % len(REPORT_STEPS))
    log("Checking PIAS")
    login(year, log=log)
    last = calendar.monthrange(year, month)[1]
    want = (dt.date(year, month, 1), dt.date(year, month, last))
    base_args = ["-Report", num, "-From", "01/%02d/%d" % (month, year), "-To", "%02d/%02d/%d" % (last, month, year)]
    if revenue_only and entry == HQKD: base_args.append("-RevenueOnly")
    start = end = None
    for attempt in range(1, tries + 1):
        if stop(): raise UserError("Stopped.")
        unminimize()
        import pias_policy as P
        P.answer_boxes(_pias_pid(), rounds=2)   # the late "new version" box after a fresh login
        close_report_dialogs()
        log("[[P]] 2|%d|Opening the report list" % len(REPORT_STEPS))
        if not report_list_open(entry):          # reopening the screen costs ten seconds a report
            open_report_list(entry)
        hwnd, x, y = row_position(num)
        args = base_args + ["-X", x, "-Y", y]
        since = time.time() - 1
        log("Running report %d (%s) in PIAS" % (num, report_name(entry, num))
            + (" (try %d)" % attempt if attempt > 1 else ""))
        log("[[P]] 3|%d|PIAS is running the report" % len(REPORT_STEPS))
        import pias_policy as P
        import pias_report as R
        pid = _pias_pid()
        lp = (int(y) << 16) | (int(x) & 0xFFFF)
        for m, w in ((0x0200, 0), (0x0201, 1), (0x0202, 0)):
            user32.PostMessageW(hwnd, m, w, lp)          # open the report from the list
        try:
            R.run(pid, num, "01/%02d/%d" % (month, year), "%02d/%02d/%d" % (last, month, year),
                  revenue_only=bool(revenue_only), log=lambda s: log("  " + str(s)))
        except P.StepError as ex:
            raise UserError(str(ex))
        log("[[P]] 4|%d|Waiting for the file" % len(REPORT_STEPS))
        src = wait_file(since, stop, log=log)
        log("PIAS wrote " + os.path.basename(src))
        log("[[P]] 5|%d|Checking the period" % len(REPORT_STEPS))
        if entry != HQKD:
            break                      # only the business performance files carry a period line we can check
        start, end = file_period(src)
        if (start, end) == want:
            key = TITLE_KEY.get(num)
            title = file_title(src)
            if key and key not in title or (num == 1 and "CAT" in title):
                close_in_excel(src)
                raise UserError("PIAS ran another report (%s), file not saved. The report list may be numbered "
                                "differently in this PIAS year." % title.title())
            break
        close_in_excel(src)
        log("PIAS ran %s-%s instead of the chosen month, running again"
            % (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")))
    else:
        raise UserError("PIAS kept running the wrong period %s-%s, file not saved."
                        % (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")))
    os.makedirs(folder, exist_ok=True)
    dst = target(folder, year, month, num, entry)
    log("[[P]] 6|%d|Saving" % len(REPORT_STEPS))
    shutil.copy2(src, dst)
    close_in_excel(src)
    return dst


# ---------------------------------------------------------------- policy lists (Tìm kiếm đơn Bảo hiểm)
SEARCH_SCREEN = "Tìm kiếm đơn Bảo hiểm"
LINES = {"01": "Con người", "02": "Tài sản kỹ thuật", "03": "Hàng hải", "05": "Xe cơ giới",
         "06": "Trách nhiệm", "07": "Hàng hóa", "08": "Thiệt hại khác"}


def _pias_pid():
    h, _ = pias_main()
    pid = wintypes.DWORD()
    if h: user32.GetWindowThreadProcessId(h, byref(pid))
    return pid.value


def _rect(h):
    r = wintypes.RECT(); user32.GetWindowRect(h, byref(r))
    return r.left, r.top, r.right - r.left, r.bottom - r.top


def search_windows():
    """The policy search screen is a window of its own, not a child of the main window, and at that level
    its title can be read straight from Windows: no UIA, so a grid with 150k rows cannot stall us."""
    pid = _pias_pid()
    out = []

    def cb(h, _):
        p = wintypes.DWORD(); user32.GetWindowThreadProcessId(h, byref(p))
        if p.value != pid or not user32.IsWindowVisible(h): return True
        t2 = ctypes.create_unicode_buffer(200); user32.GetWindowTextW(h, t2, 200)
        if t2.value.strip() == SEARCH_SCREEN: out.append(h)
        return True
    user32.EnumWindows(EnumProc(cb), 0)
    return out


def close_search_screen():
    """Close it between jobs: a grid holding a big result makes PIAS slow to answer anything."""
    closed = 0
    for h in search_windows():
        user32.PostMessageW(h, 0x0010, 0, 0)          # WM_CLOSE
        closed += 1
    for _ in range(20):
        if not search_windows(): break
        time.sleep(0.5)
    return closed


def open_search_screen():
    """Menu Nghiệp vụ bảo hiểm > Tìm kiếm đơn Bảo hiểm, then wait for the window to show up."""
    if search_windows(): return
    for _ in range(3):
        menu_click(["Nghiệp vụ bảo hiểm", SEARCH_SCREEN])
        for _ in range(25):
            time.sleep(1)
            if search_windows():
                time.sleep(1.5)
                return
    raise UserError("Could not open '%s' in PIAS." % SEARCH_SCREEN)


def search_screen_open():
    return bool(search_windows())


def message_boxes():
    """PIAS message boxes ("xuất dữ liệu thành công", "không có dữ liệu"...) as (hwnd, ok button hwnd)."""
    out = []
    h, _ = pias_main()
    pid = wintypes.DWORD()
    if h: user32.GetWindowThreadProcessId(h, byref(pid))

    def cb(w, _):
        p2 = wintypes.DWORD(); user32.GetWindowThreadProcessId(w, byref(p2))
        if p2.value != pid.value or not user32.IsWindowVisible(w): return True
        t2 = ctypes.create_unicode_buffer(200); user32.GetWindowTextW(w, t2, 200)
        if t2.value.strip() != "Insurace Solution 2022": return True
        ok = []

        def cb2(c, _):
            tx = ctypes.create_unicode_buffer(200); user32.GetWindowTextW(c, tx, 200)
            if tx.value.replace("&", "").strip() in ("OK", "Ok", "Đồng ý"): ok.append(c)
            return True
        user32.EnumChildWindows(w, EnumProc(cb2), 0)
        out.append((w, ok[0] if ok else 0))
        return True
    user32.EnumWindows(EnumProc(cb), 0)
    return out


def dismiss_boxes(rounds=6):
    """Press OK on every message box PIAS has up; left alone they block the whole window."""
    pressed = 0
    for _ in range(rounds):
        boxes = message_boxes()
        if not boxes: break
        for w, ok in boxes:
            if ok:
                lp = (11 << 16) | 37
                user32.PostMessageW(ok, 0x0201, 1, lp)
                user32.PostMessageW(ok, 0x0202, 0, lp)
                user32.PostMessageW(ok, 0x00F5, 0, 0)
            else:
                user32.PostMessageW(w, 0x0010, 0, 0)
            pressed += 1
        time.sleep(1.0)
    return pressed


def policy_file(folder, line, year, m1, m2, d1=1, d2=None):
    """01-01_03-2025.xlsx for a range, 02-2025.xlsx for a whole year, 05-07-2025.xlsx for one month,
    01-01.11_15.11-2026.xlsx for dates that do not cover whole months."""
    folder = os.path.normpath(folder)    # the Save As box rejects '/'; a normalised path is all backslashes
    os.makedirs(folder, exist_ok=True)
    last = calendar.monthrange(year, m2)[1]
    if d1 != 1 or (d2 or last) != last: base = "%s-%02d.%02d_%02d.%02d-%d" % (line, d1, m1, d2, m2, year)
    elif (m1, m2) == (1, 12): base = "%s-%d" % (line, year)
    elif m1 == m2: base = "%s-%02d-%d" % (line, m1, year)
    else: base = "%s-%02d_%02d-%d" % (line, m1, m2, year)
    path = os.path.join(folder, base + ".xlsx")
    today = dt.date.today()
    if os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d.xlsx" % (base, today.day, today.month))
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d (%d).xlsx" % (base, today.day, today.month, i))
        i += 1
    return path


POLICY_STEPS = ["PIAS", "Screen", "Period", "Search", "Export", "File"]


def _step(log, n, what=""):
    log("[[P]] %d|%d|%s" % (n, len(POLICY_STEPS), what or POLICY_STEPS[n - 1]))


def export_policies(year, m1, m2, line, folder, log=print, stop=lambda: False, d1=1, d2=None):
    r"""The policies of one line over a range of months of one year, out of 'Tìm kiếm đơn Bảo hiểm',
    saved as <folder>\<line>-<months>-<year>.xlsx."""
    import pias_policy as P
    line = str(line).strip()
    _step(log, 1, "Opening PIAS")
    log("Checking PIAS")
    login(year, log=log)
    if stop(): raise UserError("Stopped.")
    unminimize()
    pid = _pias_pid()

    _step(log, 2, "Opening the search screen")
    P.answer_boxes(pid)                      # also the late "new version" box after a fresh login
    P.dismiss_boxes(pid)
    P.close_screen(pid)                      # always start from an empty grid
    open_search_screen()

    last = calendar.monthrange(year, m2)[1]
    d2 = min(int(d2 or last), last); d1 = max(1, int(d1 or 1))
    dst = policy_file(folder, line, year, m1, m2, d1, d2)
    _step(log, 3, "Setting the period")
    log("Listing %s policies, %02d/%02d - %02d/%02d/%d" % (LINES.get(line, line), d1, m1, d2, m2, year))

    def say(s):
        s = str(s)
        if s.startswith("period"): _step(log, 3, "Period set")
        elif s.startswith("line"): _step(log, 3, "Line set")
        elif s.startswith("searching"): _step(log, 4, "PIAS is searching")
        elif s.startswith("search took"): _step(log, 4, s)
        elif s.startswith("export clicked"): _step(log, 5, "Exporting")
        elif s.startswith("saving to"): _step(log, 5, "Writing the file")
        log("  " + s)

    try:
        P.run(pid, "%02d/%02d/%d" % (d1, m1, year), "%02d/%02d/%d" % (d2, m2, year), line, dst,
              log=say, stop=stop)
    except P.StepError as ex:
        raise UserError(str(ex))

    _step(log, 6, "Waiting for the file")
    said = 0
    for tick in range(900):
        P.dismiss_boxes(pid, rounds=1)
        if tick and tick % 15 == 0 and os.path.exists(dst):
            mb = os.path.getsize(dst) / 1048576.0
            if mb > said + 0.9:
                said = mb
                _step(log, 6, "Writing the file: %.1f MB so far" % mb)
        if os.path.exists(dst) and os.path.getsize(dst) > 0:
            size = -1
            while size != os.path.getsize(dst):
                size = os.path.getsize(dst)
                time.sleep(2)
            time.sleep(1.5)
            if P.dismiss_boxes(pid): log("pressed OK on the box PIAS shows when it is done")
            log("Saved %s (%.1f MB)" % (os.path.basename(dst), os.path.getsize(dst) / 1048576.0))
            P.close_screen(pid)
            return dst
        if stop(): raise UserError("Stopped.")
        time.sleep(2)
    raise UserError("PIAS did not write %s." % os.path.basename(dst))
