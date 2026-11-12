"""Drive the PIAS desktop app (Pias2022.exe) to export one 'Báo cáo hiệu quả kinh doanh' report for one month.

Per report: make sure PIAS is open and logged in (pias_login.ps1, saved account), open the report list through
the menu (MSAA, no mouse), run the report for one calendar month (pias_export.ps1), wait for PIAS to write
BCHQKD_*.xlsx into its 'Excel Files' folder, check the period printed in the file, copy it to the save folder as
'T10.2026 - BH Kỹ thuật.xlsx' (a re-run gets the run date: 'T10.2026 - BH Kỹ thuật 12-11.xlsx'), and close the
copy PIAS opened in Excel. Clicks are posted to window handles: the user's mouse and keyboard are not touched.
"""
import calendar
import ctypes
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

EXE_CANDIDATES = [r"C:\PVI\Pias2022new\Pias2022.exe", r"C:\PVI\Pias2022\Pias2022.exe"]


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
def ps(script, *args, timeout=None):
    r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", os.path.join(HERE, script),
                        *map(str, args)], capture_output=True, timeout=timeout, creationflags=0x08000000)
    out = (r.stdout + r.stderr).decode("utf-8", "replace") if isinstance(r.stdout, bytes) else r.stdout + r.stderr
    if "\ufffd" in out:                               # console code page instead of UTF-8
        out = (r.stdout + r.stderr).decode("mbcs", "replace")
    out = out.strip()
    if r.returncode:
        lines = [x for x in out.splitlines() if x.strip()]
        msg = next((x for x in lines if not x.startswith(("At ", "+ ", "    +"))), lines[-1] if lines else "")
        raise UserError("%s: %s" % (script, msg or "failed"))
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


def unminimize():
    """A minimised PIAS hides its title from Windows and its controls have no size: restore it without
    activating it (SW_SHOWNOACTIVATE), so the user's current window keeps the keyboard."""
    h, _ = pias_main()
    if h and user32.IsIconic(h):
        user32.ShowWindow(h, 4)
        time.sleep(1.5)


def open_report_list():
    """Menu Báo cáo > Báo cáo hiệu quả kinh doanh, invoked through MSAA (DevExpress menus ignore posted clicks)."""
    import comtypes.client
    from comtypes import automation
    comtypes.client.GetModule("oleacc.dll")
    from comtypes.gen.Accessibility import IAccessible

    h, title = pias_main()
    if not h: raise UserError("PIAS main window not found.")
    if "[Báo cáo]" in title: return
    oleacc = ctypes.oledll.oleacc

    def acc(hwnd):
        p = POINTER(IAccessible)()
        oleacc.AccessibleObjectFromWindow(hwnd, -4, byref(IAccessible._iid_), byref(p))
        return p

    def kids(a):
        n = a.accChildCount
        arr = (automation.VARIANT * max(n, 1))()
        got = ctypes.c_long()
        oleacc.AccessibleChildren(a, 0, n, arr, byref(got))
        return [arr[i].value.QueryInterface(IAccessible) for i in range(got.value) if not isinstance(arr[i].value, int)]

    bars = []
    user32.EnumChildWindows(h, EnumProc(lambda c, _: bars.append(c) or True), 0)
    for c in bars:
        try:
            a = acc(c)
            if a.accName(0) != "MenuBar": continue
        except Exception:
            continue
        for name in ("Báo cáo", "Báo cáo hiệu quả kinh doanh"):
            a = next(k for k in kids(a) if k.accName(0) == name)
        a.accDoDefaultAction(0)
        for _ in range(30):
            time.sleep(1)
            if "[Báo cáo]" in pias_main()[1]:
                time.sleep(2)
                return
    raise UserError("Could not open the report list in PIAS.")


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


def wait_file(since, stop=lambda: False):
    """New BCHQKD_* file in PIAS's Excel Files folder, written after `since`, size settled."""
    deadline = time.time() + TIMEOUT_MIN * 60
    folder = pias_out()
    while time.time() < deadline:
        if stop(): raise UserError("Stopped.")
        new = [f for f in glob.glob(os.path.join(folder, "BCHQKD_*.xls*")) if os.path.getmtime(f) >= since]
        if new:
            f = max(new, key=os.path.getmtime)
            size = -1
            while size != os.path.getsize(f):
                size = os.path.getsize(f)
                time.sleep(2)
            return f
        time.sleep(3)
    raise UserError("No file from PIAS after %d min." % TIMEOUT_MIN)


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
def target(folder, year, month, num):
    base = "T%d.%d - %s" % (month, year, REPORTS[num])
    today = dt.date.today()
    path = os.path.join(folder, base + ".xlsx")
    if os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d.xlsx" % (base, today.day, today.month))
    i = 2
    while os.path.exists(path):
        path = os.path.join(folder, "%s %d-%d (%d).xlsx" % (base, today.day, today.month, i))
        i += 1
    return path


def export(year, month, num, folder, revenue_only=False, log=print, stop=lambda: False, tries=3):
    if num not in REPORTS: raise UserError("Unknown report %s." % num)
    log("Checking PIAS")
    unminimize()
    log(ps("pias_login.ps1", "-Exe", pias_exe(), "-Year", year, timeout=300).splitlines()[-1])
    last = calendar.monthrange(year, month)[1]
    want = (dt.date(year, month, 1), dt.date(year, month, last))
    args = ["-Report", num, "-From", "01/%02d/%d" % (month, year), "-To", "%02d/%02d/%d" % (last, month, year)]
    if revenue_only: args.append("-RevenueOnly")
    for attempt in range(1, tries + 1):
        if stop(): raise UserError("Stopped.")
        unminimize()
        open_report_list()
        since = time.time() - 1
        log("Running report %d in PIAS" % num + (" (try %d)" % attempt if attempt > 1 else ""))
        for line in ps("pias_export.ps1", *args, timeout=900).splitlines():
            log("  " + line)
        src = wait_file(since, stop)
        log("PIAS wrote " + os.path.basename(src))
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
        log("PIAS ran %s-%s instead of the chosen month, running again" % (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")))
    else:
        raise UserError("PIAS kept running the wrong period %s-%s, file not saved." % (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y")))
    os.makedirs(folder, exist_ok=True)
    dst = target(folder, year, month, num)
    shutil.copy2(src, dst)
    close_in_excel(src)
    return dst
