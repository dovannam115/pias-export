"""Drive the PIAS screen "Tìm kiếm đơn Bảo hiểm" to list policies and save them to a file.

Everything talks to window handles directly (win32, plus MSAA on the screen window itself), because
reading PIAS through UIA stalls for minutes once its grid holds a large result. Clicks are posted and the
dates are typed into the focused box, so the mouse and the keyboard stay with the user.
"""
import ctypes
import os
import time
from ctypes import POINTER, byref, wintypes

user32 = ctypes.windll.user32
EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

WM_CLOSE, WM_KEYDOWN, WM_KEYUP, WM_CHAR = 0x0010, 0x0100, 0x0101, 0x0102
WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP = 0x0200, 0x0201, 0x0202
WM_SETTEXT, BM_CLICK, EM_SETSEL = 0x000C, 0x00F5, 0x00B1
VK_HOME = 0x24
SCREEN_TITLE = "Tìm kiếm đơn Bảo hiểm"
BUSY_TITLE = "Form1"                      # the little "please wait" window PIAS shows while it works


class StepError(Exception):
    """Something on the screen was not where it should be: the message is shown to the user."""


# ---------------------------------------------------------------- window helpers
MONTHS = {m: i + 1 for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June",
     "July", "August", "September", "October", "November", "December"])}


def date_value(edit):
    """What a date box holds, as dd/MM/yyyy. Some of these boxes keep their value out of the window text
    and only show it through MSAA, written out in full ("30 September 2026")."""
    import re as _re
    s = (text_of(edit) or "").strip()
    if _re.match(r"^\d\d/\d\d/\d{4}$", s): return s
    s = (acc_name(edit) or "").strip()
    if _re.match(r"^\d\d/\d\d/\d{4}$", s): return s
    m = _re.match(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$", s)
    if m and m.group(2) in MONTHS:
        return "%02d/%02d/%s" % (int(m.group(1)), MONTHS[m.group(2)], m.group(3))
    return s


def field_value(h):
    """What a box holds. These WinForms boxes answer GetWindowText with nothing: the value only comes
    back through MSAA, so ask the window first and fall back to the accessible name."""
    s = (text_of(h) or "").strip()
    return s or (acc_name(h) or "").strip()


def rect(h):
    r = wintypes.RECT(); user32.GetWindowRect(h, byref(r))
    return r.left, r.top, r.right - r.left, r.bottom - r.top


def children(h):
    out = []
    user32.EnumChildWindows(h, EnumProc(lambda c, _: out.append(c) or True), 0)
    return out


def class_of(h):
    b = ctypes.create_unicode_buffer(80); user32.GetClassNameW(h, b, 80)
    return b.value


def text_of(h):
    b = ctypes.create_unicode_buffer(300); user32.GetWindowTextW(h, b, 300)
    return b.value


def acc_name(h):
    import comtypes.client
    comtypes.client.GetModule("oleacc.dll")
    from comtypes.gen.Accessibility import IAccessible
    p = POINTER(IAccessible)()
    try:
        ctypes.oledll.oleacc.AccessibleObjectFromWindow(h, -4, byref(IAccessible._iid_), byref(p))
        return (p.accName(0) or "").strip()
    except Exception:
        return ""


def windows_of(pid, title=None):
    out = []

    def cb(h, _):
        p = wintypes.DWORD(); user32.GetWindowThreadProcessId(h, byref(p))
        if p.value != pid or not user32.IsWindowVisible(h): return True
        if title is None or text_of(h).strip() == title: out.append(h)
        return True
    user32.EnumWindows(EnumProc(cb), 0)
    return out


def click(h, x=None, y=None):
    _, _, w, ht = rect(h)
    x = w // 2 if x is None else x
    y = ht // 2 if y is None else y
    lp = (y << 16) | (x & 0xFFFF)
    user32.PostMessageW(h, WM_MOUSEMOVE, 0, lp)
    user32.PostMessageW(h, WM_LBUTTONDOWN, 1, lp)
    user32.PostMessageW(h, WM_LBUTTONUP, 0, lp)


def type_into(h, text, clear=False):
    """Type into a box: PIAS reads what was typed, not what was set with WM_SETTEXT."""
    tid = user32.GetWindowThreadProcessId(h, None)
    me = ctypes.windll.kernel32.GetCurrentThreadId()
    user32.AttachThreadInput(me, tid, True)
    try:
        user32.SetFocus(h)
        time.sleep(0.2)
        if clear:
            # Home then Delete: selecting everything with EM_SETSEL(-1) makes the DevExpress masked
            # editors throw "arithmetic operation resulted in an overflow" and take PIAS down with them
            user32.PostMessageW(h, WM_KEYDOWN, VK_HOME, 0)
            user32.PostMessageW(h, WM_KEYUP, VK_HOME, 0)
            for _ in range(14):
                user32.PostMessageW(h, WM_KEYDOWN, 0x2E, 0)      # Delete
                user32.PostMessageW(h, WM_KEYUP, 0x2E, 0)
                time.sleep(0.03)
            time.sleep(0.15)
        user32.PostMessageW(h, WM_KEYDOWN, VK_HOME, 0)
        user32.PostMessageW(h, WM_KEYUP, VK_HOME, 0)
        for ch in "".join(c for c in str(text) if c.isdigit()):
            user32.PostMessageW(h, WM_CHAR, ord(ch), 0)
            time.sleep(0.07)
        time.sleep(0.3)
    finally:
        user32.AttachThreadInput(me, tid, False)


# ---------------------------------------------------------------- the screen
class Screen:
    def __init__(self, hwnd):
        self.h = hwnd
        self.left, self.top, self.w, self.ht = rect(hwnd)
        self.kids = children(hwnd)
        self.named = {}
        for c in self.kids:
            n = acc_name(c)
            if n: self.named.setdefault(n, c)

    def button(self, name):
        h = self.named.get(name)
        if not h: raise StepError("button '%s' is not on the search screen" % name)
        return h

    def _edits(self):
        out = []
        for c in self.kids:
            if "EDIT" not in class_of(c).upper(): continue
            l, t, w, h = rect(c)
            out.append((c, l - self.left, t - self.top, w, h))
        return out

    def date_boxes(self):
        """The two boxes of the period, found by what they hold: a date. Going by the labels is not safe
        here, the screen carries an empty 'Den ngay' label parked far from its own box, and following
        that one lands on the contract-code field instead."""
        import re as _re
        dated = []
        for c, x, y, w, h in self._edits():
            if not 50 < w < 90: continue
            if _re.match(r"^\d\d/\d\d/\d{4}$", date_value(c) or ""): dated.append((y, x, c))
        if len(dated) >= 2:
            dated.sort()
            same = sorted(d for d in dated if abs(d[0] - dated[0][0]) <= 8)
            if len(same) >= 2: return [same[0][2], same[1][2]]
        a = self.named.get("Từ ngày"); b = self.named.get("Đến ngày")
        if a and b:
            fa, fb = self._right_of(a), self._right_of(b)
            if fa and fb: return [fa, fb]
        tops = [e for e in self._edits() if e[2] < 60 and 55 < e[3] < 80]
        tops.sort(key=lambda e: e[1])
        return [e[0] for e in tops[:2]]

    def _right_of(self, label_hwnd):
        ll, lt, lw, lh = rect(label_hwnd)
        best = 0
        for c, x, y, w, h in self._edits():
            l, t = self.left + x, self.top + y
            if abs(t - lt) > 10 or l <= ll: continue
            if not best or l < rect(best)[0]: best = c
        return best

    def field_right_of(self, label):
        lab = self.named.get(label)
        if not lab: raise StepError("the field '%s' is not on the search screen" % label)
        ll, lt, lw, lh = rect(lab)
        best = None
        for c, x, y, w, h in self._edits():
            l, t = self.left + x, self.top + y
            if abs(t - lt) > 8 or l <= ll: continue
            if best is None or l < rect(best)[0]: best = c
        if best is None: raise StepError("no box next to '%s'" % label)
        return best


def find_screen(pid):
    ws = windows_of(pid, SCREEN_TITLE)
    return Screen(ws[0]) if ws else None


def busy(pid):
    return bool(windows_of(pid, BUSY_TITLE))


def message_boxes(pid):
    out = []
    for h in windows_of(pid, "Insurace Solution 2022"):
        ok = [c for c in children(h) if text_of(c).replace("&", "").strip() in ("OK", "Ok", "Đồng ý")]
        out.append((h, ok[0] if ok else 0))
    return out


def answer_boxes(pid, rounds=4, new_template=True):
    """Answer the questions PIAS asks by itself:
       - "Bấm Yes để chạy báo cáo theo mẫu mới, No để chạy báo cáo theo mẫu cũ" -> Yes
       - "a new version is on the server" (its title is the install folder) -> No
       - anything with just OK (no data, server error) -> OK
    Returns how many were answered."""
    answered = 0
    for _ in range(rounds):
        found = False
        for h in windows_of(pid):
            title = text_of(h).strip()
            if title.startswith("C:\\"):
                no = next((c for c in children(h) if text_of(c).replace("&", "").strip() in ("No", "Không")), 0)
                if no:
                    click(no, 20, 10); user32.PostMessageW(no, BM_CLICK, 0, 0)
                    answered += 1; found = True
                continue
            if title != "Insurace Solution 2022": continue
            btns = {text_of(c).replace("&", "").strip(): c for c in children(h)}
            pick = None
            if "Yes" in btns and "No" in btns: pick = btns["Yes"] if new_template else btns["No"]
            elif "OK" in btns: pick = btns["OK"]
            if pick:
                click(pick, 20, 10); user32.PostMessageW(pick, BM_CLICK, 0, 0)
                answered += 1; found = True
        if not found: break
        time.sleep(0.9)
    return answered


def dismiss_boxes(pid, rounds=4):
    seen = 0
    for _ in range(rounds):
        boxes = message_boxes(pid)
        if not boxes: break
        for h, ok in boxes:
            if ok:
                click(ok, 30, 10)
                user32.PostMessageW(ok, BM_CLICK, 0, 0)
            else:
                user32.PostMessageW(h, WM_CLOSE, 0, 0)
            seen += 1
        time.sleep(0.8)
    return seen


def crash_box(pid):
    """PIAS throws a .NET box ("Arithmetic operation resulted in an overflow") when the result is too
    big for its own export. Report it plainly instead of waiting for a file that will never come."""
    for h in windows_of(pid, "Microsoft .NET Framework"):
        texts = [text_of(c) for c in children(h)]
        why = next((x for x in texts if x and "exception" not in x.lower() and len(x) > 12), "")
        # Continue keeps PIAS alive for the next job; Quit would close it
        btns = {text_of(c).replace("&", "").strip(): c for c in children(h)}
        pick = btns.get("Continue") or btns.get("Quit")
        if pick:
            click(pick, 30, 12)
            user32.PostMessageW(pick, BM_CLICK, 0, 0)
            time.sleep(1.0)
        return why or "PIAS stopped with an error"
    return ""


def save_dialog(pid):
    for h in windows_of(pid):
        if class_of(h) == "#32770": return h
    return 0


def answer_save_dialog(pid, path, wait=180):
    """Fill the Windows 'Save As' box PIAS opens and press Save."""
    t0 = time.time()
    dlg = 0
    while time.time() - t0 < wait:
        dlg = save_dialog(pid)
        if dlg: break
        why = crash_box(pid)
        if why:
            raise StepError("PIAS could not export this period: %s. Ask for fewer months at a time." % why)
        time.sleep(0.7)
    if not dlg:
        raise StepError("PIAS did not ask where to save (no rows to export?)")
    edit = next((c for c in children(dlg) if class_of(c) == "Edit"), 0)
    if not edit: raise StepError("the Save As box has no file name field")
    user32.SendMessageW(edit, WM_SETTEXT, 0, ctypes.c_wchar_p(path))
    time.sleep(0.4)
    btn = next((c for c in children(dlg)
                if class_of(c) == "Button" and text_of(c).replace("&", "").strip() in ("Save", "Lưu", "OK")), 0)
    if btn: user32.PostMessageW(btn, BM_CLICK, 0, 0)
    else: user32.PostMessageW(dlg, 0x0111, 1, 0)      # IDOK
    return True


# ---------------------------------------------------------------- one export
def run(pid, date_from, date_to, line, save_path, log=print, stop=lambda: False, search_wait=3600):
    """Set the period and the line, search, export, answer the Save As box. Returns when PIAS is writing."""
    dismiss_boxes(pid)
    sc = find_screen(pid)
    if not sc: raise StepError("the policy search screen is not open")

    for attempt in range(1, 4):
        boxes = sc.date_boxes()
        if len(boxes) < 2: raise StepError("the two date boxes were not found")
        # these boxes insert what is typed instead of overwriting, so clear them first
        if date_value(boxes[0]) != date_from: type_into(boxes[0], date_from, clear=True)
        if date_value(boxes[1]) != date_to: type_into(boxes[1], date_to, clear=True)
        time.sleep(0.6)
        got = [date_value(b) for b in sc.date_boxes()]
        if got == [date_from, date_to]: break
    log("period: %s - %s" % tuple(got))
    if got != [date_from, date_to]:
        raise StepError("the period did not take (shows %s)" % " - ".join(got))

    if line:
        for attempt in range(1, 4):
            box = sc.field_right_of("Sản phẩm")
            type_into(box, line, clear=True)
            time.sleep(0.4)
            now = field_value(sc.field_right_of("Sản phẩm"))
            if now == str(line): break
        log("line: %s" % now)
        if now != str(line): raise StepError("the line did not take (shows '%s')" % now)

    click(sc.button("Tìm kiếm"), 30, 10)
    log("searching")
    t0 = time.time()
    time.sleep(2)
    while busy(pid) and time.time() - t0 < search_wait:
        if stop(): raise StepError("Stopped.")
        why = crash_box(pid)
        if why:
            raise StepError("PIAS could not search this period: %s. Ask for fewer months at a time." % why)
        time.sleep(2)
    log("search took %d s" % int(time.time() - t0))
    time.sleep(1.5)

    said = dismiss_boxes(pid)
    if said: log("PIAS showed a message after the search")

    click(sc.button("Xuất Exel"), 30, 10)
    log("export clicked")
    answer_save_dialog(pid, save_path)
    log("saving to %s" % save_path)
    return sc


def close_screen(pid, log=print, limit=20):
    """Press Thoát on every copy of the screen: it only closes through its own button, and one left open
    with a big grid makes everything after it slow."""
    closed = 0
    for _ in range(limit):
        ws = windows_of(pid, SCREEN_TITLE)
        if not ws: break
        h = ws[0]
        try:
            sc = Screen(h)
            click(sc.button("Thoát"), 30, 10)
        except StepError:
            user32.PostMessageW(h, WM_CLOSE, 0, 0)
        gone = False
        for _ in range(16):
            time.sleep(0.4)
            if h not in windows_of(pid, SCREEN_TITLE):
                gone = True
                break
        if not gone: break
        closed += 1
    return closed
