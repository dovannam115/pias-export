"""Drive one PIAS report: open its settings box, set the period, pick the unit, press Chấp nhận.

Like pias_policy.py this works on window handles with win32 and MSAA, never UIA: reading PIAS through
UIA stalls for minutes whenever it is busy.
"""
import ctypes
import re
import time

import pias_policy as P

user32 = P.user32
DIM_UNIT = "DM_DONVI"              # the unit list: PVI Digital has to be moved across or the report is empty
DLG = re.compile(r"^\d+\.\s")


class StepError(P.StepError):
    pass


def dialogs(pid):
    return [h for h in P.windows_of(pid) if DLG.match(P.text_of(h))]


def wait_dialog(pid, num, seconds=90):
    want = "%d. " % num
    t0 = time.time()
    while time.time() - t0 < seconds:
        for h in dialogs(pid):
            if P.text_of(h).startswith(want): return h
        P.answer_boxes(pid, rounds=1)
        time.sleep(0.6)
    return 0


def acc_name(h, tries=3):
    """MSAA sometimes answers with an empty name on the first ask; give it another go."""
    for i in range(tries):
        n = P.acc_name(h)
        if n: return n
        time.sleep(0.15)
    return ""


def named(dlg):
    """Every control of the box by name, falling back to the text drawn on it (buttons carry '&Chấp nhận')."""
    out = {}
    for c in P.children(dlg):
        n = acc_name(c, tries=1)
        if not n:
            n = (P.text_of(c) or "").replace("&", "").strip()
        if n: out.setdefault(n, c)
    return out


def unit_combo(dlg):
    """The list picker at the bottom left ('DM_...'); by name, else by where it sits."""
    kids = P.children(dlg)
    for c in kids:
        if acc_name(c, tries=1).startswith("DM_"): return c
    for c in kids:
        if acc_name(c, tries=3).startswith("DM_"): return c
    l, t, w, h = P.rect(dlg)
    best = 0
    for c in kids:
        cl, ct, cw, ch = P.rect(c)
        if "Window.b" not in P.class_of(c): continue
        if ct - t < h - 110 or cl - l > 220 or not (100 < cw < 220): continue
        if not best or ct > P.rect(best)[1]: best = c
    return best


def _edit_right_of(dlg, label):
    """The box next to a label, on the same line: the layout moves from report to report, the labels do not."""
    lab = next((c for c in P.children(dlg) if P.acc_name(c) == label), 0)
    if not lab: return 0
    ll, lt, lw, lh = P.rect(lab)
    best = 0
    for c in P.children(dlg):
        if "EDIT" not in P.class_of(c).upper(): continue
        cl, ct, cw, ch = P.rect(c)
        if abs(ct - lt) > 10 or cl <= ll: continue
        if not best or cl < P.rect(best)[0]: best = c
    return best


MONTHS = {m: i + 1 for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June",
     "July", "August", "September", "October", "November", "December"])}


def date_value(edit):
    """What a date box holds, as dd/MM/yyyy. These boxes keep their value out of the window text:
    it only shows up through MSAA, written out in full ("30 September 2026")."""
    s = (P.text_of(edit) or "").strip()
    if re.match(r"^\d\d/\d\d/\d{4}$", s): return s
    s = (P.acc_name(edit) or "").strip()
    if re.match(r"^\d\d/\d\d/\d{4}$", s): return s
    m = re.match(r"^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$", s)
    if m and m.group(2) in MONTHS:
        return "%02d/%02d/%s" % (int(m.group(1)), MONTHS[m.group(2)], m.group(3))
    return s


def date_boxes(dlg):
    a = _edit_right_of(dlg, "Từ ngày")
    b = _edit_right_of(dlg, "Đến ngày")
    if a and b: return [a, b]
    l, t, w, h = P.rect(dlg)            # fall back to the usual place, the top row of the box
    boxes = []
    for c in P.children(dlg):
        if "EDIT" not in P.class_of(c).upper(): continue
        cl, ct, cw, ch = P.rect(c)
        if ct - t < 90 and 55 < cw < 80: boxes.append((cl, c))
    boxes.sort()
    return [c for _, c in boxes]


def set_period(dlg, date_from, date_to, log=print):
    for _ in range(3):
        b = date_boxes(dlg)
        if len(b) < 2: raise StepError("the two date boxes were not found in the settings box")
        shown = [date_value(x) for x in b]
        if not all(re.match(r"^\d\d/\d\d/\d{4}$", s or "") for s in shown):
            raise StepError("those are not the date boxes (they show %s): the report is left alone"
                            % " and ".join("'%s'" % s for s in shown))
        if date_value(b[0]) != date_from: P.type_into(b[0], date_from)
        if date_value(b[1]) != date_to: P.type_into(b[1], date_to)
        time.sleep(0.5)
        got = [date_value(x) for x in date_boxes(dlg)]
        if got == [date_from, date_to]:
            log("period: %s - %s" % (date_from, date_to))
            return
    raise StepError("the period did not take (shows %s)" % " - ".join(got))


def pick_unit(dlg, log=print, dim=DIM_UNIT):
    """Set the list at the bottom left to DM_DONVI and move the unit to 'Nội dung được chọn'."""
    names = named(dlg)
    combo = unit_combo(dlg)
    if not combo:
        log("no unit list in this report, leaving the filter alone")
        return False
    if acc_name(combo) != dim:
        tid = user32.GetWindowThreadProcessId(combo, None)
        me = ctypes.windll.kernel32.GetCurrentThreadId()
        user32.AttachThreadInput(me, tid, True)
        try:
            user32.SetFocus(combo)
            time.sleep(0.25)
            # no Escape here: with the drop-down closed the box passes it on and the whole settings
            # window takes it as Hủy bỏ
            for key in (0x28, 0x26):                     # down through the list, then back up
                for _ in range(20):
                    if acc_name(combo, tries=1) == dim: break
                    user32.PostMessageW(combo, 0x0100, key, 0)
                    user32.PostMessageW(combo, 0x0101, key, 0)
                    time.sleep(0.35)
                if acc_name(combo, tries=1) == dim: break
        finally:
            user32.AttachThreadInput(me, tid, False)
    if acc_name(combo) != dim:
        raise StepError("could not switch the filter to %s (it shows %s)" % (dim, acc_name(combo) or "nothing"))
    log("filter: %s" % dim)
    arrow = names.get("SimpleButton1")
    if not arrow: raise StepError("the arrow that moves the unit across was not found")
    P.click(arrow, 14, 16)                               # the top arrow moves everything across
    time.sleep(1.0)
    log("unit moved across")
    return True


def accept(dlg, log=print):
    names = named(dlg)
    ok = names.get("Chấp nhận") or names.get("Chấp nhận ") or next(
        (c for c in P.children(dlg) if "Chấp nhận" in (P.text_of(c) or "")), 0)
    if not ok: raise StepError("the 'Chấp nhận' button was not found")
    P.click(ok, 49, 11)
    log("accepted")


def print_kind(pid, revenue_only=True, log=print, seconds=25):
    """'Chọn kiểu in' only comes up for the business performance reports: Doanh thu, Thực thu, Bồi thường."""
    t0 = time.time()
    win = 0
    while time.time() - t0 < seconds:
        for h in P.windows_of(pid):
            if P.text_of(h).strip().startswith("Chọn kiểu in"): win = h; break
        if win: break
        time.sleep(0.5)
    if not win:
        log("no print-kind box, the report is running")
        return False
    if revenue_only:
        boxes = []
        for c in P.children(win):
            cl, ct, cw, ch = P.rect(c)
            if "Window.b" in P.class_of(c) and cw < 30 and ch < 30: boxes.append((ct, cl, c))
        boxes.sort()
        if boxes:
            row_top = boxes[0][0]
            second = [b for b in boxes if b[0] > row_top + 5]
            for _, _, c in second[1:]:                  # keep the first of the second row (Doanh thu)
                P.click(c, 8, 9)
            time.sleep(0.5)
    ok = next((c for c in P.children(win) if P.acc_name(c) == "OK" or P.text_of(c).replace("&", "").strip() == "OK"), 0)
    if ok: P.click(ok, 41, 11)
    log("print kind: revenue only" if revenue_only else "print kind: all")
    return True


def run(pid, num, date_from, date_to, revenue_only=False, unit_filter=True, log=print):
    """From a clicked row to a running report."""
    dlg = wait_dialog(pid, num)
    if not dlg: raise StepError("the settings box of report %d did not open" % num)
    log("opened: %s" % P.text_of(dlg)[:60])
    time.sleep(1.0)
    set_period(dlg, date_from, date_to, log=log)
    if unit_filter:
        try:
            pick_unit(dlg, log=log)
        except StepError as ex:
            log("unit filter: %s" % ex)
    accept(dlg, log=log)
    time.sleep(1.5)
    why = P.crash_box(pid)
    if why:
        raise StepError("PIAS stopped on this report: %s" % why)
    print_kind(pid, revenue_only=revenue_only, log=log)
    return True
