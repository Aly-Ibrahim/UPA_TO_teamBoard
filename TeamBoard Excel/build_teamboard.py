"""
Builds TeamBoard.xlsx — a fully interactive, macro-free team board for Excel (Microsoft 365 / Excel for the web).
Drives a real Excel instance through COM so every formula, table, slicer and chart is native.

Run:  python build_teamboard.py
"""
import os, random, datetime as dt
import pythoncom
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "TeamBoard.xlsx")
M = pythoncom.Missing
TODAY = dt.date.today()
FONT = "Segoe UI"

# ---------- palette ----------
NAVY, INK, MUTED, LINE = "1E2134", "1F2330", "6B7285", "E3E6EE"
BG, LANE, WHITE = "F4F6FB", "E9ECF3", "FFFFFF"
INDIGO, INDIGO_SOFT = "5B5BD6", "ECECFB"
RED, RED_SOFT, AMBER, AMBER_SOFT = "E5484D", "FFF1F1", "F19A1A", "FEF3E2"
GREEN, GREEN_SOFT, TEAL, PINK, SLATE = "2F9E6A", "E6F6EE", "0E9AB5", "C2489D", "9AA1B5"
LANE_COLORS = [SLATE, INDIGO, AMBER, RED, GREEN]          # To Do, In Progress, Review, Blocked, Done
PRIO_COLORS = {"Urgent": RED, "High": AMBER, "Normal": INDIGO, "Low": SLATE}

# ---------- Excel enums ----------
xlCenter, xlLeft, xlRight, xlTop, xlBottom = -4108, -4131, -4152, -4160, -4107
xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight = 7, 8, 9, 10
xlContinuous, xlThin, xlMedium, xlThick = 1, 2, -4138, 4
xlExpression, xlCellValue, xlEqual = 2, 1, 3
xlDoughnut, xlBarClustered, xlBarStacked, xlColumnClustered = -4120, 57, 58, 51


def rgb(h):
    h = h.lstrip("#")
    return int(h[0:2], 16) + int(h[2:4], 16) * 256 + int(h[4:6], 16) * 65536


def serial(d):
    if d is None:
        return None
    if isinstance(d, dt.datetime):
        return (d - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    return (d - dt.date(1899, 12, 30)).days


def day(n):
    return TODAY + dt.timedelta(days=n)


# =====================================================================
#  Sample data (replace with your own — just overwrite / delete rows)
# =====================================================================
TEAM = [
    ("Sara Ali", "sara@company.com", "Team Lead", "Management", 8),
    ("Omar Haddad", "omar@company.com", "Mobile Developer", "Engineering", 6),
    ("Lina Park", "lina@company.com", "Marketing Manager", "Marketing", 6),
    ("Jack Moore", "jack@company.com", "Data Analyst", "Analytics", 5),
    ("Maya Chen", "maya@company.com", "Customer Success", "Operations", 6),
    ("Ravi Kumar", "ravi@company.com", "Finance Analyst", "Finance", 5),
    ("Nora Fischer", "nora@company.com", "UX Designer", "Design", 6),
    ("Diego Santos", "diego@company.com", "Web Developer", "Engineering", 6),
    ("Aisha Bello", "aisha@company.com", "Content Writer", "Marketing", 5),
    ("Tom Becker", "tom@company.com", "QA Engineer", "Engineering", 6),
]
PROJECTS = [  # name, code, owner, start, target, description, added by, added on
    ("Website Redesign", "WEB", "Sara Ali", -40, 25, "New brand look, faster pages and a smoother signup flow", "Sara Ali", -42),
    ("Q4 Marketing Campaign", "MKT", "Lina Park", -30, 45, "Holiday campaign across email, social and paid ads", "Lina Park", -31),
    ("Mobile App v2", "APP", "Omar Haddad", -60, 10, "Offline mode, push notifications and dark theme", "Omar Haddad", -61),
    ("Customer Onboarding", "ONB", "Maya Chen", -20, 60, "Cut time-to-first-value from 14 to 5 days", "Maya Chen", -20),
    ("General / Ops", "OPS", "Sara Ali", -90, 90, "Recurring team work that isn't part of a project", "Sara Ali", -90),
    ("Annual Report 2026", "RPT", "Ravi Kumar", -1, 70, "Collect the year's KPIs and design the annual report", "Sara Ali", 0),
]
P = {p[1]: p[0] for p in PROJECTS}
# title, project code, owner, status, priority, due offset, progress, tags, added offset, completed offset
TASKS = [
    ("Finalize homepage wireframes", "WEB", "Nora Fischer", "Done", "High", -12, 1, "design", -38, -13),
    ("Build new pricing page", "WEB", "Diego Santos", "In Progress", "High", 3, .6, "frontend", -25, None),
    ("Write pricing page copy", "WEB", "Aisha Bello", "Review", "Normal", 1, .9, "copy", -24, None),
    ("QA signup flow on mobile", "WEB", "Tom Becker", "To Do", "Urgent", -1, 0, "qa, mobile", -10, None),
    ("Migrate blog to new CMS", "WEB", "Diego Santos", "Blocked", "Normal", 8, .3, "backend", -18, None),
    ("Accessibility audit", "WEB", "Nora Fischer", "To Do", "Normal", 14, 0, "a11y", -6, None),
    ("Set up analytics events", "WEB", "Jack Moore", "In Progress", "Normal", 5, .4, "data", -9, None),
    ("Optimize image loading", "WEB", "Diego Santos", "Done", "Low", -3, 1, "performance", -20, -4),
    ("Launch-day checklist", "WEB", "Sara Ali", "To Do", "High", 22, 0, "launch", -2, None),
    ("Campaign brief & goals", "MKT", "Lina Park", "Done", "High", -20, 1, "strategy", -30, -21),
    ("Holiday email sequence", "MKT", "Aisha Bello", "In Progress", "High", 6, .5, "email", -15, None),
    ("Social media calendar", "MKT", "Lina Park", "Review", "Normal", 2, .8, "social", -14, None),
    ("Paid ads creative", "MKT", "Nora Fischer", "To Do", "Normal", 10, 0, "ads, design", -8, None),
    ("Influencer shortlist", "MKT", "Lina Park", "To Do", "Low", 18, 0, "partners", -5, None),
    ("Landing page A/B test", "MKT", "Jack Moore", "Blocked", "High", -2, .2, "data, web", -12, None),
    ("Budget approval", "MKT", "Ravi Kumar", "Done", "Urgent", -16, 1, "finance", -28, -17),
    ("Press release draft", "MKT", "Aisha Bello", "To Do", "Normal", 12, 0, "pr", -1, None),
    ("Offline sync engine", "APP", "Omar Haddad", "In Progress", "Urgent", 4, .7, "core", -50, None),
    ("Push notification service", "APP", "Omar Haddad", "Review", "High", 0, .9, "backend", -40, None),
    ("Dark theme UI", "APP", "Nora Fischer", "Done", "Normal", -8, 1, "design", -45, -9),
    ("Regression test suite", "APP", "Tom Becker", "In Progress", "High", 7, .5, "qa", -21, None),
    ("App store screenshots", "APP", "Nora Fischer", "To Do", "Low", 9, 0, "store", -4, None),
    ("Crash reporting setup", "APP", "Tom Becker", "Done", "Normal", -11, 1, "qa", -35, -10),
    ("Beta tester feedback round", "APP", "Maya Chen", "To Do", "High", -3, 0, "feedback", -7, None),
    ("Release notes v2.0", "APP", "Aisha Bello", "To Do", "Normal", 10, 0, "docs", 0, None),
    ("Map current onboarding journey", "ONB", "Maya Chen", "Done", "High", -10, 1, "research", -19, -11),
    ("Welcome email series", "ONB", "Aisha Bello", "In Progress", "Normal", 9, .3, "email", -12, None),
    ("In-app product tour", "ONB", "Diego Santos", "To Do", "High", 20, 0, "frontend", -6, None),
    ("Onboarding success metrics", "ONB", "Jack Moore", "Review", "Normal", 4, .8, "data", -9, None),
    ("Customer interview round", "ONB", "Maya Chen", "In Progress", "Normal", 2, .5, "research", -8, None),
    ("Help center refresh", "ONB", "Aisha Bello", "To Do", "Low", 30, 0, "docs", -3, None),
    ("Kickoff & timeline", "RPT", "Sara Ali", "Done", "Normal", 0, 1, "planning", 0, 0),
    ("Collect department KPIs", "RPT", "Ravi Kumar", "To Do", "High", 14, 0, "finance", 0, None),
    ("Report design template", "RPT", "Nora Fischer", "To Do", "Normal", 35, 0, "design", 0, None),
    ("Weekly team sync notes", None, "Sara Ali", "Done", "Low", -7, 1, "ops", -9, -7),
    ("Vendor contract renewal", None, "Ravi Kumar", "In Progress", "Urgent", 1, .5, "finance", -11, None),
    ("Q4 hiring plan", None, "Sara Ali", "Review", "High", 5, .7, "people", -13, None),
    ("Dashboard data clean-up", None, "Jack Moore", "Done", "Normal", -5, 1, "data", -16, -2),
    ("Security awareness training", None, "Tom Becker", "To Do", "Normal", 16, 0, "it", -4, None),
    ("Customer churn analysis", None, "Jack Moore", "In Progress", "High", -4, .6, "data", -14, None),
]
CHAT = [  # minutes ago, who, message, about
    (60 * 26, "Sara Ali", "Welcome to the Team Room! Post updates here so everyone sees them live.", ""),
    (60 * 24, "Omar Haddad", "Offline sync is passing all tests on Android, iOS next.", "Offline sync engine"),
    (60 * 20, "Lina Park", "Need final sign-off on the holiday email subject lines by Monday.", "Holiday email sequence"),
    (60 * 5, "Tom Becker", "Signup flow breaks on small screens - logging details now.", "QA signup flow on mobile"),
    (60 * 3, "Ravi Kumar", "Kicked off the Annual Report 2026 project - KPIs due in two weeks.", "Annual Report 2026"),
    (95, "Maya Chen", "Interviews booked with 6 customers this week.", "Customer interview round"),
    (40, "Diego Santos", "Pricing page is on staging, feedback welcome!", "Build new pricing page"),
    (12, "Nora Fischer", "Dark theme shipped. Screenshots for the store are next.", "App store screenshots"),
]
STATUSES = [("To Do", "Not started yet"), ("In Progress", "Someone is working on it"),
            ("Review", "Waiting for review / approval"), ("Blocked", "Stuck - needs help"),
            ("Done", "Finished (keep this LAST - it counts as complete)")]
PRIORITIES = [("Urgent", "Drop everything"), ("High", "This week"), ("Normal", "Planned"), ("Low", "Nice to have")]
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


# =====================================================================
#  Excel helpers
# =====================================================================
class B:
    """Small formatting toolkit around a worksheet."""

    def __init__(self, ws):
        self.ws = ws

    def r(self, a):
        return self.ws.Range(a)

    def fmt(self, a, fill=None, color=None, bold=None, size=None, h=None, v=None, wrap=None, nf=None,
            merge=False, italic=None, indent=None, name=None):
        rg = self.r(a)
        if merge:
            rg.Merge()
        if fill:
            rg.Interior.Color = rgb(fill)
        f = rg.Font
        if name:
            f.Name = name
        if color:
            f.Color = rgb(color)
        if bold is not None:
            f.Bold = bold
        if italic is not None:
            f.Italic = italic
        if size:
            f.Size = size
        if h:
            rg.HorizontalAlignment = h
        if v:
            rg.VerticalAlignment = v
        if wrap is not None:
            rg.WrapText = wrap
        if nf:
            rg.NumberFormat = nf
        if indent is not None:
            rg.IndentLevel = indent
        return rg

    def border(self, a, color=LINE, weight=xlThin, edges=(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)):
        rg = self.r(a)
        for e in edges:
            b = rg.Borders(e)
            b.LineStyle = xlContinuous
            b.Weight = weight
            b.Color = rgb(color)

    def widths(self, spec):
        for col, w in spec.items():
            self.ws.Columns(col).ColumnWidth = w

    def heights(self, rows, hgt):
        for rr in rows:
            self.ws.Rows(rr).RowHeight = hgt

    def link(self, a, target, text, color=None, size=None, bold=None):
        cell = self.r(a)
        self.ws.Hyperlinks.Add(cell, "", target, "", text)
        cell.Value = text
        cell.Font.Underline = False
        cell.Font.Name = FONT
        if color:
            cell.Font.Color = rgb(color)
        if size:
            cell.Font.Size = size
        if bold is not None:
            cell.Font.Bold = bold

    def cf_expr(self, a, formula, fill=None, color=None, bold=None, border=None, strike=None):
        rg = self.r(a)
        self.ws.Activate()
        rg.Cells(1, 1).Select()          # relative refs in CF formulas are anchored on the active cell
        fc = rg.FormatConditions.Add(xlExpression, 1, formula)
        if fill:
            fc.Interior.Color = rgb(fill)
        if color:
            fc.Font.Color = rgb(color)
        if bold is not None:
            fc.Font.Bold = bold
        if strike is not None:
            fc.Font.Strikethrough = strike
        if border:
            for e in (xlLeft, xlRight, xlTop, xlBottom):
                bb = fc.Borders(e)
                bb.LineStyle = xlContinuous
                bb.Color = rgb(border)
        fc.StopIfTrue = False
        return fc

    def databar(self, a, color, maxv=None):
        db = self.r(a).FormatConditions.AddDatabar()
        db.BarColor.Color = rgb(color)
        db.BarFillType = 0  # solid
        db.MinPoint.Modify(0, 0)
        if maxv is not None:
            db.MaxPoint.Modify(0, maxv)
        try:
            db.BarBorder.Type = 0
        except Exception:
            pass
        return db

    def validate_list(self, a, source, msg=None):
        v = self.r(a).Validation
        v.Delete()
        v.Add(3, 1, 1, source)
        v.InCellDropdown = True
        v.IgnoreBlank = True
        if msg:
            v.InputTitle = msg[0]
            v.InputMessage = msg[1]


def banner(b, last_col, title, subtitle, home=True, rows=(1, 2, 3)):
    ws = b.ws
    b.fmt(f"A1:{last_col}3", fill=NAVY)
    b.heights([1], 8)
    b.heights([2], 34)
    b.heights([3], 20)
    b.fmt("B2", color=WHITE, bold=True, size=20, v=xlCenter).Value = title
    b.fmt("B3", color="AEB4D0", size=10, v=xlTop).Value = subtitle
    if home:
        b.link(f"{last_col}2", "'Home'!A1", "⌂  Home", color="C9CDF5", size=10, bold=True)
        b.fmt(f"{last_col}2", h=xlRight, v=xlCenter)


def page_bg(b, area="A1:AZ300"):
    b.fmt(area, fill=BG)


def window(xl, ws, zoom=90, headings=False, grid=False, freeze=None):
    ws.Activate()
    w = xl.ActiveWindow
    w.DisplayGridlines = grid
    w.DisplayHeadings = headings
    w.Zoom = zoom
    if freeze:
        w.FreezePanes = False
        w.SplitRow, w.SplitColumn = freeze
        w.FreezePanes = True
    ws.Range("A1").Select()


def kpi_card(b, a_strip, a_label, a_value, a_sub, accent, label, value_formula, sub_formula, nf=None,
             value_color=INK):
    """A KPI tile built from cells: colored top strip, label, big value and a sub-line."""
    b.fmt(a_strip, fill=accent, merge=True)
    b.fmt(a_label, fill=WHITE, color=MUTED, size=9, bold=True, merge=True, h=xlLeft, indent=1).Value = label
    vc = b.fmt(a_value, fill=WHITE, color=value_color, size=24, bold=True, merge=True, h=xlLeft, v=xlCenter, indent=1)
    vc.Formula2 = value_formula
    if nf:
        vc.NumberFormat = nf
    sc = b.fmt(a_sub, fill=WHITE, color=MUTED, size=9, merge=True, h=xlLeft, v=xlTop, indent=1)
    sc.Formula2 = sub_formula


def style_chart(ch, title, legend=True):
    ch.HasTitle = True
    ch.ChartTitle.Text = title
    tr = ch.ChartTitle.Format.TextFrame2.TextRange.Font
    tr.Size, tr.Bold = 11, True
    tr.Fill.ForeColor.RGB = rgb(INK)
    ch.ChartTitle.Left = 8
    ca = ch.ChartArea
    ca.Format.Fill.ForeColor.RGB = rgb(WHITE)
    ca.Format.Line.Visible = True
    ca.Format.Line.ForeColor.RGB = rgb(LINE)
    ca.RoundedCorners = True
    ca.Format.TextFrame2.TextRange.Font.Name = FONT
    ca.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(MUTED)
    ch.HasLegend = legend
    if legend:
        ch.Legend.Position = -4107  # bottom
        ch.Legend.Format.TextFrame2.TextRange.Font.Size = 9


def add_chart(ws, ctype, top_left, bottom_right):
    tl, br = ws.Range(top_left), ws.Range(bottom_right)
    left, top = tl.Left + 4, tl.Top + 4
    width = br.Left + br.Width - tl.Left - 8
    height = br.Top + br.Height - tl.Top - 8
    shp = ws.Shapes.AddChart2(-1, ctype, left, top, width, height)
    return shp.Chart


def grid_lines_light(ch):
    try:
        ax = ch.Axes(2)
        ax.MajorGridlines.Format.Line.ForeColor.RGB = rgb("EEF0F5")
        ax.Format.Line.Visible = False
    except Exception:
        pass


# =====================================================================
#  Build
# =====================================================================
def build():
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    wb = xl.Workbooks.Add()
    try:
        _build(xl, wb)
        if os.path.exists(OUT):
            os.remove(OUT)
        wb.SaveAs(OUT, FileFormat=51)
        print("saved", OUT)
    finally:
        wb.Close(False)
        xl.Quit()


def _build(xl, wb):
    st = wb.Styles("Normal").Font
    st.Name, st.Size = FONT, 10

    names = ["Home", "Dashboard", "Board", "Room", "My Work", "Calendar", "Scorecards",
             "Tasks", "Projects", "Team", "Settings", "Calc"]
    while wb.Worksheets.Count < len(names):
        wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    S = {}
    for i, n in enumerate(names, 1):
        ws = wb.Worksheets(i)
        ws.Name = n
        S[n] = ws
    tabs = {"Home": NAVY, "Dashboard": INDIGO, "Board": INDIGO, "Room": PINK, "My Work": TEAL,
            "Calendar": TEAL, "Scorecards": AMBER, "Tasks": GREEN, "Projects": GREEN, "Team": GREEN,
            "Settings": SLATE, "Calc": SLATE}
    for n, c in tabs.items():
        S[n].Tab.Color = rgb(c)

    build_settings(xl, wb, S["Settings"])
    finish = [build_team(xl, wb, S["Team"]), build_projects(xl, wb, S["Projects"]), build_tasks(xl, wb, S["Tasks"])]
    build_room_table(xl, wb, S["Room"])        # every table exists before any formula points at it
    names_and_lists(wb)
    for f in finish:
        f()
    validations(wb, S)
    build_calc(xl, wb, S["Calc"])
    build_dashboard(xl, wb, S["Dashboard"])
    build_board(xl, wb, S["Board"])
    build_room(xl, wb, S["Room"])
    build_mywork(xl, wb, S["My Work"])
    build_calendar(xl, wb, S["Calendar"])
    build_scorecards(xl, wb, S["Scorecards"])
    build_home(xl, wb, S["Home"])
    slicers(xl, wb, S)
    areas = {"Home": "A1:N32", "Dashboard": "A1:P87", "Board": "A1:P66", "Room": "A1:L53", "My Work": "A1:L37",
             "Calendar": "A1:I48", "Scorecards": "A1:U47"}
    xl.PrintCommunication = False
    for n in names:
        ps = S[n].PageSetup
        ps.Orientation = 2
        if n in areas:
            ps.PrintArea = areas[n]
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
    xl.PrintCommunication = True
    S["Calc"].Visible = 0
    xl.CalculateFull()
    S["Home"].Activate()


# ---------------------------------------------------------------------
def make_table(ws, top_left_cell, headers, rows, name, style="TableStyleLight1"):
    r0 = ws.Range(top_left_cell).Row
    c0 = ws.Range(top_left_cell).Column
    data = [tuple(headers)] + [tuple(r) for r in rows]
    rg = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + len(data) - 1, c0 + len(headers) - 1))
    rg.Value = tuple(data)
    lo = ws.ListObjects.Add(1, rg, None, 1)
    lo.Name = name
    lo.TableStyle = style
    lo.ShowTableStyleRowStripes = False
    if lo.DataBodyRange is not None:
        body = lo.DataBodyRange
        for e in (12,):  # inside horizontal hairlines
            bb = body.Borders(e)
            bb.LineStyle = xlContinuous
            bb.Color = rgb("E6E8EF")
        bb = body.Borders(xlEdgeBottom)
        bb.LineStyle = xlContinuous
        bb.Color = rgb("E6E8EF")
    return lo


def header_colors(lo, inputs, calcs):
    for c in inputs:
        h = lo.ListColumns(c).Range.Cells(1, 1)
        h.Interior.Color = rgb(INDIGO)
        h.Font.Color = rgb(WHITE)
        h.Font.Bold = True
    for c in calcs:
        h = lo.ListColumns(c).Range.Cells(1, 1)
        h.Interior.Color = rgb("D9DCE6")
        h.Font.Color = rgb(INK)
        h.Font.Bold = True
        lo.ListColumns(c).DataBodyRange.Font.Color = rgb("4A5068")
        lo.ListColumns(c).DataBodyRange.Interior.Color = rgb("F6F7FA")


def sheet_intro(b, text_cells):
    for addr, text in text_cells:
        b.fmt(addr, color=MUTED, size=9).Value = text


def build_settings(xl, wb, ws):
    b = B(ws)
    banner(b, "H", "Settings", "Statuses (board columns) and priorities used in every dropdown")
    b.widths({"A": 2, "B": 18, "C": 44, "D": 3, "E": 14, "F": 30, "G": 3, "H": 16})
    b.fmt("B5", bold=True, size=11, color=INK).Value = "Statuses  (= board lanes, top to bottom = left to right)"
    lo = make_table(ws, "B6", ["Status", "Meaning"], STATUSES, "tblStatus")
    b.fmt("B13", color=MUTED, size=9, italic=True).Value = \
        "Rename freely. Keep exactly 5 rows; the LAST row is treated as 'done'."
    b.fmt("E5", bold=True, size=11, color=INK).Value = "Priorities"
    make_table(ws, "E6", ["Priority", "Meaning"], PRIORITIES, "tblPriority")
    b.fmt("B15", bold=True, size=11, color=INK).Value = "Months (used by the Calendar)"
    make_table(ws, "B16", ["Month"], [(m,) for m in MONTHS], "tblMonths")
    window(xl, ws, headings=True, grid=False)


def build_team(xl, wb, ws):
    b = B(ws)
    banner(b, "J", "Team", "Everyone who can own tasks. Add a row at the bottom to add a member.")
    b.widths({"A": 2, "B": 18, "C": 24, "D": 20, "E": 16, "F": 11, "G": 12, "H": 11, "I": 14, "J": 12})
    headers = ["Name", "Email", "Role", "Department", "Capacity", "Open Tasks", "Overdue", "Load"]
    lo = make_table(ws, "B5", headers, [t + (None, None, None) for t in TEAM], "tblTeam")
    header_colors(lo, headers[:5], headers[5:])
    b.fmt("B4", color=MUTED, size=9).Value = \
        "Blue headers = you type.  Grey headers = calculated.  Capacity = how many open tasks this person can comfortably hold."
    window(xl, ws, headings=True, freeze=(5, 2))

    def finish():
        team_formulas(b, lo)
    return finish


def team_formulas(b, lo):
    lo.ListColumns("Open Tasks").DataBodyRange.Formula2 = \
        '=COUNTIFS(tblTasks[Owner],[@Name],tblTasks[Status],"<>"&DoneStatus)'
    lo.ListColumns("Overdue").DataBodyRange.Formula2 = '=COUNTIFS(tblTasks[Owner],[@Name],tblTasks[State],"Overdue")'
    lo.ListColumns("Load").DataBodyRange.Formula2 = '=IF(N([@Capacity])=0,"",[@[Open Tasks]]/[@Capacity])'
    lo.ListColumns("Load").DataBodyRange.NumberFormat = "0%"
    b.databar(lo.ListColumns("Load").DataBodyRange.Address, AMBER, 1)


def build_projects(xl, wb, ws):
    b = B(ws)
    banner(b, "P", "Projects", "Every project the team runs. Add a row at the bottom - it appears in the Room instantly.")
    b.widths({"A": 2, "B": 24, "C": 8, "D": 15, "E": 12, "F": 12, "G": 44, "H": 15, "I": 12,
              "J": 8, "K": 8, "L": 11, "M": 9, "N": 10, "O": 15, "P": 3})
    headers = ["Project", "Code", "Owner", "Start Date", "Target Date", "Description", "Added By", "Added On",
               "Tasks", "Done", "Progress", "Overdue", "Days Left", "Health"]
    rows = [(p[0], p[1], p[2], serial(day(p[3])), serial(day(p[4])), p[5], p[6], serial(day(p[7])),
             None, None, None, None, None, None) for p in PROJECTS]
    lo = make_table(ws, "B5", headers, rows, "tblProjects")
    header_colors(lo, headers[:8], headers[8:])
    for c in ["Start Date", "Target Date", "Added On"]:
        lo.ListColumns(c).DataBodyRange.NumberFormat = "d mmm yyyy"
    window(xl, ws, headings=True, freeze=(5, 2))
    b.fmt("B4", color=MUTED, size=9).Value =         "Blue headers = you type (Code = 2-4 letters shown on cards).  Grey headers = calculated automatically."
    return lambda: project_formulas(b, lo)


def project_formulas(b, lo):
    lc = lambda n: lo.ListColumns(n).DataBodyRange
    lc("Tasks").Formula2 = "=COUNTIFS(tblTasks[Project],[@Project])"
    lc("Done").Formula2 = "=COUNTIFS(tblTasks[Project],[@Project],tblTasks[Status],DoneStatus)"
    lc("Progress").Formula2 = "=IF([@Tasks]=0,0,[@Done]/[@Tasks])"
    lc("Progress").NumberFormat = "0%"
    lc("Overdue").Formula2 = '=COUNTIFS(tblTasks[Project],[@Project],tblTasks[State],"Overdue")'
    lc("Days Left").Formula2 = '=IF([@[Target Date]]="","",[@[Target Date]]-TODAY())'
    lc("Health").Formula2 = (
        '=IF([@Project]="","",IF(AND([@Tasks]>0,[@Done]=[@Tasks]),"🏁 Complete",'
        'LET(expected,IF(OR([@[Start Date]]="",[@[Target Date]]="",N([@[Target Date]])<=N([@[Start Date]])),0,'
        'MIN(1,MAX(0,(TODAY()-[@[Start Date]])/([@[Target Date]]-[@[Start Date]])))),'
        'IF(OR([@Overdue]>=3,AND(ISNUMBER([@[Days Left]]),N([@[Days Left]])<0)),"⛔ Off Track",'
        'IF(OR([@Overdue]>=1,[@Progress]<expected-0.25),"⚠ At Risk","✅ On Track")))))')
    b.databar(lc("Progress").Address, GREEN, 1)
    health = lc("Health").Address
    b.cf_expr(health, f'=LEFT({lc("Health").Cells(1,1).Address.replace("$", "")},1)="⛔"', color=RED, bold=True)
    b.cf_expr(health, f'=LEFT({lc("Health").Cells(1,1).Address.replace("$", "")},1)="⚠"', color="B26A00", bold=True)
    b.cf_expr(health, f'=LEFT({lc("Health").Cells(1,1).Address.replace("$", "")},1)="✅"', color=GREEN, bold=True)


def build_tasks(xl, wb, ws):
    b = B(ws)
    banner(b, "W", "Tasks", "The team's single source of truth. Type in the empty row under the table to add a task.")
    widths = {"A": 2, "B": 7, "C": 34, "D": 21, "E": 14, "F": 12, "G": 9, "H": 11, "I": 11, "J": 9, "K": 16,
              "L": 26, "M": 14, "N": 11, "O": 12, "P": 9, "Q": 10, "R": 8, "S": 30, "T": 9, "U": 9, "V": 8, "W": 3}
    b.widths(widths)
    headers = ["ID", "Task", "Project", "Owner", "Status", "Priority", "Start Date", "Due Date", "Progress",
               "Tags", "Notes", "Added By", "Added On", "Completed On",
               "Days Left", "State", "On Time", "Card", "SortKey", "DueKey", "Visible"]
    rows = []
    rnd = random.Random(7)
    for t in TASKS:
        title, code, owner, status, prio, due, prog, tags, added, comp = t
        start = day(added + rnd.randint(0, 2))
        rows.append((None, title, P.get(code, "General / Ops") if code else "General / Ops", owner, status, prio,
                     serial(start), serial(day(due)), prog, tags, None, owner if rnd.random() < .5 else "Sara Ali",
                     serial(day(added)), serial(day(comp)) if comp is not None else None,
                     None, None, None, None, None, None, None))
    lo = make_table(ws, "B6", headers, rows, "tblTasks", style="TableStyleLight1")
    lc = lambda n: lo.ListColumns(n).DataBodyRange
    header_colors(lo, headers[1:14], ["ID"] + headers[14:])
    for c in ["Start Date", "Due Date", "Added On", "Completed On"]:
        lo.ListColumns(c).DataBodyRange.NumberFormat = "d mmm yyyy"
    b.fmt("B4", color=MUTED, size=9).Value = ("Blue headers = you type (use the dropdowns).  Grey headers = automatic - "
                                             "don't type there.  Dates: press Ctrl+; for today.  "
                                             "Don't sort/insert in the middle - add new tasks at the bottom so IDs stay stable.")
    b.link("C5", "'Board'!A1", "→ See it on the Board", color=INDIGO, size=9, bold=True)
    b.link("D5", "'Dashboard'!A1", "→ Dashboard", color=INDIGO, size=9, bold=True)
    window(xl, ws, headings=True, freeze=(6, 3))
    return lambda: task_formulas(b, lo)


def task_formulas(b, lo):
    lc = lambda n: lo.ListColumns(n).DataBodyRange
    lc("ID").Formula2 = '="T-"&TEXT(ROW()-ROW(tblTasks[[#Headers],[ID]]),"000")'
    lc("Days Left").Formula2 = '=IF(OR([@Task]="",[@[Due Date]]="",[@Status]=DoneStatus),"",[@[Due Date]]-TODAY())'
    lc("State").Formula2 = ('=IF([@Task]="","",IF([@Status]=DoneStatus,"Done",IF([@[Due Date]]="","No date",'
                            'IF([@[Due Date]]<TODAY(),"Overdue",IF([@[Due Date]]-TODAY()<=7,"Due soon","On track")))))')
    lc("On Time").Formula2 = ('=IF(OR([@Status]<>DoneStatus,[@[Completed On]]="",[@[Due Date]]=""),"",'
                              'IF([@[Completed On]]<=[@[Due Date]],1,0))')
    lc("Card").Formula2 = (
        '=IF([@Task]="","",[@Task]&CHAR(10)&"👤 "&IF([@Owner]="","Unassigned",[@Owner])'
        '&IF([@[Due Date]]="","","    📅 "&TEXT([@[Due Date]],"d mmm"))&CHAR(10)'
        '&IF([@State]="Overdue","⚠ OVERDUE   ","")'
        '&IFERROR("["&XLOOKUP([@Project],tblProjects[Project],tblProjects[Code])&"]   ","")'
        '&IF([@Tags]="","","#"&SUBSTITUTE(TRIM([@Tags]),", "," #"))'
        '&IF(AND(ISNUMBER([@Progress]),N([@Progress])>0,[@Status]<>DoneStatus),"   · "&TEXT([@Progress],"0%")&" done",""))')
    lc("SortKey").Formula2 = ('=IFERROR(MATCH([@Priority],lstPriority,0),9)*100000'
                              '+IF([@[Due Date]]="",99999,[@[Due Date]]-DATE(2000,1,1))')
    lc("DueKey").Formula2 = '=IF([@[Due Date]]="",99999999,[@[Due Date]])'
    lc("Visible").Formula2 = "=SUBTOTAL(103,[@Task])"
    for c in ["Start Date", "Due Date", "Added On", "Completed On"]:
        lc(c).NumberFormat = "d mmm yyyy"
    lc("Progress").NumberFormat = "0%"
    lc("On Time").NumberFormat = '"✓";;"✗"'
    lc("On Time").HorizontalAlignment = xlCenter
    lc("Days Left").NumberFormat = '0;[Red]-0;"today"'
    lc("Task").Font.Bold = True
    b.databar(lc("Progress").Address, INDIGO, 1)
    a = lambda n: lc(n).Cells(1, 1).Address.replace("$", "")
    b.cf_expr(lc("State").Address, f'={a("State")}="Overdue"', fill=RED_SOFT, color=RED, bold=True)
    b.cf_expr(lc("State").Address, f'={a("State")}="Due soon"', fill=AMBER_SOFT, color="B26A00")
    b.cf_expr(lc("State").Address, f'={a("State")}="Done"', fill=GREEN_SOFT, color=GREEN)
    for p, c in PRIO_COLORS.items():
        b.cf_expr(lc("Priority").Address, f'={a("Priority")}="{p}"', color=c, bold=True)
    for c in ["Card", "SortKey", "DueKey", "Visible"]:
        lo.ListColumns(c).Range.EntireColumn.Hidden = True


def build_room_table(xl, wb, ws):
    b = B(ws)
    rows = []
    now = dt.datetime.combine(TODAY, dt.time(16, 0))
    for mins, who, msg, about in CHAT:
        rows.append((serial(now - dt.timedelta(minutes=mins)), who, msg, about))
    lo = make_table(ws, "H12", ["When", "Who", "Message", "About"], rows, "tblChat", style="TableStyleLight1")
    header_colors(lo, ["When", "Who", "Message", "About"], [])
    lo.ListColumns("When").DataBodyRange.NumberFormat = "d mmm h:mm AM/PM"
    lo.ListColumns("When").DataBodyRange.HorizontalAlignment = xlLeft
    lo.ListColumns("Message").DataBodyRange.WrapText = True


def names_and_lists(wb):
    n = wb.Names
    n.Add("lstStatus", "=tblStatus[Status]")
    n.Add("lstPriority", "=tblPriority[Priority]")
    n.Add("lstMembers", "=tblTeam[Name]")
    n.Add("lstProjects", "=tblProjects[Project]")
    n.Add("lstMonths", "=tblMonths[Month]")
    n.Add("DoneStatus", "=INDEX(tblStatus[Status],ROWS(tblStatus[Status]))")
    n.Add("MaxCards", "=25")


def validations(wb, S):
    t = S["Tasks"].ListObjects("tblTasks")
    bt = B(S["Tasks"])
    dv = lambda col, src, msg=None: bt.validate_list(t.ListColumns(col).DataBodyRange.Address, src, msg)
    dv("Project", "=lstProjects", ("Project", "Pick a project (add new ones on the Projects sheet)"))
    dv("Owner", "=lstMembers", ("Owner", "Who is responsible?"))
    dv("Status", "=lstStatus", ("Status", "Changing this moves the card on the Board"))
    dv("Priority", "=lstPriority")
    dv("Added By", "=lstMembers")
    v = t.ListColumns("Progress").DataBodyRange.Validation
    v.Delete()
    v.Add(2, 1, 1, "0", "1")
    v.ErrorMessage = "Enter a percentage between 0% and 100%"
    p = S["Projects"].ListObjects("tblProjects")
    bp = B(S["Projects"])
    bp.validate_list(p.ListColumns("Owner").DataBodyRange.Address, "=lstMembers")
    bp.validate_list(p.ListColumns("Added By").DataBodyRange.Address, "=lstMembers")
    c = S["Room"].ListObjects("tblChat")
    br = B(S["Room"])
    br.validate_list(c.ListColumns("Who").DataBodyRange.Address, "=lstMembers")


# ---------------------------------------------------------------------
def build_calc(xl, wb, ws):
    b = B(ws)
    b.fmt("A1", bold=True).Value = "Helper calculations for charts - do not edit (sheet is hidden)"
    vis = "tblTasks[Visible],1"
    ws.Range("B2").Value = "Status"
    ws.Range("C2").Value = "Tasks"
    ws.Range("B3").Formula2 = "=lstStatus"
    ws.Range("C3").Formula2 = f"=COUNTIFS(tblTasks[Status],B3#,{vis})"
    ws.Range("B9").Value = "Priority"
    ws.Range("C9").Value = "Open"
    ws.Range("B10").Formula2 = "=lstPriority"
    ws.Range("C10").Formula2 = f'=COUNTIFS(tblTasks[Priority],B10#,{vis},tblTasks[Status],"<>"&DoneStatus)'
    ws.Range("B19:E19").Value = ("Member", "Done", "Open", "Overdue")
    ws.Range("B20").Formula2 = "=lstMembers"
    ws.Range("C20").Formula2 = f"=COUNTIFS(tblTasks[Owner],B20#,{vis},tblTasks[Status],DoneStatus)"
    ws.Range("E20").Formula2 = f'=COUNTIFS(tblTasks[Owner],B20#,{vis},tblTasks[State],"Overdue")'
    ws.Range("D20").Formula2 = f'=COUNTIFS(tblTasks[Owner],B20#,{vis},tblTasks[Status],"<>"&DoneStatus)-E20#'
    ws.Range("B44:E44").Value = ("Week of", "Label", "Added", "Completed")
    ws.Range("B45").Formula2 = "=TODAY()-WEEKDAY(TODAY(),2)+1-7*SEQUENCE(8,1,7,-1)"
    ws.Range("C45").Formula2 = '=TEXT(B45#,"d mmm")'
    ws.Range("D45").Formula2 = f'=COUNTIFS(tblTasks[Added On],">="&B45#,tblTasks[Added On],"<"&(B45#+7),{vis})'
    ws.Range("E45").Formula2 = f'=COUNTIFS(tblTasks[Completed On],">="&B45#,tblTasks[Completed On],"<"&(B45#+7),{vis})'
    ws.Range("B45:B52").NumberFormat = "d mmm"
    ws.Range("B58").Value = "Total visible"
    ws.Range("C58").Formula2 = '=SUMPRODUCT((tblTasks[Visible]=1)*(tblTasks[Task]<>""))'
    ws.Range("B59").Value = "Done visible"
    ws.Range("C59").Formula2 = f"=COUNTIFS(tblTasks[Status],DoneStatus,{vis})"
    ws.Range("B60").Value = "Done"
    ws.Range("C60").Formula2 = "=IF(C58=0,0,C59/C58)"
    ws.Range("B61").Value = "Remaining"
    ws.Range("C61").Formula2 = "=1-C60"
    ws.Range("B63").Value = "Ring title"
    ws.Range("C63").Formula2 = '="Overall completion  ·  "&TEXT(C60,"0%")'


# ---------------------------------------------------------------------
def build_dashboard(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    cols = "BCDEFGHIJKLMNO"
    b.widths({"A": 2, **{c: 11 for c in cols}, "P": 2})
    banner(b, "O", "Team Dashboard", "Live from the Tasks sheet  ·  click the slicers to filter every chart, card and list")
    b.fmt("L3:O3", color="AEB4D0", size=9, h=xlRight, merge=True).Formula2 = \
        '="Updated "&TEXT(NOW(),"ddd d mmm, h:mm AM/PM")'
    # filter strip
    b.heights([4], 10)
    b.fmt("B5", color=MUTED, size=9, bold=True).Value = "FILTERS"
    b.fmt("J5:O5", size=9, h=xlRight, merge=True, bold=True, color=INDIGO).Formula2 = \
        '=IF(Calc!C58<COUNTA(tblTasks[Task]),"🔎 Filtered: showing "&Calc!C58&" of "&COUNTA(tblTasks[Task])&" tasks - clear slicers to see all","Showing all "&COUNTA(tblTasks[Task])&" tasks")'
    b.heights(range(6, 10), 26)
    b.heights([10], 10)

    # KPI cards (rows 11-15)
    b.heights([11], 5)
    b.heights([12], 18)
    b.heights([13], 36)
    b.heights([14], 18)
    b.heights([15], 6)
    vis = "tblTasks[Visible],1"
    cards = [
        ("TOTAL TASKS", INDIGO, "=Calc!C58", '="across "&COUNTA(UNIQUE(FILTER(tblTasks[Project],(tblTasks[Visible]=1)*(tblTasks[Task]<>""),"")))&" projects"', None, INK),
        ("OPEN", TEAL, f'=COUNTIFS({vis},tblTasks[Status],"<>"&DoneStatus,tblTasks[Task],"<>")',
         f'=COUNTIFS({vis},tblTasks[Status],INDEX(lstStatus,2))&" in progress"', None, INK),
        ("COMPLETED", GREEN, "=Calc!C60", '=Calc!C59&" tasks done"', "0%", GREEN),
        ("OVERDUE", RED, f'=COUNTIFS({vis},tblTasks[State],"Overdue")', '="need attention now"', None, RED),
        ("DUE IN 7 DAYS", AMBER, f'=COUNTIFS({vis},tblTasks[State],"Due soon")', '="coming up this week"', None, INK),
        ("BLOCKED", PINK, f'=COUNTIFS({vis},tblTasks[Status],INDEX(lstStatus,4))', '="waiting on someone"', None, INK),
        ("TEAM SCORE", NAVY, '=IFERROR(ROUND(AVERAGE(FILTER(Scorecards!I8#,Scorecards!I8#<>"")),0),0)',
         '=REPT("★",ROUND(D13/20,0))&REPT("☆",5-ROUND(D13/20,0))', None, INK),
    ]
    for i, (label, accent, vf, sf, nf, vcol) in enumerate(cards):
        c1, c2 = cols[i * 2], cols[i * 2 + 1]
        if label == "TEAM SCORE":
            sf = f'=REPT("★",ROUND({c1}13/20,0))&REPT("☆",5-ROUND({c1}13/20,0))'
        kpi_card(b, f"{c1}11:{c2}11", f"{c1}12:{c2}12", f"{c1}13:{c2}13", f"{c1}14:{c2}14", accent, label, vf, sf, nf, vcol)
        b.border(f"{c1}11:{c2}15", color=BG, weight=xlThick, edges=(xlEdgeLeft, xlEdgeRight))
        b.fmt(f"{c1}15:{c2}15", fill=WHITE)

    # charts
    b.heights(range(17, 33), 15)
    ch = add_chart(ws, xlDoughnut, "B17", "E32")
    calc = wb.Worksheets("Calc")
    ch.SetSourceData(calc.Range("B60:C61"))
    style_chart(ch, "Overall completion", legend=False)
    try:
        ch.ChartTitle.Formula = "=Calc!$C$63"
    except Exception:
        pass
    ch.ChartGroups(1).DoughnutHoleSize = 72
    s = ch.SeriesCollection(1)
    s.Points(1).Format.Fill.ForeColor.RGB = rgb(GREEN)
    s.Points(2).Format.Fill.ForeColor.RGB = rgb("E6E8EF")
    s.Format.Line.Visible = False

    ch = add_chart(ws, xlDoughnut, "F17", "J32")
    ch.SetSourceData(calc.Range("B3:C7"))
    style_chart(ch, "Tasks by status")
    ch.ChartGroups(1).DoughnutHoleSize = 55
    s = ch.SeriesCollection(1)
    for i, c in enumerate(LANE_COLORS, 1):
        s.Points(i).Format.Fill.ForeColor.RGB = rgb(c)
    s.Format.Line.ForeColor.RGB = rgb(WHITE)
    s.HasDataLabels = True
    s.DataLabels().ShowValue = True
    s.DataLabels().Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(WHITE)
    s.DataLabels().Format.TextFrame2.TextRange.Font.Bold = True
    ch.Legend.Position = -4152  # right

    ch = add_chart(ws, xlBarClustered, "K17", "O32")
    ch.SetSourceData(calc.Range("B10:C13"))
    style_chart(ch, "Open tasks by priority", legend=False)
    s = ch.SeriesCollection(1)
    for i, p in enumerate(["Urgent", "High", "Normal", "Low"], 1):
        s.Points(i).Format.Fill.ForeColor.RGB = rgb(PRIO_COLORS[p])
    ch.ChartGroups(1).GapWidth = 60
    ch.Axes(1).ReversePlotOrder = True
    s.HasDataLabels = True
    grid_lines_light(ch)

    b.heights(range(34, 52), 15)
    ch = add_chart(ws, xlBarStacked, "B34", "H51")
    ch.SetSourceData(calc.Range("B19:E29"))
    style_chart(ch, "Workload by member")
    for i, c in enumerate([GREEN, INDIGO, RED], 1):
        ch.SeriesCollection(i).Format.Fill.ForeColor.RGB = rgb(c)
    ch.ChartGroups(1).GapWidth = 50
    ch.Axes(1).ReversePlotOrder = True
    grid_lines_light(ch)
    try:  # grow with the team: point the chart at the spilled member list
        ch.SeriesCollection(1).XValues = "=Calc!$B$20#"
    except Exception:
        pass

    ch = add_chart(ws, xlColumnClustered, "I34", "O51")
    ch.SetSourceData(calc.Range("C44:E52"))
    style_chart(ch, "Weekly pace: tasks added vs completed (last 8 weeks)")
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(INDIGO)
    ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = rgb(GREEN)
    ch.ChartGroups(1).GapWidth = 80
    ch.ChartGroups(1).Overlap = -10
    grid_lines_light(ch)

    # Needs attention
    r = 54
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "🔥  Needs attention"
    b.fmt(f"E{r}", color=MUTED, size=9, v=xlCenter).Value = "Overdue or Urgent tasks that are still open (top 10)"
    hdr = ["ID", "PRIORITY", "DUE", "STATE", "OWNER", "PROJECT", "TASK"]
    for i, h in enumerate(hdr):
        b.fmt(f"{cols[i]}{r+1}", bold=True, size=9, color=MUTED, fill=WHITE).Value = h
    b.fmt(f"H{r+1}:O{r+1}", fill=WHITE)
    b.r(f"B{r+2}").Formula2 = (
        '=LET(m,(tblTasks[Visible]=1)*(tblTasks[Task]<>"")*(((tblTasks[State]="Overdue")+(tblTasks[Priority]="Urgent")*(tblTasks[Status]<>DoneStatus))>0),'
        'IFERROR(TAKE(SORTBY(FILTER(HSTACK(tblTasks[ID],tblTasks[Priority],tblTasks[Due Date],tblTasks[State],tblTasks[Owner],'
        'IFERROR(XLOOKUP(tblTasks[Project],tblProjects[Project],tblProjects[Code]),"OPS"),tblTasks[Task]),m),FILTER(tblTasks[DueKey],m),1),10),'
        '"🎉 Nothing needs attention right now"))')
    body = f"B{r+2}:O{r+11}"
    b.fmt(body, fill=WHITE, v=xlCenter)
    b.heights(range(r + 1, r + 12), 20)
    b.fmt(f"D{r+2}:D{r+11}", nf="d mmm", h=xlLeft)
    b.fmt(f"H{r+2}:H{r+11}", bold=True, color=INK)
    b.border(f"B{r+1}:O{r+11}", color=LINE, edges=(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight))
    b.border(f"B{r+2}:O{r+11}", color="EEF0F5", edges=(12,))  # inside horizontal
    for p, c in PRIO_COLORS.items():
        b.cf_expr(f"C{r+2}:C{r+11}", f'=C{r+2}="{p}"', color=c, bold=True)
    b.cf_expr(f"E{r+2}:E{r+11}", f'=E{r+2}="Overdue"', color=RED, bold=True)

    # Project progress + heatmap
    r = 68
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "📁  Project health"
    b.fmt(f"I{r}", bold=True, size=13, color=INK).Value = "🔥  Who's working on what"
    for i, h in enumerate(["HEALTH", "PROGRESS", "DONE", "OVERDUE", "PROJECT"]):
        b.fmt(f"{cols[i]}{r+1}", bold=True, size=9, color=MUTED, fill=WHITE).Value = h
    b.fmt(f"G{r+1}:H{r+1}", fill=WHITE)
    b.r(f"B{r+2}").Formula2 = ('=IFERROR(HSTACK(tblProjects[Health],tblProjects[Progress],'
                              'tblProjects[Done]&" / "&tblProjects[Tasks],tblProjects[Overdue],tblProjects[Project]),"")')
    b.fmt(f"B{r+2}:H{r+16}", fill=WHITE, v=xlCenter)
    b.fmt(f"C{r+2}:C{r+16}", nf="0%")
    b.fmt(f"F{r+2}:F{r+16}", bold=True, color=INK)
    b.databar(f"C{r+2}:C{r+16}", GREEN, 1)
    b.cf_expr(f"B{r+2}:B{r+16}", f'=LEFT(B{r+2},1)="⛔"', color=RED, bold=True)
    b.cf_expr(f"B{r+2}:B{r+16}", f'=LEFT(B{r+2},1)="⚠"', color="B26A00", bold=True)
    b.cf_expr(f"B{r+2}:B{r+16}", f'=LEFT(B{r+2},1)="✅"', color=GREEN, bold=True)
    b.cf_expr(f"E{r+2}:E{r+16}", f'=N(E{r+2})>0', color=RED, bold=True)
    b.border(f"B{r+1}:H{r+16}", color=LINE)
    b.heights(range(r + 1, r + 17), 20)

    b.fmt(f"I{r+1}", bold=True, size=9, color=MUTED, fill=WHITE).Value = "MEMBER"
    b.r(f"J{r+1}").Formula2 = "=UPPER(TRANSPOSE(lstStatus))"
    b.fmt(f"J{r+1}:N{r+1}", bold=True, size=9, color=MUTED, fill=WHITE, h=xlCenter)
    b.fmt(f"O{r+1}", bold=True, size=9, color=MUTED, fill=WHITE, h=xlCenter).Value = "OPEN"
    b.r(f"I{r+2}").Formula2 = "=lstMembers"
    b.r(f"J{r+2}").Formula2 = f"=COUNTIFS(tblTasks[Owner],I{r+2}#,tblTasks[Status],TRANSPOSE(lstStatus),{vis})"
    b.r(f"O{r+2}").Formula2 = f'=COUNTIFS(tblTasks[Owner],I{r+2}#,tblTasks[Status],"<>"&DoneStatus,{vis})'
    b.fmt(f"I{r+2}:O{r+16}", fill=WHITE, v=xlCenter)
    b.fmt(f"J{r+2}:O{r+16}", h=xlCenter, nf='0;-0;""')
    cs = b.r(f"J{r+2}:N{r+16}").FormatConditions.AddColorScale(2)
    cs.ColorScaleCriteria(1).Type = 0
    cs.ColorScaleCriteria(1).Value = 0
    cs.ColorScaleCriteria(1).FormatColor.Color = rgb(WHITE)
    cs.ColorScaleCriteria(2).Type = 2  # highest value
    cs.ColorScaleCriteria(2).FormatColor.Color = rgb("8E8CF0")
    b.databar(f"O{r+2}:O{r+16}", AMBER)
    b.border(f"I{r+1}:O{r+16}", color=LINE)

    b.fmt("B86", color=MUTED, size=9, italic=True).Value = \
        "Tip: slicers filter the Tasks table for everyone editing at the same moment - clear them when you're done (the ⨯ icon on each slicer)."
    window(xl, ws, zoom=85)


# ---------------------------------------------------------------------
LANE_START_ROW = 13


def build_board(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    lane_cols = [("B", "C", "D"), ("E", "F", "G"), ("H", "I", "J"), ("K", "L", "M"), ("N", "O", "P")]
    spec = {"A": 2}
    for s, c, g in lane_cols:
        spec.update({s: 0.9, c: 30, g: 2})
    b.widths(spec)
    banner(b, "O", "Board", "Kanban view of every task  ·  change a task's Status on the Tasks sheet and its card moves here instantly")
    b.heights([4], 8)
    b.heights(range(5, 9), 19)
    b.heights([9], 26)
    b.fmt("C5", color=MUTED, size=9, bold=True).Value = "🔍  SEARCH CARDS"
    sc = b.fmt("C6:C7", fill=WHITE, size=11, merge=True, v=xlCenter, indent=1)
    b.border("C6:C7", color=INDIGO, weight=xlMedium)
    b.link("C8", "'Tasks'!B7", "＋ Add or edit tasks on the Tasks sheet", color=INDIGO, size=9, bold=True)
    b.fmt("F5", color=MUTED, size=9, bold=True).Value = "FILTERS  (shared with the Dashboard - type a word in Search and press Enter)"
    b.heights([10], 12)
    b.heights([11], 24)
    b.heights([12], 5)
    lanes_end = LANE_START_ROW + 2 * 25 - 2
    for k, (s, c, g) in enumerate(lane_cols, 1):
        hdr = b.fmt(f"{s}11:{c}11", merge=True, bold=True, size=10, color=INK, v=xlCenter, indent=1, fill=LANE)
        hdr.Formula2 = (f'=UPPER(INDEX(lstStatus,{k}))&"    "&COUNTIFS(tblTasks[Status],INDEX(lstStatus,{k}),'
                        f'tblTasks[Visible],1,tblTasks[Task],"<>")')
        b.fmt(f"{s}12:{c}12", fill=LANE_COLORS[k - 1])
        b.fmt(f"{s}{LANE_START_ROW - 0}:{c}{lanes_end + 1}", fill=LANE)
        b.fmt(f"{c}{LANE_START_ROW}:{c}{lanes_end}", wrap=True, v=xlTop, size=9.5, color=INK, indent=1)
        b.fmt(f"{s}{LANE_START_ROW}:{s}{lanes_end}", size=1, color=LANE)
        b.r(f"{s}{LANE_START_ROW}").Formula2 = (
            f'=LET(st,INDEX(lstStatus,{k}),kw,TRIM($C$6),'
            'm,(tblTasks[Status]=st)*(tblTasks[Visible]=1)*(tblTasks[Task]<>"")'
            '*IF(kw="",1,ISNUMBER(SEARCH(kw,tblTasks[Task]&" "&tblTasks[Tags]&" "&tblTasks[Owner]&" "&tblTasks[Project]))),'
            'd,IFERROR(SORTBY(FILTER(HSTACK(tblTasks[Priority],tblTasks[Card]),m),FILTER(tblTasks[SortKey],m),1),""),'
            'n,MIN(ROWS(d),MaxCards),'
            'IF(INDEX(d,1,1)="","",LET(r,SEQUENCE(2*n-1),IF(MOD(r,2)=1,INDEX(d,(r+1)/2,{1,2}),""))))')
        card = f"{c}{LANE_START_ROW}:{c}{lanes_end}"
        stripe = f"{s}{LANE_START_ROW}:{s}{lanes_end}"
        top_card, top_stripe = f"{c}{LANE_START_ROW}", f"{s}{LANE_START_ROW}"
        b.cf_expr(card, f'=ISNUMBER(SEARCH("OVERDUE",{top_card}))', fill="FFF5F5", border="F4C7C9")
        if k == 5:
            b.cf_expr(card, f'=LEN({top_card})>0', fill="FBFCFD", color="8A90A6", border="E1E4EC")
        else:
            b.cf_expr(card, f'=LEN({top_card})>0', fill=WHITE, border="DADDE6")
        for p, col in PRIO_COLORS.items():
            b.cf_expr(stripe, f'={top_stripe}="{p}"', fill=col, color=col)
    for i in range(25):
        ws.Rows(LANE_START_ROW + 2 * i).RowHeight = 62
        if i < 24:
            ws.Rows(LANE_START_ROW + 2 * i + 1).RowHeight = 6
    b.fmt(f"C{lanes_end + 3}", color=MUTED, size=9, italic=True).Value = \
        "Cards are sorted by priority, then due date. Colored edge = priority (red Urgent · amber High · indigo Normal · grey Low). Up to 25 cards per lane."
    window(xl, ws, zoom=85)


# ---------------------------------------------------------------------
def build_room(xl, wb, ws):
    b = B(ws)
    page_bg(b, "A1:L11")
    page_bg(b, "A12:G300")
    page_bg(b, "L12:L300")
    b.widths({"A": 2, "B": 12, "C": 5, "D": 15, "E": 62, "F": 15, "G": 3, "H": 17, "I": 15, "J": 46, "K": 24, "L": 2})
    banner(b, "K", "Team Room", "Everything that happens, live: new projects, new tasks, completed work and team messages")
    b.heights([4], 10)
    # pulse ticker
    b.heights([5], 40)
    pulse = b.fmt("B5:K5", merge=True, fill=WHITE, size=13, bold=True, color=INK, v=xlCenter, indent=1)
    pulse.Formula2 = (
        '="TODAY   ·   🆕 "&COUNTIFS(tblTasks[Added On],TODAY())&" new tasks      ✅ "'
        '&COUNTIFS(tblTasks[Completed On],TODAY())&" completed      🚀 "&COUNTIFS(tblProjects[Added On],TODAY())'
        '&" new projects      💬 "&COUNTIFS(tblChat[When],">="&TODAY(),tblChat[When],"<"&(TODAY()+1))'
        '&" messages      📅 "&COUNTIFS(tblTasks[Due Date],TODAY(),tblTasks[Status],"<>"&DoneStatus)&" due today"')
    b.border("B5:K5", color=LINE)
    b.fmt("B5:K5", fill=WHITE)
    b.r("B5").Interior.Color = rgb(WHITE)
    b.heights([6], 10)
    # spotlight boxes
    b.heights([7], 18)
    b.heights([8], 50)
    b.heights([9], 10)
    b.fmt("B7:E7", merge=True, fill=PINK, color=WHITE, bold=True, size=9, indent=1, v=xlCenter).Value = "🚀  NEWEST PROJECT"
    sp = b.fmt("B8:E8", merge=True, fill=WHITE, wrap=True, v=xlCenter, indent=1, size=10)
    sp.Formula2 = ('=IFERROR(TAKE(tblProjects[Project],-1)&"   ["&TAKE(tblProjects[Code],-1)&"]"&CHAR(10)'
                   '&"Owner: "&TAKE(tblProjects[Owner],-1)&"   ·   Target: "&TEXT(TAKE(tblProjects[Target Date],-1),"d mmm yyyy")'
                   '&CHAR(10)&TAKE(tblProjects[Description],-1),"No projects yet")')
    b.fmt("H7:K7", merge=True, fill=INDIGO, color=WHITE, bold=True, size=9, indent=1, v=xlCenter).Value = "🆕  LATEST TASK"
    lt = b.fmt("H8:K8", merge=True, fill=WHITE, wrap=True, v=xlCenter, indent=1, size=10)
    lt.Formula2 = ('=IFERROR(TAKE(tblTasks[ID],-1)&"  ·  "&TAKE(tblTasks[Task],-1)&CHAR(10)&"Owner: "&TAKE(tblTasks[Owner],-1)'
                   '&"   ·   "&TAKE(tblTasks[Project],-1)&CHAR(10)&"Status: "&TAKE(tblTasks[Status],-1)&"   ·   Priority: "&TAKE(tblTasks[Priority],-1),"No tasks yet")')
    b.border("B7:E8", color=LINE)
    b.border("H7:K8", color=LINE)

    # feed
    b.fmt("B10", bold=True, size=13, color=INK).Value = "📡  Activity feed  (newest first)"
    b.heights([10], 26)
    b.heights([11], 6)
    for i, h in enumerate(["WHEN", "", "WHO", "WHAT HAPPENED", "DATE"]):
        b.fmt(f"{'BCDEF'[i]}12", bold=True, size=9, color=MUTED, fill=WHITE, v=xlCenter).Value = h
    b.r("B13").Formula2 = (
        '=LET(icon,LAMBDA(col,txt,IF(LEN(col)>=0,txt)),'
        'tA,IFERROR(FILTER(HSTACK(tblTasks[Added On],icon(tblTasks[Task],"🆕"),tblTasks[Added By],'
        '"added task “"&tblTasks[Task]&"” to "&tblTasks[Project]),(tblTasks[Task]<>"")*ISNUMBER(tblTasks[Added On])),{0,"","",""}),'
        'tC,IFERROR(FILTER(HSTACK(tblTasks[Completed On],icon(tblTasks[Task],"✅"),tblTasks[Owner],'
        '"completed “"&tblTasks[Task]&"”"),(tblTasks[Status]=DoneStatus)*ISNUMBER(tblTasks[Completed On])),{0,"","",""}),'
        'tP,IFERROR(FILTER(HSTACK(tblProjects[Added On],icon(tblProjects[Project],"🚀"),tblProjects[Added By],'
        '"launched project “"&tblProjects[Project]&"” · target "&TEXT(tblProjects[Target Date],"d mmm")),'
        '(tblProjects[Project]<>"")*ISNUMBER(tblProjects[Added On])),{0,"","",""}),'
        'tM,IFERROR(FILTER(HSTACK(tblChat[When],icon(tblChat[Message],"💬"),tblChat[Who],'
        'tblChat[Message]&IF(tblChat[About]="",""," · re: "&tblChat[About])),(tblChat[Message]<>"")*ISNUMBER(tblChat[When])),{0,"","",""}),'
        'all,VSTACK(tA,tC,tP,tM),f,FILTER(all,INDEX(all,,1)>0,{0,"","",""}),'
        't,TAKE(SORTBY(f,INDEX(f,,1),-1),40),w,INDEX(t,,1),ago,TODAY()-INT(w),'
        'IF(w=0,"",HSTACK(IF(ago=0,"Today",IF(ago=1,"Yesterday",ago&" days ago")),DROP(t,,1),'
        'TEXT(w,IF(MOD(w,1)=0,"ddd d mmm","ddd d mmm, h:mm")))))')
    feed = "B13:F52"
    b.fmt(feed, fill=WHITE, v=xlCenter, wrap=True)
    b.fmt("C13:C52", h=xlCenter, size=12)
    b.fmt("D13:D52", bold=True, color=INK)
    b.fmt("B13:B52", color=MUTED, size=9)
    b.fmt("F13:F52", color=MUTED, size=9)
    b.heights(range(12, 53), 30)
    b.border("B12:F52", color=LINE)
    b.border("B13:F52", color="EEF0F5", edges=(12,))
    b.cf_expr("B13:F52", '=$B13="Today"', fill="F7F7FE")
    b.cf_expr("E13:E52", '=$C13="✅"', color=GREEN)
    b.cf_expr("E13:E52", '=$C13="🚀"', color=PINK, bold=True)

    # chat
    b.fmt("H10", bold=True, size=13, color=INK).Value = "✍  Post to the room"
    b.fmt("H11:K11", merge=True, color=MUTED, size=9).Value = \
        "Type in the first empty row under this table. When: press Ctrl+; then space then Ctrl+Shift+;  (date + time)."
    b.heights([11], 18)
    window(xl, ws, zoom=90)


# ---------------------------------------------------------------------
def build_mywork(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    b.widths({"A": 2, "B": 8, "C": 38, "D": 22, "E": 13, "F": 10, "G": 11, "H": 9, "I": 11, "J": 3, "K": 26, "L": 2})
    banner(b, "K", "My Work", "Your personal to-do list across every project  ·  pick your name below")
    b.heights([4], 10)
    b.heights([5], 30)
    b.fmt("B5", color=MUTED, size=9, bold=True, v=xlCenter).Value = "I AM"
    me = b.fmt("C5", fill=WHITE, size=13, bold=True, color=INDIGO, v=xlCenter, indent=1)
    me.Value = TEAM[0][0]
    b.border("C5", color=INDIGO, weight=xlMedium)
    b.validate_list("C5", "=lstMembers", ("Who are you?", "Pick your name to see your tasks"))
    b.fmt("D5:I5", merge=True, color=MUTED, size=9, v=xlCenter, indent=1).Value = \
        "◀ click the cell and use the dropdown arrow. (Everyone shares this cell - each person just picks their own name.)"
    b.heights([6], 10)
    b.fmt("B7", bold=True, size=13, color=INK).Value = "📋  My open tasks"
    b.heights([7], 26)
    hdr = ["ID", "TASK", "PROJECT", "STATUS", "PRIORITY", "DUE", "DAYS", "STATE"]
    for i, h in enumerate(hdr):
        b.fmt(f"{'BCDEFGHI'[i]}8", bold=True, size=9, color=MUTED, fill=WHITE, v=xlCenter).Value = h
    b.r("B9").Formula2 = (
        '=LET(me,$C$5,m,(tblTasks[Owner]=me)*(tblTasks[Status]<>DoneStatus)*(tblTasks[Task]<>""),'
        'IFERROR(SORTBY(FILTER(HSTACK(tblTasks[ID],tblTasks[Task],tblTasks[Project],tblTasks[Status],tblTasks[Priority],'
        'tblTasks[Due Date],tblTasks[Days Left],tblTasks[State]),m),FILTER(tblTasks[DueKey],m),1),"🎉 Nothing open - great job!"))')
    b.fmt("B9:I28", fill=WHITE, v=xlCenter)
    b.heights(range(8, 29), 22)
    b.fmt("C9:C28", bold=True, color=INK)
    b.fmt("G9:G28", nf="ddd d mmm", h=xlLeft)
    b.fmt("H9:H28", nf='0;[Red]-0;"today"', h=xlCenter)
    b.border("B8:I28", color=LINE)
    b.border("B9:I28", color="EEF0F5", edges=(12,))
    b.cf_expr("I9:I28", '=I9="Overdue"', fill=RED_SOFT, color=RED, bold=True)
    b.cf_expr("I9:I28", '=I9="Due soon"', fill=AMBER_SOFT, color="B26A00")
    for p, c in PRIO_COLORS.items():
        b.cf_expr("F9:F28", f'=F9="{p}"', color=c, bold=True)

    b.fmt("B31", bold=True, size=13, color=INK).Value = "✅  Recently completed"
    b.heights([31], 26)
    b.r("B32").Formula2 = (
        '=LET(me,$C$5,m,(tblTasks[Owner]=me)*(tblTasks[Status]=DoneStatus),'
        'IFERROR(TAKE(SORTBY(FILTER(HSTACK(tblTasks[ID],tblTasks[Task],tblTasks[Project],tblTasks[Completed On]),m),'
        'FILTER(tblTasks[Completed On],m),-1),5),"Nothing completed yet"))')
    b.fmt("B32:E36", fill=WHITE, color=MUTED, v=xlCenter)
    b.fmt("E32:E36", nf="ddd d mmm", h=xlLeft)
    b.heights(range(32, 37), 22)
    b.border("B32:E36", color=LINE)

    # stat sidebar
    stats = [
        ("OPEN TASKS", INDIGO, '=COUNTIFS(tblTasks[Owner],$C$5,tblTasks[Status],"<>"&DoneStatus)', '="across "&IFERROR(ROWS(UNIQUE(FILTER(tblTasks[Project],(tblTasks[Owner]=$C$5)*(tblTasks[Status]<>DoneStatus)))),0)&" projects"', INK),
        ("OVERDUE", RED, '=COUNTIFS(tblTasks[Owner],$C$5,tblTasks[State],"Overdue")', '="tackle these first"', RED),
        ("DUE THIS WEEK", AMBER, '=COUNTIFS(tblTasks[Owner],$C$5,tblTasks[State],"Due soon")', '="in the next 7 days"', INK),
        ("COMPLETED", GREEN, '=COUNTIFS(tblTasks[Owner],$C$5,tblTasks[Status],DoneStatus)', '="tasks finished"', GREEN),
        ("MY SCORE", NAVY, '=IFERROR(XLOOKUP($C$5,Scorecards!B8#,Scorecards!I8#),"-")', '=IFERROR(XLOOKUP($C$5,Scorecards!B8#,Scorecards!J8#)&"   rank #"&XLOOKUP($C$5,Scorecards!B8#,Scorecards!K8#),"")', INK),
    ]
    r = 8
    for label, accent, vf, sf, vcol in stats:
        b.fmt(f"K{r}", fill=accent)
        ws.Rows(r).RowHeight = max(ws.Rows(r).RowHeight, 22)
        kpi_card(b, f"K{r}", f"K{r+1}", f"K{r+2}", f"K{r+3}", accent, label, vf, sf, None, vcol)
        b.fmt(f"K{r}", fill=accent)
        r += 5
    window(xl, ws)


# ---------------------------------------------------------------------
def build_calendar(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    b.widths({"A": 2, **{c: 23 for c in "BCDEFGH"}, "I": 2})
    banner(b, "H", "Calendar", "Tasks by due date  ·  ⚠ overdue   ✓ done   • open  ·  Dashboard slicers apply here too")
    b.heights([4], 10)
    b.heights([5], 28)
    b.fmt("B5", color=MUTED, size=9, bold=True, v=xlCenter, h=xlRight).Value = "MONTH  ▸"
    m = b.fmt("C5", fill=WHITE, size=12, bold=True, color=INDIGO, v=xlCenter, indent=1)
    m.Formula2 = "=INDEX(lstMonths,MONTH(TODAY()))"
    b.validate_list("C5", "=lstMonths", ("Month", "Pick a month"))
    b.fmt("D5", color=MUTED, size=9, bold=True, v=xlCenter, h=xlRight).Value = "YEAR  ▸"
    y = b.fmt("E5", fill=WHITE, size=12, bold=True, color=INDIGO, v=xlCenter, indent=1, nf="0")
    y.Formula2 = "=YEAR(TODAY())"
    b.border("C5", color=INDIGO, weight=xlMedium)
    b.border("E5", color=INDIGO, weight=xlMedium)
    b.fmt("F5:H5", merge=True, color=MUTED, size=9, v=xlCenter, indent=1).Value = \
        "Pick another month/year to browse. Type =YEAR(TODAY()) back to follow today."
    wb.Names.Add("calMonthNum", "=MATCH(Calendar!$C$5,lstMonths,0)")
    wb.Names.Add("CalFirst", "=DATE(Calendar!$E$5,calMonthNum,1)")
    wb.Names.Add("CalStart", "=CalFirst-WEEKDAY(CalFirst,2)+1")
    b.heights([6], 10)
    b.heights([7], 22)
    for i, d in enumerate(["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]):
        b.fmt(f"{'BCDEFGH'[i]}7", fill=NAVY, color=WHITE, bold=True, size=9, h=xlCenter, v=xlCenter).Value = d
    grid = "B8:H13"
    b.r(grid).Formula2 = (
        '=LET(d,CalStart+(ROW()-ROW($B$8))*7+COLUMN()-COLUMN($B$8),'
        'm,(tblTasks[Due Date]=d)*(tblTasks[Visible]=1)*(tblTasks[Task]<>""),n,SUM(m),'
        't,IF(n=0,"",TEXTJOIN(CHAR(10),TRUE,TAKE(FILTER(IF(tblTasks[State]="Overdue","⚠ ",'
        'IF(tblTasks[Status]=DoneStatus,"✓ ","• "))&tblTasks[Task],m),3))),'
        'DAY(d)&IF(d=TODAY(),"   ·  TODAY","")&IF(n=0,"",CHAR(10)&t&IF(n>3,CHAR(10)&"+"&(n-3)&" more","")))')
    b.fmt(grid, fill=WHITE, wrap=True, v=xlTop, size=9, color=INK)
    b.heights(range(8, 14), 88)
    b.border(grid, color=LINE, edges=(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, 11, 12))
    dexpr = "CalStart+(ROW()-ROW($B$8))*7+COLUMN()-COLUMN($B$8)"
    b.cf_expr(grid, f"={dexpr}=TODAY()", fill=INDIGO_SOFT, bold=True, border=INDIGO)
    b.cf_expr(grid, f'=ISNUMBER(SEARCH("⚠",B8))', fill="FFF5F5")
    b.cf_expr(grid, f"=MONTH({dexpr})<>calMonthNum", fill="F7F8FB", color="B3B8C8")
    b.cf_expr(grid, "=COLUMN()>=COLUMN($G$8)", fill="FAFAFD")

    r = 16
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "🗓  Coming up - next 14 days"
    b.heights([r], 26)
    for i, h in enumerate(["DUE", "PRIORITY", "OWNER", "STATUS", "PROJECT", "TASK"]):
        b.fmt(f"{'BCDEFG'[i]}{r+1}", bold=True, size=9, color=MUTED, fill=WHITE).Value = h
    b.fmt(f"H{r+1}", fill=WHITE)
    b.r(f"B{r+2}").Formula2 = (
        '=LET(m,(tblTasks[Visible]=1)*(tblTasks[Status]<>DoneStatus)*ISNUMBER(tblTasks[Due Date])'
        '*(tblTasks[Due Date]>=TODAY())*(tblTasks[Due Date]<=TODAY()+14),'
        'IFERROR(SORTBY(FILTER(HSTACK(tblTasks[Due Date],tblTasks[Priority],tblTasks[Owner],tblTasks[Status],'
        'tblTasks[Project],tblTasks[Task]),m),FILTER(tblTasks[DueKey],m),1),"Nothing due in the next two weeks"))')
    b.fmt(f"B{r+2}:H{r+31}", fill=WHITE, v=xlCenter)
    b.fmt(f"B{r+2}:B{r+31}", nf="ddd d mmm", h=xlLeft, bold=True)
    b.fmt(f"G{r+2}:G{r+31}", bold=True, color=INK)
    b.heights(range(r + 1, r + 32), 20)
    b.border(f"B{r+1}:H{r+31}", color=LINE)
    for p, c in PRIO_COLORS.items():
        b.cf_expr(f"C{r+2}:C{r+31}", f'=C{r+2}="{p}"', color=c, bold=True)
    window(xl, ws, zoom=90)


# ---------------------------------------------------------------------
def build_scorecards(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    b.widths({"A": 2, "B": 24, "C": 10, "D": 13, "E": 10, "F": 10, "G": 12, "H": 12, "I": 10, "J": 13, "K": 8,
              "L": 12, "M": 3, **{c: 10 for c in "NOPQRSTU"}})
    banner(b, "U", "Scorecards", "Who's delivering, which projects are healthy  ·  score = 50% completion + 30% on-time + 20% not overdue")
    b.heights([4], 10)
    b.fmt("B6", bold=True, size=13, color=INK).Value = "🏆  Team scorecard"
    b.heights([6], 26)
    hdr = ["MEMBER", "ASSIGNED", "DONE", "OPEN", "OVERDUE", "COMPLETION", "ON-TIME", "SCORE", "RATING", "RANK", "BADGE"]
    for i, h in enumerate(hdr):
        b.fmt(f"{'BCDEFGHIJKL'[i]}7", bold=True, size=9, color=MUTED, fill=WHITE, v=xlCenter,
              h=xlLeft if i == 0 else xlCenter).Value = h
    vis = "tblTasks[Visible],1"
    f = {
        "B8": "=lstMembers",
        "C8": f"=COUNTIFS(tblTasks[Owner],B8#,{vis})",
        "D8": f"=COUNTIFS(tblTasks[Owner],B8#,{vis},tblTasks[Status],DoneStatus)",
        "E8": "=C8#-D8#",
        "F8": f'=COUNTIFS(tblTasks[Owner],B8#,{vis},tblTasks[State],"Overdue")',
        "G8": '=IF(C8#=0,"",D8#/C8#)',
        "H8": f'=IFERROR(SUMIFS(tblTasks[On Time],tblTasks[Owner],B8#,{vis})/COUNTIFS(tblTasks[Owner],B8#,{vis},tblTasks[On Time],">=0"),"")',
        "I8": ('=LET(a,C8#,comp,IFERROR(D8#/a,0),ot,IFERROR(H8#*1,comp),od,IFERROR(1-F8#/E8#,1),'
               'IF(a=0,"",ROUND(100*(0.5*comp+0.3*ot+0.2*od),0)))'),
        "J8": '=IF(I8#="","",REPT("★",ROUND(I8#/20,0))&REPT("☆",5-ROUND(I8#/20,0)))',
        "K8": '=IF(I8#="","",XMATCH(I8#,SORT(FILTER(I8#,I8#<>""),,-1)))',
        "L8": '=IF(K8#="","",IF(K8#=1,"🥇 MVP",IF(K8#=2,"🥈 2nd",IF(K8#=3,"🥉 3rd",IF(N(I8#)>=80,"🔥 Strong","")))))',
    }
    for a, fx in f.items():
        b.r(a).Formula2 = fx
    body = "B8:L22"
    b.fmt(body, fill=WHITE, v=xlCenter)
    b.fmt("C8:L22", h=xlCenter)
    b.fmt("B8:B22", bold=True, color=INK)
    b.fmt("G8:H22", nf="0%")
    b.fmt("I8:I22", bold=True, size=11, color=INK)
    b.fmt("J8:J22", color=AMBER, size=11)
    b.heights(range(7, 23), 22)
    b.border("B7:L22", color=LINE)
    b.border("B8:L22", color="EEF0F5", edges=(12,))
    b.databar("I8:I22", INDIGO, 100)
    cs = b.r("G8:G22").FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).FormatColor.Color = rgb("F8D7D8")
    cs.ColorScaleCriteria(2).FormatColor.Color = rgb("FDF0D5")
    cs.ColorScaleCriteria(3).FormatColor.Color = rgb("CDEFDD")
    b.cf_expr("F8:F22", "=N(F8)>0", color=RED, bold=True)
    b.cf_expr("B8:L22", "=$K8=1", fill="FFF9E6")

    ch = add_chart(ws, xlBarClustered, "N6", "U22")
    calc = wb.Worksheets("Calc")
    ch.SetSourceData(ws.Range("I8:I17"))
    ch.SeriesCollection(1).XValues = ws.Range("B8:B17")
    try:
        ch.SeriesCollection(1).XValues = "=Scorecards!$B$8#"
        ch.SeriesCollection(1).Values = "=Scorecards!$I$8#"
    except Exception:
        pass
    style_chart(ch, "Score leaderboard", legend=False)
    ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgb(INDIGO)
    ch.SeriesCollection(1).HasDataLabels = True
    ch.ChartGroups(1).GapWidth = 45
    ch.Axes(1).ReversePlotOrder = True
    try:
        ch.Axes(2).MinimumScale = 0
        ch.Axes(2).MaximumScale = 100
    except Exception:
        pass
    grid_lines_light(ch)

    r = 25
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "📁  Project scorecard"
    b.heights([r], 26)
    hdr = ["PROJECT", "CODE", "OWNER", "TASKS", "DONE", "PROGRESS", "OVERDUE", "DAYS LEFT", "TARGET", "HEALTH"]
    for i, h in enumerate(hdr):
        b.fmt(f"{'BCDEFGHIJK'[i]}{r+1}", bold=True, size=9, color=MUTED, fill=WHITE, v=xlCenter,
              h=xlLeft if i in (0, 2, 9) else xlCenter).Value = h
    b.fmt(f"L{r+1}", fill=WHITE)
    b.r(f"B{r+2}").Formula2 = ('=IFERROR(HSTACK(tblProjects[Project],tblProjects[Code],tblProjects[Owner],tblProjects[Tasks],'
                              'tblProjects[Done],tblProjects[Progress],tblProjects[Overdue],tblProjects[Days Left],'
                              'tblProjects[Target Date],tblProjects[Health]),"")')
    b.fmt(f"B{r+2}:L{r+21}", fill=WHITE, v=xlCenter)
    b.fmt(f"C{r+2}:J{r+21}", h=xlCenter)
    b.fmt(f"D{r+2}:D{r+21}", h=xlLeft)
    b.fmt(f"B{r+2}:B{r+21}", bold=True, color=INK)
    b.fmt(f"G{r+2}:G{r+21}", nf="0%")
    b.fmt(f"I{r+2}:I{r+21}", nf='0;[Red]-0;"today"')
    b.fmt(f"J{r+2}:J{r+21}", nf="d mmm yy")
    b.databar(f"G{r+2}:G{r+21}", GREEN, 1)
    b.cf_expr(f"K{r+2}:K{r+21}", f'=LEFT(K{r+2},1)="⛔"', color=RED, bold=True)
    b.cf_expr(f"K{r+2}:K{r+21}", f'=LEFT(K{r+2},1)="⚠"', color="B26A00", bold=True)
    b.cf_expr(f"K{r+2}:K{r+21}", f'=LEFT(K{r+2},1)="✅"', color=GREEN, bold=True)
    b.cf_expr(f"H{r+2}:H{r+21}", f"=N(H{r+2})>0", color=RED, bold=True)
    b.heights(range(r + 1, r + 22), 22)
    b.border(f"B{r+1}:L{r+21}", color=LINE)
    b.border(f"B{r+2}:L{r+21}", color="EEF0F5", edges=(12,))
    window(xl, ws, zoom=85)


# ---------------------------------------------------------------------
def build_home(xl, wb, ws):
    b = B(ws)
    page_bg(b)
    b.widths({"A": 3, **{c: 11 for c in "BCDEFGHIJKLM"}, "N": 3})
    b.fmt("A1:N6", fill=NAVY)
    b.heights([1], 14)
    b.heights([2], 44)
    b.heights([3], 22)
    b.heights([4], 22)
    b.heights([5], 14)
    b.heights([6], 14)
    b.fmt("B2", color=WHITE, bold=True, size=28, v=xlCenter).Value = "TeamBoard"
    b.fmt("B3", color="C9CDF5", size=12).Value = "Projects, tasks, boards and live team activity - all in one shared workbook."
    b.fmt("B4:M4", color="AEB4D0", size=10, merge=True).Formula2 = (
        '=TEXT(TODAY(),"dddd, d mmmm yyyy")&"     ·     "&COUNTA(tblTasks[Task])&" tasks     ·     "'
        '&COUNTA(tblProjects[Project])&" projects     ·     "&COUNTA(tblTeam[Name])&" teammates     ·     "'
        '&COUNTIFS(tblTasks[State],"Overdue")&" overdue"')
    tiles = [
        ("📊", "Dashboard", "KPIs, charts and health", INDIGO),
        ("🗂", "Board", "Kanban lanes by status", "4747C2"),
        ("💬", "Room", "Live feed & team chat", PINK),
        ("👤", "My Work", "Your personal to-do list", TEAL),
        ("📅", "Calendar", "Tasks by due date", "0B7F95"),
        ("🏆", "Scorecards", "Team & project scores", AMBER),
        ("✅", "Tasks", "Add & update tasks", GREEN),
        ("📁", "Projects", "Add & track projects", "23855A"),
    ]
    r0 = 8
    b.fmt(f"B{r0}", bold=True, size=13, color=INK).Value = "Where do you want to go?"
    b.heights([r0], 28)
    for i, (emo, name, desc, col) in enumerate(tiles):
        row = r0 + 1 + (i // 4) * 4
        c1 = "BEHK"[i % 4]
        c2 = chr(ord(c1) + 2)
        top = b.fmt(f"{c1}{row}:{c2}{row}", merge=True, fill=col, color=WHITE, bold=True, size=15, v=xlBottom, indent=1)
        b.link(f"{c1}{row}", f"'{name}'!A1", f"{emo}  {name}", color=WHITE, size=15, bold=True)
        b.fmt(f"{c1}{row+1}:{c2}{row+1}", merge=True, fill=col, color="F0F1FF", size=9.5, v=xlTop, indent=1).Value = desc
        b.border(f"{c1}{row}:{c2}{row+1}", color=BG, weight=xlThick)
        b.heights([row], 40)
        b.heights([row + 1], 26)
        b.heights([row + 2], 10)
    r = r0 + 10
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "How it works"
    b.heights([r], 28)
    steps = [
        ("1", "Add your team", "Team sheet: one row per person (name, email, role)."),
        ("2", "Create projects", "Projects sheet: add a row - it pops up in the Room."),
        ("3", "Add tasks", "Tasks sheet: type under the table, use the dropdowns."),
        ("4", "Move work along", "Change a task's Status - its card moves on the Board."),
    ]
    for i, (n, t, d) in enumerate(steps):
        c1 = "BEHK"[i]
        c2 = chr(ord(c1) + 2)
        b.fmt(f"{c1}{r+1}:{c2}{r+1}", merge=True, fill=WHITE, color=INDIGO, bold=True, size=11, indent=1, v=xlBottom).Value = f"{n}   {t}"
        b.fmt(f"{c1}{r+2}:{c2}{r+2}", merge=True, fill=WHITE, color=MUTED, size=9, wrap=True, indent=1, v=xlTop).Value = d
        b.border(f"{c1}{r+1}:{c2}{r+2}", color=LINE)
    b.heights([r + 1], 26)
    b.heights([r + 2], 34)
    r += 4
    b.fmt(f"B{r}", bold=True, size=13, color=INK).Value = "Legend"
    b.heights([r], 28)
    b.fmt(f"B{r+1}", fill=INDIGO, color=WHITE, bold=True, h=xlCenter).Value = "Header"
    b.fmt(f"C{r+1}:G{r+1}", merge=True, color=INK, indent=1).Value = "Blue column header = you type here (dropdowns help)."
    b.fmt(f"B{r+2}", fill="D9DCE6", color=INK, bold=True, h=xlCenter).Value = "Header"
    b.fmt(f"C{r+2}:G{r+2}", merge=True, color=INK, indent=1).Value = "Grey column header = calculated automatically - don't type."
    b.fmt(f"H{r+1}:M{r+1}", merge=True, color=INK, indent=1).Value = "Dashboard, Board, Room, My Work, Calendar = views only."
    b.fmt(f"H{r+2}:M{r+2}", merge=True, color=INK, indent=1).Value = "Sample data is included - overwrite or delete those rows when you start."
    b.heights([r + 1, r + 2], 22)
    r += 4
    b.heights([r], 30)
    b.fmt(f"B{r}:M{r}", merge=True, color=MUTED, size=9, italic=True, wrap=True, v=xlTop).Value = (
        "Share it: save to OneDrive or SharePoint → Share → give your team 'Can edit'. Everyone sees changes live in "
        "Excel for the web or the Excel app (Microsoft 365). No macros - works for free in the browser.")
    window(xl, ws, zoom=90)


# ---------------------------------------------------------------------
def slicer_style(wb):
    """Flat indigo slicer buttons that match the workbook design."""
    try:
        ts = wb.TableStyles.Add("TeamBoard Slicer")
        ts.ShowAsAvailableSlicerStyle = True
        ts.ShowAsAvailableTableStyle = False
        ts.ShowAsAvailablePivotTableStyle = False
        el = ts.TableStyleElements
        whole = el(0)
        whole.Interior.Color = rgb(WHITE)
        whole.Font.Color = rgb(INK)
        hdr = el(1)
        hdr.Font.Color = rgb(MUTED)
        hdr.Font.Bold = True
        spec = {30: (INDIGO, WHITE), 33: ("4747C2", WHITE), 28: ("EEF0F6", INK), 32: (INDIGO_SOFT, INK),
                29: ("F7F8FB", "B3B8C8"), 34: ("F7F8FB", "B3B8C8"), 31: ("C9CBEF", WHITE), 35: ("C9CBEF", WHITE)}
        for k, (fill, font) in spec.items():
            e = el(k)
            e.Interior.Color = rgb(fill)
            e.Font.Color = rgb(font)
            for edge in (7, 8, 9, 10):
                bb = e.Borders(edge)
                bb.LineStyle = xlContinuous
                bb.Color = rgb(WHITE)
        return "TeamBoard Slicer"
    except Exception as ex:
        print("custom slicer style failed, using built-in:", ex)
        return "SlicerStyleLight2"


def slicers(xl, wb, S):
    lo = S["Tasks"].ListObjects("tblTasks")
    dash, board = S["Dashboard"], S["Board"]

    style = slicer_style(wb)

    def place(ws, cache, name, caption, cell_tl, width, height, cols):
        tl = ws.Range(cell_tl)
        s = cache.Slicers.Add(ws)
        s.Name, s.Caption = name, caption
        s.Top, s.Left, s.Width, s.Height = tl.Top, tl.Left, width, height
        s.NumberOfColumns = cols
        s.RowHeight = 20
        s.Style = style
        return s

    caches = {}
    for field in ["Project", "Owner", "Priority", "Status"]:
        caches[field] = wb.SlicerCaches.Add2(lo, field)
        try:
            caches[field].CrossFilterType = 2  # show items with no data last / greyed
        except Exception:
            pass
    d6 = dash.Range("B6")
    unit = dash.Range("C6").Left - dash.Range("B6").Left
    place(dash, caches["Project"], "slDashProject", "Project", "B6", unit * 3.9, 100, 2)
    place(dash, caches["Owner"], "slDashOwner", "Owner", "F6", unit * 4.9, 100, 4)
    place(dash, caches["Priority"], "slDashPriority", "Priority", "K6", unit * 1.9, 100, 2)
    place(dash, caches["Status"], "slDashStatus", "Status", "M6", unit * 2.9, 100, 2)
    lane = board.Range("I6").Left - board.Range("F6").Left
    place(board, caches["Project"], "slBoardProject", "Project", "F6", lane * 0.97, 100, 2)
    place(board, caches["Owner"], "slBoardOwner", "Owner", "I6", lane * 1.97, 100, 4)
    place(board, caches["Priority"], "slBoardPriority", "Priority", "O6", lane * 0.9, 100, 2)


if __name__ == "__main__":
    build()
