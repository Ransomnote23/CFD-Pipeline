from http.server import BaseHTTPRequestHandler
import json, io
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── helpers ────────────────────────────────────────────────────────────────────
def today_str():
    return date.today().strftime("%B %d, %Y").replace(" 0"," ")

def parse_date(s):
    if not s: return None
    try: return datetime.strptime(str(s)[:10], "%Y-%m-%d").date()
    except: return None

def fmt_date(s):
    d = parse_date(s)
    return d.strftime("%Y-%m-%d") if d else "—"

def is_overdue(due_str):
    d = parse_date(due_str)
    return d is not None and d < date.today()

def line_shipped_qty(line):
    return sum(s.get("qty",0) for s in line.get("shipped",[]))

def line_open_qty(line):
    return line["qty"] - line_shipped_qty(line)

def line_status(line):
    shipped = line_shipped_qty(line)
    open_qty = line["qty"] - shipped
    if open_qty <= 0:
        has_inv = any(s.get("inv") for s in line.get("shipped",[]))
        return "INVOICED" if has_inv else "COMPLETE"
    return "PARTIAL" if shipped > 0 else "OPEN"

# ── styles ─────────────────────────────────────────────────────────────────────
FN = "Arial"
def fill(hex): return PatternFill("solid", fgColor=hex)
def font(bold=False, color="000000", size=9, italic=False):
    return Font(name=FN, bold=bold, color=color, size=size, italic=italic)

thin = Side(style="thin",   color="CCCCCC")
med  = Side(style="medium", color="888888")
thk  = Side(style="medium", color="1F4E79")

BORDER       = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_MED   = Border(left=med,  right=med,  top=med,  bottom=med)
BORDER_HEAVY = Border(left=thk,  right=thk,  top=thk,  bottom=thk)
BORDER_BOT   = Border(bottom=Side(style="medium", color="1F4E79"))

F_HDR   = fill("1F4E79")  # navy title bar
F_COL   = fill("2E75B6")  # blue col headers
F_PO    = fill("D6E4F0")  # light blue PO grouping row
F_OD    = fill("FFD7D7")  # red overdue
F_OD_PO = fill("FFBCBC")  # darker red overdue PO header
F_ALT   = fill("F0F6FF")  # alternating row
F_TOT   = fill("E2EFDA")  # green totals
F_PROJ  = fill("FFFDE7")  # yellow proj delivery col
F_WHITE = fill("FFFFFF")
F_LGND  = fill("F5F5F5")  # legend bg

al_c = Alignment(horizontal="center", vertical="center")
al_l = Alignment(horizontal="left",   vertical="center")
al_r = Alignment(horizontal="right",  vertical="center")
al_w = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def sc(ws, r, c, v, fnt=None, fil=None, aln=None, brd=None, nf=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fnt: cell.font      = fnt
    if fil: cell.fill      = fil
    if aln: cell.alignment = aln
    if brd: cell.border    = brd
    if nf:  cell.number_format = nf
    return cell

# ── summary sheet ──────────────────────────────────────────────────────────────
SUM_HDRS   = ["Customer PO#","WO#","Date Received","Due Date",
               "Total Lines","Open Lines","Partial Lines","Open Value ($)"]
SUM_WIDTHS = [20, 12, 16, 16, 13, 12, 14, 16]

def build_summary(ws, customer, po_meta):
    ws.freeze_panes = "A6"

    # Row 1 — title bar
    ws.merge_cells("A1:H1")
    sc(ws,1,1, f"Open Purchase Order Status Report  ·  As of {today_str()}",
       fnt=font(bold=True,color="FFFFFF",size=14),
       fil=F_HDR, aln=Alignment(horizontal="center",vertical="center"))
    ws.row_dimensions[1].height = 28

    # Row 2 — customer name
    ws.merge_cells("A2:H2")
    sc(ws,2,1, customer,
       fnt=font(bold=True,color="1F4E79",size=14),
       aln=al_l)
    ws.row_dimensions[2].height = 22

    # Row 3 — spacer
    ws.row_dimensions[3].height = 5

    # Row 4 — col headers
    for ci,h in enumerate(SUM_HDRS,1):
        sc(ws,4,ci,h, fnt=font(bold=True,color="FFFFFF",size=9),
           fil=F_COL, aln=al_c, brd=BORDER)
    ws.row_dimensions[4].height = 18

    total_open = total_pd = 0
    r = 5
    for i, po in enumerate(po_meta):
        od  = po["od"]
        fil = F_OD if od else (F_ALT if i%2==0 else F_WHITE)
        fc  = "BB0000" if od else "000000"
        vals = [po["cpo"], po["wo"],
                parse_date(po["rcvd"]), parse_date(po["due"]),
                po["total"], po["open_l"], po["part_l"], po["open_v"]]
        for ci, v in enumerate(vals, 1):
            nf=None; aln=al_c
            if ci==1: aln=al_l
            if ci==8: nf='#,##0.00'; aln=al_r
            if ci in (3,4): nf='YYYY-MM-DD'
            sc(ws,r,ci,v, fnt=font(color=fc,size=9),
               fil=fil, aln=aln, brd=BORDER, nf=nf)
        total_open += po["open_v"]
        if od: total_pd += po["open_v"]
        r += 1

    # Totals rows
    def tot_row(row, label, val, fil, fc):
        sc(ws,row,1,label, fnt=font(bold=True,color=fc,size=9), fil=fil, aln=al_l, brd=BORDER)
        for ci in range(2,8):
            sc(ws,row,ci,"", fil=fil, brd=BORDER)
        sc(ws,row,8,val, fnt=font(bold=True,color=fc,size=9),
           fil=fil, aln=al_r, brd=BORDER, nf='#,##0.00')

    tot_row(r,   "TOTAL",          total_open, F_TOT, "1F6B00"); r+=1
    tot_row(r,   "TOTAL PAST DUE", total_pd,   F_OD,  "BB0000"); r+=2

    # Legend
    sc(ws,r,1,"Status Legend:", fnt=font(bold=True,size=8,color="555555"), fil=F_LGND, aln=al_l)
    for ci in range(2,9): sc(ws,r,ci,"", fil=F_LGND)
    r+=1
    for code, desc, fc in [
        ("OPEN",     "Items not yet shipped",  "000000"),
        ("PARTIAL",  "Partially shipped",       "CC5500"),
        ("Red row",  "Due date has passed",     "BB0000"),
    ]:
        sc(ws,r,1,code, fnt=font(bold=True,size=8,color=fc)); 
        sc(ws,r,2,desc, fnt=font(size=8,color="555555"))
        r+=1

    for ci,w in enumerate(SUM_WIDTHS,1):
        ws.column_dimensions[get_column_letter(ci)].width = w


# ── detail sheet ───────────────────────────────────────────────────────────────
DET_HDRS = ["Customer PO#","WO#","Line","Description","Ordered Qty",
            "Shipped Qty","Open Qty","Status","Proj. Delivery Date ✎",
            "Labor Hours","Due Date","Date Received"]
DET_WIDTHS = [18,11,7,46,13,12,10,10,21,13,14,14]

STATUS_COLORS = {
    "OPEN":"CC0000","PARTIAL":"CC5500",
    "INVOICED":"1A7A1A","COMPLETE":"1A7A1A"
}

def build_detail(ws, customer, po_list, skus_map):
    ws.freeze_panes = "A6"

    # Title
    ws.merge_cells("A1:L1")
    sc(ws,1,1, f"{customer} — Open PO Line Item Detail",
       fnt=font(bold=True,color="FFFFFF",size=14),
       fil=F_HDR, aln=Alignment(horizontal="center",vertical="center"))
    ws.row_dimensions[1].height = 28

    # Subtitle
    ws.merge_cells("A2:L2")
    sc(ws,2,1,
       f"Report Date: {today_str()}    ·    "
       f"Labor Hours = BOM hrs/unit × Open Qty × 1.2 safety factor    ·    "
       f"'Proj. Delivery Date' column is blank — fill in by hand",
       fnt=font(size=8,italic=True,color="555555"), aln=al_w)
    ws.row_dimensions[2].height = 14

    ws.row_dimensions[3].height = 5

    # Column headers
    for ci,h in enumerate(DET_HDRS,1):
        f_use = F_PROJ if ci==9 else F_COL
        fc_use= "666666" if ci==9 else "FFFFFF"
        sc(ws,4,ci,h, fnt=font(bold=True,color=fc_use,size=9),
           fil=f_use, aln=al_c, brd=BORDER)
    ws.row_dimensions[4].height = 30

    r = 5
    for po in po_list:
        od = is_overdue(po.get("due",""))
        po_start_row = r  # remember where this PO block starts

        # PO header row — bold navy or dark crimson, white text
        ws.merge_cells(f"A{r}:L{r}")
        po_lbl = (f"PO# {po['cpo']}    ·    WO# {po.get('wo') or '—'}    ·    "
                  f"Due: {fmt_date(po.get('due'))}    ·    "
                  f"Received: {fmt_date(po.get('rcvd'))}")
        po_hdr_fill = fill("8B0000") if od else fill("1F3864")  # dark crimson : dark navy
        sc(ws,r,1, po_lbl,
           fnt=font(bold=True, color="FFFFFF", size=10),
           fil=po_hdr_fill,
           aln=al_l, brd=BORDER_MED)
        ws.row_dimensions[r].height = 18
        r += 1

        for li, line in enumerate(po.get("lines",[])):
            shq = line_shipped_qty(line)
            oq  = line_open_qty(line)
            st  = line_status(line)
            bom = skus_map.get(line["desc"],0)
            lh  = round(bom*oq*1.2,2) if (bom and oq>0) else None

            row_od = od and oq>0
            fil = F_OD if row_od else (F_ALT if li%2==0 else F_WHITE)

            vals=[po["cpo"], po.get("wo") or "—", line["id"], line["desc"],
                  line["qty"], shq, oq, st, None, lh,
                  parse_date(po.get("due")), parse_date(po.get("rcvd"))]

            for ci, v in enumerate(vals,1):
                nf=None; aln=al_c
                fnt_use = font(color=STATUS_COLORS.get(st,"000000") if ci==8 else
                               ("BB0000" if row_od and ci!=9 else "000000"), size=9)
                if ci==4:  aln=al_l
                if ci in (5,6,7): nf="0"; aln=al_r
                if ci==10: nf="0.00"; aln=al_r
                if ci in (11,12): nf="YYYY-MM-DD"; aln=al_c
                if ci==9:  aln=al_c; fnt_use=font(color="AAAAAA",size=8,italic=True)
                fil_use = F_PROJ if ci==9 else fil
                sc(ws,r,ci,v, fnt=fnt_use, fil=fil_use, aln=aln, brd=BORDER, nf=nf)

            ws.row_dimensions[r].height = 14
            r += 1

        po_end_row = r - 1  # last line row of this PO

        # ── Bold border box around entire PO block (header + all lines) ──────
        box_color = "8B0000" if od else "1F3864"
        s_thick = Side(style="medium", color=box_color)
        s_none  = Side(style=None)
        for row_i in range(po_start_row, po_end_row + 1):
            is_top    = (row_i == po_start_row)
            is_bottom = (row_i == po_end_row)
            for col_i in range(1, 13):
                is_left  = (col_i == 1)
                is_right = (col_i == 12)
                cell = ws.cell(row=row_i, column=col_i)
                # Build border: thick on box edges, keep existing thin between cells
                top_side    = s_thick if is_top    else cell.border.top
                bottom_side = s_thick if is_bottom else cell.border.bottom
                left_side   = s_thick if is_left   else cell.border.left
                right_side  = s_thick if is_right  else cell.border.right
                cell.border = Border(top=top_side, bottom=bottom_side,
                                     left=left_side, right=right_side)

        # Spacer between POs
        ws.merge_cells(f"A{r}:L{r}")
        sc(ws,r,1,"", fil=F_WHITE)
        ws.row_dimensions[r].height = 8
        r += 1

    for ci,w in enumerate(DET_WIDTHS,1):
        ws.column_dimensions[get_column_letter(ci)].width = w


# ── main builder ───────────────────────────────────────────────────────────────
def build_workbook(data, customer_filter=None):
    pos      = data["pos"]
    skus     = data.get("skus",[])
    skus_map = {s["desc"]: s.get("laborHours",0) for s in skus}

    from collections import defaultdict
    groups = defaultdict(list)
    for po in pos:
        if not any(line_open_qty(l)>0 for l in po.get("lines",[])):
            continue
        c = po.get("customer","Unknown")
        if customer_filter and c != customer_filter:
            continue
        groups[c].append(po)

    customers = sorted(groups.keys())
    if not customers:
        raise ValueError("No open POs found.")

    for c in customers:
        groups[c].sort(key=lambda p:(
            0 if is_overdue(p.get("due","")) else 1,
            p.get("due","9999") or "9999"
        ))

    wb = Workbook()
    wb.remove(wb.active)

    # Build metadata for all customers first
    all_meta = {}
    for cust in customers:
        po_list = groups[cust]
        po_meta = []
        for po in po_list:
            total=len(po.get("lines",[])); openl=partl=0; openv=0.0
            for l in po.get("lines",[]):
                oq=line_open_qty(l)
                if oq>0:
                    s=line_status(l)
                    if s=="PARTIAL": partl+=1
                    else: openl+=1
                    openv+=oq*l.get("unitPrice",0)
            po_meta.append(dict(
                cpo=po["cpo"], wo=po.get("wo") or "—",
                rcvd=po.get("rcvd",""), due=po.get("due",""),
                total=total, open_l=openl, part_l=partl, open_v=openv,
                od=is_overdue(po.get("due",""))
            ))
        all_meta[cust] = po_meta

    # Pass 1 — all Summary tabs
    for cust in customers:
        safe = cust[:24]
        ws_sum = wb.create_sheet(f"{safe[:27]} Sum")
        build_summary(ws_sum, cust, all_meta[cust])

    # Pass 2 — all Detail tabs
    for cust in customers:
        safe = cust[:24]
        ws_det = wb.create_sheet(f"{safe[:24]} Detail")
        build_detail(ws_det, cust, groups[cust], skus_map)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Labor productivity PDF ─────────────────────────────────────────────────────
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Table, TableStyle,
                                Spacer, KeepTogether)
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.widgets.markers import makeMarker

NAVY = colors.HexColor("#1F3864")
BLUE = colors.HexColor("#2E75B6")
STEEL = colors.HexColor("#9DB9D9")
AMBER = colors.HexColor("#C77700")
GREEN = colors.HexColor("#1A7A1A")
RED = colors.HexColor("#B00020")
GRID = colors.HexColor("#CCCCCC")
ALT = colors.HexColor("#F2F6FB")
MUTED = colors.HexColor("#666666")
SEL = colors.HexColor("#E2EAF4")

S = {
    "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=17, textColor=colors.white, leading=21),
    "h1s": ParagraphStyle("h1s", fontName="Helvetica", fontSize=9, textColor=colors.HexColor("#D6E4F0"), leading=12),
    "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=11, textColor=NAVY, leading=14, spaceBefore=10, spaceAfter=4),
    "body": ParagraphStyle("body", fontName="Helvetica", fontSize=8.5, leading=11),
    "small": ParagraphStyle("small", fontName="Helvetica", fontSize=7.5, leading=9.5, textColor=MUTED),
    "kl": ParagraphStyle("kl", fontName="Helvetica", fontSize=7.5, textColor=MUTED, leading=9),
    "kv": ParagraphStyle("kv", fontName="Helvetica-Bold", fontSize=16, textColor=NAVY, leading=19),
    "warn": ParagraphStyle("warn", fontName="Helvetica", fontSize=8, leading=10.5, textColor=AMBER),
}


def h(v):
    return "—" if v is None else f"{v:,.1f}"


def pct(v):
    return "—" if v is None else f"{round(v * 100)}%"


def money(v, dec=0):
    return "—" if v is None else f"${v:,.{dec}f}"


def eff_color(v):
    if v is None:
        return MUTED
    return GREEN if v >= 0.85 else AMBER if v >= 0.65 else RED


def delta(cur, prev, kind):
    if cur is None or prev is None:
        return '<font color="#888888">no prior month</font>'
    d = cur - prev
    col = "#1A7A1A" if d >= 0 else "#B00020"
    arrow = "+" if d >= 0 else "−"
    if kind == "pct":
        txt = f"{abs(round(d * 100))} pts"
    elif kind == "money":
        txt = money(abs(d))
    else:
        txt = h(abs(d)) + " h"
    return f'<font color="{col}"><b>{arrow}{txt}</b></font> vs prior month'


def base_table_style(header_rows=1):
    return [
        ("FONTNAME", (0, 0), (-1, header_rows - 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, header_rows - 1), BLUE),
        ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, GRID),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]


def trend_chart(series, width):
    """Bars: direct vs earned hours. Line: efficiency % on its own scale."""
    d = Drawing(width, 150)
    labels = [m["short"].split(" ")[0] + (" (MTD)" if m.get("mtd") else "") for m in series]
    direct = [m["direct"] or 0 for m in series]
    earned = [m["earned"] or 0 for m in series]
    top = max(direct + earned + [1])

    bc = VerticalBarChart()
    bc.x, bc.y, bc.width, bc.height = 36, 28, width - 80, 100
    bc.data = [direct, earned]
    bc.categoryAxis.categoryNames = labels
    bc.categoryAxis.labels.fontSize = 6.5
    bc.categoryAxis.labels.fontName = "Helvetica"
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = top * 1.1
    bc.valueAxis.labels.fontSize = 6.5
    bc.valueAxis.strokeColor = GRID
    bc.valueAxis.gridStrokeColor = colors.HexColor("#EEEEEE")
    bc.valueAxis.visibleGrid = True
    bc.categoryAxis.strokeColor = GRID
    bc.bars[0].fillColor = STEEL
    bc.bars[1].fillColor = NAVY
    bc.bars.strokeColor = None
    bc.groupSpacing = 6
    bc.barSpacing = 1
    d.add(bc)

    # efficiency line, right-hand 0–150% scale
    effs = [m["eff"] for m in series]
    pts = [(i, min(e, 1.5)) for i, e in enumerate(effs) if e is not None]
    if pts:
        lp = LinePlot()
        lp.x, lp.y, lp.width, lp.height = bc.x, bc.y, bc.width, bc.height
        n = len(series)
        step = 1.0
        lp.data = [[(i * step + 0.5, v) for i, v in pts]]
        lp.xValueAxis.valueMin = 0
        lp.xValueAxis.valueMax = n * step
        lp.xValueAxis.visible = False
        lp.yValueAxis.valueMin = 0
        lp.yValueAxis.valueMax = 1.5
        lp.yValueAxis.visible = False
        lp.lines[0].strokeColor = AMBER
        lp.lines[0].strokeWidth = 1.6
        lp.lines[0].symbol = makeMarker("FilledCircle")
        lp.lines[0].symbol.size = 3.2
        lp.lines[0].symbol.fillColor = AMBER
        lp.lines[0].symbol.strokeColor = AMBER
        d.add(lp)
        for tick in (0.5, 1.0, 1.5):
            y = bc.y + bc.height * tick / 1.5
            d.add(String(bc.x + bc.width + 4, y - 2, f"{int(tick * 100)}%", fontSize=6.5,
                         fillColor=AMBER, fontName="Helvetica"))

    # legend
    lx = bc.x
    for i, (lbl, col) in enumerate([("Direct labor hrs", STEEL), ("Earned hrs", NAVY), ("Efficiency (right axis)", AMBER)]):
        from reportlab.graphics.shapes import Rect
        d.add(Rect(lx, 140, 8, 6, fillColor=col, strokeColor=None))
        d.add(String(lx + 11, 140, lbl, fontSize=7, fontName="Helvetica", fillColor=MUTED))
        lx += 105
    return d


def build_labor_pdf(p):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=0.55 * inch, rightMargin=0.55 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.55 * inch,
                            title=f"Labor Productivity Report — {p['current']['label']}")
    W = letter[0] - 1.1 * inch
    cur, prv = p["current"], p.get("prior") or {}
    story = []

    # ── header band
    sub = " · ".join(x for x in [p.get("company"), p.get("division")] if x)
    meta = f"Generated {p.get('generated', '')}" + (f" · Prepared by {p['preparedBy']}" if p.get("preparedBy") else "")
    title = f"Labor Productivity Report — {cur['label']}" + (" (month to date)" if cur.get("mtd") else "")
    band = Table([[Paragraph(title, S["h1"])], [Paragraph(f"{sub}<br/>{meta}", S["h1s"])]], colWidths=[W])
    band.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY),
                              ("LEFTPADDING", (0, 0), (-1, -1), 12), ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                              ("TOPPADDING", (0, 0), (0, 0), 10), ("BOTTOMPADDING", (0, -1), (-1, -1), 10)]))
    story += [band, Spacer(1, 10)]

    # ── KPI tiles
    def tile(label, value, sub, color=NAVY):
        vs = ParagraphStyle("v", parent=S["kv"], textColor=color)
        return [Paragraph(label.upper(), S["kl"]), Paragraph(value, vs), Paragraph(sub, S["small"])]

    tiles = [
        tile("Direct labor hours", h(cur["direct"]), delta(cur["direct"], prv.get("direct"), "h")),
        tile("Earned hours", h(cur["earned"]), delta(cur["earned"], prv.get("earned"), "h")
             + (f"<br/>incl. {h(cur.get('earnedOutside'))} h outside clients" if cur.get("earnedOutside") else "")),
        tile("Efficiency", pct(cur["eff"]), delta(cur["eff"], prv.get("eff"), "pct"), eff_color(cur["eff"])),
        tile("3-month rolling eff.", pct(cur["roll3"]), "smooths shipment timing", eff_color(cur["roll3"])),
        tile("Billed $ / direct hr", money(cur["perHr"]), f"{money(cur['shipped'])} billed"),
    ]
    tw = W / 5
    kt = Table([[t for t in tiles]], colWidths=[tw] * 5)
    kt.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, GRID), ("INNERGRID", (0, 0), (-1, -1), 0.6, GRID),
                            ("BACKGROUND", (0, 0), (-1, -1), ALT), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                            ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                            ("LEFTPADDING", (0, 0), (-1, -1), 7)]))
    story += [kt]

    if p.get("noBom"):
        names = "; ".join(p["noBom"][:5]) + ("…" if len(p["noBom"]) > 5 else "")
        story += [Spacer(1, 5), Paragraph(
            f"<b>Data note:</b> {len(p['noBom'])} SKU(s) shipped this month have no BOM hours, so earned hours "
            f"and efficiency are understated: {names}", S["warn"])]

    # ── 12-month trend
    story += [Paragraph("12-month trend", S["h2"]), trend_chart(p["series"], W)]
    rows = [["Month", "Worked\nhrs", "Time\noff", "Direct\nhrs", "Earned:\ntracker", "Earned:\noutside", "Earned\ntotal",
             "Eff.", "3-mo\neff.", "Billed $", "$ per\ndirect hr"]]
    for m in p["series"]:
        rows.append([m["short"] + (" (MTD)" if m.get("mtd") else ""), h(m["worked"]), h(m["off"]), h(m["direct"]),
                     h(m.get("earnedTracked")), h(m.get("earnedOutside")) if m.get("earnedOutside") else "—", h(m["earned"]),
                     pct(m["eff"]), pct(m["roll3"]), money(m["shipped"]), money(m["perHr"])])
    t = Table(rows, colWidths=[W * 0.12] + [W * 0.088] * 10, repeatRows=1)
    st = base_table_style()
    for i, m in enumerate(p["series"], start=1):
        if i % 2 == 0:
            st.append(("BACKGROUND", (0, i), (-1, i), ALT))
        if m["mk"] == cur["mk"]:
            st += [("BACKGROUND", (0, i), (-1, i), SEL), ("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")]
        st.append(("TEXTCOLOR", (7, i), (7, i), eff_color(m["eff"])))
        st.append(("TEXTCOLOR", (8, i), (8, i), eff_color(m["roll3"])))
        st.append(("FONTNAME", (6, i), (6, i), "Helvetica-Bold"))
    t.setStyle(TableStyle(st))
    story += [Spacer(1, 4), t]

    # ── staffing
    staff = p.get("staff") or []
    srows = [["Employee", "Wk hrs", "Scheduled", "Holiday", "Time off", "Overtime", "Worked", "Direct %", "Direct hrs", "Source"]]
    tot = {k: 0 for k in ("sched", "hol", "off", "ot", "worked", "direct")}
    for s in staff:
        srows.append([Paragraph(f"<b>{s['name']}</b>" + (f" <font color='#888888' size='7'>{s['role']}</font>" if s.get("role") else ""), S["body"]),
                      h(s["weekly"]), h(s["sched"]), h(s["hol"]), h(s["off"]) if s["off"] else "—",
                      ("+" + h(s["ot"])) if s["ot"] else "—", h(s["worked"]), f"{s['pct']:g}%", h(s["direct"]),
                      "Payroll" if s["source"] == "actual" else "Calculated"])
        for k in tot:
            tot[k] += s[k]
    srows.append(["Total", "", h(tot["sched"]), h(tot["hol"]), h(tot["off"]), h(tot["ot"]), h(tot["worked"]), "", h(tot["direct"]), ""])
    stbl = Table(srows, colWidths=[W * 0.22, W * 0.07, W * 0.09, W * 0.08, W * 0.08, W * 0.09, W * 0.08, W * 0.08, W * 0.09, W * 0.12], repeatRows=1)
    sst = base_table_style() + [("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2EFDA")),
                                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                                ("ALIGN", (-1, 1), (-1, -1), "CENTER")]
    for i in range(1, len(srows) - 1):
        if i % 2 == 0:
            sst.append(("BACKGROUND", (0, i), (-1, i), ALT))
    stbl.setStyle(TableStyle(sst))
    story += [KeepTogether([Paragraph(f"Staffing — {cur['label']}", S["h2"]), stbl])]

    # ── time off + top SKUs side by side
    to = p.get("timeOff") or []
    torows = [["Employee", "Dates", "Hrs", "Type"]]
    for x in to:
        dates = x["start"][5:] if x["start"] == x["end"] else f"{x['start'][5:]} – {x['end'][5:]}"
        torows.append([x["name"], dates, h(x["hours"]), Paragraph(x["type"] + (f" · {x['notes']}" if x.get("notes") else ""), S["small"])])
    if len(torows) == 1:
        torows.append(["None logged", "", "", ""])
    half = (W - 12) / 2
    tot_t = Table(torows, colWidths=[half * 0.28, half * 0.27, half * 0.12, half * 0.33], repeatRows=1)
    tot_t.setStyle(TableStyle(base_table_style() + [("ALIGN", (1, 0), (1, -1), "LEFT"), ("ALIGN", (3, 0), (3, -1), "LEFT")]))

    skrows = [["Top SKUs by earned hours", "Qty", "Earned hrs", "Shipped $"]]
    for s in p.get("skus") or []:
        skrows.append([Paragraph(s["desc"], S["small"]), f"{s['qty']:g}", h(s["hrs"]), money(s["value"])])
    if len(skrows) == 1:
        skrows.append(["No shipments recorded", "", "", ""])
    sk_t = Table(skrows, colWidths=[half * 0.47, half * 0.1, half * 0.21, half * 0.22], repeatRows=1)
    sk_t.setStyle(TableStyle(base_table_style()))

    pair = Table([[Paragraph("Time off &amp; overtime", S["h2"]), Paragraph("What was built", S["h2"])],
                  [tot_t, sk_t]], colWidths=[half + 6, half + 6])
    pair.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (0, -1), 12), ("RIGHTPADDING", (1, 0), (1, -1), 0)]))
    story += [pair]

    # ── outside clients
    outside = p.get("outside") or []
    if outside:
        orows = [["Client (not tracked in pipeline app)", "Earned hrs", "Billed $", "Notes"]]
        for o in outside:
            orows.append([o["client"], h(o["hours"]), money(o["revenue"]) if o.get("revenue") else "—",
                          Paragraph(o.get("notes") or "", S["small"])])
        orows.append(["Total", h(sum(o["hours"] for o in outside)),
                      money(sum(o.get("revenue") or 0 for o in outside)) if any(o.get("revenue") for o in outside) else "—", ""])
        ot = Table(orows, colWidths=[W * 0.36, W * 0.13, W * 0.15, W * 0.36], repeatRows=1)
        ot.setStyle(TableStyle(base_table_style() + [
            ("ALIGN", (3, 0), (3, -1), "LEFT"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2EFDA")),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]))
        story += [KeepTogether([Paragraph(f"Outside clients — {cur['label']}", S["h2"]), ot])]

    # ── method
    story += [Spacer(1, 10), Paragraph(
        "<b>How these numbers are calculated.</b> Scheduled hours = each employee's weekly hours × working days "
        "(company holidays excluded). Worked hours = scheduled − time off + overtime, or payroll actuals where entered. "
        "Direct hours = worked × the employee's direct-labor share. Earned hours = quantity shipped in the month × BOM "
        "labor hours per unit, plus hours billed to outside clients not yet tracked in the pipeline app. Efficiency = earned ÷ direct hours. Because earned hours are recognized when product "
        "ships, single months swing with shipment timing; the 3-month rolling figure is the steadier trend. "
        "Green ≥ 85%, amber 65–84%, red &lt; 65%.", S["small"])]

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(MUTED)
        canvas.drawString(doc_.leftMargin, 0.35 * inch, f"{p.get('company') or ''} · Labor Productivity · {cur['label']}")
        canvas.drawRightString(letter[0] - doc_.rightMargin, 0.35 * inch, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return buf.getvalue()


# ── Vercel handler ─────────────────────────────────────────────────────────────
class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def do_POST(self):
        try:
            length  = int(self.headers.get("Content-Length",0))
            body    = self.rfile.read(length)
            payload = json.loads(body)

            if payload.get("type") == "labor":
                pdf = build_labor_pdf(payload)
                self.send_response(200)
                self._cors()
                self.send_header("Content-Type", "application/pdf")
                self.send_header("Content-Disposition", 'attachment; filename="Labor_Productivity_Report.pdf"')
                self.send_header("Content-Length", str(len(pdf)))
                self.end_headers()
                self.wfile.write(pdf)
                return

            cust    = payload.get("customer") or None
            data    = payload["data"]

            xlsx_bytes = build_workbook(data, cust)

            self.send_response(200)
            self._cors()
            self.send_header("Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition",
                f'attachment; filename="Open_PO_Report.xlsx"')
            self.send_header("Content-Length", str(len(xlsx_bytes)))
            self.end_headers()
            self.wfile.write(xlsx_bytes)

        except Exception as e:
            self.send_response(500)
            self._cors()
            self.send_header("Content-Type","application/json")
            self.end_headers()
            self.wfile.write(json.dumps({"error":str(e)}).encode())

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin","*")
        self.send_header("Access-Control-Allow-Methods","POST,OPTIONS")
        self.send_header("Access-Control-Allow-Headers","Content-Type")

    def log_message(self, *a): pass
