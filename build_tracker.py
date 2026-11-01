"""Builds Debt-Tracker.xlsx: a formula-driven credit card debt tracker
with a snowball payoff plan (minimums on every card, extra to the lowest balance).

Run:  python3 build_tracker.py
"""
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "Debt-Tracker.xlsx"
MAX_CARDS = 20          # card slots on the Card List / Dashboard
REG_ROWS = 240          # register rows per card (20 years of monthly entries)
SIM_MONTHS = 600        # snowball simulation length (50 years)

# ---- styles -------------------------------------------------------------
FONT = "Arial"
DARK = "2F4F4F"
ACCENT = "5B8A72"
LIGHT = "EEF4F0"
INPUT_FILL = PatternFill("solid", fgColor="FFF6CC")
HEAD_FILL = PatternFill("solid", fgColor=DARK)
SUB_FILL = PatternFill("solid", fgColor=LIGHT)
thin = Side(style="thin", color="C9D3CD")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

MONEY = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.00%;-0.00%;"-"'
DATE = "mm/dd/yyyy"
MONTHS = '0" mo";-0;"-"'


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = f(18, True, DARK)
    ws["A2"] = sub
    ws["A2"].font = f(10, italic=True, color="555555")


def header_row(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font = f(10, True, "FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[row].height = 32


def label(ws, ref, text):
    ws[ref] = text
    ws[ref].font = f(10, True, DARK)
    ws[ref].fill = SUB_FILL
    ws[ref].border = BOX


def value(ws, ref, v, fmt=None, is_input=False, bold=False):
    ws[ref] = v
    ws[ref].font = f(11 if bold else 10, bold, "0000FF" if is_input else "000000")
    ws[ref].border = BOX
    ws[ref].alignment = Alignment(horizontal="right")
    if fmt:
        ws[ref].number_format = fmt
    if is_input:
        ws[ref].fill = INPUT_FILL


# ---- card sheet ---------------------------------------------------------
R0 = 11                    # first register row
RN = R0 + REG_ROWS - 1     # last register row


def build_card(wb, sheet_name, info, entries):
    ws = wb.create_sheet(sheet_name)
    title(ws, "Card Details", "Yellow cells are yours to fill in. Everything else calculates.")

    # top section: inputs (left)
    rows = [
        ("Card Name", info.get("name"), None),
        ("Starting Balance", info.get("start"), MONEY),
        ("Credit Limit", info.get("limit"), MONEY),
        ("Interest Rate (APR)", info.get("apr"), PCT),
        ("Starting Minimum Payment", info.get("min"), MONEY),
    ]
    for i, (lab, v, fmt) in enumerate(rows):
        r = 3 + i
        label(ws, f"A{r}", lab)
        value(ws, f"B{r}", v, fmt, is_input=True)
    ws["B3"].alignment = Alignment(horizontal="left")
    ws["B4"].comment = Comment("The balance on the day you started tracking. "
                               "The register builds from here.", "Tracker")
    ws["B7"].comment = Comment("Used until you enter a minimum payment in the register below.",
                               "Tracker")

    # top section: results (right), all from the register
    dates = f"$A${R0}:$A${RN}"
    results = [
        ("Current Balance",
         f"=IF(COUNT({dates})=0,N(B4),INDEX($G${R0}:$G${RN},COUNT({dates})))", MONEY),
        ("Minimum Payment",
         f"=IF(E3<=0,0,MIN(E3,IFERROR(INDEX($E${R0}:$E${RN},"
         f"MATCH(9.99E+307,$E${R0}:$E${RN})),N(B7))))", MONEY),
        ("Available Credit", "=N(B5)-E3", MONEY),
        ("Utilization", "=IF(N(B5)>0,E3/B5,0)", PCT),
        ("Months to Pay Off (Minimum Only)",
         '=IF(E3<=0,0,IF(E4<=E3*N(B6)/12,"Never",ROUNDUP(NPER(N(B6)/12,-E4,E3),0)))', MONTHS),
        ("Payoff Date (Minimum Only)",
         '=IF(E3<=0,"Paid off",IF(ISNUMBER(E7),EDATE(TODAY(),E7),"Never"))', DATE),
    ]
    for i, (lab, frm, fmt) in enumerate(results):
        r = 3 + i
        label(ws, f"D{r}", lab)
        value(ws, f"E{r}", frm, fmt, bold=(i == 0))
    ws["E7"].comment = Comment('"Never" means the minimum payment does not cover the '
                               "monthly interest, so the balance will not go down.", "Tracker")

    # register
    ws[f"A{R0 - 2}"] = "Register"
    ws[f"A{R0 - 2}"].font = f(13, True, DARK)
    ws[f"C{R0 - 2}"] = ("Enter one row per statement or payment, top to bottom with no gaps. "
                        "Interest is estimated from the APR; type the real amount from "
                        "your statement over it any time.")
    ws[f"C{R0 - 2}"].font = f(9, italic=True, color="555555")
    header_row(ws, R0 - 1, ["Date", "Starting Balance", "Interest", "New Charges",
                            "Minimum Payment", "Payment Made", "Ending Balance", "Notes"])
    for r in range(R0, RN + 1):
        prev = "$B$4" if r == R0 else f"G{r - 1}"
        ws[f"A{r}"].number_format = DATE
        ws[f"B{r}"] = f'=IF(A{r}="","",N({prev}))'
        ws[f"C{r}"] = f'=IF(A{r}="","",ROUND(B{r}*N($B$6)/12,2))'
        ws[f"G{r}"] = f'=IF(A{r}="","",B{r}+N(C{r})+N(D{r})-N(F{r}))'
        for c in "ABCDEFGH":
            cell = ws[f"{c}{r}"]
            cell.font = f(10, color="0000FF" if c in "ADEFH" else "000000")
            cell.border = BOX
            if c in "BCDEFG":
                cell.number_format = MONEY
            if c in "ADEFH":
                cell.fill = INPUT_FILL
    for i, e in enumerate(entries):
        r = R0 + i
        ws[f"A{r}"] = e[0]
        ws[f"D{r}"] = e[1]
        ws[f"E{r}"] = e[2]
        ws[f"F{r}"] = e[3]
        ws[f"H{r}"] = e[4] if len(e) > 4 else None

    for c, w in zip("ABCDEFGH", [26, 16, 13, 14, 32, 14, 16, 30]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = f"A{R0}"
    return ws


# ---- workbook -----------------------------------------------------------
wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
howto = wb.create_sheet("How To Use")
cl = wb.create_sheet("Card List")

samples = [
    ("Visa", dict(name="Visa (example)", start=1310, limit=3000, apr=0.2499, min=35),
     [(date(2026, 9, 1), 0, 35, 100, "Example row: replace with your own"),
      (date(2026, 10, 1), 45.20, 35, 100),
      (date(2026, 11, 1), 0, 35, 0, "Statement posted, payment not made yet")]),
    ("Store Card", dict(name="Store Card (example)", start=520, limit=1000, apr=0.2999, min=25),
     [(date(2026, 9, 1), 0, 25, 25, "Example row: replace with your own"),
      (date(2026, 10, 1), 0, 25, 25)]),
    ("Mastercard", dict(name="Mastercard (example)", start=4850, limit=8000, apr=0.2199, min=120),
     [(date(2026, 9, 1), 0, 120, 150, "Example row: replace with your own"),
      (date(2026, 10, 1), 0, 120, 150)]),
]
for sheet_name, info, entries in samples:
    build_card(wb, sheet_name, info, entries)
build_card(wb, "Template", {}, [])
sim = wb.create_sheet("Snowball Plan")

# ---- Card List ----------------------------------------------------------
CL0, CLN = 5, 5 + MAX_CARDS - 1
title(cl, "Card List", "Type the exact tab name of each card sheet in column A. "
      "One card per row. Order does not matter, the Dashboard sorts them.")
header_row(cl, 4, ["Card Sheet Name", "Current Balance", "Minimum Payment", "APR",
                   "Credit Limit", "Months (Min Only)", "Card Name", "Status", "Sort Key"])
for i, r in enumerate(range(CL0, CLN + 1)):
    cl[f"A{r}"] = samples[i][0] if i < len(samples) else None
    ref = lambda cell: f'IFERROR(INDIRECT("\'"&$A{r}&"\'!{cell}"),"")'
    cl[f"B{r}"] = f'=IF($A{r}="","",{ref("E3")})'
    cl[f"C{r}"] = f'=IF($A{r}="","",{ref("E4")})'
    cl[f"D{r}"] = f'=IF($A{r}="","",{ref("B6")})'
    cl[f"E{r}"] = f'=IF($A{r}="","",{ref("B5")})'
    cl[f"F{r}"] = f'=IF($A{r}="","",{ref("E7")})'
    cl[f"G{r}"] = f'=IF($A{r}="","",IF({ref("B3")}="",$A{r},{ref("B3")}))'
    cl[f"H{r}"] = (f'=IF($A{r}="","",IF(ISERROR(INDIRECT("\'"&$A{r}&"\'!E3")),'
                   f'"Tab not found: check spelling","OK"))')
    cl[f"I{r}"] = (f'=IF(OR($A{r}="",NOT(ISNUMBER(B{r}))),"",'
                   f'IF(B{r}<=0.005,1000000000+ROW(),B{r}+ROW()/1000000))')
    for c, fmt in zip("ABCDEFGHI", [None, MONEY, MONEY, PCT, MONEY, MONTHS, None, None, "0.000000"]):
        cell = cl[f"{c}{r}"]
        cell.border = BOX
        cell.font = f(10, color="0000FF" if c == "A" else "000000")
        if c == "A":
            cell.fill = INPUT_FILL
        if fmt:
            cell.number_format = fmt
for c, w in zip("ABCDEFGHI", [24, 16, 16, 10, 14, 16, 24, 26, 12]):
    cl.column_dimensions[c].width = w
cl.column_dimensions["I"].hidden = True
cl.freeze_panes = "A5"

# ---- Dashboard ----------------------------------------------------------
title(dash, "Debt Payoff Dashboard",
      "Pay the minimum on every card. Put every extra dollar on the top card. "
      "When it is paid off, roll its payment to the next one.")

summary = [
    ("Total Debt", f"=SUM('Card List'!B{CL0}:B{CLN})", MONEY, False),
    ("Total Minimum Payments", f"=SUM('Card List'!C{CL0}:C{CLN})", MONEY, False),
    ("Extra Available Each Month", 100, MONEY, True),
    ("Total Monthly Payment", "=B4+B5", MONEY, False),
    ("Debt-Free In (Your Plan)", "='Snowball Plan'!B4", None, False),
    ("Debt-Free Date (Your Plan)",
     "=IF(ISNUMBER(B7),EDATE(TODAY(),B7),B7)", DATE, False),
    ("Estimated Interest (Your Plan)", "='Snowball Plan'!B5", MONEY, False),
    ("Debt-Free In (Minimums Only)",
     f"=IF(B3<=0,0,IF(COUNTIF('Card List'!F{CL0}:F{CLN},\"Never\")>0,\"Never\","
     f"MAX('Card List'!F{CL0}:F{CLN})))", None, False),
    ("Debt-Free Date (Minimums Only)",
     "=IF(ISNUMBER(B10),EDATE(TODAY(),B10),B10)", DATE, False),
]
for i, (lab, v, fmt, inp) in enumerate(summary):
    r = 3 + i
    label(dash, f"A{r}", lab)
    value(dash, f"B{r}", v, fmt, is_input=inp, bold=(r in (3, 7, 8)))
for r in (7, 10):
    dash[f"B{r}"].number_format = '0" months";-0;"0 months"'
dash["B3"].font = f(14, True, DARK)
dash["B5"].comment = Comment("Any money you can add on top of all the minimums. "
                             "It goes to the lowest balance first.", "Tracker")
dash["B10"].comment = Comment("Assumes each minimum payment stays the same. "
                              "Real minimums often shrink as the balance drops, "
                              "which makes this take longer.", "Tracker")
dash["D5"] = "← Change this number to see how much faster you can be debt free."
dash["D5"].font = f(9, italic=True, color=ACCENT)

D0 = 14
header_row(dash, D0 - 1, ["Order", "Card", "Current Balance", "Minimum Payment", "APR",
                          "Credit Limit", "Utilization", "Months (Min Only)",
                          "Paid Off In (Your Plan)", "Payoff Date (Your Plan)", "Focus"])
for k in range(1, MAX_CARDS + 1):
    r = D0 + k - 1
    # hidden helpers: M = sort key, N = position in Card List
    dash[f"M{r}"] = f"=IFERROR(SMALL('Card List'!$I${CL0}:$I${CLN},{k}),\"\")"
    dash[f"N{r}"] = f"=IF(M{r}=\"\",\"\",MATCH(M{r},'Card List'!$I${CL0}:$I${CLN},0))"
    pick = lambda c: f"=IF($N{r}=\"\",\"\",INDEX('Card List'!${c}${CL0}:${c}${CLN},$N{r}))"
    dash[f"A{r}"] = f'=IF(N{r}="","",{k})'
    dash[f"B{r}"] = pick("G")
    dash[f"C{r}"] = pick("B")
    dash[f"D{r}"] = pick("C")
    dash[f"E{r}"] = pick("D")
    dash[f"F{r}"] = pick("E")
    dash[f"G{r}"] = f'=IF(N{r}="","",IF(N(F{r})>0,C{r}/F{r},0))'
    dash[f"H{r}"] = f'=IF(N{r}="","",IF(C{r}<=0,"Paid off",{pick("F")[1:]}))'
    dash[f"I{r}"] = (f"=IF(N{r}=\"\",\"\",IF(C{r}<=0,\"Paid off\","
                     f"INDEX('Snowball Plan'!$E$7:${col(4 + MAX_CARDS)}$7,{k})))")
    dash[f"J{r}"] = f'=IF(N{r}="","",IF(ISNUMBER(I{r}),EDATE(TODAY(),I{r}),I{r}))'
    dash[f"K{r}"] = (f'=IF(N{r}="","",IF(C{r}<=0,"Done!",'
                     f'IF(COUNTIF($C${D0}:C{r},">0")=1,"Put extra here","Minimum only")))')
    for c, fmt in zip("ABCDEFGHIJK", ["0", None, MONEY, MONEY, PCT, MONEY, PCT,
                                      MONTHS, MONTHS, DATE, None]):
        cell = dash[f"{c}{r}"]
        cell.border = BOX
        cell.font = f(10)
        if fmt:
            cell.number_format = fmt
        if c in "AK":
            cell.alignment = Alignment(horizontal="center")
dash.conditional_formatting.add(
    f"A{D0}:K{D0 + MAX_CARDS - 1}",
    FormulaRule(formula=[f'$K{D0}="Put extra here"'],
                fill=PatternFill("solid", fgColor="D9EAD3"), font=Font(bold=True, color=DARK)))
dash.conditional_formatting.add(
    f"A{D0}:K{D0 + MAX_CARDS - 1}",
    FormulaRule(formula=[f'$K{D0}="Done!"'], font=Font(color="999999")))
for c, w in zip("ABCDEFGHIJK", [32, 24, 16, 16, 9, 14, 12, 14, 15, 15, 16]):
    dash.column_dimensions[c].width = w
dash.column_dimensions["M"].hidden = True
dash.column_dimensions["N"].hidden = True
dash.freeze_panes = "A3"

# ---- Snowball Plan (simulation) -----------------------------------------
title(sim, "Snowball Plan (calculations)",
      "No need to edit this sheet. It simulates paying the Dashboard plan month by month.")
S0 = 10
SN = S0 + SIM_MONTHS - 1
OWED = 5                         # first Owed column (E)
MINC = OWED + MAX_CARDS + 1      # first MinCap column
PAY = MINC + MAX_CARDS + 1       # first Pay column


def ocol(k): return col(OWED + k - 1)
def mcol(k): return col(MINC + k - 1)
def pcol(k): return col(PAY + k - 1)


label(sim, "A4", "Months to debt free")
sim["B4"] = (f'=IF(B6<=0.005,0,IF(COUNTIF($C${S0}:$C${SN},">0.01")>={SIM_MONTHS},'
             f'"50+ years",COUNTIF($C${S0}:$C${SN},">0.01")+1))')
label(sim, "A5", "Total interest")
sim["B5"] = f"=SUM($D${S0}:$D${SN})"
sim["B5"].number_format = MONEY
label(sim, "A6", "Starting total")
sim["B6"] = f"=SUM({ocol(1)}4:{ocol(MAX_CARDS)}4)"
sim["B6"].number_format = MONEY
label(sim, "A7", "Monthly budget")
sim["B7"] = "=Dashboard!B6"
sim["B7"].number_format = MONEY

sim[f"{col(OWED - 1)}3"] = "Card (Dashboard order)"
sim[f"{col(OWED - 1)}4"] = "Starting balance"
sim[f"{col(OWED - 1)}5"] = "Monthly rate"
sim[f"{col(OWED - 1)}6"] = "Minimum payment"
sim[f"{col(OWED - 1)}7"] = "Paid off in month"
for k in range(1, MAX_CARDS + 1):
    dr = D0 + k - 1
    c = ocol(k)
    sim[f"{c}3"] = f'=IF(Dashboard!B{dr}="","",Dashboard!B{dr})'
    sim[f"{c}4"] = f"=N(Dashboard!C{dr})"
    sim[f"{c}5"] = f"=N(Dashboard!E{dr})/12"
    sim[f"{c}6"] = f"=N(Dashboard!D{dr})"
    sim[f"{c}7"] = (f'=IF({c}4<=0.005,0,IF(SUMPRODUCT(--(({c}{S0}:{c}{SN}-'
                    f'{pcol(k)}{S0}:{pcol(k)}{SN})>0.01))>={SIM_MONTHS},"50+ years",'
                    f'SUMPRODUCT(--(({c}{S0}:{c}{SN}-{pcol(k)}{S0}:{pcol(k)}{SN})>0.01))+1))')
    sim[f"{c}4"].number_format = MONEY
    sim[f"{c}5"].number_format = "0.000%"
    sim[f"{c}6"].number_format = MONEY

header_row(sim, S0 - 1, ["Month", "Budget", "Remaining", "Interest"])
for k in range(1, MAX_CARDS + 1):
    for start, name in ((OWED, "Owed"), (MINC, "Min"), (PAY, "Paid")):
        cc = sim.cell(row=S0 - 1, column=start + k - 1, value=f"{name} {k}")
        cc.font = f(9, True, "FFFFFF")
        cc.fill = PatternFill("solid", fgColor=ACCENT)

for t in range(SIM_MONTHS):
    r = S0 + t
    sim[f"A{r}"] = t + 1
    sim[f"B{r}"] = "=$B$7"
    sim[f"C{r}"] = f"=SUM({ocol(1)}{r}:{ocol(MAX_CARDS)}{r})-SUM({pcol(1)}{r}:{pcol(MAX_CARDS)}{r})"
    prev_total = "$B$6" if t == 0 else f"C{r - 1}"
    sim[f"D{r}"] = f"=SUM({ocol(1)}{r}:{ocol(MAX_CARDS)}{r})-{prev_total}"
    for k in range(1, MAX_CARDS + 1):
        o, m, p = ocol(k), mcol(k), pcol(k)
        prev = f"{o}$4" if t == 0 else f"({o}{r - 1}-{p}{r - 1})"
        sim[f"{o}{r}"] = f"=IF({prev}<=0.005,0,{prev}*(1+{o}$5))"
        sim[f"{m}{r}"] = f"=MIN({o}{r},{o}$6)"
        paid_before = f"-SUM({pcol(1)}{r}:{pcol(k - 1)}{r})" if k > 1 else ""
        mins_after = (f"-SUM({mcol(k + 1)}{r}:{mcol(MAX_CARDS)}{r})"
                      if k < MAX_CARDS else "")
        sim[f"{p}{r}"] = f"=MIN({o}{r},MAX(0,$B{r}{paid_before}{mins_after}))"
sim.freeze_panes = sim[f"E{S0}"]
sim.column_dimensions["A"].width = 22
sim.column_dimensions["D"].width = 20

# ---- How To Use ---------------------------------------------------------
title(howto, "How To Use This Tracker", "Set it up once. Update it when a statement arrives.")
steps = [
    ("Set up a card", ""),
    ("1.", "Right-click the Template tab and choose Duplicate (Google Sheets) or "
           "Move or Copy > Create a copy (Excel)."),
    ("2.", "Rename the new tab to the card's name, for example Chase or Target."),
    ("3.", "Fill in the yellow cells at the top: name, starting balance, credit limit, "
           "APR, and minimum payment."),
    ("4.", "Go to the Card List tab and type the tab name exactly as it appears on the tab."),
    ("", ""),
    ("Each month", ""),
    ("1.", "Open the card's tab and add a row to the register: date, any new charges, "
           "the minimum payment on the statement, and what you paid."),
    ("2.", "Interest fills in automatically from the APR. If your statement shows a different "
           "amount, type the real number over it."),
    ("3.", "Check the Dashboard. The card marked Put extra here gets any extra money."),
    ("", ""),
    ("Your plan", ""),
    ("•", "Pay the minimum on every card."),
    ("•", "Put the Extra Available Each Month (Dashboard) on the card at the top of the list."),
    ("•", "When that card is paid off, add its whole payment to the next card. "
          "The Dashboard updates the order for you."),
    ("", ""),
    ("Good to know", ""),
    ("•", "Yellow cells are inputs. Everything else is a formula, so avoid typing over it "
          "(except the Interest column, which is meant to be adjustable)."),
    ("•", "Fill register rows top to bottom with no blank rows in between."),
    ("•", "Payoff estimates assume today's minimum payments and APRs stay the same and no new "
          "charges are added. Treat them as a guide, not a promise."),
    ("•", f"Room for {MAX_CARDS} cards and {REG_ROWS} register rows per card."),
    ("•", "To start fresh, delete the example tabs (Visa, Store Card, Mastercard) and clear "
          "their names from the Card List."),
]
for i, (a, b) in enumerate(steps):
    r = 4 + i
    howto[f"A{r}"] = a
    howto[f"B{r}"] = b
    if b == "" and a:
        howto[f"A{r}"].font = f(12, True, DARK)
    else:
        howto[f"A{r}"].font = f(10, True, ACCENT)
        howto[f"B{r}"].font = f(10)
        howto[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
howto.column_dimensions["A"].width = 16
howto.column_dimensions["B"].width = 100
label(howto, "D4", "Legend")
howto["D5"] = "Input"
howto["D5"].fill = INPUT_FILL
howto["D5"].font = f(10, color="0000FF")
howto["D6"] = "Formula"
howto["D6"].font = f(10)
howto.column_dimensions["D"].width = 12

# data validation on APR and money inputs
for name in [s[0] for s in samples] + ["Template"]:
    ws = wb[name]
    dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                        showErrorMessage=True, errorTitle="APR",
                        error="Enter the APR as a percent, for example 24.99%.")
    ws.add_data_validation(dv)
    dv.add("B6")

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
dash.sheet_properties.tabColor = ACCENT
wb.save(OUT)
print("saved", OUT)
