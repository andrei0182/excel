"""Patch v2 in place at XML level (keeps LibreOffice-cached formula results) and add a Cost Summary sheet."""
import re, shutil, zipfile, openpyxl
from pathlib import Path
from xml.sax.saxutils import escape

W = Path(__file__).parent
SRC, X, OUT = W / "v2.xlsx", W / "x", W / "Fiserv_Statement_parsed_v3.xlsx"
shutil.rmtree(X, ignore_errors=True)
zipfile.ZipFile(SRC).extractall(X)
val = openpyxl.load_workbook(SRC, data_only=True)


def cv(ref):  # cached value in v2, e.g. "CardType!D9"
    sh, c = ref.split("!")
    return val[sh][c].value


def rw(path, fn):
    p = X / path
    s = p.read_text(encoding="utf-8")
    new = fn(s)
    assert new != s, f"no change in {path}"
    p.write_text(new, encoding="utf-8")


def sub1(old, new):
    def f(s):
        assert s.count(old) == 1, (old, s.count(old))
        return s.replace(old, new)
    return f


# ---------------------------------------------------------------- 1. text fixes
fixes = [
    ("(all 1,683 dollar amounts on pages 2-16 match in order, plus the page 1 values)",
     "(every dollar amount on pages 2-16 matches in order, plus the page 1 values)"),
    ("Page 1 shows it truncated (4961); the full number is printed in the header of pages 2-16",
     "Page 1 and the page headers show it truncated (4961); the full number is visible as the first "
     "Location ID in Summary By Location (page 3)"),
    ("take the corporate number from the page headers (pages 2-16).",
     "take the corporate number from the first Location ID in Summary By Location on page 3 (the page headers "
     "visibly show only 4961; the full number there sits in the text layer)."),
    ("by $0.01 to $0.03 (printed rates", "by $0.01 to $0.04 (printed rates"),
]
for old, new in fixes:
    rw("xl/sharedStrings.xml", sub1(escape(old, {'"': "&quot;"}), escape(new, {'"': "&quot;"})))

# ---------------------------------------------------------------- 2. new sheet
S_TITLE, S_SECTION, S_LABEL, S_TEXT, S_NOTE = 1, 6, 2, 4, 8
S_INT, S_MONEY, S_PCT, S_HEAD, S_MONEY_B = 10, 11, 12, 27, 26

rows = {}  # row -> list of cell xml


def cell(ref, style, value=None, formula=None):
    r = int(re.sub(r"[A-Z]", "", ref))
    if formula is not None:
        f = f"<f>{escape(formula)}</f>"
        if isinstance(value, str):
            x = f'<c r="{ref}" s="{style}" t="str">{f}<v>{escape(value)}</v></c>'
        else:
            x = f'<c r="{ref}" s="{style}">{f}<v>{repr(float(value))}</v></c>'
    elif value is None:
        x = f'<c r="{ref}" s="{style}"/>'
    elif isinstance(value, str):
        x = f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(value)}</t></is></c>'
    else:
        x = f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    rows.setdefault(r, []).append(x)


def section(r, title):
    cell(f"A{r}", S_SECTION, title)
    for c in "BCDE":
        cell(f"{c}{r}", S_SECTION)


def line(r, label, formula, value, style=S_MONEY, note=None):
    cell(f"A{r}", S_LABEL, label)
    cell(f"B{r}", style, value, formula)
    if note:
        cell(f"C{r}", S_NOTE, note)


gross, refunds, net = cv("CardType!D9"), cv("CardType!F9"), cv("Summary!B16")
items, avg = cv("CardType!G9"), cv("Summary!B27")
fees = -cv("Summary!B20")
amex_vol, amex_fees = cv("CardType!H7"), -cv("FeeSummary!D8")
vol_x = net - amex_vol
fees_x = fees - amex_fees
passthru = -cv("FeeSummary!H6")
markup = -(cv("FeeSummary!H5") + cv("FeeSummary!H7"))

cell("A1", S_TITLE, "Cost summary: effective rate, pass-through and processor fees")
cell("A2", S_NOTE, "Fees are shown here as positive costs. Every value is a live formula on the Summary, "
                   "CardType, FeeSummary and FeeDetail sheets.")

section(4, "Volume")
line(5, "Gross sales", "CardType!D9", gross)
line(6, "Refunds", "CardType!F9", refunds)
line(7, "Amount submitted (net)", "Summary!B16", net)
line(8, "Transactions (sales + refunds)", "CardType!G9", items, S_INT)
line(9, "Average ticket", "Summary!B27", avg)
line(10, "Refund rate (refunds / gross sales)", "-CardType!F9/CardType!D9", -refunds / gross, S_PCT)

section(12, "Cost")
line(13, "Total fees", "-Summary!B20", fees)
line(14, "Effective rate (total fees / amount submitted)", "B13/B7", fees / net, S_PCT)
line(15, "Amount submitted excluding AMEX", "B7-CardType!H7", vol_x,
     note="AMEX is funded and priced directly by American Express (Paid by Others)")
line(16, "Fees excluding AMEX", "B13+FeeSummary!D8", fees_x,
     note="Only $170.00 of AMEX-related processor fees appear on this statement")
line(17, "Effective rate excluding AMEX", "B16/B15", fees_x / vol_x, S_PCT)
line(18, "Pass-through: interchange + card-brand assessments (IC/PF)", "-FeeSummary!H6", passthru,
     note="Includes $9,965.45 of assessments/network fees and $3,863.23 of debit interchange (see FeeSummary)")
line(19, "Pass-through share of total fees", "B18/B13", passthru / fees, S_PCT)
line(20, "Processor fees + service charges (markup)", "-(FeeSummary!H5+FeeSummary!H7)", markup)
line(21, "Markup share of total fees", "B20/B13", markup / fees, S_PCT)
line(22, "Markup as % of volume excluding AMEX", "B20/B15", markup / vol_x, S_PCT)

section(24, "Effective rate by card brand")
for c, h in zip("ABCDE", ["Card brand", "Fees", "Net volume", "Effective rate", "Note"]):
    cell(f"{c}25", S_HEAD, h)
brands = [
    ("Visa", "B8", "H5", None),
    ("Mastercard", "C8", "H4", None),
    ("Discover", "E8", "H6", None),
    ("Debit (PIN networks)", "F8", "H8", "Visa / Mastercard / Discover debit cards sit in their brand rows"),
    ("American Express", "D8", "H7", "AMEX discount is billed by American Express, not on this statement"),
]
r = 26
for name, fcol, vcol, note in brands:
    f, v = -cv(f"FeeSummary!{fcol}"), cv(f"CardType!{vcol}")
    cell(f"A{r}", S_TEXT, name)
    cell(f"B{r}", S_MONEY, f, f"-FeeSummary!{fcol}")
    cell(f"C{r}", S_MONEY, v, f"CardType!{vcol}")
    cell(f"D{r}", S_PCT, f / v, f"B{r}/C{r}")
    if note:
        cell(f"E{r}", S_NOTE, note)
    r += 1
others = -cv("FeeSummary!G8")
cell(f"A{r}", S_TEXT, "Others (account-level, no brand)")
cell(f"B{r}", S_MONEY, others, "-FeeSummary!G8")
cell(f"E{r}", S_NOTE, "e.g. ANNUAL COMPLIANCE SVC FEE; counted in the total rate, not in any brand rate")
r += 1
cell(f"A{r}", S_LABEL, "Total")
cell(f"B{r}", S_MONEY_B, fees, "SUM(B26:B31)")
cell(f"C{r}", S_MONEY_B, net, "SUM(C26:C30)")
cell(f"D{r}", S_PCT, fees / net, f"B{r}/C{r}")
assert r == 32
brand_sum = sum(-cv(f"FeeSummary!{c}8") for c in "BCDEFG")
assert abs(brand_sum - fees) < 0.005, brand_sum

section(34, "Largest individual fee lines (excluding IC/PF)")
for c, h in zip("ABCDE", ["Description (FeeDetail)", "Amount", "Product", "Type", "FeeDetail row"]):
    cell(f"{c}35", S_HEAD, h)
fd = val["FeeDetail"]
cand = [row[0].row for row in fd.iter_rows(min_row=4, max_row=360) if row[4].value != "Interchange Charges"]
top = sorted(cand, key=lambda i: fd[f"F{i}"].value)[:10]
for k, i in enumerate(top):
    r = 36 + k
    cell(f"A{r}", S_TEXT, fd[f"D{i}"].value, f"FeeDetail!D{i}")
    cell(f"B{r}", S_MONEY, -fd[f"F{i}"].value, f"-FeeDetail!F{i}")
    cell(f"C{r}", S_TEXT, fd[f"C{i}"].value, f"FeeDetail!C{i}")
    cell(f"D{r}", S_TEXT, fd[f"E{i}"].value, f"FeeDetail!E{i}")
    cell(f"E{r}", S_INT, i)
last = 45

col = lambda letter: (ord(letter) - 64)
body = "".join(f'<row r="{n}">{"".join(sorted(rows[n], key=lambda x: col(re.match(r"<c r=.([A-Z])", x)[1])))}</row>'
               for n in sorted(rows))
sheet = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    f'<dimension ref="A1:E{last}"/>'
    '<sheetViews><sheetView showGridLines="false" workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    '<cols><col min="1" max="1" width="62" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/>'
    '<col min="3" max="3" width="16" customWidth="1"/><col min="4" max="4" width="16" customWidth="1"/>'
    '<col min="5" max="5" width="16" customWidth="1"/></cols>'
    f'<sheetData>{body}</sheetData>'
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    '</worksheet>'
)
(X / "xl/worksheets/sheet15.xml").write_text(sheet, encoding="utf-8")

# register: after Summary (index 1) -> new index 2; shift localSheetId of later sheets
rw("xl/workbook.xml", sub1('<sheet name="Summary" sheetId="2" state="visible" r:id="rId4"/>',
                           '<sheet name="Summary" sheetId="2" state="visible" r:id="rId4"/>'
                           '<sheet name="Cost Summary" sheetId="15" state="visible" r:id="rId18"/>'))
rw("xl/workbook.xml", lambda s: re.sub(r'localSheetId="(\d+)"',
                                       lambda m: f'localSheetId="{int(m[1]) + 1 if int(m[1]) >= 2 else m[1]}"', s))
rw("xl/workbook.xml", lambda s: s.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1))
rw("xl/_rels/workbook.xml.rels", sub1(
    "</Relationships>",
    '<Relationship Id="rId18" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet15.xml"/></Relationships>'))
rw("[Content_Types].xml", sub1(
    '<Override PartName="/xl/worksheets/sheet1.xml"',
    '<Override PartName="/xl/worksheets/sheet15.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml"'))
app = X / "docProps/app.xml"
if "Summary" in app.read_text(encoding="utf-8"):
    print("NOTE: app.xml lists sheet names")

# ---------------------------------------------------------------- 3. README: add row after "Summary"
def readme(s):
    m = re.search(r'<row r="11"[^>]*>.*?</row>', s, re.S)
    assert m and 'r="A11"' in m[0]
    def shift(mm):
        n = int(mm[2])
        return f'{mm[1]}{n + 1 if n >= 12 else n}'
    head, tail = s[:m.end()], s[m.end():]
    tail = re.sub(r'(<row r=")(\d+)', shift, tail)
    tail = re.sub(r'(<c r="[A-Z]+)(\d+)', shift, tail)
    a_style = re.search(r'<c r="A11" s="(\d+)"', m[0])[1]
    b_style = re.search(r'<c r="B11" s="(\d+)"', m[0])[1]
    new_row = (f'<row r="12"><c r="A12" s="{a_style}" t="inlineStr"><is><t>Cost Summary</t></is></c>'
               f'<c r="B12" s="{b_style}" t="inlineStr"><is><t xml:space="preserve">Effective rate (total and '
               'excluding AMEX), pass-through vs processor fees, effective rate by card brand and the 10 largest '
               'fee lines, all as live formulas.</t></is></c></row>')
    out = head + new_row + tail
    return out.replace('<dimension ref="A1:B23"/>', '<dimension ref="A1:B24"/>')
rw("xl/worksheets/sheet1.xml", readme)

# ---------------------------------------------------------------- zip (mimetype order: content types first)
OUT.unlink(missing_ok=True)
with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    names = zin.namelist()
    names.insert(names.index("xl/worksheets/sheet14.xml") + 1, "xl/worksheets/sheet15.xml")
    names.remove("[Content_Types].xml")
    for n in ["[Content_Types].xml", *names]:
        zout.write(X / n, n)
print("written", OUT)
