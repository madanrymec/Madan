"""
aop_quarterly_report.py
=======================
Quarter-wise AOP report (ORDER BOOK and BILLING) in the layout of AOP_REPORT.xlsx

    Table 1 : Sales Office  x  Product group  x  (months, Total-Qn ... , Total)  + AOP + Achieved
    Table 2 : Unit / Product x (months, Total-Qn ... , Total)                      + AOP + Achieved

Rules
  * ALL data of the financial year is used - no product / unit / sales office / region filter.
  * Only the quarters that exist in the data are shown (a quarter is "available" when at least one row
    of that fiscal year falls in one of its months); every available quarter shows its 3 months.
  * Sales Office comes from the `Sales_office` column of the actual table and of the target table.
  * AOP = SUM(Target) of the respective sales office (Table 1) / respective product (Table 2) from
        order_target   (order book report)
        billing_target (billing report)
    Targets are already in Lacs, actuals are Net_Value / 100000  ->  "All Values in Rs Lacs".
  * Product grouping is exactly the one in the sample sheet (verified: Table 1 group totals and Table 2
    product totals of the sample reconcile to the paisa; Switchgear is part of "Transformer (Mysore)").
  * Values that cannot be placed in one of the known products are NOT dropped: they appear in an extra
    "Others" column / row, so the grand totals always equal the real totals of the table.

Routes (add `app.include_router(aop_router)` in app.py, after the other routers)
    GET /export_aop_quarterly_order_excel?fin_year=2026-2027
    GET /export_aop_quarterly_billing_excel?fin_year=2026-2027
    GET /export_aop_quarterly_consolidated_excel?fin_year=2026-2027      (both reports, 2 sheets)
`fin_year` is optional: the latest fin_year found in the actual table is used.
"""
import io
import re
from collections import defaultdict

import pymysql
from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter

from db_config import get_db_connection
from fastapi_compat import send_file

aop_router = APIRouter()

XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

# ----------------------------------------------------------------------------------------------
# Business rules (change here only)
# ----------------------------------------------------------------------------------------------
# Same net-sales definition as every billing dashboard in the project:
#   Inter Unit Transfer is not a sale, Unit UN07 is scrap sales which the dashboards report separately.
BILLING_WHERE_EXTRA = "AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'"
# The order dashboards do not exclude anything at overall level -> nothing excluded here.
ORDER_WHERE_EXTRA = ""

FISCAL_MONTHS = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]
QUARTER_MONTHS = {1: [4, 5, 6], 2: [7, 8, 9], 3: [10, 11, 12], 4: [1, 2, 3]}
MONTH_ABBR = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
              9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'}


def month_quarter(month):
    return 1 if month in (4, 5, 6) else 2 if month in (7, 8, 9) else 3 if month in (10, 11, 12) else 4


# Table 1 columns (order of the sample sheet) -> products they contain
GROUPS = [
    ('LV Motors (Non-Std)', ['LV Motors (Non-Std)']),
    ('LV Motors (Std)', ['LV Motors (Std)']),
    ('LV Generators', ['LV Generators']),
    ('EVM', ['EVM']),
    ('HVM,DCM,RRM,HVG', ['HV Motors', 'DC Machines', 'RRM', 'HV Generators']),
    ('Transformer (Pune)', ['Transformer Pune']),
    ('DG Sets', ['DG Sets']),
    ('Transformer (Mysore)', ['Transformer Mysore', 'Switchgear']),
    ('SPARES & SERVICE', ['SPARES & SERVICE']),
]
# Table 2 rows (order of the sample sheet)
UNITS = [
    ('Unit-01', ['DC Machines', 'HV Generators', 'HV Motors', 'RRM']),
    ('Unit-02', ['LV Motors (Non-Std)']),
    ('Unit-05', ['Switchgear', 'Transformer Mysore']),
    ('Unit-06', ['DG Sets']),
    ('Unit-16', ['Transformer Pune']),
    ('Unit-20', ['SPARES & SERVICE']),
    ('Unit-25', ['EVM', 'LV Generators', 'LV Motors (Std)']),
]
OTHERS = 'Others'
GROUP_OF = {p: g for g, prods in GROUPS for p in prods}


def _norm(text):
    """'LV MOTORS(S)' / 'LV Motors (S)' -> 'lvmotorss' : case, spaces and punctuation do not matter."""
    return re.sub(r'[^a-z0-9]', '', (text or '').lower())


_COMMON_ALIASES = {
    'dcmachines': 'DC Machines', 'hvgenerators': 'HV Generators', 'hvmotors': 'HV Motors', 'rrm': 'RRM',
    'evm': 'EVM', 'lvgenerators': 'LV Generators', 'dgsets': 'DG Sets',
    'switchgear': 'Switchgear', 'switchgearmysore': 'Switchgear',
    'transformermysore': 'Transformer Mysore', 'transformersmysore': 'Transformer Mysore',
    'transformerpune': 'Transformer Pune', 'transformerspune': 'Transformer Pune',
    'sparesservice': 'SPARES & SERVICE',
    'lvmotorsns': 'LV Motors (Non-Std)', 'lvmotorsnonstd': 'LV Motors (Non-Std)',
    'lvmotorss': 'LV Motors (Std)', 'lvmotorsstd': 'LV Motors (Std)', 'lvmotorsstandard': 'LV Motors (Std)',
}
# order tables name the NON-standard LV motors plainly 'LV Motors' (the standard ones 'LV MOTORS(S)')
ORDER_ALIASES = dict(_COMMON_ALIASES, lvmotors='LV Motors (Non-Std)')
# billing: plain 'LV MOTORS' (a plant other than AC02 / AC25) cannot be classified -> goes to Others
BILLING_ALIASES = dict(_COMMON_ALIASES)

# ----------------------------------------------------------------------------------------------
# SQL  (one grouped query per table: a few thousand rows come back even for 50 lakh+ source rows)
# ----------------------------------------------------------------------------------------------
ORDER_SPEC = {
    'key': 'order',
    'sheet': 'Order Book',
    'title': 'ORDER BOOK',
    'aliases': ORDER_ALIASES,
    'latest_fy_sql': "SELECT fin_year FROM order_data WHERE fin_year IS NOT NULL "
                     "GROUP BY fin_year ORDER BY fin_year DESC LIMIT 1",
    'actual_sql': f"""
        SELECT Sales_office AS office,
               CASE WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                    ELSE Product1 END AS prod_key,
               MONTH(Created_Date) AS month_no,
               SUM(Net_Value) / 100000 AS val
        FROM order_data
        WHERE fin_year = %s {ORDER_WHERE_EXTRA}
        GROUP BY office, prod_key, month_no
    """,
    'target_sql': """
        SELECT Sales_office AS office, Product1 AS prod_key, SUM(`Target`) AS val
        FROM order_target
        WHERE fin_year = %s
        GROUP BY office, prod_key
    """,
}
BILLING_SPEC = {
    'key': 'billing',
    'sheet': 'Billing',
    'title': 'BILLING',
    'aliases': BILLING_ALIASES,
    'latest_fy_sql': "SELECT fin_year FROM billing_data WHERE fin_year IS NOT NULL "
                     "GROUP BY fin_year ORDER BY fin_year DESC LIMIT 1",
    'actual_sql': f"""
        SELECT Sales_office AS office,
               CASE WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                    WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                    ELSE Prod END AS prod_key,
               MONTH(Billing_Date) AS month_no,
               SUM(Net_Value) / 100000 AS val
        FROM billing_data
        WHERE fin_year = %s {BILLING_WHERE_EXTRA}
        GROUP BY office, prod_key, month_no
    """,
    'target_sql': """
        SELECT Sales_office AS office, Product AS prod_key, SUM(`Target`) AS val
        FROM billing_target
        WHERE fin_year = %s
        GROUP BY office, prod_key
    """,
}


# ----------------------------------------------------------------------------------------------
# data layer
# ----------------------------------------------------------------------------------------------
def _office_key(name):
    return re.sub(r'\s+', ' ', (name or '').strip()).lower()


def _office_label(name):
    label = re.sub(r'\s+', ' ', (name or '').strip())
    return label if label else '(No Sales Office)'


def parse_fin_year(fin_year):
    """'2026-2027' / '2026-27' / '2026' -> 2026 (calendar year in which April of the fin year falls)"""
    m = re.match(r'\s*(\d{4})', str(fin_year or ''))
    if not m:
        raise ValueError(f"Invalid fin_year '{fin_year}' (expected like 2026-2027)")
    return int(m.group(1))


def month_label(month, start_year):
    year = start_year if month >= 4 else start_year + 1
    return f"{MONTH_ABBR[month]}'{year % 100:02d}"


def period_label(quarters, start_year):
    first = QUARTER_MONTHS[quarters[0]][0]
    last = QUARTER_MONTHS[quarters[-1]][-1]
    y1 = start_year if first >= 4 else start_year + 1
    y2 = start_year if last >= 4 else start_year + 1
    if y1 == y2:
        return f"{MONTH_ABBR[first]}-{MONTH_ABBR[last]}'{y2 % 100:02d}"
    return f"{MONTH_ABBR[first]}'{y1 % 100:02d}-{MONTH_ABBR[last]}'{y2 % 100:02d}"


def latest_fin_year(cursor, spec):
    cursor.execute(spec['latest_fy_sql'])
    row = cursor.fetchone()
    return row['fin_year'] if row else None


def load_rows(cursor, spec, fin_year):
    cursor.execute(spec['actual_sql'], (fin_year,))
    actual = cursor.fetchall()
    cursor.execute(spec['target_sql'], (fin_year,))
    target = cursor.fetchall()
    return actual, target


def build_report_data(spec, fin_year, actual_rows, target_rows):
    """
    Pure function (no DB, no Excel): turns the grouped rows into everything the sheet needs.
    Returns None when the fin_year has no actual data at all.
    """
    aliases = spec['aliases']

    def canonical(prod_key):
        return aliases.get(_norm(prod_key), OTHERS)

    offices = {}                                            # office_key -> label (first spelling seen)
    office_group = defaultdict(lambda: defaultdict(float))  # office_key -> (group, month) -> value
    product_month = defaultdict(lambda: defaultdict(float))  # product -> month -> value
    months_present = set()
    undated_total, undated_rows = 0.0, 0

    for r in actual_rows:
        val = float(r['val'] or 0)
        okey = _office_key(r['office'])
        offices.setdefault(okey, _office_label(r['office']))
        prod = canonical(r['prod_key'])
        month = r['month_no']
        if month is None:
            undated_total += val
            undated_rows += 1
            continue
        month = int(month)
        months_present.add(month)
        group = GROUP_OF.get(prod, OTHERS)
        office_group[okey][(group, month)] += val
        product_month[prod][month] += val

    if not months_present:
        return None

    office_aop = defaultdict(float)
    product_aop = defaultdict(float)
    for r in target_rows:
        val = float(r['val'] or 0)
        okey = _office_key(r['office'])
        offices.setdefault(okey, _office_label(r['office']))
        office_aop[okey] += val
        product_aop[canonical(r['prod_key'])] += val

    quarters = sorted({month_quarter(m) for m in months_present})
    start_year = parse_fin_year(fin_year)

    has_others_actual = any(g == OTHERS for per in office_group.values() for (g, _m) in per)
    has_others = has_others_actual or product_aop.get(OTHERS, 0) != 0
    groups = [g for g, _ in GROUPS] + ([OTHERS] if has_others else [])

    unit_rows = [(u, list(p)) for u, p in UNITS]
    if has_others:
        unit_rows.append((OTHERS, [OTHERS]))

    ordered = sorted(offices, key=lambda k: offices[k].lower())
    return {
        'spec': spec, 'fin_year': fin_year, 'start_year': start_year,
        'quarters': quarters, 'period': period_label(quarters, start_year),
        'offices': [(k, offices[k]) for k in ordered],
        'groups': groups, 'unit_rows': unit_rows,
        'office_group': office_group, 'product_month': product_month,
        'office_aop': office_aop, 'product_aop': product_aop,
        'undated_total': undated_total, 'undated_rows': undated_rows,
    }


# ----------------------------------------------------------------------------------------------
# Excel layer (same look as AOP_REPORT.xlsx)
# ----------------------------------------------------------------------------------------------
_THIN = Side(style='thin')
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
FILL_Q = PatternFill('solid', fgColor='D9E1F2')       # "Total-Qn"  (light blue in the sample)
FILL_TOTAL = PatternFill('solid', fgColor='E2EFDA')   # "Total"     (light green in the sample)
CENTER = Alignment(horizontal='center', vertical='center')
FMT_MONTH = '0.00'
FMT_INT = '0'


def _put(ws, row, col, value=None, bold=False, fmt=None, fill=None, align=None, size=11, italic=False):
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.font = Font(name='Calibri', size=size, bold=bold, italic=italic)
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    return c


def _round(v):
    return round(v, 6)


class _Cols:
    """Column layout of the month area: [m m m Q] [m m m Q] ... Total"""

    def __init__(self, quarters, first_col):
        self.quarters = quarters
        self.first_col = first_col
        self.width = 4 * len(quarters) + 1
        self.month_col = {}    # month -> offset inside the block
        self.q_col = {}        # quarter -> offset
        off = 0
        for q in quarters:
            for m in QUARTER_MONTHS[q]:
                self.month_col[m] = off
                off += 1
            self.q_col[q] = off
            off += 1
        self.total_col = off   # == width - 1

    def month(self, base, m):
        return base + self.month_col[m]

    def q(self, base, q):
        return base + self.q_col[q]

    def total(self, base):
        return base + self.total_col


def _headers_months(ws, row, base, cols, start_year):
    for q in cols.quarters:
        for m in QUARTER_MONTHS[q]:
            _put(ws, row, cols.month(base, m), month_label(m, start_year), bold=True, align=CENTER)
        _put(ws, row, cols.q(base, q), f'Total-Q{q}', bold=True, fill=FILL_Q, align=CENTER)
    _put(ws, row, cols.total(base), 'Total', bold=True, fill=FILL_TOTAL, align=CENTER)


def _row_formulas(ws, row, base, cols, bold=False):
    """Total-Qn and Total formulas of one row (identical to the sample sheet)."""
    for q in cols.quarters:
        ms = QUARTER_MONTHS[q]
        a, b = col_letter(cols.month(base, ms[0])), col_letter(cols.month(base, ms[-1]))
        _put(ws, row, cols.q(base, q), f'=SUM({a}{row}:{b}{row})', bold=bold, fmt=FMT_INT, fill=FILL_Q)
    qcells = '+'.join(f'{col_letter(cols.q(base, q))}{row}' for q in cols.quarters)
    _put(ws, row, cols.total(base), f'={qcells}', bold=bold, fmt=FMT_INT, fill=FILL_TOTAL)


def write_sheet(ws, data):
    spec, start_year = data['spec'], data['start_year']
    quarters, groups, offices = data['quarters'], data['groups'], data['offices']
    cols = _Cols(quarters, 4)
    gw = cols.width
    last_col1 = 3 + gw * len(groups)

    # ---------------- Table 1 : Sales Office x product group
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col1)
    for c in range(1, last_col1 + 1):
        _put(ws, 1, c, bold=True, size=12, align=CENTER)
    ws.cell(row=1, column=1).value = f"{spec['title']} for {data['period']}"
    _put(ws, 2, 1, 'All Values in ₹ Lacs', bold=True, size=10)

    for c, text in ((1, 'Sales Office'), (2, 'AOP'), (3, 'Achieved')):
        ws.merge_cells(start_row=3, start_column=c, end_row=4, end_column=c)
        _put(ws, 3, c, text, bold=True, align=CENTER)
        _put(ws, 4, c, bold=True, align=CENTER)

    for gi, g in enumerate(groups):
        base = 4 + gi * gw
        ws.merge_cells(start_row=3, start_column=base, end_row=3, end_column=base + gw - 1)
        for c in range(base, base + gw):
            _put(ws, 3, c, bold=True, align=CENTER)
        ws.cell(row=3, column=base).value = g
        _headers_months(ws, 4, base, cols, start_year)

    first_row = 5
    for i, (okey, label) in enumerate(offices):
        r = first_row + i
        _put(ws, r, 1, label)
        _put(ws, r, 2, _round(data['office_aop'].get(okey, 0.0)))
        for gi, g in enumerate(groups):
            base = 4 + gi * gw
            for q in quarters:
                for m in QUARTER_MONTHS[q]:
                    _put(ws, r, cols.month(base, m), _round(data['office_group'][okey].get((g, m), 0.0)),
                         fmt=FMT_MONTH)
            _row_formulas(ws, r, base, cols)
        total_cells = '+'.join(f'{col_letter(cols.total(4 + gi * gw))}{r}' for gi in range(len(groups)))
        _put(ws, r, 3, f'={total_cells}', fmt=FMT_INT)
    last_row = first_row + len(offices) - 1
    gt = last_row + 1                      # Grand Total row
    _put(ws, gt, 1, 'Grand Total', bold=True)
    for c in (2, 3):
        L = col_letter(c)
        _put(ws, gt, c, f'=SUM({L}{first_row}:{L}{last_row})', bold=True, fmt=FMT_INT if c == 3 else None)
    for gi, g in enumerate(groups):
        base = 4 + gi * gw
        for q in quarters:
            for m in QUARTER_MONTHS[q]:
                L = col_letter(cols.month(base, m))
                _put(ws, gt, cols.month(base, m), f'=SUM({L}{first_row}:{L}{last_row})', bold=True, fmt=FMT_INT)
        _row_formulas(ws, gt, base, cols, bold=True)

    # ---------------- Table 2 : Unit / Product
    t2 = gt + 3
    cols2 = _Cols(quarters, 5)
    last_col2 = 4 + cols2.width
    for c, text in ((1, 'Unit'), (2, 'Products'), (3, 'AOP'), (4, 'Achieved')):
        _put(ws, t2, c, text, bold=True)
    _headers_months(ws, t2, 5, cols2, start_year)

    r = t2 + 1
    unit_total_rows = []                   # rows that make the grand total (unit totals / single rows)
    for unit, prods in data['unit_rows']:
        unit_first = r
        for p in prods:
            _put(ws, r, 2, p)
            _put(ws, r, 3, _round(data['product_aop'].get(p, 0.0)))
            for q in quarters:
                for m in QUARTER_MONTHS[q]:
                    _put(ws, r, cols2.month(5, m), _round(data['product_month'][p].get(m, 0.0)), fmt=FMT_INT)
            _row_formulas(ws, r, 5, cols2)
            q_total_cells = '+'.join(f'{col_letter(cols2.q(5, q))}{r}' for q in quarters)
            _put(ws, r, 4, f'={q_total_cells}', fmt=FMT_INT)
            r += 1
        unit_last = r - 1
        if len(prods) > 1:
            _put(ws, r, 2, 'Total', bold=True)
            for c in (3, 4):
                L = col_letter(c)
                _put(ws, r, c, f'=SUM({L}{unit_first}:{L}{unit_last})', bold=True, fmt=FMT_INT if c == 4 else None)
            for q in quarters:
                for m in QUARTER_MONTHS[q]:
                    L = col_letter(cols2.month(5, m))
                    _put(ws, r, cols2.month(5, m), f'=SUM({L}{unit_first}:{L}{unit_last})', bold=True, fmt=FMT_INT)
            _row_formulas(ws, r, 5, cols2, bold=True)
            unit_total_rows.append(r)
            r += 1
        else:
            unit_total_rows.append(unit_last)
        end = r - 1
        if end > unit_first:
            ws.merge_cells(start_row=unit_first, start_column=1, end_row=end, end_column=1)
        for rr in range(unit_first, end + 1):
            _put(ws, rr, 1, align=Alignment(horizontal='left', vertical='center'))
        ws.cell(row=unit_first, column=1).value = unit

    g2 = r                                  # Grand Total of table 2
    ws.merge_cells(start_row=g2, start_column=1, end_row=g2, end_column=2)
    _put(ws, g2, 1, 'Grand Total', bold=True, align=Alignment(horizontal='left'))
    _put(ws, g2, 2, bold=True)
    for c in (3, 4):
        L = col_letter(c)
        _put(ws, g2, c, '=' + '+'.join(f'{L}{x}' for x in unit_total_rows), bold=True,
             fmt=FMT_INT if c == 4 else None)
    for q in quarters:
        for m in QUARTER_MONTHS[q]:
            L = col_letter(cols2.month(5, m))
            _put(ws, g2, cols2.month(5, m), '=' + '+'.join(f'{L}{x}' for x in unit_total_rows), bold=True,
                 fmt=FMT_INT)
    _row_formulas(ws, g2, 5, cols2, bold=True)

    if data['undated_rows']:
        n = g2 + 2
        ws.cell(row=n, column=1).value = (
            f"Note: {data['undated_rows']} grouped row(s) with no date (₹ {data['undated_total']:,.2f} Lacs) "
            f"are not part of any month and are excluded from this report.")
        ws.cell(row=n, column=1).font = Font(name='Calibri', size=10, italic=True)

    # widths like the sample: 15 / 19.3 / 13 ...
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 19.3
    for c in range(3, max(last_col1, last_col2) + 1):
        ws.column_dimensions[col_letter(c)].width = 13
    ws.row_dimensions[1].height = 15.75
    return ws


# ----------------------------------------------------------------------------------------------
# orchestration
# ----------------------------------------------------------------------------------------------
def _valid_fin_year(value):
    return bool(re.fullmatch(r'\s*\d{4}\s*-\s*\d{2,4}\s*|\s*\d{4}\s*', value or ''))


def generate_report_data(cursor, spec, fin_year=None):
    fin_year = (fin_year or '').strip() or latest_fin_year(cursor, spec)
    if not fin_year:
        return None
    actual, target = load_rows(cursor, spec, fin_year)
    return build_report_data(spec, fin_year, actual, target)


def build_workbook(datasets, empty=()):
    """datasets: report data dicts; empty: (spec, fin_year) of reports without data (sheet with a message)"""
    wb = Workbook()
    wb.remove(wb.active)
    for data in datasets:
        ws = wb.create_sheet(data['spec']['sheet'])
        write_sheet(ws, data)
    for spec, fin_year in empty:
        ws = wb.create_sheet(spec['sheet'])
        ws['A1'] = f"{spec['title']}: no data found for fin_year {fin_year or '(latest available)'}."
        ws['A1'].font = Font(name='Calibri', size=11, bold=True)
    return wb


def serve_excel(request, specs, base_name):
    fin_year = request.query_params.get('fin_year')
    if fin_year and not _valid_fin_year(fin_year):
        return HTMLResponse(f"Invalid fin_year '{fin_year}'. Use the format 2026-2027.", status_code=400)
    try:
        datasets, empty = [], []
        with get_db_connection() as conn:
            with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                for spec in specs:
                    data = generate_report_data(cursor, spec, fin_year)
                    if data:
                        datasets.append(data)
                    else:
                        empty.append((spec, fin_year))
        if not datasets:
            return HTMLResponse('No data found for the selected financial year.', status_code=404)
        wb = build_workbook(datasets, empty)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        first = datasets[0]
        safe = lambda text: re.sub(r'[^0-9A-Za-z_-]+', '', str(text))      # no quotes / spaces in the file name
        filename = f"{base_name}_{safe(first['period'])}_FY{safe(first['fin_year'])}.xlsx"
        return send_file(output, mimetype=XLSX_MIME, as_attachment=True, download_name=filename)
    except Exception as e:
        print(f"AOP quarterly report error: {e}")
        return HTMLResponse(f'Error exporting data: {str(e)}', status_code=500)


@aop_router.get('/export_aop_quarterly_order_excel')
def export_aop_quarterly_order_excel(request: Request):
    return serve_excel(request, [ORDER_SPEC], 'ORDER_BOOK_AOP')


@aop_router.get('/export_aop_quarterly_billing_excel')
def export_aop_quarterly_billing_excel(request: Request):
    return serve_excel(request, [BILLING_SPEC], 'BILLING_AOP')


@aop_router.get('/export_aop_quarterly_consolidated_excel')
def export_aop_quarterly_consolidated_excel(request: Request):
    return serve_excel(request, [ORDER_SPEC, BILLING_SPEC], 'ORDER_BOOK_AND_BILLING_AOP')
