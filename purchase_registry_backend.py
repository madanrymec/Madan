"""
Purchase Registry - Backend (Optimized)
========================================
This contains everything from your original snippet, with the performance
fixes applied. Merge this into your existing app.py — keep your existing
Flask app instance, other blueprints/routes (home, ytd_dashboard, etc.),
session config, and imports; just swap in the pieces below.

Run this migration once, BEFORE deploying this file:
    -> see purchase_registry_indexes.sql

New dependency (for connection pooling):
    pip install DBUtils
"""

import io
import time
import pymysql
import pandas as pd
from flask import Flask, request, jsonify, render_template, session, send_file
from dbutils.pooled_db import PooledDB

# ----------------------------------------------------------------------------
# Remove this line if you're merging into an existing app.py that already
# creates the Flask app instance.
# ----------------------------------------------------------------------------
app = Flask(__name__)


# ==============================================================================
# 1. CONNECTION POOL
#    Replaces "open a new DB connection per request" with a reusable pool.
#    Without this, every dashboard load pays connection-setup cost ~7-8 times.
# ==============================================================================
DB_CONFIG = dict(
    host="localhost",
    user="your_user",
    password="your_password",
    database="your_database",
    cursorclass=pymysql.cursors.DictCursor,
    charset="utf8mb4",
)

pool = PooledDB(
    creator=pymysql,
    maxconnections=20,   # tune to (gunicorn workers * threads) + headroom
    mincached=2,
    maxcached=5,
    blocking=True,
    ping=1,               # auto-reconnect stale/dropped connections
    **DB_CONFIG,
)


def get_db_connection():
    """Returns a pooled connection. Use as: `with get_db_connection() as conn:`"""
    return pool.connection()


# ==============================================================================
# 2. TINY IN-MEMORY TTL CACHE
#    fin_years / plants barely change — no need to hit the DB for them on
#    every single filter-panel load. No extra dependency required.
# ==============================================================================
_cache_store = {}
_CACHE_TTL_SECONDS = 600  # 10 minutes


def cache_get(key):
    entry = _cache_store.get(key)
    if not entry:
        return None
    value, expires_at = entry
    if time.time() > expires_at:
        _cache_store.pop(key, None)
        return None
    return value


def cache_set(key, value, ttl=_CACHE_TTL_SECONDS):
    _cache_store[key] = (value, time.time() + ttl)


# ==============================================================================
# ROUTES
# ==============================================================================

@app.route('/purchase_registry')
def purchase_registry():
    username = session.get('username', 'Guest')
    return render_template('purchase_reg.html', username=username)


def build_purchase_filter(exclude_vendor=False):
    """
    Builds the SQL WHERE clause for Purchase queries, supporting date ranges
    and the cascading vendor filter.

    CHANGED vs. original:
      - Month filtering now uses the generated/stored `Post_Month` column
        instead of MONTH(Post_Date). Wrapping an indexed column in a function
        makes it "non-sargable" — MySQL can't use an index for it, forcing a
        full table scan. Post_Month is a plain indexed column, so range/IN
        filters on it use the index normally.
      - Exact-date / date-range filtering now uses a sargable range
        (Post_Date >= x AND Post_Date < x + 1 day) instead of DATE(Post_Date),
        which had the same function-on-indexed-column problem.

    Requires the one-time migration in purchase_registry_indexes.sql
    (adds the Post_Month generated column + new covering indexes).
    """
    filters = []
    params = []

    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')
    plants = request.args.getlist('plant')
    vendors = request.args.getlist('vendor')
    exact_date = request.args.get('date')

    if fin_years:
        placeholders = ','.join(['%s'] * len(fin_years))
        filters.append(f"fin_Year IN ({placeholders})")
        params.extend(fin_years)

    if months:
        month_map = {
            "January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
            "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12
        }
        month_nums = [str(month_map[m]) for m in months if m in month_map]
        if month_nums:
            placeholders = ','.join(['%s'] * len(month_nums))
            filters.append(f"Post_Month IN ({placeholders})")  # was: MONTH(Post_Date) IN (...)
            params.extend(month_nums)

    if plants:
        placeholders = ','.join(['%s'] * len(plants))
        filters.append(f"ValA IN ({placeholders})")
        params.extend(plants)

    # Omit vendor constraint when updating the vendor dropdown itself
    if not exclude_vendor and vendors:
        placeholders = ','.join(['%s'] * len(vendors))
        filters.append(f"Vendor_Name IN ({placeholders})")
        params.extend(vendors)

    # Date Range Logic
    if exact_date:
        if " to " in exact_date:
            start_date, end_date = exact_date.split(" to ")
            filters.append("Post_Date >= %s AND Post_Date < DATE_ADD(%s, INTERVAL 1 DAY)")  # was: DATE(Post_Date) BETWEEN %s AND %s
            params.extend([start_date, end_date])
        else:
            filters.append("Post_Date >= %s AND Post_Date < DATE_ADD(%s, INTERVAL 1 DAY)")  # was: DATE(Post_Date) = %s
            params.extend([exact_date, exact_date])

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@app.route('/api/filters/purchase', methods=['GET'])
def get_purchase_filters():
    """
    Fetches unique values for the Purchase Dashboard filters with cascading
    vendor lookup.

    CHANGED: fin_years and plants are cached for 10 minutes (they change
    rarely), removing 2 of the 3 queries on most calls to this endpoint.
    Vendors stay uncached since they're filter-dependent.
    """
    vendor_where, vendor_params = build_purchase_filter(exclude_vendor=True)

    fin_years = cache_get('fin_years')
    plants = cache_get('plants')

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            if fin_years is None:
                cursor.execute(
                    "SELECT DISTINCT fin_Year FROM purchase_registry WHERE fin_Year IS NOT NULL ORDER BY fin_Year DESC")
                fin_years = [r['fin_Year'] for r in cursor.fetchall() if r['fin_Year']]
                cache_set('fin_years', fin_years)

            if plants is None:
                cursor.execute(
                    "SELECT DISTINCT ValA FROM purchase_registry WHERE ValA IS NOT NULL ORDER BY ValA ASC")
                plants = [r['ValA'] for r in cursor.fetchall() if r['ValA']]
                cache_set('plants', plants)

            # Dependent: Vendors matching selected Year, Month, Date Range, and Plant
            if vendor_where:
                vendor_query = f"SELECT DISTINCT Vendor_Name FROM purchase_registry {vendor_where} AND Vendor_Name IS NOT NULL ORDER BY Vendor_Name ASC"
            else:
                vendor_query = "SELECT DISTINCT Vendor_Name FROM purchase_registry WHERE Vendor_Name IS NOT NULL ORDER BY Vendor_Name ASC"

            cursor.execute(vendor_query, vendor_params)
            vendors = [r['Vendor_Name'] for r in cursor.fetchall() if r['Vendor_Name']]

    return jsonify({"fin_years": fin_years, "plants": plants, "vendors": vendors})


def execute_bar_query(query, params):
    """Helper function to execute and format standard Bar Chart API responses."""
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            data = cursor.fetchall()
    return jsonify({"labels": [r['label'] for r in data], "values": [r['val'] or 0 for r in data]})


@app.route('/api/purchase/dashboard/kpis', methods=['GET'])
def get_purchase_kpis():
    """
    CHANGED: collapsed 5 separate full-table-scan queries (SUM + 4x
    COUNT DISTINCT, each scanning the same filtered rows independently)
    into a single query / single pass. Backed by the covering index
    idx_pr_kpi (fin_Year, Post_Month, Total, po_number, Material,
    Vendor_Name, ValA) from the SQL migration.
    """
    where, params = build_purchase_filter()
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(f"""
                SELECT
                    SUM(Total) AS total_spend,
                    COUNT(DISTINCT po_number) AS po_count,
                    COUNT(DISTINCT Material) AS material_count,
                    COUNT(DISTINCT Vendor_Name) AS vendor_count,
                    COUNT(DISTINCT ValA) AS plant_count
                FROM purchase_registry {where}
            """, params)
            row = cursor.fetchone() or {}

    total_spend = row.get('total_spend') or 0
    po_count = row.get('po_count') or 0
    avg_po = (total_spend / po_count) if po_count else 0

    return jsonify({
        "total_spend": total_spend, "po_count": po_count,
        "material_count": row.get('material_count') or 0,
        "vendor_count": row.get('vendor_count') or 0,
        "plant_count": row.get('plant_count') or 0,
        "avg_po_value": avg_po
    })


@app.route('/api/purchase/dashboard/trend', methods=['GET'])
def get_purchase_trend():
    """CHANGED: month grouping now uses Post_Month instead of MONTH(Post_Date)."""
    where, params = build_purchase_filter()
    months = request.args.getlist('month')
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            if len(months) == 1:
                query = f"""SELECT FLOOR((DAY(Post_Date)-1)/7)+1 AS week_num, SUM(Total) AS val
                            FROM purchase_registry {where}
                            GROUP BY FLOOR((DAY(Post_Date)-1)/7)+1 ORDER BY FLOOR((DAY(Post_Date)-1)/7)+1"""
            else:
                query = f"""SELECT Post_Month AS month_num, SUM(Total) AS val
                            FROM purchase_registry {where}
                            GROUP BY Post_Month ORDER BY Post_Month"""  # was: MONTH(Post_Date)
            cursor.execute(query, params)
            raw_data = cursor.fetchall()

    trend_data = []
    if len(months) == 1:
        for r in raw_data:
            trend_data.append({"label": f"Week {r['week_num']}", "val": r['val']})
    else:
        month_names = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sep",
                       10: "Oct", 11: "Nov", 12: "Dec"}
        for r in raw_data:
            trend_data.append({"label": month_names.get(r['month_num'], "Unknown"), "val": r['val']})

    return jsonify({"labels": [r['label'] for r in trend_data], "values": [r['val'] or 0 for r in trend_data]})


@app.route('/api/purchase/dashboard/plants', methods=['GET'])
def get_purchase_plants():
    where, params = build_purchase_filter()
    return execute_bar_query(
        f"SELECT ValA as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY ValA ORDER BY SUM(Total) DESC",
        params)


@app.route('/api/purchase/dashboard/potypes', methods=['GET'])
def get_purchase_potypes():
    where, params = build_purchase_filter()
    return execute_bar_query(
        f"SELECT po_type as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY po_type ORDER BY SUM(Total) DESC",
        params)


@app.route('/api/purchase/dashboard/vendors', methods=['GET'])
def get_purchase_vendors():
    where, params = build_purchase_filter()
    return execute_bar_query(
        f"SELECT Vendor_Name as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY Vendor_Name ORDER BY SUM(Total) DESC LIMIT 10",
        params)


@app.route('/api/purchase/dashboard/materials', methods=['GET'])
def get_purchase_materials():
    where, params = build_purchase_filter()
    return execute_bar_query(
        f"SELECT Short_Text as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY Short_Text ORDER BY SUM(Total) DESC LIMIT 10",
        params)


@app.route('/api/purchase/dashboard/table', methods=['GET'])
def get_purchase_table():
    """
    CHANGED: removed the separate SUM(Total) query that ran just to get the
    grand total for share%. It's now derived in Python from the grouped
    result — this endpoint went from 2 full scans down to 1.
    """
    where, params = build_purchase_filter()
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(f"""
                SELECT ValA as plant, SUM(Total) as total, COUNT(DISTINCT po_number) as pos, COUNT(DISTINCT Vendor_Name) as vendors
                FROM purchase_registry {where} GROUP BY ValA ORDER BY SUM(Total) DESC
            """, params)
            table_data = cursor.fetchall()

    total_spend = sum((row['total'] or 0) for row in table_data)  # was a separate SELECT SUM(Total) query
    for row in table_data:
        row['total'] = row['total'] or 0
        row['share'] = (row['total'] / total_spend * 100) if total_spend else 0
        row['avgpo'] = (row['total'] / row['pos']) if row['pos'] else 0

    return jsonify(table_data)


@app.route('/api/export/purchase_excel', methods=['GET'])
def export_purchase_excel():
    """Exports the filtered purchase registry data to an Excel file."""
    where_clause, params = build_purchase_filter()
    query = f"SELECT * FROM purchase_registry {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

        df = pd.DataFrame(results)

        if 'Post_Date' in df.columns:
            df['Post_Date'] = pd.to_datetime(df['Post_Date'], errors='coerce').dt.strftime('%d-%m-%Y')

        if 'Post_Month' in df.columns:
            df = df.drop(columns=['Post_Month'])  # internal helper column — not useful in the export

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Purchase_Export')
        output.seek(0)

        fin_years = request.args.getlist('fin_year')
        fy_str = "_".join(fin_years) if fin_years else "All_Years"
        filename = f"Purchase_Registry_Export_{fy_str}.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return f"Error exporting data: {str(e)}", 500


# ----------------------------------------------------------------------------
# NOTE: your other routes (home, order_ytd_dashboard, pending_orders_page,
# ytd_dashboard, collection_report, mis_report, mis_unit_report, etc.)
# referenced by the sidebar/template weren't part of the original snippet,
# so they aren't reproduced here — keep them as-is in your existing app.py.
# ----------------------------------------------------------------------------

if __name__ == '__main__':
    # For production, run under gunicorn with multiple workers/threads instead,
    # e.g.:  gunicorn -w 4 --threads 4 -b 0.0.0.0:5000 app:app
    # The Flask dev server handles one request at a time by default, which
    # will make your 7 "parallel" dashboard fetches queue up sequentially.
    app.run(threaded=True, debug=True)
