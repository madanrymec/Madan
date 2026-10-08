import calendar
import io
import os
import hashlib
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
import requests
from openpyxl.styles import PatternFill, Side, Border, Font, Alignment
from openpyxl.utils import get_column_letter
from requests.auth import HTTPBasicAuth
from apscheduler.executors.pool import ThreadPoolExecutor
from contextlib import nullcontext
import shutil
import smtplib
import openpyxl
from datetime import datetime
from datetime import timedelta
from email.message import EmailMessage
from io import BytesIO
import traceback
import pandas as pd
import pymysql
import smbclient
from apscheduler.schedulers.background import BackgroundScheduler
import json
import mimetypes
import uvicorn
from datetime import timezone
from decimal import Decimal as _Decimal
from email.utils import format_datetime
from urllib.parse import quote, urlencode
from fastapi import FastAPI, Request, Form
from fastapi.responses import JSONResponse, HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
)
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt
from sqlalchemy import create_engine, text
from Products.product_head import prod_app
from Products.unit_head import unit_app
from Products.finance_head import fin_app
from functools import wraps


DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'dashboard',
}
USER_DB_CONFIG = {
    'host': '192.7.200.47',
    'user': 'root',
    'password': '',
    'database': 'enquiry_portal'
}
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = FastAPI()
# NOTE: Products/product_head.py, unit_head.py, finance_head.py must expose fastapi.APIRouter objects
app.include_router(prod_app)
app.include_router(unit_app)
app.include_router(fin_app)
app.add_middleware(SessionMiddleware, secret_key='Kec@2025@11')
app.mount('/static', StaticFiles(directory=os.path.join(BASE_DIR, 'static'), check_dir=False), name='static')
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, 'templates'))


# ------------------------------------------------------------------
# Flask-compatibility helpers (keep every route body identical to the Flask version)
# ------------------------------------------------------------------
def _json_default(o):
    """Same fallbacks Flask's jsonify used (Decimal -> str, date/datetime -> HTTP date)."""
    if isinstance(o, datetime):
        o = o.replace(tzinfo=timezone.utc) if o.tzinfo is None else o.astimezone(timezone.utc)
        return format_datetime(o, usegmt=True)
    if isinstance(o, date):
        return format_datetime(datetime(o.year, o.month, o.day, tzinfo=timezone.utc), usegmt=True)
    if isinstance(o, _Decimal):
        return str(o)
    return str(o)


class FlaskJSONResponse(JSONResponse):
    def render(self, content) -> bytes:
        return json.dumps(content, default=_json_default, sort_keys=True,
                          ensure_ascii=True, separators=(',', ':')).encode('utf-8')


def jsonify(data=None, status_code=200):
    return FlaskJSONResponse(content=data, status_code=status_code)


def redirect(location, code=302):
    return RedirectResponse(url=location, status_code=code)


def url_for(endpoint, **values):
    """Flask style url_for: 'static' + filename, blueprint.endpoint names, extra values -> query string."""
    if endpoint == 'static':
        filename = values.get('filename', values.get('path', ''))
        return '/static/' + str(filename).lstrip('/')
    name = endpoint.split('.')[-1]
    values.pop('_external', None)
    for route in app.routes:
        if getattr(route, 'name', None) == name and hasattr(route, 'param_convertors'):
            path_values = {k: values.pop(k) for k in list(values) if k in route.param_convertors}
            url = str(route.url_path_for(name, **path_values))
            return url + ('?' + urlencode(values, doseq=True) if values else '')
    raise KeyError(f"No route named '{endpoint}'")


templates.env.globals['url_for'] = url_for


def render_template(request, template_name, **context):
    context.setdefault('session', request.session)
    return templates.TemplateResponse(request, template_name, context)


def send_file(path_or_file, mimetype=None, as_attachment=False, download_name=None):
    if isinstance(path_or_file, (str, os.PathLike)):
        with open(path_or_file, 'rb') as fh:
            data = fh.read()
        if download_name is None:
            download_name = os.path.basename(path_or_file)
    elif hasattr(path_or_file, 'getvalue'):
        data = path_or_file.getvalue()
    else:
        data = path_or_file.read()
    if mimetype is None and download_name:
        mimetype = mimetypes.guess_type(download_name)[0]
        if mimetype is None and download_name.lower().endswith('.xlsx'):
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    if mimetype is None:
        mimetype = 'application/octet-stream'
    headers = {'Content-Length': str(len(data))}
    if as_attachment and download_name:
        ascii_name = download_name.encode('ascii', 'ignore').decode().replace('"', '')
        headers['Content-Disposition'] = f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(download_name, safe='')}"
    return Response(content=data, media_type=mimetype, headers=headers)


@app.get('/test')
def test():
    return HTMLResponse("TEST OK")


def get_db_connection():
    """Connects to a MySQL database."""
    try:
        conn = pymysql.connect(
            cursorclass=pymysql.cursors.DictCursor,
            **DB_CONFIG
        )
        return conn
    except Exception as e:
        print(f"Database connection failed: {e}")
        return None


def get_user_db_connection():
    try:
        return pymysql.connect(cursorclass=pymysql.cursors.DictCursor, **DB_CONFIG)
    except Exception as e:
        print(f"User DB Connection failed: {e}")
        return None


def requires_login(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        request = kwargs.get('request')
        if not request.session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)

    decorated_function.__name__ = f.__name__
    return decorated_function


@app.api_route('/', methods=['GET', 'POST'])
@app.api_route('/login', methods=['GET', 'POST'])
def login(request: Request, username: str = Form(None), password: str = Form(None)):
    error = None
    if request.method == 'POST':
        emp_id_input = username
        password_input = password

        conn = get_user_db_connection()
        # Assuming you use dictionary=True or a row factory since you access by keys
        cursor = conn.cursor()

        query = """
            SELECT emp_id, emp_name, permission 
            FROM user_mast 
            WHERE emp_id = %s AND BINARY password = %s AND active_status = 'active'
        """
        cursor.execute(query, (emp_id_input, password_input))
        user = cursor.fetchone()
        conn.close()

        if user:
            # Clear old session data completely to avoid permission overlap
            request.session.clear()

            # Set basic session data
            request.session['logged_in'] = True

            # Handle tuple vs dict safely depending on your cursor type
            if isinstance(user, dict):
                request.session['emp_id'] = user['emp_id'].strip()
                request.session['username'] = user['emp_name']
                perm = user['permission']
            else:
                request.session['emp_id'] = user[0].strip()
                request.session['username'] = user[1]
                perm = user[2]

            if not perm or perm == 'NULL':
                return redirect(url_for('home'))

            # Save exact permission for Jinja logic
            request.session['permission'] = perm

            # --- ROUTING & RBAC LOGIC ---
            if perm.startswith('unit_'):
                request.session['role_type'] = 'unit_head'
                # Redirect normal users to the standard home/landing page
                target_route = 'home'

            elif perm.startswith('finance_'):
                request.session['role_type'] = 'finance_head'
                # Redirect ALL finance heads to their new central radial menu
                target_route = 'finance_head.finance_home'

            elif perm.startswith('product_'):
                request.session['role_type'] = 'product_head'
                target_route = 'product_head.product_home'
            else:
                target_route = 'home'

            try:
                return redirect(url_for(target_route))
            except Exception as e:
                print(f"Routing Error for {target_route}: {e}")
                return redirect(url_for('home'))
        else:
            error = "Invalid Employee ID or Password"

    return render_template(request,'login.html', error=error)


@app.get('/logout')
def logout(request: Request):
    request.session.pop('logged_in', None)
    request.session.pop('username', None)
    return redirect(url_for('login'))


# def get_all_fin_years(cursor):
#     cursor.execute("SELECT DISTINCT fin_year FROM billing_data ORDER BY fin_year DESC")
#     years = [row['fin_year'] for row in cursor.fetchall()]
#     return years if years else ['2024-2025']


def build_filter_query(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Optimized for MySQL using %s placeholders.
    """
    filters = []
    params = []

    # Extract multiple selections
    regions = request.query_params.getlist('region')
    offices = request.query_params.getlist('sales_office')
    products = request.query_params.getlist('product')
    units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # Region Handling
    if regions:
        placeholders = ','.join(['%s'] * len(regions))
        filters.append(f"Sales_Region_Name IN ({placeholders})")
        params.extend(regions)

    # Sales Office Handling
    if offices:
        placeholders = ','.join(['%s'] * len(offices))
        filters.append(f"Sales_office IN ({placeholders})")
        params.extend(offices)

    # Product handling
    if products:
        placeholders = ','.join(['%s'] * len(products))
        if table_type == "data":
            mapped_prod_sql = """
                CASE 
                    WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                    WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                    ELSE Prod 
                END
            """
            filters.append(f"({mapped_prod_sql}) IN ({placeholders})")
        else:
            filters.append(f"Product IN ({placeholders})")
        params.extend(products)

    # Unit Handling (NEW)
    if units:
        placeholders = ','.join(['%s'] * len(units))
        if table_type == "data":
            mapped_unit_sql = """
                CASE 
                    WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                    WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                    ELSE Unit 
                END
            """
            filters.append(f"({mapped_unit_sql}) IN ({placeholders})")
        else:
            mapped_unit_target_sql = """
                CASE 
                    WHEN Product = 'TRANSFORMERS PUNE' THEN 'UN16'
                    ELSE Unit 
                END
            """
            filters.append(f"({mapped_unit_target_sql}) IN ({placeholders})")
        params.extend(units)

    # Financial Year Handling
    if fin_years:
        placeholders = ','.join(['%s'] * len(fin_years))
        filters.append(f"fin_year IN ({placeholders})")
        params.extend(fin_years)

    # Month Handling
    if months and table_type == "data":
        placeholders = ','.join(['%s'] * len(months))
        filters.append(f"MONTHNAME(Billing_Date) IN ({placeholders})")
        params.extend(months)

    # Exclude Inter Unit Transfers strictly for billing_data
    if table_type == "data":
        filters.append("Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod IS NOT NULL")

    where_clause = " AND ".join(filters) if filters else "1=1"
    return "WHERE " + where_clause, params

# ==========================================
# PAGE ROUTE
# ==========================================


@app.get('/home')
def home(request: Request):
    # Assuming login logic is handled elsewhere, defaulting to a static render for now
    username = request.session.get('username', 'Guest')
    return render_template(request,'home.html', username=username)


# ==========================================
# API ENDPOINTS
# ==========================================

@app.get('/api/filters')
def get_filters(request: Request):
    """Fetches distinct filter options, handling cascading dependencies (Region->Office, Unit->Product)."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("""
                SELECT DISTINCT Sales_Region_Name 
                FROM billing_data 
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_Region_Name IS NOT NULL
            """)
            regions = [r['Sales_Region_Name'] for r in cursor.fetchall() if r['Sales_Region_Name']]

            cursor.execute("""
                SELECT DISTINCT fin_year 
                FROM billing_data 
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer' AND fin_year IS NOT NULL
            """)
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = """
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM billing_data 
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = """
                SELECT DISTINCT Sales_office 
                FROM billing_data 
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office IS NOT NULL
            """
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit using dynamic mapping)
            prod_query = """
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod 
                    END as mapped_prod
                FROM billing_data
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod IS NOT NULL
            """
            prod_params = []
            if units_selected:
                placeholders = ','.join(['%s'] * len(units_selected))
                mapped_unit_sql = """
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END
                """
                prod_query += f" AND ({mapped_unit_sql}) IN ({placeholders})"
                prod_params.extend(units_selected)

            cursor.execute(prod_query, prod_params)
            products = [r['mapped_prod'] for r in cursor.fetchall() if r['mapped_prod']]

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })


@app.get('/api/kpis')
def get_kpis(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query(request, "data")
    target_where, target_params = build_filter_query(request, "target")

    # Determine proration factor based on selected months
    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # 1. Query Actuals from billing_data (Optimized)
    data_query = f"""
        SELECT SUM(Net_Value) as overall_sales
        FROM billing_data
        {data_where}
    """

    # 2. Query Targets from billing_target
    target_query = f"""
        SELECT (SUM(Target) * 100000 * {proration_factor}) as total_target
        FROM billing_target1
        {target_where}
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Fetch Actuals
            cursor.execute(data_query, data_params)
            result = cursor.fetchone() or {}

            # Fetch Targets
            cursor.execute(target_query, target_params)
            target_result = cursor.fetchone() or {}

    # Map results and safely handle NULLs
    result['overall_sales'] = result.get('overall_sales') or 0
    result['total_target'] = target_result.get('total_target') or 0

    return jsonify(result)


@app.get('/api/sales-trend')
def get_sales_trend(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query(request, "data")

    # 1. Monthly Query (For MoM Growth)
    monthly_query = f"""
        SELECT 
            DATE_FORMAT(Billing_Date, '%%Y-%%m') as time_period,
            SUM(Net_Value) as total_revenue
        FROM billing_data
        {where_clause}
        GROUP BY time_period
        ORDER BY time_period ASC
    """

    weekly_query = f"""
        SELECT 
            DATE_FORMAT(MIN(Billing_Date), '%%d %%b %%y') as time_period,
            SUM(Net_Value) as total_revenue
        FROM billing_data
        {where_clause}
        GROUP BY YEARWEEK(Billing_Date, 1)
        ORDER BY MIN(Billing_Date) ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Execute Monthly
            cursor.execute(monthly_query, params)
            monthly_results = cursor.fetchall()

            # Execute Weekly
            cursor.execute(weekly_query, params)
            weekly_results = cursor.fetchall()

    # Return both datasets cleanly
    return jsonify({
        "monthly": monthly_results,
        "weekly": weekly_results
    })

@app.get('/api/products-revenue')
def get_products_revenue(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query(request, "data")

    query = f"""
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                ELSE Prod 
            END as mapped_product,
            SUM(Net_Value) as revenue,
            SUM(Billing_Qty) as quantity
        FROM billing_data
        {where_clause}
        GROUP BY mapped_product
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            results = cursor.fetchall()

    return jsonify(results)


@app.get('/api/sales-by-office')
def get_sales_by_office(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query(request, "data")

    # Added LIMIT 12 to optimize DB performance for large datasets
    office_query = f"""
        SELECT Sales_office, SUM(Net_Value) as revenue
        FROM billing_data {where_clause}
        GROUP BY Sales_office 
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(office_query, params)
            offices = cursor.fetchall()

    return jsonify(offices)


@app.get('/api/sales-by-region')
def get_sales_by_region(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query(request, "data")

    region_query = f"""
        SELECT Sales_Region_Name, SUM(Net_Value) as revenue
        FROM billing_data {where_clause}
        GROUP BY Sales_Region_Name 
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(region_query, params)
            regions = cursor.fetchall()

    return jsonify(regions)



@app.get('/api/aop-achievement')
def get_aop_achievement(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query(request, "data")
    target_where, target_params = build_filter_query(request, "target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # 1. Fetch Actuals from billing_data
    # (Corrected 'TRANSFORMERS PUNE' spelling to match the database/image)
    actuals_query = f"""
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                ELSE Prod 
            END as mapped_product,
            CASE 
                WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                ELSE Unit 
            END as mapped_unit,
            SUM(Net_Value) as actual_revenue
        FROM billing_data
        {data_where}
        GROUP BY mapped_product, mapped_unit
        ORDER BY mapped_unit ASC
    """

    # 2. Fetch Targets from billing_target
    # (Applied proration_factor to the target revenue calculation)
    targets_query = f"""
        SELECT 
            Product,
            CASE 
                WHEN Product = 'TRANSFORMERS PUNE' THEN 'UN16'
                ELSE Unit 
            END as mapped_unit,
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue
        FROM billing_target1
        {target_where}
        GROUP BY Product, mapped_unit
        ORDER BY mapped_unit ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    # 3. Python Mapping (O(N) Complexity - No SQL JOIN used)
    merged_data = {}

    # Initialize dictionary with targets
    for row in targets_raw:
        product = row['Product']
        unit = row['mapped_unit']
        key = f"{product}|{unit}"

        merged_data[key] = {
            "product": product,
            "unit": unit,
            "target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map actuals onto the dictionary
    for row in actuals_raw:
        product = row['mapped_product']
        unit = row['mapped_unit']
        key = f"{product}|{unit}"

        if key not in merged_data:
            merged_data[key] = {
                "product": product,
                "unit": unit,
                "target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[key]["actual"] += float(row['actual_revenue'] or 0)

    # 4. Calculate Final Achievement Percentages
    final_results = []
    for data in merged_data.values():
        if data["target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)


@app.get('/api/un07-sales')
def get_un07_sales(request: Request):
    filters = ["Unit = 'UN07'"]
    params = []

    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    if fin_years:
        placeholders = ','.join(['%s'] * len(fin_years))
        filters.append(f"fin_year IN ({placeholders})")
        params.extend(fin_years)

    if months:
        placeholders = ','.join(['%s'] * len(months))
        filters.append(f"MONTHNAME(Billing_Date) IN ({placeholders})")
        params.extend(months)

    where_clause = " AND ".join(filters)

    query = f"""
        SELECT SUM(Net_Value) as un07_sales 
        FROM billing_data 
        WHERE {where_clause}
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            result = cursor.fetchone()

    un07_sales = result.get('un07_sales') if result and result.get('un07_sales') else 0

    return jsonify({"un07_sales": float(un07_sales)})

# ==========================================
# QUERY BUILDER (Fixed & Corrected)
# ==========================================
def build_filter_orders_query(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters specifically for Orders based on frontend filters.
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')  # NEW: Extract Unit filter
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # Region
    if order_regions:
        placeholders = ','.join(['%s'] * len(order_regions))
        filters.append(f"Sales_Region IN ({placeholders})")
        params.extend(order_regions)

    # Sales Office
    if order_offices:
        placeholders = ','.join(['%s'] * len(order_offices))
        filters.append(f"Sales_office IN ({placeholders})")
        params.extend(order_offices)

    # Product handling
    if order_products:
        placeholders = ','.join(['%s'] * len(order_products))
        if table_type == "order_data":
            mapped_prod_sql = """
                CASE 
                    WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                    ELSE Product1 
                END
            """
            filters.append(f"({mapped_prod_sql}) IN ({placeholders})")
        else:
            filters.append(f"Product1 IN ({placeholders})")
        params.extend(order_products)

    # Unit Handling (NEW)
    if order_units:
        placeholders = ','.join(['%s'] * len(order_units))
        if table_type == "order_data":
            mapped_unit_sql = """
                CASE 
                    WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                    WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                    ELSE Unit 
                END
            """
            filters.append(f"({mapped_unit_sql}) IN ({placeholders})")
        else:
            # order_target mapping
            mapped_unit_target_sql = """
                CASE 
                    WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                    ELSE Unit 
                END
            """
            filters.append(f"({mapped_unit_target_sql}) IN ({placeholders})")
        params.extend(order_units)

    if fin_years:
        placeholders = ','.join(['%s'] * len(fin_years))
        filters.append(f"fin_year IN ({placeholders})")
        params.extend(fin_years)

    if months and table_type == "order_data":
        placeholders = ','.join(['%s'] * len(months))
        filters.append(f"MONTHNAME(Created_Date) IN ({placeholders})")
        params.extend(months)

    if table_type == "order_data":
        filters.append("Product1 IS NOT NULL")

    where_clause = " AND ".join(filters) if filters else "1=1"
    return "WHERE " + where_clause, params

# ==========================================
# PAGE ROUTE
# ==========================================
@app.get('/orders')
def order(request: Request):
    username = request.session.get('username', 'Guest')
    return render_template(request,'orders_dashboard.html', username=username)


# ==========================================
# API ENDPOINTS (Points to build_filter_orders_query)
# ==========================================

@app.get('/api/filters/orders')
def get_orders_filters(request: Request):
    """Fetches distinct filter options for Orders, handling cascading dependencies (Region->Office, Unit->Product)."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters (Regions, Fin Years, Units)
            cursor.execute("SELECT DISTINCT Sales_Region FROM order_data WHERE Sales_Region IS NOT NULL")
            regions = [r['Sales_Region'] for r in cursor.fetchall() if r['Sales_Region']]

            cursor.execute("SELECT DISTINCT fin_year FROM order_data WHERE fin_year IS NOT NULL")
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = """
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM order_data 
                WHERE Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = "SELECT DISTINCT Sales_office FROM order_data WHERE Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit)
            prod_query = """
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1 
                    END as mapped_prod
                FROM order_data
                WHERE Product1 IS NOT NULL
            """
            prod_params = []
            if units_selected:
                placeholders = ','.join(['%s'] * len(units_selected))
                mapped_unit_sql = """
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END
                """
                prod_query += f" AND ({mapped_unit_sql}) IN ({placeholders})"
                prod_params.extend(units_selected)

            cursor.execute(prod_query, prod_params)
            products = [r['mapped_prod'] for r in cursor.fetchall() if r['mapped_prod']]

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@app.get('/api/order/kpis')
def get_orders_kpis(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query(request, "order_data")
    target_where, target_params = build_filter_orders_query(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # 1. Query Actuals
    data_query = f"""
        SELECT 
            SUM(Net_Value) as overall_sales,
            COUNT(DISTINCT Customer) as customers_billed
        FROM order_data
        {data_where}
    """

    # 2. Query Targets
    target_query = f"""
        SELECT 
            (SUM(Target) * 100000 * {proration_factor}) as total_target
        FROM order_target1
        {target_where}
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(data_query, data_params)
            result = cursor.fetchone() or {}

            cursor.execute(target_query, target_params)
            target_result = cursor.fetchone() or {}

    result['overall_sales'] = result.get('overall_sales') or 0
    result['total_target'] = target_result.get('total_target') or 0
    result['customers_billed'] = result.get('customers_billed') or 0

    return jsonify(result)


@app.get('/api/orders-trend')
def get_orders_trend(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query(request, "order_data")

    # 1. Monthly Query (For MoM Growth)
    monthly_query = f"""
        SELECT 
            DATE_FORMAT(Created_Date, '%%Y-%%m') as time_period,
            SUM(Net_Value) as total_revenue
        FROM order_data
        {where_clause}
        GROUP BY time_period
        ORDER BY time_period ASC
    """

    weekly_query = f"""
        SELECT 
            DATE_FORMAT(MIN(Created_Date), '%%d %%b %%y') as time_period,
            SUM(Net_Value) as total_revenue
        FROM order_data
        {where_clause}
        GROUP BY YEARWEEK(Created_Date, 1)
        ORDER BY MIN(Created_Date) ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Execute Monthly Query
            cursor.execute(monthly_query, params)
            monthly_results = cursor.fetchall()

            # Execute Weekly Query
            cursor.execute(weekly_query, params)
            weekly_results = cursor.fetchall()

    return jsonify({
        "monthly": monthly_results,
        "weekly": weekly_results
    })


@app.get('/api/orders/products-revenue')
def get_orders_products_revenue(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query(request, "order_data")

    query = f"""
        SELECT 
            CASE 
                WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as revenue,
            SUM(Order_Qty) as quantity
        FROM order_data
        {where_clause}
        GROUP BY mapped_product
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            results = cursor.fetchall()

    return jsonify(results)


@app.get('/api/orders/sales-by-office')
def get_orders_sales_by_office(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query(request, "order_data")

    office_query = f"""
        SELECT Sales_office, SUM(Net_Value) as revenue
        FROM order_data {where_clause}
        GROUP BY Sales_office 
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(office_query, params)
            offices = cursor.fetchall()

    return jsonify(offices)


@app.get('/api/order/sales-by-region')
def get_order_sales_by_region(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query(request, "order_data")

    region_query = f"""
        SELECT Sales_Region, SUM(Net_Value) as revenue
        FROM order_data {where_clause}
        GROUP BY Sales_Region 
        ORDER BY revenue DESC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(region_query, params)
            regions = cursor.fetchall()

    return jsonify(regions)


@app.get('/api/orders/aop-achievement')
def get_orders_aop_achievement(request: Request):
    # FIXED FUNCTION CALLS HERE
    data_where, data_params = build_filter_orders_query(request, "order_data")
    target_where, target_params = build_filter_orders_query(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    actuals_query = f"""
            SELECT 
                CASE 
                   WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                    ELSE Product1
                END AS mapped_product,
                CASE 
                    WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                    WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                    ELSE Unit 
                END as mapped_unit,
                SUM(Net_Value) as actual_revenue
            FROM order_data
            {data_where}
            GROUP BY mapped_product, mapped_unit
            ORDER BY mapped_unit ASC
        """

    targets_query = f"""
            SELECT 
                Product1, 
                Unit AS mapped_unit, 
                (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
            FROM order_target1
            {target_where}
            GROUP BY Product1, mapped_unit
            ORDER BY mapped_unit ASC
        """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    for row in targets_raw:
        product = row['Product1']
        unit = row['mapped_unit']
        key = f"{product}|{unit}"

        merged_data[key] = {
            "product": product,
            "unit": unit,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    for row in actuals_raw:
        product = row['mapped_product']
        unit = row['mapped_unit']
        key = f"{product}|{unit}"

        if key not in merged_data:
            merged_data[key] = {
                "product": product,
                "unit": unit,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[key]["actual"] += float(row['actual_revenue'] or 0)

    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

#-----------------------------------------
#PURCHASE REGISTRY
#------------------------------------------

@app.get('/purchase_registry')
def purchase_registry(request: Request):
    username = request.session.get('username', 'Guest')
    return render_template(request,'purchase_reg.html', username=username)


def build_purchase_filter(request, exclude_vendor=False):
    """Helper function to build SQL WHERE clause for Purchase queries, supporting date ranges and cascading vendor filter."""
    filters = []
    params = []

    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')
    plants = request.query_params.getlist('plant')
    vendors = request.query_params.getlist('vendor')
    exact_date = request.query_params.get('date')

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
            filters.append(f"MONTH(Post_Date) IN ({placeholders})")
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
            filters.append("Post_Date >= %s AND Post_Date < DATE_ADD(%s, INTERVAL 1 DAY)")
            params.extend([start_date, end_date])
        else:
            filters.append("Post_Date >= %s AND Post_Date < DATE_ADD(%s, INTERVAL 1 DAY)")
            params.extend([exact_date, exact_date])

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@app.get('/api/filters/purchase')
def get_purchase_filters(request: Request):
    """Fetches unique values for the Purchase Dashboard filters with cascading vendor lookup."""
    # Build filter using date parameters and plant, excluding vendor
    vendor_where, vendor_params = build_purchase_filter(request, exclude_vendor=True)

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Independent: Financial Years & Plants
            cursor.execute("SELECT DISTINCT fin_Year FROM purchase_registry WHERE fin_Year IS NOT NULL ORDER BY fin_Year DESC")
            fin_years = [r['fin_Year'] for r in cursor.fetchall() if r['fin_Year']]

            cursor.execute("SELECT DISTINCT ValA FROM purchase_registry WHERE ValA IS NOT NULL ORDER BY ValA ASC")
            plants = [r['ValA'] for r in cursor.fetchall() if r['ValA']]

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


@app.get('/api/purchase/dashboard/kpis')
def get_purchase_kpis(request: Request):
    where, params = build_purchase_filter(request)
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT SUM(Total) as val FROM purchase_registry {where}", params)
            total_spend = (cursor.fetchone() or {}).get('val') or 0

            cursor.execute(f"SELECT COUNT(DISTINCT po_number) as val FROM purchase_registry {where}", params)
            po_count = (cursor.fetchone() or {}).get('val') or 0

            cursor.execute(f"SELECT COUNT(DISTINCT Material) as val FROM purchase_registry {where}", params)
            material_count = (cursor.fetchone() or {}).get('val') or 0

            cursor.execute(f"SELECT COUNT(DISTINCT Vendor_Name) as val FROM purchase_registry {where}", params)
            vendor_count = (cursor.fetchone() or {}).get('val') or 0

            cursor.execute(f"SELECT COUNT(DISTINCT ValA) as val FROM purchase_registry {where}", params)
            plant_count = (cursor.fetchone() or {}).get('val') or 0

    avg_po = (total_spend / po_count) if po_count else 0
    return jsonify({
        "total_spend": total_spend, "po_count": po_count,
        "material_count": material_count, "vendor_count": vendor_count,
        "plant_count": plant_count, "avg_po_value": avg_po
    })


@app.get('/api/purchase/dashboard/trend')
def get_purchase_trend(request: Request):
    where, params = build_purchase_filter(request)
    months = request.query_params.getlist('month')
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            if len(months) == 1:
                query = f"""SELECT FLOOR((DAY(Post_Date)-1)/7)+1 AS week_num, SUM(Total) AS val 
                            FROM purchase_registry {where} 
                            GROUP BY FLOOR((DAY(Post_Date)-1)/7)+1 ORDER BY FLOOR((DAY(Post_Date)-1)/7)+1"""
            else:
                query = f"""SELECT MONTH(Post_Date) AS month_num, SUM(Total) AS val 
                            FROM purchase_registry {where} 
                            GROUP BY MONTH(Post_Date) ORDER BY MONTH(Post_Date)"""
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


@app.get('/api/purchase/dashboard/plants')
def get_purchase_plants(request: Request):
    where, params = build_purchase_filter(request)
    return execute_bar_query(
        f"SELECT ValA as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY ValA ORDER BY SUM(Total) DESC",
        params)


@app.get('/api/purchase/dashboard/potypes')
def get_purchase_potypes(request: Request):
    where, params = build_purchase_filter(request)
    return execute_bar_query(
        f"SELECT po_type as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY po_type ORDER BY SUM(Total) DESC",
        params)


@app.get('/api/purchase/dashboard/vendors')
def get_purchase_vendors(request: Request):
    where, params = build_purchase_filter(request)
    return execute_bar_query(
        f"SELECT Vendor_Name as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY Vendor_Name ORDER BY SUM(Total) DESC LIMIT 10",
        params)


@app.get('/api/purchase/dashboard/materials')
def get_purchase_materials(request: Request):
    where, params = build_purchase_filter(request)
    return execute_bar_query(
        f"SELECT Short_Text as label, SUM(Total) as val FROM purchase_registry {where} GROUP BY Short_Text ORDER BY SUM(Total) DESC LIMIT 10",
        params)


@app.get('/api/purchase/dashboard/table')
def get_purchase_table(request: Request):
    where, params = build_purchase_filter(request)
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(f"SELECT SUM(Total) as val FROM purchase_registry {where}", params)
            total_spend = (cursor.fetchone() or {}).get('val') or 0

            cursor.execute(f"""
                SELECT ValA as plant, SUM(Total) as total, COUNT(DISTINCT po_number) as pos, COUNT(DISTINCT Vendor_Name) as vendors 
                FROM purchase_registry {where} GROUP BY ValA ORDER BY SUM(Total) DESC
            """, params)
            table_data = cursor.fetchall()

    for row in table_data:
        row['total'] = row['total'] or 0
        row['share'] = (row['total'] / total_spend * 100) if total_spend else 0
        row['avgpo'] = (row['total'] / row['pos']) if row['pos'] else 0

    return jsonify(table_data)

@app.get('/api/export/purchase_excel')
def export_purchase_excel(request: Request):
    """Exports the filtered purchase registry data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_purchase_filter(request)
    query = f"SELECT * FROM purchase_registry {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        # 2. Load dictionary into Pandas DataFrame
        df = pd.DataFrame(results)

        # 3. Clean up Date formats to prevent Excel serialization errors
        if 'Post_Date' in df.columns:
            df['Post_Date'] = pd.to_datetime(df['Post_Date'], errors='coerce').dt.strftime('%d-%m-%Y')

        # 4. Write to BytesIO stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Purchase_Export')
        output.seek(0)

        # 5. Generate dynamic filename
        fin_years = request.query_params.getlist('fin_year')
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
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


 ###--- collection Data ----#####

@app.get('/api/filters/collections')
def get_collections_filters(request: Request):
    """Fetches unique values for the Collections Dashboard filters, with cascading Unit -> Profit Center."""
    unit_selected = request.query_params.getlist('unit')

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Independent Filters
            cursor.execute(
                "SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data WHERE Posting_Date IS NOT NULL ORDER BY yr DESC")
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                "SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data WHERE Posting_Date IS NOT NULL")
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute("SELECT DISTINCT Unit_Code FROM collections_data WHERE Unit_Code IS NOT NULL")
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute("SELECT DISTINCT Coll_BR_DESC FROM collections_data WHERE Coll_BR_DESC IS NOT NULL")
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit)
            pc_query = "SELECT DISTINCT Profit_Centre FROM collections_data WHERE Profit_Centre IS NOT NULL"
            pc_params = []
            if unit_selected:
                placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({placeholders})"
                pc_params.extend(unit_selected)

            cursor.execute(pc_query, pc_params)
            profit_centers = [r['Profit_Centre'] for r in cursor.fetchall()]

    return jsonify({
        "years": years,
        "months": months,
        "units": units,
        "profit_centers": profit_centers,
        "sales_offices": sales_offices
    })


def build_collections_filter(request):
    """Builds SQL WHERE clause for Collections queries, supporting date ranges and units."""
    filters = []
    params = []

    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')  # NEW
    profit_centers = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7, "August": 8,
                     "September": 9, "October": 10, "November": 11, "December": 12}
        month_nums = [str(month_map[m]) for m in months if m in month_map]
        if month_nums:
            placeholders = ','.join(['%s'] * len(month_nums))
            filters.append(f"MONTH(Posting_Date) IN ({placeholders})")
            params.extend(month_nums)

    if exact_date:
        if " to " in exact_date:
            start_date, end_date = exact_date.split(" to ")
            filters.append("Posting_Date >= %s AND Posting_Date < DATE_ADD(%s, INTERVAL 1 DAY)")
            params.extend([start_date, end_date])
        else:
            filters.append("Posting_Date >= %s AND Posting_Date < DATE_ADD(%s, INTERVAL 1 DAY)")
            params.extend([exact_date, exact_date])

    if units:
        placeholders = ','.join(['%s'] * len(units))
        filters.append(f"Unit_Code IN ({placeholders})")
        params.extend(units)

    if profit_centers:
        placeholders = ','.join(['%s'] * len(profit_centers))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(profit_centers)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@app.get('/api/collections/dashboard')
def get_collections_dashboard_data(request: Request):
    where, params = build_collections_filter(request)

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. KPIs
            cursor.execute(f"""
                SELECT 
                    SUM(Cheque_Amt) as total_amt,
                    COUNT(Document_Number) as receipts,
                    COUNT(DISTINCT Unit_Code) as units,
                    COUNT(DISTINCT Coll_BR_DESC) as offices
                FROM collections_data {where}
            """, params)
            kpis = cursor.fetchone() or {}
            total_amt = kpis.get('total_amt') or 0

            # 2. Spend Trend (Fixed pymysql %d formatting error by escaping with %%)
            months = request.query_params.getlist('month')
            exact_date = request.query_params.get('date')

            if exact_date or months:
                trend_query = f"""
                    SELECT DATE_FORMAT(Posting_Date, '%%d %%b') as label, SUM(Cheque_Amt) as val 
                    FROM collections_data {where} 
                    GROUP BY DATE(Posting_Date) 
                    ORDER BY DATE(Posting_Date)
                """
            else:
                trend_query = f"""
                    SELECT MONTH(Posting_Date) as month_num, SUM(Cheque_Amt) as val 
                    FROM collections_data {where} 
                    GROUP BY MONTH(Posting_Date) 
                    ORDER BY MONTH(Posting_Date)
                """
            cursor.execute(trend_query, params)
            trend_data_raw = cursor.fetchall()

            trend_data = []
            if exact_date or months:
                trend_data = trend_data_raw
            else:
                month_names = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sep",
                               10: "Oct", 11: "Nov", 12: "Dec"}
                for r in trend_data_raw:
                    trend_data.append({"label": month_names.get(r['month_num'], "Unknown"), "val": r['val']})

            # 3. Categorical Splits
            queries = {
                "paytype": f"SELECT PAYMENT_TYPE as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY PAYMENT_TYPE ORDER BY val DESC",
                "doctype": f"SELECT Doc_Type as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Doc_Type ORDER BY val DESC",
                "unit": f"SELECT Unit_Code as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Unit_Code ORDER BY val DESC",
                "branch": f"SELECT Coll_BR_DESC as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Coll_BR_DESC ORDER BY val DESC",
                "prod": f"SELECT Prod_GRP_Code as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Prod_GRP_Code ORDER BY val DESC",
                "bank": f"SELECT Bank_Desc as label, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Bank_Desc ORDER BY val DESC",
                "payers": f"SELECT Cust_Name as label, COUNT(Document_Number) as txns, SUM(Cheque_Amt) as val FROM collections_data {where} GROUP BY Cust_Name ORDER BY val DESC LIMIT 10"
            }

            results = {}
            for key, q in queries.items():
                cursor.execute(q, params)
                results[key] = cursor.fetchall()

    return jsonify({
        "kpis": {
            "total_amt": total_amt,
            "receipts": kpis.get('receipts') or 0,
            "units": kpis.get('units') or 0,
            "offices": kpis.get('offices') or 0
        },
        "trend": {"labels": [r['label'] for r in trend_data], "values": [r['val'] or 0 for r in trend_data]},
        "paytype": {"labels": [r['label'] for r in results['paytype']],
                    "values": [r['val'] or 0 for r in results['paytype']]},
        "doctype": {"labels": [r['label'] for r in results['doctype']],
                    "values": [r['val'] or 0 for r in results['doctype']]},
        "unit": {"labels": [r['label'] for r in results['unit']], "values": [r['val'] or 0 for r in results['unit']]},
        "branch": {"labels": [r['label'] for r in results['branch']],
                   "values": [r['val'] or 0 for r in results['branch']]},
        "prod": {"labels": [r['label'] for r in results['prod']], "values": [r['val'] or 0 for r in results['prod']]},
        "bank": results['bank'],
        "payers": results['payers']
    })

@app.get('/api/export/collections_excel')
def export_collections_excel(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter(request)
    query = f"SELECT * FROM collections_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        # 2. Load dictionary into Pandas DataFrame
        df = pd.DataFrame(results)

        # 3. Clean up Date formats to prevent Excel serialization errors
        if 'Posting_Date' in df.columns:
            df['Posting_Date'] = pd.to_datetime(df['Posting_Date'], errors='coerce').dt.strftime('%d-%m-%Y')

        # 4. Write to BytesIO stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Collections_Export')
        output.seek(0)

        # 5. Generate dynamic filename based on the selected year
        years = request.query_params.getlist('year')
        yr_str = "_".join(years) if years else "All_Years"
        filename = f"Collections_Export_{yr_str}.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


@app.get('/ytd')
@requires_login
def ytd_dashboard(request: Request):
    return render_template(request,'ytd.html',
                           username=request.session['username'],
                           )


@app.get('/get_available_years')
def get_available_years():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT DISTINCT fin_year FROM billing_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    years = [r['fin_year'] for r in cursor.fetchall()]
    conn.close()
    return jsonify(years)


@app.get('/api/ytd/kpis')
def get_ytd_kpi_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    # Actual
    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS net FROM billing_data WHERE fin_year=%s AND Dist_Channel_TEXT != 'Inter Unit Transfer' ",
        (year,))
    ytd_sales = cursor.fetchone()['net'] or 0
    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS total FROM billing_data WHERE fin_year=%s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit!= 'UN07' ",
        (year,))
    actual = cursor.fetchone()['total'] or 0
    cursor.execute("SELECT SUM(Target) AS TARGET FROM billing_target where fin_year=%s", (year,))
    target = cursor.fetchone()['TARGET'] or 0
    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS value FROM billing_data WHERE fin_year=%s AND Dist_Channel_TEXT = 'Inter Unit Transfer' ",
        (year,))
    iut = cursor.fetchone()['value'] or 0
    cursor.execute("SELECT SUM(Net_Value)/100000 AS scrap FROM billing_data WHERE fin_year=%s AND Unit= 'UN07'",
                   (year,))
    scrap = cursor.fetchone()['scrap'] or 0
    prev_y_start = int(year.split('-')[0]) - 1
    prev_y_end = int(year.split('-')[0])
    prev_year_str = f"{prev_y_start}-{prev_y_end}"
    cursor.execute("SELECT SUM(Net_Value)/100000 AS total FROM Billing_data WHERE fin_year=%s ", (prev_year_str,))
    last_year_actual = cursor.fetchone()['total'] or 0
    conn.close()
    percent_val = (actual / target * 100) if target > 0 else 0
    return jsonify({
        "net_value": f"{float(ytd_sales):,.5f}",
        "actual": f"{float(actual):,.5f}L",
        "target": f"{float(target):,.0f}L",
        "iut": f"{float(iut): ,.5f}L",
        "scrap": f"{float(scrap):,.5f}L",
        "percent": percent_val,
        "raw_percent": percent_val,
        "last_year": f"{float(last_year_actual):,.5f}L"
    })


@app.get('/api/ytd/products')
def get_ytd_product_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_ORDER = [
        ('DC MACHINES', 'DCM'),
        ('HV MOTORS', 'HVM'),
        ('HV GENERATORS', 'HVG'),
        ('RRM', 'RRM'),
        ('LV MOTORS (NS)', 'LVM(NS)'),
        ('LV MOTORS (S)', 'LVM(S)'),
        ('EVM', 'EVM'),
        ('LV GENERATORS', 'LVG'),
        ('DG SETS', 'DGS'),
        ('TRANSFORMERS PUNE', 'TRF(P)'),
        ('TRANSFORMERS MYSORE', 'TRF(M)'),
        ('SWITCHGEAR MYSORE', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]

    DATABASE_VALID_PRODS = (
        'DC MACHINES', 'HV MOTORS', 'HV GENERATORS', 'RRM', 'LV MOTORS',
        'EVM', 'LV GENERATORS', 'DG SETS', 'TRANSFORMERS PUNE',
        'TRANSFORMERS MYSORE', 'SWITCHGEAR MYSORE', 'SPARES & SERVICE'
    )

    query = """
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                ELSE Prod
            END AS display_prod, 
            SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
          AND Prod IN %s
        GROUP BY display_prod
    """
    cursor.execute(query, (year, DATABASE_VALID_PRODS))
    rows = cursor.fetchall()

    target_query = "SELECT Product, SUM(Target) as t_val FROM billing_target WHERE fin_year = %s GROUP BY Product"
    cursor.execute(target_query, (year,))
    target_rows = cursor.fetchall()
    conn.close()

    results_dict = {r['display_prod']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['Product']: float(r['t_val'] or 0) for r in target_rows}

    final_labels = []
    final_actuals = []
    final_targets = []
    for full_name, abbr in PRODUCT_ORDER:
        final_labels.append(abbr)
        final_actuals.append(results_dict.get(full_name, 0))
        final_targets.append(target_dict.get(full_name, 0))

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/api/ytd/regions')
def get_ytd_region_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India', 'InternationalMarket')

    query = """
        SELECT 
            CASE 
                WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket'
                ELSE Sales_Region_Name
            END AS display_region, 
            SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY display_region
        HAVING display_region IN %s
    """
    cursor.execute(query, (year, valid_regions))
    rows = cursor.fetchall()

    target_query = """

               SELECT Sales_Region_Name AS region, SUM(Target) AS annual_target
               FROM billing_target
               WHERE Sales_Region_Name IN %s AND fin_year = %s
               GROUP BY Sales_Region_Name

    """
    cursor.execute(target_query, (valid_regions, year))
    target_rows = cursor.fetchall()

    conn.close()

    # Convert rows to dictionaries for easy lookup
    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['region']: float(r['annual_target'] or 0) for r in target_rows}

    region_labels = []
    region_actuals = []
    region_targets = []

    # Iterate through valid_regions to maintain specific order
    for region in valid_regions:
        region_labels.append(region)
        region_actuals.append(actual_dict.get(region, 0))
        region_targets.append(target_dict.get(region, 0))

    return jsonify({
        "labels": region_labels,
        "actuals": region_actuals,
        "targets": region_targets
    })


@app.get('/api/ytd/sector')
def get_ytd_sector_data(request: Request):
    year = request.query_params.get('fin_year')
    products = request.query_params.getlist('product')

    if not year:
        return jsonify({"error": "fin_year parameter is required"}, status_code=400)

    valid_sectors = (
        'Direct Export',
        'Domestic Direct',
        'Domestic Dealers',
        'Deemed Export',
        'SEZ'
    )

    query = """
        SELECT Dist_Channel_TEXT AS channel,
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
          AND Dist_Channel_TEXT IN %s
    """
    params = [year, valid_sectors]

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS (NS)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
            elif p == 'LV MOTORS (S)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
            elif p == 'TRANSFORMERS MYSORE':
                product_clauses.append(
                    "(Prod = 'TRANSFORMERS MYSORE' AND (`VTEXT1` != 'Switchgear' OR `VTEXT1` IS NULL))")
            else:
                product_clauses.append("Prod = %s")
                params.append(p)

        if product_clauses:
            query += " AND (" + " OR ".join(product_clauses) + ")"

    query += """
        GROUP BY Dist_Channel_TEXT
    """

    try:
        conn = get_db_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(query, params)
            rows = cursor.fetchall()

            processed_data = {}

            for row in rows:
                channel = row['channel']
                actual = float(row['actual'] or 0)

                # Merge SEZ + Domestic Direct
                if channel in ('SEZ', 'Domestic Direct'):
                    label = 'Domestic Direct'
                    processed_data[label] = processed_data.get(label, 0) + actual
                else:
                    processed_data[channel] = actual

            return jsonify({
                "labels": list(processed_data.keys()),
                "values": list(processed_data.values())
            })

        finally:
            cursor.close()
            conn.close()

    except Exception as e:
        return jsonify({
            "error": "Database error",
            "message": str(e)
        }, status_code=500)


@app.get('/api/ytd/Unit')
def get_ytd_Unit_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    VALID_PRODUCTS = (
        'DC MACHINES', 'DG SETS', 'EVM', 'HV GENERATORS', 'HV MOTORS',
        'LV GENERATORS', 'LV MOTORS', 'SPARES & SERVICE',
        'RRM', 'SWITCHGEAR MYSORE', 'TRANSFORMERS PUNE', 'TRANSFORMERS MYSORE'
    )

    query = """
        SELECT 
            CASE 
                WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                ELSE Unit 
            END AS display_unit, 
            SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE Prod IN %s AND fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_unit
    """

    cursor.execute(query, (VALID_PRODUCTS, year))

    rows = cursor.fetchall()
    conn.close()

    return jsonify({

        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })


@app.get('/get_ytd_name_data')
def get_ytd_name(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')

    conn = None
    cursor = None

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = """
            SELECT 
                Name,
                SUM(Net_Value) / 100000 AS total_value
            FROM billing_data
            WHERE fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        """
        params = [fin_year]
        if months:
            placeholders = ', '.join(['%s'] * len(months))
            query += f" AND MONTHNAME(Billing_Date) IN ({placeholders})"
            params.extend(months)

        query += """
            GROUP BY Name
            ORDER BY total_value DESC
            LIMIT 10
        """
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
        NAME_REPLACEMENTS = {
            "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
            "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
        }

        processed_labels = []
        for r in rows:
            original_name = r['Name'].strip() if r['Name'] else ""
            short_name = original_name
            for long_name, alias in NAME_REPLACEMENTS.items():
                if long_name in original_name.upper():
                    short_name = alias
                    break
            processed_labels.append(short_name)

        return jsonify({
            "labels": processed_labels,
            "values": [float(r['total_value'] or 0) for r in rows]
        })

    except Exception as e:
        print(f"Error fetching top customers: {str(e)}")
        return jsonify({"error": "Internal Server Error", "details": str(e)}, status_code=500)

    finally:

        if cursor:
            cursor.close()
        if conn:
            conn.close()

@app.get('/comparison')
def data_compare(request: Request):
    conn = get_db_connection()
    cursor = conn.cursor()
    # Fetch years dynamically for the filter checkboxes
    cursor.execute(
        "SELECT DISTINCT fin_year FROM billing_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    years = [row['fin_year'] for row in cursor.fetchall()]
    conn.close()
    return render_template(request,'compare.html', years=years, username=request.session.get('username'))


@app.get('/get_comparison_product_data')
def get_comparison_product_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    DATABASE_VALID_PRODS = (
        'DC MACHINES', 'HV MOTORS', 'HV GENERATORS', 'RRM', 'LV MOTORS',
        'EVM', 'LV GENERATORS', 'DG SETS', 'TRANSFORMERS PUNE',
        'TRANSFORMERS MYSORE', 'SWITCHGEAR MYSORE', 'SPARES & SERVICE'
    )

    PRODUCT_ORDER = [
        ('DC MACHINES', 'DCM'),
        ('HV MOTORS', 'HVM'),
        ('HV GENERATORS', 'HVG'),
        ('RRM', 'RRM'),
        ('LV MOTORS (NS)', 'LVM(NS)'),
        ('LV MOTORS (S)', 'LVM(S)'),
        ('EVM', 'EVM'),
        ('LV GENERATORS', 'LVG'),
        ('DG SETS', 'DGS'),
        ('TRANSFORMERS PUNE', 'TRF(P)'),
        ('TRANSFORMERS MYSORE', 'TRF(M)'),
        ('SWITCHGEAR MYSORE', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:
        clean_year = year_str.strip()
        parts = clean_year.split('-')

        if len(parts) == 2:
            start_yr_full = parts[0].strip()
            end_yr_full = parts[1].strip()
            start_yr = start_yr_full[-2:]
            end_yr = end_yr_full[-2:]
        else:

            start_yr = end_yr = "XX"

        for month in months:
            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP.get(month, month)} {suffix}"

            query = """
                SELECT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod
                    END AS display_prod, 
                    SUM(Net_Value)/100000 as actual
                FROM billing_data
                WHERE Prod IN %s AND fin_year = %s AND MONTHNAME(Billing_Date) = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
                GROUP BY display_prod
            """

            cursor.execute(query, (DATABASE_VALID_PRODS, year_str, month))
            rows = cursor.fetchall()

            month_data_map = {item[0]: 0 for item in PRODUCT_ORDER}
            for r in rows:
                month_data_map[r['display_prod']] = float(r['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [month_data_map[item[0]] for item in PRODUCT_ORDER]
            })

    conn.close()

    return jsonify({
        "labels": [item[1] for item in PRODUCT_ORDER],
        "datasets": datasets
    })


@app.get('/get_comparison_region_data')
def get_comparison_region_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India', 'IMD')

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:
        clean_year = year_str.strip()
        parts = clean_year.split('-')

        if len(parts) == 2:
            start_yr_full = parts[0].strip()
            end_yr_full = parts[1].strip()

            start_yr = start_yr_full[-2:]  # "24"
            end_yr = end_yr_full[-2:]  # "25"

        for month in months:
            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP[month]} {suffix}"

            # Modified Query with CASE logic for Sales_Office
            query = """
                SELECT 
                    CASE 
                        WHEN Sales_office = 'Bangalore - IMD' THEN 'IMD'
                        ELSE Sales_Region_Name 
                    END AS region, 
                    SUM(Net_Value)/100000 AS actual
                FROM billing_data
                WHERE fin_year = %s 
                  AND MONTHNAME(Billing_Date) = %s 
                  AND Dist_Channel_TEXT != 'Inter Unit Transfer'
                GROUP BY region
                HAVING region IN %s
            """

            cursor.execute(query, (year_str, month, valid_regions))
            rows = cursor.fetchall()

            region_data_map = {r: 0 for r in valid_regions}
            for row in rows:
                region_data_map[row['region']] = float(row['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [region_data_map[r] for r in valid_regions]
            })
    conn.close()
    return jsonify({
        "labels": valid_regions,
        "datasets": datasets
    })


@app.get('/get_comparison_branch_data')
def get_comparison_branch_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT Sales_office FROM billing_data WHERE Sales_office IS NOT NULL")
    valid_branches = [row['Sales_office'] for row in cursor.fetchall()]

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:

        clean_year = year_str.strip()
        parts = clean_year.split('-')
        start_yr = parts[0].strip()[-2:]  # '24'
        end_yr = parts[1].strip()[-2:]  # '25'

        for month in months:
            # April-Dec uses start year; Jan-Mar uses end year
            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP[month]} {suffix}"

            query = """
                SELECT Sales_office, SUM(Net_Value)/100000 AS actual
                FROM billing_data
                WHERE fin_year = %s AND MONTHNAME(Billing_Date) = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
                GROUP BY Sales_office
            """
            cursor.execute(query, (year_str, month))
            rows = cursor.fetchall()

            branch_data_map = {b: 0 for b in valid_branches}
            for row in rows:
                branch_data_map[row['Sales_office']] = float(row['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [branch_data_map[b] for b in valid_branches]
            })

    conn.close()
    return jsonify({
        "labels": valid_branches,
        "datasets": datasets
    })


@app.get('/data_entry')
@requires_login
def data_entry(request: Request):
    """Renders the main upload page."""
    conn = get_db_connection()
    cursor = conn.cursor()
    print("Scheduler is Running...!")
    conn.close()
    return render_template(request,'data_entry.html', username=request.session['username'])


def get_month_order():
    return {
        'April': 1, 'May': 2, 'June': 3, 'July': 4, 'August': 5, 'September': 6,
        'October': 7, 'November': 8, 'December': 9, 'January': 10, 'February': 11, 'March': 12
    }


@app.get('/order_mtd')
@requires_login
def order_mtd(request: Request):
    conn = get_db_connection()
    years = []
    if conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT fin_year FROM order_data ORDER BY fin_year DESC")
        years = [row['fin_year'] for row in cursor.fetchall()]
        conn.close()
    selected_year = request.query_params.get('fin_year', years[0] if years else "")
    return render_template(request,'order_mtd.html', years=years, selected_year=selected_year, username=request.session['username'])


@app.get('/get_orders_kpi')
def get_orders_kpi(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    # Base query for current year
    actual_sql = "SELECT SUM(Net_Value)/100000 AS total_val FROM order_data WHERE fin_year=%s"
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        actual_sql += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    cursor.execute(actual_sql, tuple(params))
    res = cursor.fetchone()
    actual_val = res['total_val'] if res and res['total_val'] else 0

    target_sql = "SELECT SUM(`Target`) AS total_target FROM order_target WHERE fin_year=%s"
    cursor.execute(target_sql, (fin_year,))
    t_res = cursor.fetchone()
    annual_target = t_res['total_target'] if t_res and t_res['total_target'] else 0

    num_months = len(months) if months else 12
    pro_rata_target = (annual_target / 12) * num_months

    try:
        # Last Month
        current_month_name = months[0]
        current_year_start = int(fin_year.split('-')[0])  # 2025 (from '2025-2026')
        current_year_end = int(fin_year.split('-')[1])  # 2026

        month_number = datetime.strptime(current_month_name, "%B").month

        if month_number == 4:

            last_month_num = 3
            lm_y_start = current_year_start - 1
            lm_y_end = current_year_start
        else:

            last_month_num = 12 if month_number == 1 else month_number - 1
            lm_y_start = current_year_start
            lm_y_end = current_year_end

        last_month_name = datetime(2000, last_month_num, 1).strftime('%B')
        last_month_fin_year = f"{lm_y_start}-{lm_y_end}"

        lm_sql = """
            SELECT SUM(`Net_Value`)/100000 AS total 
            FROM order_data 
            WHERE fin_year=%s 
            AND MONTHNAME(`Created_Date`) = %s
        """
        print(f" Current: {current_month_name} ({fin_year}) -> Last: {last_month_name} ({last_month_fin_year})")

        cursor.execute(lm_sql, (last_month_fin_year, last_month_name))
        lm_res = cursor.fetchone()
        last_month_actual = lm_res['total'] if lm_res and lm_res['total'] else 0

    except Exception as e:
        print(f"Error calculating last month: {e}")
        last_month_actual = 0

    conn.close()
    achievement_pct = (actual_val / pro_rata_target * 100) if pro_rata_target > 0 else 0
    return jsonify({
        "actual_value": f"{float(actual_val):,.5f}L",
        "target_value": f"{float(pro_rata_target):,.0f}L",
        "achievement": f"{achievement_pct:.1f}%",
        "last_year": f"{float(last_month_actual):,.5f}L"
    })


@app.get('/get_orders_product_data')
def get_orders_product_data(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('DC Machines', 'DCM'),
        ('HV Motors', 'HVM'),
        ('HV Generators', 'HVG'),
        ('RRM', 'RRM'),
        ('LV Motors', 'LVM(NS)'),
        ('LV MOTORS(S)', 'LVM(S)'),
        ('EVM', 'EVM'),
        ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'),
        ('Transformer Pune', 'TRF(P)'),
        ('Transformer Mysore', 'TRF(M)'),
        ('Switchgear', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]
    DB_Valid = ('DC Machines', 'HV Generators', 'HV Motors', 'RRM', 'EVM', 'LV Generators', 'LV Motors',
                'DG Sets', 'Switchgear', 'Transformer Mysore', 'Transformer Pune', 'SPARES & SERVICE'
                )

    actual_query = """
         SELECT 
              CASE 
                  WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                  ELSE Product1
              END AS display_product, 
              SUM(`Net_Value`)/100000 AS actual 
          FROM order_data 
          WHERE fin_year = %s AND Product1 IN %s
      """
    actual_params = [fin_year, DB_Valid]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        actual_query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        actual_params.extend(months)

    actual_query += " GROUP BY display_product"
    cursor.execute(actual_query, tuple(actual_params))
    actual_rows = cursor.fetchall()
    actual_dict = {r['display_product']: float(r['actual'] or 0) for r in actual_rows}

    target_query = """
          SELECT Product1, SUM(Target) as annual_target 
          FROM order_target 
          WHERE fin_year = %s 
          GROUP BY Product1
      """
    cursor.execute(target_query, (fin_year,))
    target_rows = cursor.fetchall()
    month_factor = len(months) / 12.0 if months else 1.0
    target_dict = {r['Product1']: float(r['annual_target'] or 0) * month_factor for r in target_rows}

    conn.close()

    final_labels = []
    final_actuals = []
    final_targets = []

    for full_name, abbr in product_abbr:

        if full_name in actual_dict or full_name in target_dict:
            final_labels.append(abbr)
            final_actuals.append(actual_dict.get(full_name, 0))
            final_targets.append(target_dict.get(full_name, 0))

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/get_orders_unit_data')
def get_orders_unit_data(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()
    VALID_PRODUCTS = (
        'DC Machines', 'DG Sets', 'EVM', 'HV Generators', 'HV Motors', 'LV Generators', 'LV Motors',
        'RRM', 'Switchgear', 'Transformer Pune', 'Transformer Mysore', 'SPARES & SERVICE'
    )
    query = """
           SELECT 
               CASE 
                   WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END AS display_unit, 
               SUM(Net_Value)/100000 as actual 
           FROM order_data
           WHERE Product1 IN %s AND fin_year = %s AND Unit!=''
       """
    params = [VALID_PRODUCTS, fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    query += """ 
            GROUP BY 
            CASE 
                WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                ELSE Unit 
            END 
    """
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['actual'] or 0) for r in rows]
    })


@app.get('/get_orders_region_data')
def get_orders_region_data(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()
    valid_regions = ('Central India', 'East India', 'InternationalMarket', 'North India', 'South India', 'West India')

    query = "SELECT `Sales_Region` as region, SUM(Net_Value)/100000 as actual FROM order_data WHERE fin_year = %s"
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)
    query += " GROUP BY `Sales_Region`"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    actual_dict = {r['region']: float(r['actual'] or 0) for r in rows}

    target_query = """
        SELECT Sales_Region AS region, SUM(Target) AS target
        FROM order_target 
        WHERE Sales_Region IN %s AND fin_year = %s
        GROUP BY Sales_Region
    """
    cursor.execute(target_query, (valid_regions, fin_year))
    target_rows = cursor.fetchall()
    num_months = len(months) if months else 12
    month_factor = num_months / 12.0

    target_dict = {r['region']: float(r['target'] or 0) * month_factor for r in target_rows}

    final_labels = []
    final_actuals = []
    final_targets = []

    for region in valid_regions:
        final_labels.append(region)
        final_actuals.append(actual_dict.get(region, 0))
        final_targets.append(target_dict.get(region, 0))

    conn.close()

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/get_orders_sector_data')
def get_orders_sector_data(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')
    products = request.query_params.getlist('product')

    if not fin_year:
        return jsonify({"error": "fin_year is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    # Base query for Sector Data
    query = """
        SELECT `Dist_Channel` AS channel, 
               SUM(`Net_Value`)/100000 as actual 
        FROM order_data 
        WHERE fin_year = %s
    """
    params = [fin_year]

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS(S)':
                product_clauses.append("(`Product1` = 'LV Motors' AND `Product2` = 'LV Motors (Standard)')")
            elif p == 'LV Motors':
                product_clauses.append("(`Product1` = 'LV Motors' AND `Product2` != 'LV Motors (Standard)')")
            else:
                product_clauses.append("`Product1` = %s")
                params.append(p)

        if product_clauses:
            query += " AND (" + " OR ".join(product_clauses) + ")"

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    query += " GROUP BY `Dist_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        conn.close()
        return jsonify({"error": str(e)}, status_code=500)

    conn.close()

    processed_data = {}
    for r in rows:
        channel = r['channel']
        actual = float(r['actual'] or 0)

        if channel == 'SZ' or channel == 'DD':
            label = 'DD'
            processed_data[label] = processed_data.get(label, 0) + actual
        else:
            processed_data[channel] = actual

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


@app.get('/get_order_name_data')
def get_order_name(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')

    conn = None
    cursor = None

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = """
            SELECT 
                Name,
                SUM(Net_Value) / 100000 AS total_value
            FROM order_data
            WHERE fin_year = %s
        """
        params = [fin_year]
        if months:
            placeholders = ', '.join(['%s'] * len(months))
            query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
            params.extend(months)

        query += """
            GROUP BY Name
            ORDER BY total_value DESC
            LIMIT 10
        """
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
        NAME_REPLACEMENTS = {
            "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
            "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
        }

        processed_labels = []
        for r in rows:
            original_name = r['Name'].strip() if r['Name'] else ""
            short_name = original_name
            for long_name, alias in NAME_REPLACEMENTS.items():
                if long_name in original_name.upper():
                    short_name = alias
                    break
            processed_labels.append(short_name)

        return jsonify({
            "labels": processed_labels,
            "values": [float(r['total_value'] or 0) for r in rows]
        })

    except Exception as e:
        print(f"Error fetching top customers: {str(e)}")
        return jsonify({"error": "Internal Server Error", "details": str(e)}, status_code=500)

    finally:

        if cursor:
            cursor.close()
        if conn:
            conn.close()


@app.get('/order_ytd')
@requires_login
def order_ytd_dashboard(request: Request):
    return render_template(request,'order_ytd.html',
                           username=request.session['username'],
                           )


@app.get('/get_order_available_years')
def get_order_available_years():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT DISTINCT fin_year FROM order_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    years = [r['fin_year'] for r in cursor.fetchall()]
    conn.close()
    return jsonify(years)


@app.get('/api/ytd/order_kpis')
def get_order_ytd_kpi_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    # Actual
    cursor.execute("SELECT SUM(Net_Value)/100000 AS total FROM order_data WHERE fin_year=%s", (year,))
    actual = cursor.fetchone()['total'] or 0

    target_sql = "SELECT SUM(`Target`) AS total_target FROM order_target WHERE fin_year=%s"
    cursor.execute(target_sql, (year,))
    t_res = cursor.fetchone()
    target = t_res['total_target'] if t_res and t_res['total_target'] else 0
    achievement_pct = (actual / target * 100) if target > 0 else 0

    prev_y_start = int(year.split('-')[0]) - 1
    prev_y_end = int(year.split('-')[0])
    prev_year_str = f"{prev_y_start}-{prev_y_end}"
    cursor.execute("SELECT SUM(Net_Value)/100000 AS total FROM order_data WHERE fin_year=%s",
                   (prev_year_str,))
    last_year_actual = cursor.fetchone()['total'] or 0

    conn.close()

    return jsonify({
        "actuals": f"{float(actual):,.5f}L",
        "last_year": f"{float(last_year_actual):,.5f}L",
        "targets": f"{float(target):,.0f}L",
        "achievement": f"{achievement_pct:.1f}%",
    })


@app.get('/api/ytd/order_products')
def get_order_ytd_product_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('DC Machines', 'DCM'),
        ('HV Motors', 'HVM'),
        ('HV Generators', 'HVG'),
        ('RRM', 'RRM'),
        ('LV Motors', 'LVM(NS)'),
        ('LV MOTORS(S)', 'LVM(S)'),
        ('EVM', 'EVM'),
        ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'),
        ('Transformer Pune', 'TRF(P)'),
        ('Transformer Mysore', 'TRF(M)'),
        ('Switchgear', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]
    DB_Valid = ('DC Machines', 'HV Generators', 'HV Motors', 'RRM', 'EVM', 'LV Generators', 'LV Motors',
                'DG Sets', 'Switchgear', 'Transformer Mysore', 'Transformer Pune', 'SPARES & SERVICE'
                )

    query = """
       SELECT 
            CASE 
                WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS display_product, 
            SUM(`Net_Value`)/100000 AS actual 
        FROM order_data 
        WHERE fin_year = %s AND Product1 IN %s
        GROUP BY display_product
    """

    cursor.execute(query, (year, DB_Valid))
    rows = cursor.fetchall()
    actual_dict = {r['display_product']: float(r['actual'] or 0) for r in rows}

    target_query = """
              SELECT Product1, SUM(Target) as annual_target 
              FROM order_target 
              WHERE fin_year = %s 
              GROUP BY Product1
          """
    cursor.execute(target_query, (year,))
    target_rows = cursor.fetchall()
    target_dict = {r['Product1']: float(r['annual_target'] or 0) for r in target_rows}
    conn.close()
    final_labels = []
    final_actuals = []
    final_targets = []

    for full_name, abbr in product_abbr:
        if full_name in actual_dict or full_name in target_dict:
            final_labels.append(abbr)
            final_actuals.append(actual_dict.get(full_name, 0))
            final_targets.append(target_dict.get(full_name, 0))

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/api/ytd/order_regions')
def get_order_ytd_region_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'InternationalMarket', 'North India', 'South India', 'West India')

    query = """
        SELECT s.Sales_Region AS region,
               SUM(s.Net_Value)/100000 AS actual
        FROM order_data s
        WHERE s.fin_year = %s AND s.Sales_Region IN %s
        GROUP BY s.Sales_Region
        ORDER BY s.Sales_Region
    """

    cursor.execute(query, (year, valid_regions))
    rows = cursor.fetchall()
    actual_dict = {r['region']: float(r['actual'] or 0) for r in rows}

    target_query = """
            SELECT Sales_Region AS region, SUM(Target) AS target
            FROM order_target 
            WHERE Sales_Region IN %s AND fin_year = %s
            GROUP BY Sales_Region
            ORDER BY Sales_Region
        """
    cursor.execute(target_query, (valid_regions, year))
    target_rows = cursor.fetchall()

    target_dict = {r['region']: float(r['target'] or 0) for r in target_rows}

    final_labels = []
    final_actuals = []
    final_targets = []

    for region in valid_regions:
        final_labels.append(region)
        final_actuals.append(actual_dict.get(region, 0))
        final_targets.append(target_dict.get(region, 0))

    conn.close()

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/api/ytd/order_sector')
def get_order_ytd_sector_data(request: Request):
    year = request.query_params.get('fin_year')
    products = request.query_params.getlist('product')

    if not year:
        return jsonify({"error": "fin_year is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
            SELECT `Dist_Channel` AS channel, 
                   SUM(`Net_Value`)/100000 as actual 
            FROM order_data 
            WHERE fin_year = %s
        """
    params = [year]

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS(S)':
                product_clauses.append("(`Product1` = 'LV Motors' AND `Product2` = 'LV Motors (Standard)')")
            elif p == 'LV Motors':
                product_clauses.append("(`Product1` = 'LV Motors' AND `Product2` != 'LV Motors (Standard)')")
            else:
                product_clauses.append("`Product1` = %s")
                params.append(p)

        if product_clauses:
            query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY `Dist_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        conn.close()
        return jsonify({"error": str(e)}, status_code=500)

    conn.close()

    processed_data = {}
    for r in rows:
        channel = r['channel']
        actual = float(r['actual'] or 0)

        if channel == 'SZ' or channel == 'DD':
            label = 'DD'
            processed_data[label] = processed_data.get(label, 0) + actual
        else:
            processed_data[channel] = actual

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


@app.get('/api/ytd/order_Unit')
def get_order_ytd_Unit_data(request: Request):
    year = request.query_params.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    VALID_PRODUCTS = (
        'DC Machines', 'DG Sets', 'EVM', 'HV Generators', 'HV Motors',
        'LV Generators', 'LV Motors', 'RRM', 'Switchgear',
        'Transformer Pune', 'Transformer Mysore', 'SPARES & SERVICE', ' '
    )

    query = """
        SELECT 
            CASE 
                WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                ELSE Unit 
            END AS display_unit, 
            SUM(Net_Value)/100000 as total 
        FROM order_data
        WHERE Product1 IN %s AND fin_year = %s  AND Unit !='UN07'
        GROUP BY display_unit
    """

    cursor.execute(query, (VALID_PRODUCTS, year))

    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })


@app.get('/get_order_ytd_name_data')
def get_order_ytd__name(request: Request):
    months = request.query_params.getlist('month')
    fin_year = request.query_params.get('fin_year')

    conn = None
    cursor = None

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        query = """
            SELECT 
                Name,
                SUM(Net_Value) / 100000 AS total_value
            FROM order_data
            WHERE fin_year = %s
        """
        params = [fin_year]
        if months:
            placeholders = ', '.join(['%s'] * len(months))
            query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
            params.extend(months)

        query += """
            GROUP BY Name
            ORDER BY total_value DESC
            LIMIT 10
        """
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
        NAME_REPLACEMENTS = {
            "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
            "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
        }

        processed_labels = []
        for r in rows:
            original_name = r['Name'].strip() if r['Name'] else ""
            short_name = original_name
            for long_name, alias in NAME_REPLACEMENTS.items():
                if long_name in original_name.upper():
                    short_name = alias
                    break
            processed_labels.append(short_name)

        return jsonify({
            "labels": processed_labels,
            "values": [float(r['total_value'] or 0) for r in rows]
        })

    except Exception as e:
        print(f"Error fetching top customers: {str(e)}")
        return jsonify({"error": "Internal Server Error", "details": str(e)}, status_code=500)

    finally:

        if cursor:
            cursor.close()
        if conn:
            conn.close()



@app.get('/order_comparison')
def order_data_compare(request: Request):
    conn = get_db_connection()
    cursor = conn.cursor()
    # Fetch years dynamically for the filter checkboxes
    cursor.execute(
        "SELECT DISTINCT fin_year FROM order_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    years = [row['fin_year'] for row in cursor.fetchall()]
    conn.close()
    return render_template(request,'order_compare.html', years=years, username=request.session.get('username'))


@app.get('/get_comparison_order_product_data')
def get_comparison_order_product_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_MAP = {
        'DC Machines': 'DCM',
        'HV Motors': 'HVM',
        'HV Generators': 'HVG',
        'RRM': 'RRM',
        'LV Motors': 'LVM(NS)',
        'LV MOTORS(S)': 'LVM(S)',
        'EVM': 'EVM',
        'LV Generators': 'LVG',
        'DG Sets': 'DGS',
        'Transformer Pune': 'TRF(P)',
        'Transformer Mysore': 'TRF(M)',
        'Switchgear': 'SWG',
        'SPARES & SERVICE': 'SP&S'
    }

    DB_Valid = ('DC Machines', 'HV Generators', 'HV Motors', 'RRM', 'EVM', 'LV Generators', 'LV Motors',
                'DG Sets', 'Switchgear', 'Transformer Mysore', 'Transformer Pune', 'SPARES & SERVICE')

    ordered_display_products = [
        'DC Machines', 'HV Motors', 'HV Generators', 'RRM', 'LV Motors',
        'LV MOTORS(S)', 'EVM', 'LV Generators', 'DG Sets',
        'Transformer Pune', 'Transformer Mysore', 'Switchgear', 'SPARES & SERVICE'
    ]

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:
        clean_year = year_str.strip()
        parts = clean_year.split('-')

        start_yr = parts[0].strip()[-2:] if len(parts) == 2 else ""
        end_yr = parts[1].strip()[-2:] if len(parts) == 2 else ""

        for month in months:
            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP[month]} {suffix}"

            query = """
                SELECT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1
                    END AS display_product, 
                    SUM(Net_Value)/100000 AS actual 
                FROM order_data
                WHERE Product1 IN %s 
                  AND fin_year = %s 
                  AND MONTHNAME(Created_Date) = %s
                GROUP BY display_product
            """
            cursor.execute(query, (DB_Valid, year_str, month))
            rows = cursor.fetchall()

            month_data_map = {p: 0 for p in ordered_display_products}
            for r in rows:
                p_name = r['display_product']
                if p_name in month_data_map:
                    month_data_map[p_name] = float(r['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [month_data_map[p] for p in ordered_display_products]
            })

    conn.close()
    return jsonify({
        "labels": [PRODUCT_MAP[p] for p in ordered_display_products],
        "datasets": datasets
    })


@app.get('/get_comparison_order_region_data')
def get_comparison_order_region_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India', 'InternationalMarket')

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:
        clean_year = year_str.strip()
        parts = clean_year.split('-')

        if len(parts) == 2:
            start_yr_full = parts[0].strip()
            end_yr_full = parts[1].strip()

            start_yr = start_yr_full[-2:]  # "24"
            end_yr = end_yr_full[-2:]  # "25"

        for month in months:
            # April-Dec uses start year; Jan-Mar uses end year
            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP[month]} {suffix}"

            query = """
                SELECT Sales_Region AS region, SUM(Net_Value)/100000 AS actual
                FROM order_data
                WHERE Sales_Region IN %s AND fin_year = %s AND MONTHNAME(Created_Date) = %s
                GROUP BY Sales_Region
            """
            cursor.execute(query, (valid_regions, year_str, month))
            rows = cursor.fetchall()

            # Map results to regions in order
            region_data_map = {r: 0 for r in valid_regions}
            for row in rows:
                region_data_map[row['region']] = float(row['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [region_data_map[r] for r in valid_regions]
            })

    conn.close()
    return jsonify({
        "labels": valid_regions,
        "datasets": datasets
    })


@app.get('/get_comparison_order_branch_data')
def get_comparison_order_branch_data(request: Request):
    years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT `Sales_office` FROM order_data WHERE `Sales_office` IS NOT NULL")
    valid_branches = [row['Sales_office'] for row in cursor.fetchall()]

    MONTH_MAP = {
        'April': 'Apr', 'May': 'May', 'June': 'Jun', 'July': 'Jul',
        'August': 'Aug', 'September': 'Sep', 'October': 'Oct',
        'November': 'Nov', 'December': 'Dec', 'January': 'Jan',
        'February': 'Feb', 'March': 'Mar'
    }

    datasets = []

    for year_str in years:

        clean_year = year_str.strip()
        parts = clean_year.split('-')
        start_yr = parts[0].strip()[-2:]  # '24'
        end_yr = parts[1].strip()[-2:]  # '25'

        for month in months:

            suffix = end_yr if month in ['January', 'February', 'March'] else start_yr
            label = f"{MONTH_MAP[month]} {suffix}"

            query = """
                SELECT `Sales_office`, SUM(Net_Value)/100000 AS actual
                FROM order_data
                WHERE fin_year = %s AND MONTHNAME(Created_Date) = %s
                GROUP BY `Sales_office`
            """
            cursor.execute(query, (year_str, month))
            rows = cursor.fetchall()

            branch_data_map = {b: 0 for b in valid_branches}
            for row in rows:
                branch_data_map[row['Sales_office']] = float(row['actual'] or 0)

            datasets.append({
                "label": label,
                "data": [branch_data_map[b] for b in valid_branches]
            })

    conn.close()
    return jsonify({
        "labels": valid_branches,
        "datasets": datasets
    })


import io
import pandas as pd
from flask import send_file, request
import openpyxl
from openpyxl.styles import Font, Alignment


@app.get('/download_order_comparison_excel')
def download_order_comparison_excel(request: Request):
    years = request.query_params.getlist('fin_year')  # e.g., ['2024-25', '2025-26', '2026-27']
    months = request.query_params.getlist('month')

    if not years or not months:
        return HTMLResponse("Please select at least one year and one month", status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_DISPLAY_MAP = {
        'DC Machines': 'DC Machines',
        'HV Motors': 'HV Motors',
        'HV Generators': 'HV Generators',
        'RRM': 'RRM',
        'LV Motors': 'LV Motors (Non-Std)',
        'LV MOTORS(S)': 'LV Motors (Std)',
        'EVM': 'EVM',
        'LV Generators': 'LV Generators',
        'DG Sets': 'DG Sets ',
        'Transformer Pune': 'Transformer (Pune)',
        'Transformer Mysore': 'Transformer (Mysore)',
        'Switchgear': 'Switchgear',
        'SPARES & SERVICE': 'Spares & Service'
    }

    PRODUCT_UNITS = {
        'DC Machines': 'Unit-01',
        'HV Motors': 'Unit-01',
        'HV Generators': 'Unit-01',
        'RRM': 'Unit-01',
        'LV Motors': 'Unit-02',
        'LV MOTORS(S)': 'Unit-25',
        'EVM': 'Unit-25',
        'LV Generators': 'Unit-25',
        'DG Sets': 'Unit-06',
        'Transformer Pune': 'Unit-16',
        'Transformer Mysore': 'Unit-05',
        'Switchgear': 'Unit-10',
        'SPARES & SERVICE': 'Unit-20'
    }

    DB_Valid = ('DC Machines', 'HV Generators', 'HV Motors', 'RRM', 'EVM', 'LV Generators', 'LV Motors',
                'DG Sets', 'Switchgear', 'Transformer Mysore', 'Transformer Pune', 'SPARES & SERVICE')

    ordered_display_products = [
        'DC Machines', 'HV Motors', 'RRM', 'HV Generators', 'LV Motors',
        'Transformer Mysore', 'DG Sets', 'Switchgear', 'Transformer Pune',
        'LV MOTORS(S)', 'LV Generators', 'EVM', 'SPARES & SERVICE'
    ]

    MONTH_MAP = {
        'April': 'APR', 'May': 'MAY', 'June': 'JUN', 'July': 'JUL',
        'August': 'AUG', 'September': 'SEP', 'October': 'OCT',
        'November': 'NOV', 'December': 'DEC', 'January': 'JAN',
        'February': 'FEB', 'March': 'MAR'
    }

    cursor.execute(
        "SELECT DISTINCT Sales_office FROM order_data WHERE Sales_office IS NOT NULL AND Sales_office != '' ORDER BY Sales_office")
    sales_offices = [r['Sales_office'] for r in cursor.fetchall()]

    query = """
        SELECT 
            Sales_office,
            CASE 
                WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS display_product,
            fin_year,
            MONTHNAME(Created_Date) AS month_name,
            SUM(Net_Value)/100000 AS actual 
        FROM order_data
        WHERE Product1 IN %s 
          AND fin_year IN %s 
          AND MONTHNAME(Created_Date) IN %s
        GROUP BY Sales_office, display_product, fin_year, month_name
    """
    cursor.execute(query, (DB_Valid, tuple(years), tuple(months)))
    rows = cursor.fetchall()
    conn.close()

    data_matrix = {
        off: {p: {(y, m): 0.0 for y in years for m in months} for p in ordered_display_products}
        for off in sales_offices
    }

    for r in rows:
        off = r['Sales_office']
        prod = r['display_product']
        y_val = r['fin_year']
        m_val = r['month_name']
        if off in data_matrix and prod in data_matrix[off] and (y_val, m_val) in data_matrix[off][prod]:
            data_matrix[off][prod][(y_val, m_val)] = float(r['actual'] or 0)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name='Calibri', size=11, bold=True)
    font_header = Font(name='Calibri', size=10, bold=True)
    font_total = Font(name='Calibri', size=10, bold=True, italic=True)
    font_grand = Font(name='Calibri', size=10, bold=True)
    font_regular = Font(name='Calibri', size=10)

    years_title = " vs ".join(years)
    ws.cell(row=1, column=1, value=f"Order Booking Comparision {years_title} Sales Office Wise").font = font_title
    ws.cell(row=2, column=1, value="All Values in ₹ Lacs").font = font_title

    ws.cell(row=5, column=1, value="Sales Office").font = font_header
    ws.cell(row=5, column=1).alignment = Alignment(vertical='center')

    current_col = 2
    num_fmt = "#,##0"
    pct_fmt = "0.0%"

    sorted_years = sorted(years)
    num_years = len(sorted_years)
    num_periods = num_years - 1
    has_growth = num_years > 1
    use_cagr = num_years > 2

    growth_header_label = "GROWTH %" if use_cagr else "YOY GROWTH"

    def get_suffix(year_str, month_str):
        parts = year_str.strip().split('-')
        start_yr = parts[0][-2:] if len(parts) >= 1 else ""
        end_yr = parts[1][-2:] if len(parts) == 2 else start_yr
        return end_yr if month_str in ['January', 'February', 'March'] else start_yr

    ordered_years_desc = sorted_years[::-1]  # e.g., ['2026-27', '2025-26', '2024-25']

    for p_db in ordered_display_products:
        prod_start_col = current_col
        p_display = PRODUCT_DISPLAY_MAP.get(p_db, p_db)
        p_unit = PRODUCT_UNITS.get(p_db, None)

        if p_unit:
            ws.cell(row=3, column=current_col, value=p_unit).font = font_header
            ws.cell(row=3, column=current_col).alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=4, column=current_col, value=p_display).font = font_header
        ws.cell(row=4, column=current_col).alignment = Alignment(horizontal='center', vertical='center')

        for m in months:
            m_code = MONTH_MAP.get(m, m[:3].upper())
            for y_str in ordered_years_desc:
                suf = get_suffix(y_str, m)
                c_cell = ws.cell(row=5, column=current_col, value=f"{m_code} {suf}")
                c_cell.font = font_header
                c_cell.alignment = Alignment(horizontal='center', vertical='center')
                current_col += 1

            if has_growth:
                g_cell = ws.cell(row=5, column=current_col, value=growth_header_label)
                g_cell.font = font_header
                g_cell.alignment = Alignment(horizontal='center', vertical='center')
                current_col += 1

        for y_str in ordered_years_desc:
            tot_cell = ws.cell(row=5, column=current_col, value=f"TOTAL {get_suffix(y_str, months[0])}")
            tot_cell.font = font_header
            tot_cell.alignment = Alignment(horizontal='center', vertical='center')
            current_col += 1

        if has_growth:
            tot_g_label = "TOTAL GROWTH" if use_cagr else "TOTAL YOY"
            tot_g_cell = ws.cell(row=5, column=current_col, value=tot_g_label)
            tot_g_cell.font = font_header
            tot_g_cell.alignment = Alignment(horizontal='center', vertical='center')
            current_col += 1

        prod_end_col = current_col - 1

        ws.merge_cells(start_row=4, start_column=prod_start_col, end_row=4, end_column=prod_end_col)
        if p_unit:
            ws.merge_cells(start_row=3, start_column=prod_start_col, end_row=3, end_column=prod_end_col)

        current_col += 1

    for off_idx, off in enumerate(sales_offices):
        row_num = 6 + off_idx
        ws.cell(row=row_num, column=1, value=off).font = font_regular

        c = 2
        for p_db in ordered_display_products:
            year_month_col_map = {y: [] for y in ordered_years_desc}

            for m in months:
                end_col_let = None
                start_col_let = None

                for y_idx, y_str in enumerate(ordered_years_desc):
                    val = int(round(data_matrix[off][p_db].get((y_str, m), 0.0)))
                    cell = ws.cell(row=row_num, column=c, value=val)
                    cell.font = font_regular
                    cell.number_format = num_fmt
                    cell.alignment = Alignment(horizontal='right')

                    col_let = openpyxl.utils.get_column_letter(c)
                    year_month_col_map[y_str].append(col_let)

                    if y_idx == 0:
                        end_col_let = col_let
                    if y_idx == len(ordered_years_desc) - 1:
                        start_col_let = col_let

                    c += 1

                if has_growth:
                    if use_cagr:
                        growth_formula = f"=IF(OR({start_col_let}{row_num}<=0, {end_col_let}{row_num}<=0), 0, ({end_col_let}{row_num}/{start_col_let}{row_num})^(1/{num_periods}) - 1)"
                    else:
                        growth_formula = f"=IF({start_col_let}{row_num}=0, 0, ({end_col_let}{row_num}-{start_col_let}{row_num})/{start_col_let}{row_num})"

                    cell_g = ws.cell(row=row_num, column=c, value=growth_formula)
                    cell_g.font = font_regular
                    cell_g.number_format = pct_fmt
                    cell_g.alignment = Alignment(horizontal='right')
                    c += 1

            tot_end_let = None
            tot_start_let = None

            for y_idx, y_str in enumerate(ordered_years_desc):
                m_cols = year_month_col_map[y_str]
                # EXPLICIT CELL SUMMATION (e.g., =SUM(B6, F6, J6)) TO PREVENT SUMMING INTERVENING COLUMNS
                tot_formula = f"=SUM({', '.join([f'{col}{row_num}' for col in m_cols])})"
                tot_cell = ws.cell(row=row_num, column=c, value=tot_formula)
                tot_cell.font = font_total
                tot_cell.number_format = num_fmt
                tot_cell.alignment = Alignment(horizontal='right')

                tot_col_let = openpyxl.utils.get_column_letter(c)
                if y_idx == 0:
                    tot_end_let = tot_col_let
                if y_idx == len(ordered_years_desc) - 1:
                    tot_start_let = tot_col_let

                c += 1

            if has_growth:
                if use_cagr:
                    tot_growth_formula = f"=IF(OR({tot_start_let}{row_num}<=0, {tot_end_let}{row_num}<=0), 0, ({tot_end_let}{row_num}/{tot_start_let}{row_num})^(1/{num_periods}) - 1)"
                else:
                    tot_growth_formula = f"=IF({tot_start_let}{row_num}=0, 0, ({tot_end_let}{row_num}-{tot_start_let}{row_num})/{tot_start_let}{row_num})"

                tot_g_cell = ws.cell(row=row_num, column=c, value=tot_growth_formula)
                tot_g_cell.font = font_total
                tot_g_cell.number_format = pct_fmt
                tot_g_cell.alignment = Alignment(horizontal='right')
                c += 1

            c += 1

    gt_row = 6 + len(sales_offices)
    ws.cell(row=gt_row, column=1, value="Grand Total").font = font_grand

    c = 2
    for p_db in ordered_display_products:
        for m in months:
            end_col_let = None
            start_col_let = None

            for y_idx, y_str in enumerate(ordered_years_desc):
                col_let = openpyxl.utils.get_column_letter(c)
                gt_cell = ws.cell(row=gt_row, column=c, value=f"=SUM({col_let}6:{col_let}{gt_row - 1})")
                gt_cell.font = font_grand
                gt_cell.number_format = num_fmt
                gt_cell.alignment = Alignment(horizontal='right')

                if y_idx == 0:
                    end_col_let = col_let
                if y_idx == len(ordered_years_desc) - 1:
                    start_col_let = col_let

                c += 1

            if has_growth:
                if use_cagr:
                    gt_growth_formula = f"=IF(OR({start_col_let}{gt_row}<=0, {end_col_let}{gt_row}<=0), 0, ({end_col_let}{gt_row}/{start_col_let}{gt_row})^(1/{num_periods}) - 1)"
                else:
                    gt_growth_formula = f"=IF({start_col_let}{gt_row}=0, 0, ({end_col_let}{gt_row}-{start_col_let}{gt_row})/{start_col_let}{gt_row})"

                gt_g_cell = ws.cell(row=gt_row, column=c, value=gt_growth_formula)
                gt_g_cell.font = font_grand
                gt_g_cell.number_format = pct_fmt
                gt_g_cell.alignment = Alignment(horizontal='right')
                c += 1

        tot_end_let = None
        tot_start_let = None

        for y_idx, y_str in enumerate(ordered_years_desc):
            tot_col_let = openpyxl.utils.get_column_letter(c)
            gt_tot_cell = ws.cell(row=gt_row, column=c, value=f"=SUM({tot_col_let}6:{tot_col_let}{gt_row - 1})")
            gt_tot_cell.font = font_grand
            gt_tot_cell.number_format = num_fmt
            gt_tot_cell.alignment = Alignment(horizontal='right')

            if y_idx == 0:
                tot_end_let = tot_col_let
            if y_idx == len(ordered_years_desc) - 1:
                tot_start_let = tot_col_let

            c += 1

        if has_growth:
            if use_cagr:
                gt_tot_growth_formula = f"=IF(OR({tot_start_let}{gt_row}<=0, {tot_end_let}{gt_row}<=0), 0, ({tot_end_let}{gt_row}/{tot_start_let}{gt_row})^(1/{num_periods}) - 1)"
            else:
                gt_tot_growth_formula = f"=IF({tot_start_let}{gt_row}=0, 0, ({tot_end_let}{gt_row}-{tot_start_let}{gt_row})/{tot_start_let}{gt_row})"

            gt_tot_g_cell = ws.cell(row=gt_row, column=c, value=gt_tot_growth_formula)
            gt_tot_g_cell.font = font_grand
            gt_tot_g_cell.number_format = pct_fmt
            gt_tot_g_cell.alignment = Alignment(horizontal='right')
            c += 1

        c += 1

    ws.column_dimensions['A'].width = 22
    for col_idx in range(2, ws.max_column + 1):
        col_let = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_let].width = 12

    ws.freeze_panes = 'B6'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        download_name='Order_Comparison_Sales_Office_Wise.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.get('/export_ytd_excel')
def export_ytd_excel(request: Request):
    year = request.query_params.get('fin_year', '2024-25')
    engine_uri = f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}/{DB_CONFIG['database']}"
    engine = create_engine(engine_uri)

    try:
        query = text(
            "SELECT Unit, Plant, Prod, Billing_Qty, Net_Value, Billing_Date FROM billing_data WHERE fin_year = :yr")
        df_raw = pd.read_sql(query, engine, params={'yr': year})

        if df_raw.empty:
            return HTMLResponse("No data found.", status_code=404)

        df_raw['Net_Value'] = df_raw['Net_Value'] / 100000

        # Process Dates
        df_raw['Billing_Date'] = pd.to_datetime(df_raw['Billing_Date'])
        df_raw['Month_Label'] = df_raw['Billing_Date'].dt.strftime('%b-%y')
        df_raw['Month_Sort'] = df_raw['Billing_Date'].dt.month.map(lambda x: x - 3 if x >= 4 else x + 9)

        # 2. Pivot Table
        pivot_df = df_raw.pivot_table(
            index=['Unit', 'Plant', 'Prod'],
            columns=['Month_Sort', 'Month_Label'],
            values=['Billing_Qty', 'Net_Value'],
            aggfunc='sum',
            fill_value=0
        )
        pivot_df = pivot_df.reorder_levels([1, 2, 0], axis=1).sort_index(axis=1)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Summary")

            header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D9D9D9', 'border': 1})
            subtotal_fmt = workbook.add_format(
                {'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'num_format': '#,##0.00'})
            num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            unit_cell_fmt = workbook.add_format({'bg_color': '#FCE4D6', 'border': 1})
            head_fmt = workbook.add_format({'bold': True, 'font_size': 18, 'font_color': '#3a5b94'})

            worksheet.write('H1', 'BILLING SUMMARY.', head_fmt)
            worksheet.write(2, 0, 'Unit', header_fmt)
            worksheet.write(2, 1, 'Plant', header_fmt)
            worksheet.write(2, 2, 'Product', header_fmt)

            curr_col = 3
            unique_months = df_raw.sort_values('Month_Sort')['Month_Label'].unique()
            for month in unique_months:
                worksheet.merge_range(2, curr_col, 2, curr_col + 1, month, header_fmt)
                worksheet.write(3, curr_col, 'Qty', header_fmt)
                worksheet.write(3, curr_col + 1, 'Val(In Lakhs)', header_fmt)
                curr_col += 2

            # Total Year Header
            worksheet.merge_range(2, curr_col, 2, curr_col + 1, 'Total Year', header_fmt)
            worksheet.write(3, curr_col, 'Qty', header_fmt)
            worksheet.write(3, curr_col + 1, 'Val(In Lakhs)', header_fmt)

            row_idx = 4
            units = pivot_df.index.get_level_values('Unit').unique()
            grand_totals = [0] * (len(pivot_df.columns) + 2)

            for unit in units:
                unit_df = pivot_df.xs(unit, level='Unit')

                # Product rows for this Unit
                for (plant, prod), row in unit_df.iterrows():
                    worksheet.write(row_idx, 0, unit, unit_cell_fmt)
                    worksheet.write(row_idx, 1, plant, num_fmt)
                    worksheet.write(row_idx, 2, prod, num_fmt)

                    col_idx = 3
                    r_qty, r_val = 0, 0
                    for i in range(len(row)):
                        val = row.iloc[i]
                        worksheet.write(row_idx, col_idx, val, num_fmt)
                        grand_totals[col_idx - 3] += val
                        if i % 2 == 0:
                            r_qty += val
                        else:
                            r_val += val
                        col_idx += 1

                    # Row Totals
                    worksheet.write(row_idx, col_idx, r_qty, num_fmt)
                    worksheet.write(row_idx, col_idx + 1, r_val, num_fmt)
                    grand_totals[col_idx - 3] += r_qty
                    grand_totals[col_idx - 2] += r_val
                    row_idx += 1

                unit_sum = unit_df.sum()
                worksheet.write(row_idx, 0, f"{unit} Total", subtotal_fmt)
                worksheet.write(row_idx, 1, "", subtotal_fmt)
                worksheet.write(row_idx, 2, "", subtotal_fmt)

                col_idx = 3
                u_qty_total, u_val_total = 0, 0
                for i in range(len(unit_sum)):
                    val = unit_sum.iloc[i]
                    worksheet.write(row_idx, col_idx, val, subtotal_fmt)
                    if i % 2 == 0:
                        u_qty_total += val
                    else:
                        u_val_total += val
                    col_idx += 1

                # Subtotal Year Totals
                worksheet.write(row_idx, col_idx, u_qty_total, subtotal_fmt)
                worksheet.write(row_idx, col_idx + 1, u_val_total, subtotal_fmt)
                row_idx += 1

            # Grand Total Row
            worksheet.write(row_idx, 0, "Grand Total", header_fmt)
            for i, total_val in enumerate(grand_totals):
                worksheet.write(row_idx, i + 3, total_val, header_fmt)

            worksheet.set_column('A:C', 20)
            worksheet.set_column(3, curr_col + 1, 12)

        output.seek(0)
        return send_file(output, download_name=f"KEC_Sales_Report_{year}.xlsx", as_attachment=True)

    except Exception as e:
        print(f"Error: {e}")
        return HTMLResponse(f"Internal Error: {str(e)}", status_code=500)
    finally:
        engine.dispose()


# Daily Tracker Logic.
@app.get('/daily')
@requires_login
def daily_dashboard(request: Request):
    today_date = datetime.now().strftime('%Y-%m-%d')
    selected_date = request.query_params.get('date', today_date)
    return render_template(request,'daily_sales.html', selected_date=selected_date, username=request.session['username'])


@app.get('/get_daily_summary')
def get_daily_summary(request: Request):
    # Frontend passes selected date from calendar (e.g., '2026-02-24')
    selected_date_str = request.query_params.get('date')
    if not selected_date_str:
        return jsonify({"error": "No date selected"}, status_code=400)

    selected_date = datetime.strptime(selected_date_str, '%Y-%m-%d')
    day_of_month = selected_date.day

    # Financial Year Logic
    month = selected_date.month
    year = selected_date.year
    fin_year = f"{year}-{year + 1}" if month >= 4 else f"{year - 1}-{year}"

    # Month-to-Date range
    mtd_start = selected_date.replace(day=1).strftime('%Y-%m-%d')

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # 1. Selected Day Actual
        cursor.execute("""
                    SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
                    WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
                """, (selected_date_str,))
        today_net = cursor.fetchone()['total'] or 0

        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
            WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        """, (selected_date_str,))
        today_actual = cursor.fetchone()['total'] or 0

        # 2. MTD Actual (Cumulative from 1st to selected date)
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
            WHERE Billing_Date >= %s AND Billing_Date <= %s 
        """, (mtd_start, selected_date_str))
        mtd_actual = cursor.fetchone()['total'] or 0

        # 3. Targets (Annual divided by 365)
        cursor.execute("SELECT SUM(Target) as annual FROM billing_target WHERE fin_year=%s", (fin_year,))
        annual_target = cursor.fetchone()['annual'] or 0

        daily_target = annual_target / 365
        mtd_target = daily_target * day_of_month
        achievement = (today_actual / daily_target * 100) if daily_target > 0 else 0

        # 4. iut
        cursor.execute("""
                    SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
                    WHERE Billing_Date = %s AND Dist_Channel_TEXT = 'Inter Unit Transfer'
                """, (selected_date_str,))
        today_iut = cursor.fetchone()['total'] or 0

        # 5. Scrap
        cursor.execute("""
                    SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
                    WHERE Billing_Date = %s AND Unit = 'UN07'
                """, (selected_date_str,))
        today_scrap = cursor.fetchone()['total'] or 0

        return jsonify({
            "today_net_value": f"{float(today_net):,.5f}",
            "today_actual": f"{float(today_actual):,.5f}",
            "today_target": f"{float(daily_target):,.0f}",
            "achievement": round(achievement, 1),
            "mtd_actual": f"{float(mtd_actual):,.5f}",
            "today_iut": f"{float(today_iut):,.5f}",
            "today_scrap": f"{float(today_scrap):,.5f}"
        })
    finally:
        conn.close()


@app.get('/get_daily_products')
def get_daily_products(request: Request):
    selected_date = request.query_params.get('date')
    date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{date_obj.year}-{date_obj.year + 1}" if date_obj.month >= 4 else f"{date_obj.year - 1}-{date_obj.year}"

    conn = get_db_connection()
    cursor = conn.cursor()

    # Actuals
    cursor.execute("""
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                ELSE Prod
            END AS display_prod, 
            SUM(Net_Value)/100000 as actual
        FROM billing_data 
        WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_prod
    """, (selected_date,))
    actual_map = {r['display_prod']: float(r['actual'] or 0) for r in cursor.fetchall()}

    # Daily Targets
    cursor.execute(
        "SELECT Product, SUM(Target)/365 as daily_target FROM billing_target WHERE fin_year=%s GROUP BY Product",
        (fin_year,))
    target_map = {r['Product']: float(r['daily_target'] or 0) for r in cursor.fetchall()}
    conn.close()

    # Match your existing Product Abbreviations
    PRODUCT_MAP = [
        ('DC MACHINES', 'DCM'), ('HV MOTORS', 'HVM'), ('HV GENERATORS', 'HVG'),
        ('RRM', 'RRM'), ('LV MOTORS (NS)', 'LVM(NS)'), ('LV MOTORS (S)', 'LVM(S)'),
        ('EVM', 'EVM'), ('LV GENERATORS', 'LVG'), ('DG SETS', 'DGS'),
        ('TRANSFORMERS PUNE', 'TRF(P)'), ('TRANSFORMERS MYSORE', 'TRF(M)'),
        ('SWITCHGEAR MYSORE', 'SWG'), ('SPARES & SERVICE', 'SP&S')
    ]

    return jsonify({
        "labels": [p[1] for p in PRODUCT_MAP],
        "actuals": [actual_map.get(p[0], 0) for p in PRODUCT_MAP],
        "targets": [target_map.get(p[0], 0) for p in PRODUCT_MAP]
    })


### 3. Daily Region Chart
@app.get('/get_daily_regions')
def get_daily_regions(request: Request):
    selected_date = request.query_params.get('date')
    date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{date_obj.year}-{date_obj.year + 1}" if date_obj.month >= 4 else f"{date_obj.year - 1}-{date_obj.year}"
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India', 'InternationalMarket')

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            CASE WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket' ELSE Sales_Region_Name END AS display_region, 
            SUM(Net_Value)/100000 AS actual
        FROM billing_data
        WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY display_region
    """, (selected_date,))
    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in cursor.fetchall()}

    cursor.execute(
        "SELECT Sales_Region_Name as region, SUM(Target)/365 as daily_target FROM billing_target WHERE fin_year=%s GROUP BY region",
        (fin_year,))
    target_dict = {r['region']: float(r['daily_target'] or 0) for r in cursor.fetchall()}
    conn.close()

    return jsonify({
        "labels": valid_regions,
        "actuals": [actual_dict.get(r, 0) for r in valid_regions],
        "targets": [target_dict.get(r, 0) for r in valid_regions]
    })


@app.get('/get_daily_sectors')
def get_daily_sectors(request: Request):
    selected_date = request.query_params.get('date')
    products = request.query_params.getlist('product')

    if not selected_date:
        return jsonify({"error": "No date provided"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_sectors = ['Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ']
    sector_placeholders = ', '.join(['%s'] * len(valid_sectors))

    # 1. Base query and initial params
    query = f"""
        SELECT 
            Dist_Channel_TEXT as channel, 
            SUM(Net_Value)/100000 as actual
        FROM billing_data
        WHERE Dist_Channel_TEXT IN ({sector_placeholders}) 
        AND Billing_Date = %s 
        AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
    """
    params = list(valid_sectors)
    params.append(selected_date)

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS (NS)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
            elif p == 'LV MOTORS (S)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
            elif p == 'TRANSFORMERS MYSORE':
                product_clauses.append(
                    "(Prod = 'TRANSFORMERS MYSORE' AND (`VTEXT1` != 'Switchgear' OR `VTEXT1` IS NULL))")
            else:
                product_clauses.append("Prod = %s")
                params.append(p)

        query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY Dist_Channel_TEXT"

    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    # 4. Process labels (Grouping SEZ and Domestic Direct)
    processed_data = {}
    for r in rows:
        channel = r['channel']
        actual = float(r['actual'] or 0)

        # Mapping logic
        if channel in ['SEZ', 'Domestic Direct']:
            label = 'Domestic Direct'
        else:
            label = channel

        processed_data[label] = processed_data.get(label, 0) + actual

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


@app.get('/get_daily_units')
def get_daily_units(request: Request):
    selected_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    # Unit-mapping logic
    query = """
        SELECT 
            CASE 
                WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                ELSE Unit 
            END AS display_unit, 
            SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' and Unit!='UN07'
        GROUP BY display_unit
        ORDER BY total DESC
    """

    cursor.execute(query, (selected_date,))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })


@app.get('/get_daily_customers')
def get_daily_customers(request: Request):
    selected_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT 
            Name,
            SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE Billing_Date = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY Name
        ORDER BY total_value DESC
        LIMIT 10
    """

    cursor.execute(query, (selected_date,))
    rows = cursor.fetchall()
    conn.close()

    NAME_REPLACEMENTS = {
        "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
        "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
    }

    processed_labels = []
    processed_values = []

    for r in rows:
        original_name = r['Name'].strip() if r['Name'] else "Unknown"
        short_name = original_name
        for long_name, alias in NAME_REPLACEMENTS.items():
            if long_name in original_name.upper():
                short_name = alias
                break
        processed_labels.append(short_name)
        processed_values.append(float(r['total_value'] or 0))

    return jsonify({
        "labels": processed_labels,
        "values": processed_values
    })


### Daily Orders.

@app.get('/order_daily')
@requires_login
def order_daily(request: Request):
    today_date = datetime.now().strftime('%Y-%m-%d')
    selected_date = request.query_params.get('date', today_date)
    return render_template(request,'daily_orders.html', selected_date=selected_date, username=request.session['username'])


# --- DAILY ORDER SUMMARY (KPIs) ---

@app.get('/get_daily_orders_summary')
def get_daily_orders_summary(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "No date selected"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # 1. Today's Actual Orders
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS actual 
            FROM order_data 
            WHERE Created_Date = %s
        """, (selected_date,))
        res = cursor.fetchone()
        today_actual = float(res['actual'] or 0)

        # 2. Daily Target
        dt = datetime.strptime(selected_date, '%Y-%m-%d')
        if dt.month >= 4:
            fin_year = f"{dt.year}-{dt.year + 1}"
        else:
            fin_year = f"{dt.year - 1}-{dt.year}"

        cursor.execute("SELECT SUM(Target) AS annual_target FROM order_target WHERE fin_year=%s", (fin_year,))
        t_res = cursor.fetchone()
        annual_target = float(t_res['annual_target'] or 0)
        today_target = annual_target / 365

        # 3. Month-to-Date (MTD) Actual
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS mtd 
            FROM order_data 
            WHERE fin_year = %s 
            AND MONTH(Created_Date) = %s 
            AND Created_Date <= %s
        """, (fin_year, dt.month, selected_date))
        m_res = cursor.fetchone()
        mtd_actual = float(m_res['mtd'] or 0)

        achievement = (today_actual / today_target * 100) if today_target > 0 else 0

        return jsonify({
            "today_actual": float(today_actual, ),
            "today_target": float(today_target, ),
            "achievement": float(achievement, ),
            "mtd_actual": f"{float(mtd_actual):,.0f}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()


# --- DAILY PRODUCT WISE ORDERS ---
@app.get('/get_daily_orders_products')
def get_daily_orders_products(request: Request):
    selected_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('DC Machines', 'DCM'), ('HV Motors', 'HVM'), ('HV Generators', 'HVG'),
        ('RRM', 'RRM'),
        ('LV Motors', 'LVM(NS)'), ('LV MOTORS(S)', 'LVM(S)'), ('EVM', 'EVM'), ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'),
        ('Transformer Pune', 'TRF(P)'), ('Transformer Mysore', 'TRF(M)'), ('Switchgear', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]

    # Actuals for the day
    cursor.execute("""
        SELECT 
            CASE 
                WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS display_product, 
            SUM(Net_Value)/100000 AS actual 
        FROM order_data 
        WHERE Created_Date = %s
        GROUP BY display_product
    """, (selected_date,))
    actual_rows = cursor.fetchall()
    actual_dict = {r['display_product']: float(r['actual'] or 0) for r in actual_rows}

    # Target calculation (Annual Target for that year / 365)
    dt = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{dt.year}-{dt.year + 1}" if dt.month >= 4 else f"{dt.year - 1}-{dt.year}"

    cursor.execute(
        "SELECT Product1, SUM(Target)/365 as daily_target FROM order_target WHERE fin_year = %s GROUP BY Product1",
        (fin_year,))
    target_rows = cursor.fetchall()
    target_dict = {r['Product1']: float(r['daily_target'] or 0) for r in target_rows}

    conn.close()

    final_labels, final_actuals, final_targets = [], [], []
    for full_name, abbr in product_abbr:
        if full_name in actual_dict or full_name in target_dict:
            final_labels.append(abbr)
            final_actuals.append(actual_dict.get(full_name, 0))
            final_targets.append(target_dict.get(full_name, 0))

    return jsonify({"labels": final_labels, "actuals": final_actuals, "targets": final_targets})


# --- DAILY REGION WISE ORDERS ---
@app.get('/get_daily_orders_region_data')
def get_daily_orders_region_data(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "No date provided"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'InternationalMarket', 'North India', 'South India', 'West India')

    # 1. Fetch Actuals for the specific day
    query = """
        SELECT Sales_Region as region, SUM(Net_Value)/100000 as actual 
        FROM order_data 
        WHERE Created_Date = %s
        GROUP BY Sales_Region
    """
    cursor.execute(query, (selected_date,))
    rows = cursor.fetchall()
    actual_dict = {r['region']: float(r['actual'] or 0) for r in rows}

    # 2. Calculate Daily Target (Annual / 365)
    dt = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{dt.year}-{dt.year + 1}" if dt.month >= 4 else f"{dt.year - 1}-{dt.year}"

    target_query = """
        SELECT Sales_Region AS region, SUM(Target)/365 AS daily_target
        FROM order_target 
        WHERE Sales_Region IN %s AND fin_year = %s
        GROUP BY Sales_Region
    """
    cursor.execute(target_query, (valid_regions, fin_year))
    target_rows = cursor.fetchall()
    target_dict = {r['region']: float(r['daily_target'] or 0) for r in target_rows}

    conn.close()

    final_labels = []
    final_actuals = []
    final_targets = []

    for region in valid_regions:
        final_labels.append(region)
        final_actuals.append(actual_dict.get(region, 0))
        final_targets.append(target_dict.get(region, 0))

    return jsonify({
        "labels": final_labels,
        "actuals": final_actuals,
        "targets": final_targets
    })


# --- DAILY UNIT WISE ORDERS ---
@app.get('/get_daily_orders_unit_data')
def get_daily_orders_unit_data(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "No date provided"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    VALID_PRODUCTS = (
        'DC Machines', 'DG Sets', 'EVM', 'HV Generators', 'HV Motors', 'LV Generators', 'LV Motors',
        'RRM', 'Switchgear', 'Transformer Pune', 'Transformer Mysore', 'SPARES & SERVICE'
    )

    # Unit logic: Transformer Pune is mapped to UN16, others use the Unit column
    query = """
           SELECT 
               CASE 
                   WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END AS display_unit, 
               SUM(Net_Value)/100000 as actual 
           FROM order_data
           WHERE Product1 IN %s AND Created_Date = %s AND Unit != ''
           GROUP BY 
               CASE 
                   WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END
       """
    params = [VALID_PRODUCTS, selected_date]

    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['actual'] or 0) for r in rows]
    })


# --- DAILY SECTOR WISE ORDERS ---
@app.get('/get_daily_orders_sectors')
def get_daily_orders_sectors(request: Request):
    selected_date = request.query_params.get('date')
    products = request.query_params.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT Dist_Channel AS channel, SUM(Net_Value)/100000 as actual FROM order_data WHERE Created_Date = %s"
    params = [selected_date]

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS(S)':
                product_clauses.append("(Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)')")
            elif p == 'LV Motors':
                product_clauses.append("(Product1 = 'LV Motors' AND Product2 != 'LV Motors (Standard)')")
            else:
                product_clauses.append("Product1 = %s")
                params.append(p)
        query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY Dist_Channel"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['actual'] or 0)

    return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})


# --- DAILY TOP 10 CUSTOMERS ---
@app.get('/get_daily_order_names')
def get_daily_order_names(request: Request):
    selected_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Name, SUM(Net_Value)/100000 AS total_value
        FROM order_data
        WHERE Created_Date = %s
        GROUP BY Name
        ORDER BY total_value DESC
        LIMIT 10
    """, (selected_date,))
    rows = cursor.fetchall()
    conn.close()
    NAME_REPLACEMENTS = {
        "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
        "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
    }

    processed_labels = []
    processed_values = []

    for r in rows:
        original_name = r['Name'].strip() if r['Name'] else "Unknown"
        short_name = original_name
        for long_name, alias in NAME_REPLACEMENTS.items():
            if long_name in original_name.upper():
                short_name = alias
                break
        processed_labels.append(short_name)
        processed_values.append(float(r['total_value'] or 0))

    return jsonify({
        "labels": processed_labels,
        "values": processed_values
    })


@app.get('/get_order_details')
def get_order_details(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Fetching specific columns requested
        query = """
           SELECT Product2, Order_Number, Sales_office, Name, sum(Order_Qty) as Order_Qty, sum(Net_Value) as Net_Value FROM order_data WHERE Created_Date = %s GROUP BY Product2, Order_Number, Sales_office, Name;
        """
        cursor.execute(query, (selected_date,))
        results = cursor.fetchall()

        return jsonify(results)

    except Exception as e:
        print(f"Error in order_details: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


# Quarterly_Sales
@app.get('/quarterly_sales')
@requires_login
def quarterly_dashboard(request: Request):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT fin_year FROM billing_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    available_years = [row['fin_year'] for row in cursor.fetchall()]

    default_year = available_years[0] if available_years else "2025-2026"

    conn.close()

    return render_template(request,
        'quarterly_sales.html',
        username=request.session.get('username', 'User'),
        years=available_years,
        default_year=default_year
    )


def get_quarter_months(quarters_list):
    mapping = {
        'Q1': (4, 5, 6),
        'Q2': (7, 8, 9),
        'Q3': (10, 11, 12),
        'Q4': (1, 2, 3)
    }
    months = []
    for q in quarters_list:
        months.extend(mapping.get(q, []))
    return tuple(months) if months else (0,)


@app.get('/api/ytd/quarterly_kpis')
def get_quarterly_kpis(request: Request):
    year = request.query_params.get('fin_year')
    quarters_list = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({
            "actual": "0.00L",
            "target": "0L",
            "achievement": 0,
            "last_year": "0.00L",
            "iut": "0.00L",
            "scrap": "0.00L"
        })

    conn = get_db_connection()
    cursor = conn.cursor()

    #  mAIN Actuals
    cursor.execute("""
    SELECT SUM(Net_Value)/100000 AS total 
        FROM billing_data 
        WHERE fin_year=%s AND MONTH(Billing_Date) IN %s  AND Dist_Channel_TEXT != 'Inter Unit Transfer'
    """, (year, months))
    net_value = cursor.fetchone()['total'] or 0

    # 1.pROD Actuals
    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS total 
        FROM billing_data 
        WHERE fin_year=%s AND MONTH(Billing_Date) IN %s
        AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
    """, (year, months))
    actual = cursor.fetchone()['total'] or 0

    # 2. Inter Unit Transfer (IUT)
    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS iut_total 
        FROM billing_data 
        WHERE fin_year=%s AND MONTH(Billing_Date) IN %s
        AND Dist_Channel_TEXT = 'Inter Unit Transfer'
    """, (year, months))
    iut = cursor.fetchone()['iut_total'] or 0

    # 3. Scrap Logic (Unit UN07)
    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS scrap_total 
        FROM billing_data 
        WHERE fin_year=%s AND MONTH(Billing_Date) IN %s
        AND Unit = 'UN07'
    """, (year, months))
    scrap = cursor.fetchone()['scrap_total'] or 0

    # 4. Target (Annual target scaled by number of quarters)
    cursor.execute("SELECT SUM(Target) AS annual FROM billing_target WHERE fin_year=%s", (year,))
    annual_target = cursor.fetchone()['annual'] or 0

    target = (annual_target / 4) * len(quarters_list)

    # 5. Last Year Comparison
    try:
        start_year = int(year.split('-')[0])
        prev_year_str = f"{start_year - 1}-{start_year}"
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total 
            FROM billing_data 
            WHERE fin_year=%s AND MONTH(Billing_Date) IN %s
        """, (prev_year_str, months))
        last_year_actual = cursor.fetchone()['total'] or 0
    except Exception as e:
        print(f"Error calculating LY: {e}")
        last_year_actual = 0

    conn.close()

    percent_val = (actual / target * 100) if target > 0 else 0

    return jsonify({
        "net_value": f"{float(net_value):,.5f}L",
        "actual": f"{float(actual):,.5f}L",
        "target": f"{float(target):,.0f}L",
        "achievement": round(percent_val, 1),
        "raw_percent": percent_val,
        "last_year": f"{float(last_year_actual):,.5f}L",
        "iut": f"{float(iut):,.5f}L",
        "scrap": f"{float(scrap):,.5f}L"
    })


@app.get('/api/quarterly/products')
def get_quarterly_product_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_ORDER = [
        ('DC MACHINES', 'DCM'), ('HV MOTORS', 'HVM'), ('HV GENERATORS', 'HVG'),
        ('RRM', 'RRM'), ('LV MOTORS (NS)', 'LVM(NS)'), ('LV MOTORS (S)', 'LVM(S)'),
        ('EVM', 'EVM'), ('LV GENERATORS', 'LVG'), ('DG SETS', 'DGS'),
        ('TRANSFORMERS PUNE', 'TRF(P)'), ('TRANSFORMERS MYSORE', 'TRF(M)'),
        ('SWITCHGEAR MYSORE', 'SWG'), ('SPARES & SERVICE', 'SP&S')
    ]

    query = """
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                ELSE Prod
            END AS display_prod, 
            SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND MONTH(Billing_Date) IN %s
          AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_prod
    """
    cursor.execute(query, (year, months))
    rows = cursor.fetchall()

    # Target Logic
    target_query = "SELECT Product, (SUM(Target)/4)*%s as t_val FROM billing_target WHERE fin_year = %s GROUP BY Product"
    cursor.execute(target_query, (len(quarters), year))
    target_rows = cursor.fetchall()
    conn.close()

    results_dict = {r['display_prod']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['Product']: float(r['t_val'] or 0) for r in target_rows}

    final_labels, final_actuals, final_targets = [], [], []
    for full_name, abbr in PRODUCT_ORDER:
        final_labels.append(abbr)
        final_actuals.append(results_dict.get(full_name, 0))
        final_targets.append(target_dict.get(full_name, 0))

    return jsonify({"labels": final_labels, "actuals": final_actuals, "targets": final_targets})


@app.get('/api/quarterly/regions')
def get_quarterly_region_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India', 'InternationalMarket')

    query = """
        SELECT CASE WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket' ELSE Sales_Region_Name END AS display_region, 
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY display_region HAVING display_region IN %s
    """
    cursor.execute(query, (year, months, valid_regions))
    rows = cursor.fetchall()

    target_query = "SELECT Sales_Region_Name AS region, (SUM(Target)/4)*%s AS q_target FROM billing_target WHERE fin_year = %s GROUP BY region"
    cursor.execute(target_query, (len(quarters), year))
    target_rows = cursor.fetchall()
    conn.close()

    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['region']: float(r['q_target'] or 0) for r in target_rows}

    region_labels, region_actuals, region_targets = [], [], []
    for region in valid_regions:
        region_labels.append(region)
        region_actuals.append(actual_dict.get(region, 0))
        region_targets.append(target_dict.get(region, 0))

    return jsonify({"labels": region_labels, "actuals": region_actuals, "targets": region_targets})


@app.get('/api/quarterly/sector')
def get_quarterly_sector_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)
    products = request.query_params.getlist('product')

    if not year or not months:
        return jsonify({"labels": [], "values": []})

    valid_sectors = ('Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ')

    # Base Query
    query = """
        SELECT Dist_Channel_TEXT AS channel,
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s 
          AND MONTH(Billing_Date) IN %s
          AND Dist_Channel_TEXT IN %s 
          AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
    """

    params = [year, months, valid_sectors]

    if products:
        product_clauses = []
        for p in products:
            if p == 'LV MOTORS (NS)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
            elif p == 'LV MOTORS (S)':
                product_clauses.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
            elif p == 'TRANSFORMERS MYSORE':
                product_clauses.append(
                    "(Prod = 'TRANSFORMERS MYSORE' AND (`VTEXT1` != 'Switchgear' OR `VTEXT1` IS NULL))")
            else:
                product_clauses.append("Prod = %s")
                params.append(p)

        if product_clauses:
            query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY Dist_Channel_TEXT"

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()

        processed_data = {}
        for row in rows:
            channel = row['channel']
            actual = float(row['actual'] or 0)

            if channel in ('SEZ', 'Domestic Direct'):
                label = 'Domestic Direct'
                processed_data[label] = processed_data.get(label, 0) + actual
            else:
                processed_data[channel] = actual

        return jsonify({
            "labels": list(processed_data.keys()),
            "values": list(processed_data.values())
        })

    except Exception as e:
        return jsonify({"error": "Database error", "message": str(e)}, status_code=500)


@app.get('/api/quarterly/unit')
def get_quarterly_unit_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT CASE WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16' ELSE Unit END AS display_unit, 
               SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE fin_year = %s AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY display_unit
    """
    cursor.execute(query, (year, months))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({"labels": [r['display_unit'] for r in rows], "values": [float(r['total'] or 0) for r in rows]})


@app.get('/api/quarterly/top_customers')
def get_quarterly_customers(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT Name, SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE fin_year = %s AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit != 'UN07'
        GROUP BY Name ORDER BY total_value DESC LIMIT 10
    """
    cursor.execute(query, (year, months))
    rows = cursor.fetchall()
    conn.close()

    NAME_REPLACEMENTS = {"KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman", "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"}
    processed_labels = []
    for r in rows:
        name = r['Name'].strip() if r['Name'] else ""
        for long_name, alias in NAME_REPLACEMENTS.items():
            if long_name in name.upper():
                name = alias
                break
        processed_labels.append(name)

    return jsonify({"labels": processed_labels, "values": [float(r['total_value'] or 0) for r in rows]})


# ORDERS-QUARTERLY
@app.get('/quarterly_order')
@requires_login
def quarterly_order_dashboard(request: Request):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT fin_year FROM order_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    available_years = [row['fin_year'] for row in cursor.fetchall()]

    default_year = available_years[0] if available_years else "2025-2026"

    conn.close()

    return render_template(request,
        'quarterly_orders.html',
        username=request.session.get('username', 'User'),
        years=available_years,
        default_year=default_year
    )


@app.get('/api/order/quarterly_kpis')
def get_order_quarterly_kpis(request: Request):
    year = request.query_params.get('fin_year')
    quarters_list = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({"actuals": "0L", "targets": "0L", "achievement": "0%", "raw_percent": 0})

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS total 
        FROM order_data 
        WHERE fin_year=%s AND MONTH(Created_Date) IN %s
    """, (year, months))
    actual = cursor.fetchone()['total'] or 0

    cursor.execute("SELECT SUM(Target) AS annual FROM order_target WHERE fin_year=%s", (year,))
    annual_target = cursor.fetchone()['annual'] or 0
    target = (annual_target / 4) * len(quarters_list)

    try:
        prev_y_start = int(year.split('-')[0]) - 1
        prev_year_str = f"{prev_y_start}-{prev_y_start + 1}"
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total 
            FROM order_data 
            WHERE fin_year=%s AND MONTH(Created_Date) IN %s
        """, (prev_year_str, months))
        last_year_actual = cursor.fetchone()['total'] or 0
    except:
        last_year_actual = 0

    conn.close()
    achievement_pct = (actual / target * 100) if target > 0 else 0

    return jsonify({
        "actuals": f"{float(actual):,.5f}L",
        "last_year": f"{float(last_year_actual):,.5f}L",
        "targets": f"{float(target):,.0f}L",
        "achievement": f"{achievement_pct:.1f}%",
        "raw_percent": achievement_pct,
        "raw_actual": float(actual),
        "raw_target": float(target)
    })


@app.get('/api/order/quarterly_products')
def get_order_quarterly_products(request: Request):
    year = request.query_params.get('fin_year')
    quarters_list = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters_list)

    product_abbr = [
        ('DC Machines', 'DCM'), ('HV Motors', 'HVM'), ('HV Generators', 'HVG'),
        ('RRM', 'RRM'), ('LV Motors', 'LVM(NS)'), ('LV MOTORS(S)', 'LVM(S)'), ('EVM', 'EVM'), ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'),
        ('Transformer Pune', 'TRF(P)'), ('Transformer Mysore', 'TRF(M)'), ('Switchgear', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]

    conn = get_db_connection()
    cursor = conn.cursor()

    # Actuals
    query = """
       SELECT 
            CASE 
                WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS display_product, 
            SUM(Net_Value)/100000 AS actual 
        FROM order_data 
        WHERE fin_year = %s AND MONTH(Created_Date) IN %s
        GROUP BY display_product
    """
    cursor.execute(query, (year, months))
    rows = cursor.fetchall()
    actual_dict = {r['display_product']: float(r['actual'] or 0) for r in rows}

    # Targets
    target_query = "SELECT Product1, SUM(Target) as annual FROM order_target WHERE fin_year=%s GROUP BY Product1"
    cursor.execute(target_query, (year,))
    target_dict = {r['Product1']: (float(r['annual'] or 0) / 4) * len(quarters_list) for r in cursor.fetchall()}

    conn.close()

    # Format for Chart.js
    labels, actuals, targets = [], [], []
    for full, abbr in product_abbr:
        labels.append(abbr)
        actuals.append(actual_dict.get(full, 0))
        targets.append(target_dict.get(full, 0))

    return jsonify({"labels": labels, "actuals": actuals, "targets": targets})


@app.get('/api/order/quarterly_regions')
def get_order_quarterly_region_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters_list = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({"labels": [], "actuals": [], "targets": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'InternationalMarket', 'North India', 'South India', 'West India')

    # 1. Fetch Actuals for the specific quarters
    query = """
        SELECT Sales_Region AS region,
               SUM(Net_Value)/100000 AS actual
        FROM order_data
        WHERE fin_year = %s AND MONTH(Created_Date) IN %s AND Sales_Region IN %s
        GROUP BY Sales_Region
    """
    cursor.execute(query, (year, months, valid_regions))
    rows = cursor.fetchall()
    actual_dict = {r['region']: float(r['actual'] or 0) for r in rows}

    # 2. Fetch Annual Targets and Pro-rate them (Annual / 4 * selected_quarters)
    target_query = """
        SELECT Sales_Region AS region, SUM(Target) AS annual_target
        FROM order_target 
        WHERE fin_year = %s AND Sales_Region IN %s
        GROUP BY Sales_Region
    """
    cursor.execute(target_query, (year, valid_regions))
    target_rows = cursor.fetchall()

    # Pro-rating logic: (Annual Target / 4) * number of quarters selected
    num_q = len(quarters_list)
    target_dict = {r['region']: (float(r['annual_target'] or 0) / 4) * num_q for r in target_rows}

    conn.close()

    final_actuals = []
    final_targets = []
    # Map back to valid_regions to ensure consistent order in chart
    for region in valid_regions:
        final_actuals.append(actual_dict.get(region, 0))
        final_targets.append(target_dict.get(region, 0))

    return jsonify({
        "labels": valid_regions,
        "actuals": final_actuals,
        "targets": final_targets
    })


@app.get('/api/order/quarterly_sector')
def get_order_quarterly_sector(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)
    products = request.query_params.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Dist_Channel AS channel, SUM(Net_Value)/100000 as actual 
        FROM order_data 
        WHERE fin_year = %s AND MONTH(Created_Date) IN %s
    """
    params = [year, months]

    if products:
        clauses = []
        for p in products:
            if p == 'LV MOTORS(S)':
                clauses.append("(Product1='LV Motors' AND Product2='LV Motors (Standard)')")
            elif p == 'LV Motors':
                clauses.append("(Product1='LV Motors' AND Product2!='LV Motors (Standard)')")
            else:
                clauses.append("Product1 = %s")
                params.append(p)
        query += " AND (" + " OR ".join(clauses) + ")"

    query += " GROUP BY Dist_Channel"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    processed = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed[label] = processed.get(label, 0) + float(r['actual'] or 0)

    return jsonify({"labels": list(processed.keys()), "values": list(processed.values())})


@app.get('/api/order/quarterly_Unit')
def get_order_quarterly_Unit_data(request: Request):
    year = request.query_params.get('fin_year')
    quarters_list = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    VALID_PRODUCTS = (
        'DC Machines', 'DG Sets', 'EVM', 'HV Generators', 'HV Motors',
        'LV Generators', 'LV Motors', 'RRM', 'Switchgear',
        'Transformer Pune', 'Transformer Mysore', 'SPARES & SERVICE'
    )

    # Note: Using Created_Date for Order monitoring
    query = """
        SELECT 
            CASE 
                WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                ELSE Unit 
            END AS display_unit, 
            SUM(Net_Value)/100000 as total 
        FROM order_data
        WHERE fin_year = %s 
          AND MONTH(Created_Date) IN %s 
          AND Product1 IN %s 
          AND Unit != ''
        GROUP BY display_unit
    """

    cursor.execute(query, (year, months, VALID_PRODUCTS))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })


@app.get('/api/order/quarterly_customers')
def get_order_quarterly_customers(request: Request):
    year = request.query_params.get('fin_year')
    quarters = request.query_params.getlist('quarters')
    months = get_quarter_months(quarters)  # Assuming your helper function

    if not year or not months:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Filter by selected months and year
        query = """
            SELECT 
                Name,
                SUM(Net_Value) / 100000 AS total_value
            FROM order_data
            WHERE fin_year = %s AND MONTH(Created_Date) IN %s
            GROUP BY Name
            ORDER BY total_value DESC
            LIMIT 10
        """
        cursor.execute(query, (year, months))
        rows = cursor.fetchall()

        # Clean customer names (Handle KEC aliases)
        NAME_REPLACEMENTS = {
            "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
            "KIRLOSKAR ELECTRIC CO LTD": "KEC, Ajman"
        }

        labels = []
        values = []
        for r in rows:
            name = r['Name'].strip().upper() if r['Name'] else ""
            display_name = name
            for full, alias in NAME_REPLACEMENTS.items():
                if full in name:
                    display_name = alias
                    break
            labels.append(display_name)
            values.append(float(r['total_value'] or 0))

        return jsonify({"labels": labels, "values": values})
    finally:
        conn.close()


@app.get('/pending_orders')
def pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


@app.get('/api/pending/total')
def get_total_pending(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(`Net_Value`)/100000 AS total FROM pending_order WHERE `As_on_Date` = %s ", (order_date,))
    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


@app.get('/api/pending/products')
def get_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('DC Machines', 'DCM'), ('HV Motors', 'HVM'), ('HV Generators', 'HVG'),
        ('RRM', 'RRM'), ('Traction', 'TRC'), ('LV Motors (Non-Standard)', 'LVM(NS)'),
        ('LV Motors (Standard)', 'LVM(S)'),
        ('EVM', 'EVM'), ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'), ('Transformer Pune', 'TRF(P)'), ('Transformer Mysore', 'TRF(M)'), ('SWITCHGEARS', 'SWG'),
        ('SPARES & SERVICE', 'SP&S')
    ]

    try:

        cursor.execute("""
            SELECT 
                CASE
                    WHEN Product2 = 'Transformer Mysore' AND Level_1 = 'Switchgear' THEN 'SWITCHGEARS'
                    ELSE Product2
                END AS display_product, 
                SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s  
            GROUP BY display_product
        """, (order_date,))

        rows = cursor.fetchall()
        actual_dict = {r['display_product']: float(r['value'] or 0) for r in rows}

        final_labels = []
        final_values = []

        for full_name, abbr in product_abbr:
            if full_name in actual_dict:
                final_labels.append(abbr)
                final_values.append(float(actual_dict[full_name]))

        return jsonify({
            "labels": final_labels,
            "values": final_values
        })

    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()


@app.get('/api/pending/branches')
def get_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND Branch_Name != '' 
          AND Branch_Name NOT IN ('Hubli', 'Jaipur', 'Lucknow')
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)

        if name in ['Delhi', 'Faridabad']:
            target_name = 'Faridabad'
            processed_dict[target_name] = processed_dict.get(target_name, 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


@app.get('/api/pending/sectors')
def get_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    # Get the list of selected products from the request
    products = request.query_params.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT `Distr_Channel` AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s
    """
    params = [order_date]

    if products:
        product_clauses = []
        for p in products:
            if p == 'SWITCHGEARS':
                product_clauses.append("(`Product2` = 'Transformer Mysore' AND `Level_1` = 'Switchgear')")

            elif p == 'Transformer Mysore':
                product_clauses.append(
                    "(`Product2` = 'Transformer Mysore' AND (`Level_1` != 'Switchgear' OR `Level_1` IS NULL))")

            else:
                # Default case for all other products
                product_clauses.append("`Product2` = %s")
                params.append(p)

        query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY `Distr_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()

    processed = {}
    for r in rows:
        chan = r['label']
        val = float(r['value'] or 0)
        label = 'DD' if chan in ['SZ', 'DD'] else chan
        processed[label] = processed.get(label, 0) + val

    return jsonify({
        "labels": list(processed.keys()),
        "values": [float(v) for v in processed.values()]
    })


@app.get('/api/pending/units')
def get_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    VALID_PRODUCTS = (
        'Traction', 'DC Machines', 'DG Sets', 'EVM', 'HV Generators', 'HV Motors', 'LV Generators', 'LV Motors',
        'RRM', 'Transformer Pune', 'Transformer Mysore', 'SPARES & SERVICE'
    )

    query = """
        SELECT 
               CASE 
                   WHEN Product2 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END AS display_unit, 
               SUM(Net_Value)/100000 as value 
           FROM pending_order
           WHERE Product2 IN %s AND `As_on_Date` = %s AND Unit != ''
           GROUP BY 
               CASE 
                   WHEN Product2 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END
    """
    params = [VALID_PRODUCTS, order_date]
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


@app.get('/api/pending/customers')
def get_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))
    rows = cursor.fetchall()
    conn.close()
    NAME_REPLACEMENTS = {
        "KIRLOSKAR ELECTRIC COMPANY LIMITED": "KEC, Ajman",
        "KIRLOSKAR ELECTRIC COMPANY LIM": "KEC, Ajman"
    }

    processed_labels = []
    processed_values = []

    for r in rows:
        original_name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        short_name = original_name
        for long_name, alias in NAME_REPLACEMENTS.items():
            if long_name in original_name.upper():
                short_name = alias
                break
        processed_labels.append(short_name)
        processed_values.append(float(r['value'] or 0))
    return jsonify({
        "labels": processed_labels,
        "values": processed_values
    })


@app.get('/fg_report')
def fg_report(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'fg_report.html', username=request.session.get('username', 'User'), selected_date=today)


@app.get('/api/fg_report/stock_value')
def get_stock_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
            SELECT SUM(Stock_Value) / 100000 AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        total_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "total_value": float(total_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_2')
def get_unit2_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
             SELECT SUM(Stock_Value) / 100000 AS total_stock_lakhs 
             FROM stock_data 
            WHERE AS_ON_Date = %s AND `Product2`='LV Motors (Non-Standard)'
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit2_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit2_value": float(unit2_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_1')
def get_unit1_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
            SELECT SUM(Stock_Value) / 100000 AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s AND `Unit`='UN01'
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit1_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit1_value": float(unit1_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_25')
def get_unit25_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
           SELECT SUM(Stock_Value) / 100000  AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s AND `Product2`IN('EVM','LV Generators','LV Motors (Standard)')
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit25_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit25_value": float(unit25_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_5')
def get_unit5_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
           SELECT SUM(Stock_Value) / 100000  AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s AND `Product2`IN('Transformer Mysore')
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit5_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit5_value": float(unit5_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_6')
def get_unit6_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
           SELECT SUM(Stock_Value) / 100000  AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s AND `Unit`='UN06'
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit6_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit6_value": float(unit6_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/unit_16')
def get_unit16_value(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
           SELECT SUM(Stock_Value) / 100000 AS total_stock_lakhs 
            FROM stock_data 
            WHERE AS_ON_Date = %s AND  `Product2`='Transformer Pune'
        """
        cursor.execute(query, (selected_date,))
        result = cursor.fetchone()

        unit16_value = result['total_stock_lakhs'] if result['total_stock_lakhs'] else 0

        return jsonify({
            "unit16_value": float(unit16_value)
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/profit_center_stock')
def get_profit_center_stock(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    PROFIT_CENTER_ORDER = [
        'UN01DCM', 'UN01ACM', 'UN01MBS', 'UN02ACM', 'UN25ACM',
        'UN25EVM', 'UN25ACG', 'UN05OFT', 'UN06DGS', 'UN16CRT'
    ]

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        query = """
            SELECT 
                Profit_Center, 
                SUM(Stock_Value) / 100000 as total_value
            FROM stock_data
            WHERE AS_ON_Date = %s
            GROUP BY Profit_Center
        """
        cursor.execute(query, (selected_date,))
        results = cursor.fetchall()

        db_map = {row['Profit_Center']: row['total_value'] for row in results}

        ordered_labels = []
        ordered_values = []

        for pc in PROFIT_CENTER_ORDER:
            ordered_labels.append(pc)
            val = db_map.get(pc, 0)
            ordered_values.append(float(val) if val else 0.0)

        return jsonify({
            "labels": ordered_labels,
            "values": ordered_values
        })

    except Exception as e:
        print(f"Database Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/aging_stock')
def get_fg_aging_stock(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    PROFIT_CENTER_ORDER = [
        'UN01DCM', 'UN01ACM', 'UN01MBS', 'UN02ACM', 'UN25ACM',
        'UN25EVM', 'UN25ACG', 'UN05OFT', 'UN06DGS', 'UN16CRT'
    ]

    conn = get_db_connection()

    cursor = conn.cursor()

    try:

        query = """
            SELECT 
                Profit_Center, 
                SUM(0_90_Val) / 100000 as val_0_90,
                SUM(91_180_Val) / 100000 as val_91_180,
                SUM(Above_180_Val) / 100000 as val_above_180
            FROM stock_data
            WHERE AS_ON_Date = %s
            GROUP BY Profit_Center
        """
        cursor.execute(query, (selected_date,))
        results = cursor.fetchall()
        db_map = {row['Profit_Center']: row for row in results}

        ordered_labels = []
        dataset_0_90 = []
        dataset_91_180 = []
        dataset_above_180 = []

        for pc in PROFIT_CENTER_ORDER:
            ordered_labels.append(pc)

            if pc in db_map:
                row = db_map[pc]
                dataset_0_90.append(float(row['val_0_90'] or 0))
                dataset_91_180.append(float(row['val_91_180'] or 0))
                dataset_above_180.append(float(row['val_above_180'] or 0))
            else:
                dataset_0_90.append(0.0)
                dataset_91_180.append(0.0)
                dataset_above_180.append(0.0)

        return jsonify({
            "labels": ordered_labels,
            "datasets": [
                {"label": "0-90 Days", "data": dataset_0_90, "color": "#28a745"},
                {"label": "91-180 Days", "data": dataset_91_180, "color": "#ffc107"},
                {"label": "Above 180 Days", "data": dataset_above_180, "color": "#dc3545"}
            ]
        })

    except Exception as e:
        print(f"Database Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/fg_report/credit_table')
def get_credit_table_data(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return jsonify({"error": "Date is required"}, status_code=400)

    PROFIT_CENTER_ORDER = [
        'UN01DCM', 'UN01ACM', 'UN01MBS', 'UN02ACM', 'UN25ACM',
        'UN25EVM', 'UN25ACG', 'UN05OFT', 'UN06DGS', 'UN16CRT'
    ]

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        query = """
            SELECT 
                Profit_Center, 
                SUM(Opening_Val) as opening,
                SUM(Receipt_Val) as receipt,
                SUM(Issue_Val) as issued,
                SUM(Closing_Val) as closing
            FROM stock_analysis
            WHERE As_on_date = %s
            GROUP BY Profit_Center
        """
        cursor.execute(query, (selected_date,))
        results = cursor.fetchall()

        db_map = {row['Profit_Center']: row for row in results}

        ordered_table_data = []

        for pc in PROFIT_CENTER_ORDER:
            if pc in db_map:
                row = db_map[pc]
                ordered_table_data.append({
                    "product": pc,
                    "opening": float(row['opening'] or 0),
                    "receipt": float(row['receipt'] or 0),
                    "issued": float(row['issued'] or 0),
                    "closing": float(row['closing'] or 0)
                })
            else:
                ordered_table_data.append({
                    "product": pc,
                    "opening": 0.0,
                    "receipt": 0.0,
                    "issued": 0.0,
                    "closing": 0.0
                })

        return jsonify(ordered_table_data)

    except Exception as e:
        print(f"Error in credit_table: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()

@app.get('/api/export/sales_excel')
def export_sales_excel(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query(request, "data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        # 3. Load the dictionary list directly into Pandas
        # This completely fixes the "repeating headers" bug and the DBAPI2 warning
        df = pd.DataFrame(results)

        # 4. Clean up Date formats before exporting
        date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date', 'PO_Date']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y')

        # 5. Write to BytesIO stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sales_Export')
        output.seek(0)

        # 6. Generate dynamic filename based on selected Financial Year
        fin_years = request.query_params.getlist('fin_year')
        fy_str = "_".join(fin_years) if fin_years else "All_Years"
        filename = f"Sales_Export_{fy_str}.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

#Orders Page

@app.get('/api/export/orders_excel')
def export_orders_excel(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query(request, "order_data")

    # 2. Build the query
    query = f"SELECT * FROM order_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                # Execute the query normally
                cursor.execute(query, params)
                results = cursor.fetchall()

        # If no results, return 404
        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        # 3. Load the dictionary list directly into Pandas
        df = pd.DataFrame(results)

        # 4. Clean up Date formats (Created_Date is standard for your order_data)
        date_cols = ['Created_Date', 'Document_Date']
        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y')

        # 5. Write to BytesIO stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Orders_Export')
        output.seek(0)

        # 6. Generate dynamic filename
        fin_years = request.query_params.getlist('fin_year')
        fy_str = "_".join(fin_years) if fin_years else "All_Years"
        filename = f"Orders_Export_{fy_str}.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# Excel download
@app.get('/sales/export/ytd_excel')
def export_sales_ytd_excel(request: Request):
    fin_year = request.query_params.get('fin_year')
    page_type = request.query_params.get('type')

    if not fin_year:
        return HTMLResponse("Financial Year is required", status_code=400)

    table_name = "billing_data" if page_type == "billing" else "order_data"

    try:

        with engine.connect() as conn:
            query = text(f"SELECT * FROM {table_name} WHERE fin_Year = :fy")
            df = pd.read_sql(query, conn, params={"fy": fin_year})

        if df.empty:
            return HTMLResponse("No data found for the selected Financial Year", status_code=404)

        date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date', 'PO_Date']

        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='YTD_Export')

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"{page_type}_export_{fin_year}.xlsx"
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error: {str(e)}", status_code=500)


def serve_excel(df, filename):
    if df.empty:
        return HTMLResponse("No data found for the selected period", status_code=404)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )




@app.get('/api/export/advance_daily_orders_excel')
def export_daily_order_advance_excel(request: Request):
    selected_date = request.query_params.get('date')
    if not selected_date:
        return HTMLResponse("Please select a date", status_code=400)

    raw_query = """
            SELECT 
                Order_Number, 
                Created_Date,
                Location,
                Unit,
                Product1,
                Product2,
                Plant,
                Sales_Org,
                Sales_Region,
                Sales_office_Cd,
                Sales_office,
                Name,
                Sales_Engg_Name,
                PO_No,
                PO_Date,
                Item,
                Material_Code,
                Description,
                Order_Qty,
                Net_Value,
                Total_SO_Value,
                Advance_Value,
                Advance_Recevied,
                Advance_To_be_received,
                Payment_Terms

            FROM 
                order_data
            WHERE 
                Created_Date = :dt
        """
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(raw_query), conn, params={"dt": selected_date})

        if df.empty:
            return HTMLResponse(f"No records found for {selected_date} in orders", status_code=404)

        aggregation_rules = {
            'Created_Date': 'first',
            'Location': 'first',
            'Unit': 'first',
            'Product1': 'first',
            'Product2': 'first',
            'Plant': 'first',
            'Sales_Org': 'first',
            'Sales_Region': 'first',
            'Sales_office_Cd': 'first',
            'Sales_office': 'first',
            'Name': 'first',
            'Sales_Engg_Name': 'first',
            'PO_No': 'first',
            'PO_Date': 'first',
            'Item': 'sum',
            'Material_Code': 'first',
            'Description': 'first',
            'Order_Qty': 'sum',
            'Net_Value': 'sum',
            'Total_SO_Value': 'sum',
            'Advance_Value': 'sum',
            'Advance_Recevied': 'sum',
            'Advance_To_be_received': 'sum',
            'Payment_Terms': 'first'
        }

        df_unique_orders = df.groupby('Order_Number', as_index=False).agg(aggregation_rules)

        date_cols = ['Created_Date', 'PO_Date']
        for col in date_cols:
            if col in df_unique_orders:
                df_unique_orders[col] = pd.to_datetime(df_unique_orders[col], errors='coerce')
                df_unique_orders[col] = df_unique_orders[col].dt.strftime('%d-%m-%Y')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_unique_orders.to_excel(writer, index=False, sheet_name='Advance_Orders')

        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"AdvanceOrder_{selected_date}.xlsx"
        )
    except Exception as e:
        print(f"SQL Error: {e}")
        traceback.print_exc()
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)


@app.get('/api/export/monthly_ad_order_excel')
def export_monthly_ad_order_excel(request: Request):
    fin_year = request.query_params.get('year')
    months_param = request.query_params.get('months')

    if not fin_year or not months_param:
        return HTMLResponse("Please select Financial Year and at least one month", status_code=400)

    month_map = {
        'January': 1, 'February': 2, 'March': 3, 'April': 4,
        'May': 5, 'June': 6, 'July': 7, 'August': 8,
        'September': 9, 'October': 10, 'November': 11, 'December': 12
    }

    selected_months = [month_map[m] for m in months_param.split(',') if m in month_map]

    raw_query = """
               SELECT 
                   Order_Number, 
                   Created_Date,
                   Location,
                   Unit,
                   Product1,
                   Product2,
                   Plant,
                   Sales_Org,
                   Sales_Region,
                   Sales_office_Cd,
                   Sales_office,
                   Name,
                   Sales_Engg_Name,
                   PO_No,
                   PO_Date,
                   Item,
                   Material_Code,
                   Description,
                   Order_Qty,
                   Net_Value,
                   Total_SO_Value,
                   Advance_Value,
                   Advance_Recevied,
                   Advance_To_be_received,
                   Payment_Terms

               FROM 
                   order_data
               WHERE 
                   fin_Year = :fy AND MONTH(Created_Date) IN :mo_list
           """

    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(raw_query), conn, params={
                "fy": fin_year,
                "mo_list": tuple(selected_months)
            })
            aggregation_rules = {
                'Created_Date': 'first',
                'Location': 'first',
                'Unit': 'first',
                'Product1': 'first',
                'Product2': 'first',
                'Plant': 'first',
                'Sales_Org': 'first',
                'Sales_Region': 'first',
                'Sales_office_Cd': 'first',
                'Sales_office': 'first',
                'Name': 'first',
                'Sales_Engg_Name': 'first',
                'PO_No': 'first',
                'PO_Date': 'first',
                'Item': 'sum',
                'Material_Code': 'first',
                'Description': 'first',
                'Order_Qty': 'sum',
                'Net_Value': 'sum',
                'Total_SO_Value': 'sum',
                'Advance_Value': 'sum',
                'Advance_Recevied': 'sum',
                'Advance_To_be_received': 'sum',
                'Payment_Terms': 'first'
            }

            df_unique_orders = df.groupby('Order_Number', as_index=False).agg(aggregation_rules)
            date_cols = ['Created_Date', 'PO_Date']

            for col in date_cols:
                if col in df.columns:
                    df_unique_orders[col] = pd.to_datetime(df_unique_orders[col], errors='coerce')
                    df_unique_orders[col] = df_unique_orders[col].dt.strftime('%d-%m-%Y')

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_unique_orders.to_excel(writer, index=False, sheet_name='Advance_Order_MTD')

            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f"AdvanceOrder_MTD_{fin_year}.xlsx"
            )
    except Exception as e:
        print(f"Monthly Export Error: {e}")
        traceback.print_exc()
        return HTMLResponse(str(e), status_code=500)


@app.get('/api/export/ad_order_quarterly_excel')
def export_ad_order_quarterly_excel(request: Request):
    fin_year = request.query_params.get('year')
    quarters_param = request.query_params.get('quarters')

    if not fin_year or not quarters_param:
        return HTMLResponse("Missing parameters", status_code=400)

    quarters_list = quarters_param.split(',')
    target_months = get_quarter_months(quarters_list)
    raw_query = """
            SELECT 
                Order_Number, 
                Created_Date,
                Location,
                Unit,
                Product1,
                Product2,
                Plant,
                Sales_Org,
                Sales_Region,
                Sales_office_Cd,
                Sales_office,
                Name,
                Sales_Engg_Name,
                PO_No,
                PO_Date,
                Item,
                Material_Code,
                Description,
                Order_Qty,
                Net_Value,
                Total_SO_Value,
                Advance_Value,
                Advance_Recevied,
                Advance_To_be_received,
                Payment_Terms

            FROM 
                order_data
            WHERE 
                fin_Year = :fy AND MONTH(Created_Date) IN :months
        """
    try:
        with engine.connect() as conn:
            df = pd.read_sql(text(raw_query), conn, params={
                "fy": fin_year,
                "months": target_months
            })

            aggregation_rules = {
                'Created_Date': 'first',
                'Location': 'first',
                'Unit': 'first',
                'Product1': 'first',
                'Product2': 'first',
                'Plant': 'first',
                'Sales_Org': 'first',
                'Sales_Region': 'first',
                'Sales_office_Cd': 'first',
                'Sales_office': 'first',
                'Name': 'first',
                'Sales_Engg_Name': 'first',
                'PO_No': 'first',
                'PO_Date': 'first',
                'Item': 'sum',
                'Material_Code': 'first',
                'Description': 'first',
                'Order_Qty': 'sum',
                'Net_Value': 'sum',
                'Total_SO_Value': 'sum',
                'Advance_Value': 'sum',
                'Advance_Recevied': 'sum',
                'Advance_To_be_received': 'sum',
                'Payment_Terms': 'first'
            }

            df_unique_orders = df.groupby('Order_Number', as_index=False).agg(aggregation_rules)
            date_cols = ['Created_Date', 'PO_Date']

            for col in date_cols:
                if col in df.columns:
                    df_unique_orders[col] = pd.to_datetime(df_unique_orders[col], errors='coerce')
                    df_unique_orders[col] = df_unique_orders[col].dt.strftime('%d-%m-%Y')

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_unique_orders.to_excel(writer, index=False, sheet_name='Advance_Order_QTD')

            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f"AdvanceOrder_QTD_{fin_year}.xlsx"
            )
    except Exception as e:
        print(f"Pivot Export  Failure Error: {e}")
        traceback.print_exc()
        return HTMLResponse(f"Internal Server error: {str(e)}", status_code=500)


@app.get('/sales/export/ad_order_ytd_excel')
def export_ad_order_ytd_excel(request: Request):
    fin_year = request.query_params.get('fin_year')

    if not fin_year:
        return HTMLResponse("Financial Year is required", status_code=400)

    raw_query = """
        SELECT 
            Order_Number, 
            Created_Date,
            Location,
            Unit,
            Product1,
            Product2,
            Plant,
            Sales_Org,
            Sales_Region,
            Sales_office_Cd,
            Sales_office,
            Name,
            Sales_Engg_Name,
            PO_No,
            PO_Date,
            Item,
            Material_Code,
            Description,
            Order_Qty,
            Net_Value,
            Total_SO_Value,
            Advance_Value,
            Advance_Recevied,
            Advance_To_be_received,
            Payment_Terms

        FROM 
            order_data
        WHERE 
            fin_Year = :fy
    """

    try:
        with engine.connect() as conn:
            df_raw = pd.read_sql(text(raw_query), conn, params={"fy": fin_year})

        if df_raw.empty:
            return HTMLResponse("No data found for the selected Financial Year", status_code=404)

        aggregation_rules = {
            'Created_Date': 'first',
            'Location': 'first',
            'Unit': 'first',
            'Product1': 'first',
            'Product2': 'first',
            'Plant': 'first',
            'Sales_Org': 'first',
            'Sales_Region': 'first',
            'Sales_office_Cd': 'first',
            'Sales_office': 'first',
            'Name': 'first',
            'Sales_Engg_Name': 'first',
            'PO_No': 'first',
            'PO_Date': 'first',
            'Item': 'sum',
            'Material_Code': 'first',
            'Description': 'first',
            'Order_Qty': 'sum',
            'Net_Value': 'sum',
            'Total_SO_Value': 'sum',
            'Advance_Value': 'sum',
            'Advance_Recevied': 'sum',
            'Advance_To_be_received': 'sum',
            'Payment_Terms': 'first'
        }

        df_unique_orders = df_raw.groupby('Order_Number', as_index=False).agg(aggregation_rules)

        for date_col in ['Created_Date', 'PO_Date']:
            if date_col in df_unique_orders.columns:
                df_unique_orders[date_col] = pd.to_datetime(df_unique_orders[date_col], errors='coerce')
                df_unique_orders[date_col] = df_unique_orders[date_col].dt.strftime('%d-%m-%Y').fillna('')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_unique_orders.to_excel(writer, index=False, sheet_name='Advance_Order_YTD')

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"AdvanceOrder_YTD_{fin_year}.xlsx"
        )

    except Exception as e:
        print(f"Pivot Export  Failure Error: {e}")
        traceback.print_exc()
        return HTMLResponse(f"Internal Server error: {str(e)}", status_code=500)


@app.get('/api/pending_order_excel')
def pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date', 'Created_On']

        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')

        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)


@app.get('/api/export/combined_stock_excel')
def export_combined_stock_excel(request: Request):
    selected_date = request.query_params.get('date')

    if not selected_date:
        return HTMLResponse("Please select a date", status_code=400)

    try:
        with engine.connect() as conn:

            query_stock = text("SELECT * FROM stock_data WHERE AS_ON_Date >= :dt AND AS_ON_Date < DATE_ADD(:dt, INTERVAL 1 DAY)")
            df_stock = pd.read_sql(query_stock, conn, params={"dt": selected_date})

            query_analysis = text("SELECT * FROM stock_analysis WHERE As_on_date >= :dt AND As_on_date < DATE_ADD(:dt, INTERVAL 1 DAY)")
            df_analysis = pd.read_sql(query_analysis, conn, params={"dt": selected_date})

        if df_stock.empty and df_analysis.empty:
            return HTMLResponse(f"No stock records found for {selected_date}", status_code=404)

        if not df_stock.empty and 'AS_ON_Date' in df_stock.columns:
            df_stock['AS_ON_Date'] = pd.to_datetime(df_stock['AS_ON_Date'], errors='coerce').dt.strftime('%d-%m-%Y')

        if not df_analysis.empty and 'As_on_date' in df_analysis.columns:
            df_analysis['As_on_date'] = pd.to_datetime(df_analysis['As_on_date'], errors='coerce').dt.strftime(
                '%d-%m-%Y')

        if not df_analysis.empty and 'Posting_Date' in df_analysis.columns:
            df_analysis['Posting_Date'] = pd.to_datetime(df_analysis['Posting_Date'], errors='coerce').dt.strftime(
                '%d-%m-%Y')

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_stock.to_excel(writer, index=False, sheet_name='Stock_Data')
            df_analysis.to_excel(writer, index=False, sheet_name='Stock_FG_Report')

        output.seek(0)

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"Stock_Report_{selected_date}.xlsx"
        )

    except Exception as e:
        print(f"Combined Export Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)


@app.get('/collection_report')
def collection_report(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'collection_report.html', username=request.session.get('username', 'User'), selected_date=today)


@app.get('/api/collection/value')
def get_collection_value(request: Request):
    start_date = request.query_params.get('start_date') or request.query_params.get('date')
    end_date = request.query_params.get('end_date') or start_date
    if not start_date:
        return jsonify({"error": "Start date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        query = """
            SELECT SUM(Balance) AS total_value 
            FROM collections_set 
            WHERE AS_ON_Date BETWEEN %s AND %s 
            AND Special_Transaction IN ('COLLECTABLE DUE','COLLECTABLE NOT DUE','RETENTION NOT DUE','LC PMT DUE','LC PMT NOT DUE')
        """
        cursor.execute(query, (start_date, end_date))
        result = cursor.fetchone()
        total_value = result['total_value'] if result and result['total_value'] else 0

        return jsonify({"total_value": float(total_value)})
    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/collected/value')
def get_collected_value(request: Request):
    start_date = request.query_params.get('start_date') or request.query_params.get('date')
    end_date = request.query_params.get('end_date') or start_date
    if not start_date:
        return jsonify({"error": "Start date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        query = """
            SELECT SUM(Cheque_Amt) AS total_value 
            FROM collections_data 
            WHERE Posting_Date BETWEEN %s AND %s
        """
        cursor.execute(query, (start_date, end_date))
        result = cursor.fetchone()
        total_value = result['total_value'] if result and result['total_value'] else 0

        return jsonify({"collected_value": float(total_value)})
    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/collections/branches')
def get_branches_value(request: Request):
    start_date = request.query_params.get('start_date') or request.query_params.get('date')
    end_date = request.query_params.get('end_date') or start_date
    if not start_date:
        return jsonify({"error": "Start date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        query = """
            SELECT 
                Description, 
                SUM(Balance) / 100000 AS total_value 
            FROM collections_set 
            WHERE AS_ON_Date BETWEEN %s AND %s 
            AND Special_Transaction IN (
                'COLLECTABLE DUE', 'COLLECTABLE NOT DUE', 'RETENTION NOT DUE', 'LC PMT DUE', 'LC PMT NOT DUE'
            )
            GROUP BY Description
            ORDER BY total_value DESC
        """
        cursor.execute(query, (start_date, end_date))
        results = cursor.fetchall()
        ordered_labels = [row['Description'] for row in results]
        ordered_values = [float(row['total_value']) if row['total_value'] is not None else 0.0 for row in results]

        return jsonify({"labels": ordered_labels, "values": ordered_values})
    except Exception as e:
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/collections/branches-list')
def get_branches_list(request: Request):
    start_date = request.query_params.get('start_date') or request.query_params.get('date')
    end_date = request.query_params.get('end_date') or start_date
    if not start_date:
        return jsonify({"error": "Start date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        query = """
            SELECT DISTINCT Coll_BR_DESC 
            FROM collections_data 
            WHERE Cheque_Amt > 0 AND Posting_Date BETWEEN %s AND %s 
            ORDER BY Coll_BR_DESC
        """
        cursor.execute(query, (start_date, end_date))
        branches = [row['Coll_BR_DESC'] for row in cursor.fetchall()]
        return jsonify(branches)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/collections/mapped-details')
def get_mapped_collection_details(request: Request):
    start_date = request.query_params.get('start_date') or request.query_params.get('date')
    end_date = request.query_params.get('end_date') or start_date
    selected_branch = request.query_params.get('branch')

    if not start_date:
        return jsonify({"error": "Date range is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        branch_clause_set = ""
        branch_clause_data = ""
        params_set = [start_date, end_date]
        params_data = [start_date, end_date]

        if selected_branch and selected_branch != 'All':
            branch_clause_set = "AND Description = %s "
            branch_clause_data = "AND Coll_BR_DESC = %s "
            params_set.append(selected_branch)
            params_data.append(selected_branch)

        # 1. Fetch Outstanding records for matching lookup
        cursor.execute(f"""
            SELECT Name1, Profit_Center, Description as branch, SUM(Balance) AS outstanding 
            FROM collections_set 
            WHERE AS_ON_Date BETWEEN %s AND %s {branch_clause_set}
            AND Special_Transaction IN ('COLLECTABLE DUE', 'COLLECTABLE NOT DUE', 'RETENTION NOT DUE', 'LC PMT DUE', 'LC PMT NOT DUE', 'PMT_ADVANCE', 'RETENTION DUE') 
            GROUP BY Name1, Profit_Center, Description
        """, tuple(params_set))
        outstanding_rows = cursor.fetchall()

        # Build outstanding lookup map keyed by (Customer, Profit_Center)
        outstanding_map = {}
        for o in outstanding_rows:
            key = (o['Name1'], o['Profit_Center'])
            outstanding_map[key] = {
                "branch": o['branch'],
                "outstanding": float(o['outstanding'] or 0)
            }

        # 2. Fetch Collected records (Dominating Source)
        cursor.execute(f"""
            SELECT 
                Coll_BR_DESC AS branch,
                Cust_Name, 
                PROFIT_CENTRE_1,
                SUM(Cheque_Amt) AS collected
            FROM collections_data 
            WHERE Cheque_Amt > 0 AND Posting_Date BETWEEN %s AND %s {branch_clause_data}
            GROUP BY Coll_BR_DESC, Cust_Name, PROFIT_CENTRE_1
        """, tuple(params_data))
        collected_rows = cursor.fetchall()

        merged_results = []

        # Iterate strictly over collected_rows
        for c in collected_rows:
            key = (c['Cust_Name'], c['PROFIT_CENTRE_1'])
            col_amt = float(c['collected'] or 0)

            if key in outstanding_map:
                out_val = outstanding_map[key]['outstanding']
                branch_name = outstanding_map[key]['branch'] or c['branch'] or 'N/A'
                balance_val = out_val - col_amt
            else:
                # Unmapped collections: display collected info, keep outstanding & balance at 0
                out_val = 0.0
                branch_name = c['branch'] or 'N/A'
                balance_val = 0.0

            merged_results.append({
                "branch": branch_name,
                "customer": c['Cust_Name'],
                "PROFIT_CENTRE_1": c['PROFIT_CENTRE_1'],
                "outstanding": out_val,
                "collected": col_amt,
                "balance": balance_val
            })

        merged_results.sort(key=lambda x: str(x['branch']), reverse=True)
        return jsonify(merged_results)

    except Exception as e:
        print(f"Backend Error: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


# @app.route('/api/export/collections_excel')
# def export_collections_excel():
#     start_date = request.args.get('start_date') or request.args.get('date')
#     end_date = request.args.get('end_date') or start_date
#
#     if not start_date:
#         return "Please select a date range", 400
#
#     try:
#         with engine.connect() as conn:
#             query_coll = text("SELECT * FROM collections_data WHERE DATE(Posting_Date) BETWEEN :st AND :et")
#             df_coll = pd.read_sql(query_coll, conn, params={"st": start_date, "et": end_date})
#
#             query_out = text("SELECT * FROM collections_set WHERE DATE(AS_ON_Date) BETWEEN :st AND :et")
#             df_out = pd.read_sql(query_out, conn, params={"st": start_date, "et": end_date})
#
#         if df_coll.empty and df_out.empty:
#             return f"No records found for range {start_date} to {end_date}", 404
#
#         date_cols_coll = ['AS_ON_Date', 'Document_Date', 'Posting_Date', 'Invoice_Date']
#         for col in date_cols_coll:
#             if not df_coll.empty and col in df_coll.columns:
#                 df_coll[col] = pd.to_datetime(df_coll[col], errors='coerce').dt.strftime('%d-%m-%Y')
#
#         date_cols_out = ['AS_ON_Date', 'Invoice_Date', 'Baseline_Payment_Dte', 'Purchase_order_date', 'Posting_Date']
#         for col in date_cols_out:
#             if not df_out.empty and col in df_out.columns:
#                 df_out[col] = pd.to_datetime(df_out[col], errors='coerce').dt.strftime('%d-%m-%Y')
#
#         output = io.BytesIO()
#         with pd.ExcelWriter(output, engine='openpyxl') as writer:
#             df_out.to_excel(writer, index=False, sheet_name='OUTSTANDING')
#             df_coll.to_excel(writer, index=False, sheet_name='COLLECTION')
#
#         output.seek(0)
#         return send_file(
#             output,
#             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
#             as_attachment=True,
#             download_name=f"ZCM_Report_{start_date}_to_{end_date}.xlsx"
#         )
#     except Exception as e:
#         return f"Database Error: {str(e)}", 500


PROFIT_CENTER_MAP = {
    'CFD CORP': ['CORPEXP', 'DUBAIEXP'],
    'Bran O_S': ['KOLEXP', 'FAREXP', 'LUDEXP', 'HYDEXP', 'MUMEXP', 'AHDEXP', 'PUNEXP', 'RPREXP', 'INDEXP', 'CHEEXP',
                 'CTREXP', 'KOCEXP', 'NPREXP', 'DELEXP', 'BBHEXP', 'BPLEXP', 'BRANAHD',
                 'BRANCHE', 'BRANDEL', 'BRANHYD', 'BRANKOL', 'BRANLUK', 'BRANPUN', 'DEHEXP', 'DUREXP', 'GTYEXP',
                 'JABEXP', 'JAMEXP', 'JPREXP', 'JSREXP', 'LUCEXP', 'MADEXP', 'NASEXP', 'PATEXP', 'RUREXP', 'SRTEXP',
                 'VADEXP', 'VIZEXP'],
    'Unit-1 Govanahalli': ['UN01ACM', 'UN01DCM', 'UN01ACG', 'UN01COM', 'UN15MCS', 'UN01MBS', 'UN01TRN'],
    'Unit-2 HUBLI-Special': ['UN02ACM', 'UN02ACG', 'UN02COM', ],
    'Unit-25ACM ACM-STD Motors': ['UN25ACM'],
    'Unit-25ACG ACG&EVM': ['UN25ACG', 'UN25EVM', 'UN25COM'],
    'Unit-7 TUMKUR': ['UN07ACM', 'UN07STG', 'UN07DIE', 'UN07WST', 'UN07COM', 'UN07DCM'],
    'UN05 Mysore Unit': ['UN10SWG', 'UN05OFT', 'UN04ELE', ],
    'Unit-16 PUNE': ['UN16CRT', 'UN16OFT', 'UN17CRT', ],
    'DG Set HUBLI': ['UN06DGS'],
    'Unit-12-20 Spares&Service': ['UN12SPS', 'UN20SPA', 'UN12SSD']
}


@app.get('/mis_report')
def mis_report(request: Request):
    conn = get_db_connection()
    years = []
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Fiscal_Year FROM erm_data ORDER BY Fiscal_Year DESC")
            years = [row['Fiscal_Year'] for row in cursor.fetchall()]
        except Exception:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Fiscal_Year FROM erm_data ORDER BY Fiscal_Year DESC")
            years = [row[0] for row in cursor.fetchall()]
        finally:
            cursor.close()
            conn.close()

    selected_year = request.query_params.get('fin_year', years[0] if years else "")

    return render_template(request,
        'MIS.html',
        years=years,
        selected_year=selected_year,
        username=request.session.get('username', 'User')
    )


@app.get('/api/mis/net-sales-data')
def get_net_sales_data(request: Request):
    # Prioritize 'fin_year' to match your JavaScript getQueryString parameter naming
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    # Single value fallback parsing mechanisms
    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    # Convert text names ("June", "May") into integers (6, 5)
    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month structure value provided: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        net_sales_accounts = [
            700001, 700004, 700002, 700003, 700006, 700011, 700009,
            700021, 700018, 700006, 700014, 700007, 700034,
            700008, 700010, 700013, 700012, 700015, 700016, 700017, 700019,
            700022, 700020, 700023, 700024, 700025, 700026, 700027, 700028,
            700029, 700030, 700031, 700033,
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(net_sales_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause}, 
                Account_Number
        """

        # Secure query serialization array setup
        query_params = []
        query_params.extend(all_pcs)
        query_params.append(selected_year)  # Dynamically fills target 'WHERE Fiscal_Year = %s'
        query_params.extend(month_nums)
        query_params.extend(net_sales_accounts)
        query_params.extend(all_pcs)

        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Net Sales",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"MIS API Execution Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/scrap-sales-data')
def get_scrap_sales_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month structure value: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        scrap_sales_accounts = [700005]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(scrap_sales_accounts))

        # CRITICAL FIX: Selected Account_Number and added it to GROUP BY
        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + scrap_sales_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0

        # CRITICAL FIX: Initializing the dropdown accounts structure
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Scrap Sales",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown  # Sent correctly to client view map
        })

    except Exception as e:
        print(f"Scrap Sales API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/iut-out-data')
def get_iut_out_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        iut_accounts = [900001]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(iut_accounts))

        # CRITICAL FIX: Selected Account_Number and added it to GROUP BY
        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders}) 
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + iut_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0

        # CRITICAL FIX: Initializing the dropdown accounts structure
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "IUT - Out",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown  # Sent correctly to client view map
        })
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/other-income-data')
def get_other_income_data(request: Request):
    selected_year = request.query_params.get('fiscal_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Filters are required"}, status_code=400)

    # Base Standard Accounts for Other Income
    other_income_accounts = [459001, 711001, 712001, 712002, 713001, 714001, 715001, 715003, 715004, 715005, 717001,
                             719001, 720009, 720006, 715006, 446105, 712006, ]
    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(other_income_accounts))

        query = f"""
            SELECT {case_statement}, Account_Number, SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s AND MONTH(Posting_Date) IN ({month_placeholders}) AND Account_Number IN ({account_placeholders})
            GROUP BY {group_by_clause}, Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + other_income_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Other Income",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/rm-consumption-data')
def get_rm_consumption_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month structure value: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        rm_accounts = [
            410001, 410002, 410003, 410004, 410005, 410006, 410007, 410008, 410009,
            410010, 410011, 410014, 410015, 410016, 410017, 410018, 410019, 410020,
            410021, 410022, 410023, 410024, 410101, 410102, 410103, 410104, 410105,
            410106, 410107, 410108, 410109, 410110, 410111, 410112, 410113, 410114,
            410115, 410116, 410117, 410118, 410119, 410120, 410121, 410122, 410123,
            410124, 410125, 410126, 410127, 410128, 410129, 410201, 410202, 410203,
            410204, 410205, 410207, 410208, 410209, 410211, 410212, 410213, 410214,
            410215, 410216, 410217, 410218, 410219, 410301, 410302, 410303, 410304,
            410305, 410306, 410307, 410308, 410309, 410310, 410311, 410312, 410313,
            410314, 410315, 410316, 410317, 410401, 410402, 410403, 410404, 410405,
            410406, 410407, 410408, 410409, 410410, 410411, 410412, 410413, 410414,
            410415, 410416, 410417, 410418, 410419, 410501, 410502, 410503, 410504,
            410505, 410506, 410507, 410508, 410509, 410510, 410511, 410512, 410513,
            410514, 410515, 410601, 410602, 410603, 410604, 410605, 410606, 410607,
            410608, 410609, 410610, 410611, 410612, 410701, 410702, 410703, 410704,
            410705, 410706, 410707, 410708, 410709, 410710, 410711, 410712, 410713,
            410714, 410715, 410716, 410717, 410718, 410719, 410720, 410721, 410722,
            410723, 410724, 410725, 410726, 410727, 410728, 410729, 410730, 410731,
            410732, 410733, 410734, 410735, 410736, 410737, 410738, 410739, 410740,
            438004, 448001, 720001, 410210
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(rm_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + rm_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "RM Consumption",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"RM Consumption API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/employee-expenses-data')
def get_employee_expenses_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        ere_accounts = [
            420000, 420001, 420002, 420003, 420004, 420005, 420006, 420007, 420008, 420009,
            420010, 420011, 420012, 420013, 420016, 421214, 457002, 431006, 441001, 441008, 441021,
            421000, 421001, 421002, 421003, 421004, 421005, 421006, 421007, 421008, 421009,
            422001, 422002, 422003, 422004, 422005, 422006, 420101, 420102, 420103, 420104,
            420105, 420106, 420107, 420108, 420109, 420110, 420111, 420112, 420113, 420114,
            420115, 420116, 420117, 420118, 420119, 420120, 420121, 420122, 420123, 420124,
            420125, 420126, 420127, 420128, 420201, 420202, 420203, 420204, 420205, 420206,
            420207, 420208, 420209, 420210, 420211, 420212, 420213, 420214, 420215, 420216,
            420217, 420218, 420219, 420220, 420221, 420222, 420223, 420224, 420225, 420226,
            420227, 420228, 420229, 420230, 420234, 421102, 421103, 421104, 421105, 421106,
            421107, 421108, 421109, 421110, 421111, 421112, 421113, 421201, 421202, 421204,
            421205, 421206, 421207, 421208, 421209, 421210, 421211, 421212, 422101, 422102,
            422103, 422104, 422105, 422106, 422107, 422108, 422201, 422202, 422203, 422204,
            422205, 422206, 422207, 422208, 422209, 423101, 423102, 423103, 423106, 423108,
            423109, 423110, 423111, 423112, 423113, 423114, 420129, 420231, 420232, 421213,
            453001,
        ]

        # Also capture 'fin_year' parameter to line up with your frontend getQueryString()
        if not selected_year:
            selected_year = request.query_params.get('fin_year')

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(ere_accounts))

        query = f"""
               SELECT 
                   {case_statement},
                   Account_Number,
                   SUM(LC_Amount)/ 100000 AS total_amount
               FROM erm_data
               WHERE Fiscal_Year = %s 
                 AND MONTH(Posting_Date) IN ({month_placeholders})
                 AND Account_Number IN ({account_placeholders})
               GROUP BY 
                   {group_by_clause},
                   Account_Number
           """

        query_params = all_pcs + [selected_year] + month_nums + ere_accounts + all_pcs

        # Ensure database driver maps rows to dictionaries correctly
        # If your default connector doesn't support dictionary configuration natively,
        # change row fetching elements to tuple index syntax (row[0], row[1], row[2])
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        # FIX: Changed PROFIT_CENTER_MAP.items() to .keys() to prevent unhashable structure loop crashes
        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Employee related expenses",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/iut-in-data')
def get_iut_in_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        iut_in_accounts = [900101]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(iut_in_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + iut_in_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "IUT-In",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"IUT-In API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/transport-charges-data')
def get_transport_charges_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        transport_accounts = [
            432001, 432012, 432013, 432701, 432702,
            432703, 432704, 432705, 432706, 432707,
            432708, 432709
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(transport_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + transport_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Transport charges",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Transport Charges API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/finance-charges-data')
def get_finance_charges_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Specified Account numbers for Interest & Finance charges
        finance_accounts = [
            446001, 446002, 446003, 446004,
            454001, 454002, 454003, 454004, 454005, 454006, 454009,
            454011, 454012, 454013, 454014, 454015, 454016, 454017,
            454022, 446104, 453101, 453102, 453104,
            454101, 454102, 454103, 454104, 454105, 454106,
            454107, 454108, 454109, 454110, 454111, 454112,
            454113, 454114, 454115, 454116, 453103
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(finance_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + finance_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Interest & finance charges",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Finance Charges API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/depreciation-data')
def get_depreciation_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Specified Account numbers for Depreciation & Amortization
        depreciation_accounts = [
            458001, 458002, 458003, 458004, 458005, 458006,
            458007, 458009, 458010, 458011, 458013,
            458014, 457101, 458008
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(depreciation_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + depreciation_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Depreciation",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Depreciation API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/revenue-expenses-data')
def get_revenue_expenses_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Full master array of Revenue Expenses Account Numbers provided
        revenue_accounts = [
            430001, 420014, 420015, 430002, 430003, 430004, 430005, 431001, 431002, 431003,
            431004, 431005, 431008, 431009, 432002, 432004, 432005, 432006, 432007, 432008,
            432009, 432010, 432011, 432018, 433001, 434001, 435001, 436001, 437001, 437002,
            437004, 437005, 437006, 437007, 438001, 438003, 438005, 439002, 439003, 439005,
            439102, 440001, 440002, 440003, 441005, 441002, 441003, 441004, 441006, 441007,
            441009, 441010, 441011, 441012, 441013, 441014, 441015, 441016, 441017, 441018,
            441019, 441020, 441022, 441023, 442001, 442002, 442003, 443001, 443002, 444001,
            444002, 444004, 444005, 445001, 446009, 447001, 448002, 448003, 448005, 448006,
            448007, 450001, 456001, 454008, 446005, 446006, 446008, 446007, 720002, 720003,
            449003, 448008, 448009, 457003, 452001, 454020, 454021, 718001, 451001, 426101,
            426102, 426103, 426201, 426202, 426104, 426105, 426106, 426203, 426204, 427101,
            427102, 427201, 427202, 428101, 428102, 428103, 428104, 428105, 428106, 428107,
            428108, 430101, 430102, 430201, 430202, 430301, 430401, 430402, 430403, 430501,
            430502, 431101, 431201, 431301, 431401, 431501, 431601, 431701, 432101, 432102,
            432103, 432201, 432301, 432401, 432402, 432403, 432404, 432405, 432501, 432502,
            432504, 432601, 432602, 432604, 432801, 432802, 432803, 433101, 434101, 435101,
            435102, 435103, 435104, 436101, 437101, 437102, 437201, 437202, 437203, 437204,
            437205, 437206, 438101, 438102, 438103, 438106, 438107, 438108, 438109, 439101,
            439103, 440101, 440102, 440103, 440104, 441101, 441102, 441103, 441104, 441105,
            441106, 441107, 441108, 441109, 441110, 441111, 441112, 441113, 441114, 441115,
            441116, 441117, 441118, 441119, 441120, 441121, 441122, 441123, 442101, 442102,
            442103, 442104, 442105, 442201, 442202, 443101, 443102, 443103, 444101, 444102,
            444103, 444104, 444201, 444202, 444203, 444204, 444205, 444206, 444207, 444208,
            445101, 445102, 445103, 445104, 445105, 445106, 445107, 445201, 445202, 445203,
            445204, 445301, 445302, 446101, 446102, 446103, 447101, 447200, 448101, 449001,
            449101, 450101, 451101, 451102, 452101, 454118, 455101, 456101, 456102, 456103,
            456104, 456105, 456106, 459002, 459003, 460001, 461001, 430203, 432406
        ]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(revenue_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + revenue_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Revenue expenses",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Revenue Expenses API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/ere-allocation-data')
def get_ere_allocation_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Array of account numbers specified for ERE - allocation
        ere_alloc_accounts = [424901, 424902, 424903, 424904]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(ere_alloc_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + ere_alloc_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "ERE - allocation",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"ERE Allocation API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/ore-allocation-data')
def get_ore_allocation_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        ore_alloc_accounts = [425901, 425902, 425903, 425904]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(ore_alloc_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + ore_alloc_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "ORE - allocation",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"ORE Allocation API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/interest-allocation-data')
def get_interest_allocation_data(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Array containing specific account number for Interest allocation
        interest_alloc_accounts = [454121]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(interest_alloc_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + interest_alloc_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Interest - allocation",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Interest Allocation API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/depreciation-allocation-data')
def get_depreciation_allocation_data(request: Request):
    selected_year = request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return jsonify({"error": "Fiscal Year and Month filters are required"}, status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return jsonify({"error": f"Invalid Month name structure: {m}"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Array containing specific account number for Depreciation allocation
        depr_alloc_accounts = [458012]

        case_segments = []
        all_pcs = []
        for unit, pcs in PROFIT_CENTER_MAP.items():
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"
        group_by_clause = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(depr_alloc_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/ 100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY 
                {group_by_clause},
                Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + depr_alloc_accounts + all_pcs
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()

        units = list(PROFIT_CENTER_MAP.keys())
        unit_data = {unit: 0.0 for unit in units}
        unit_data['GRAND TOTAL'] = 0.0
        account_breakdown = {}

        for row in rows:
            b_unit = row['business_unit']
            acct = row['Account_Number']
            amt = float(row['total_amount'] or 0.0)

            if b_unit != 'OTHERS':
                unit_data[b_unit] += amt
                unit_data['GRAND TOTAL'] += amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {u: 0.0 for u in units}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][b_unit] += amt
                account_breakdown[acct]['GRAND TOTAL'] += amt

        return jsonify({
            "row_title": "Depreciation - allocation",
            "unit_totals": unit_data,
            "dropdown_accounts": account_breakdown
        })

    except Exception as e:
        print(f"Depreciation Allocation API Failure: {e}")
        return jsonify({"error": "Internal Server Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


class SchedulerConfig:
    SCHEDULER_API_ENABLED = False
    SCHEDULER_EXECUTORS = {"default": {"type": "threadpool", "max_workers": 1}}


scheduler = BackgroundScheduler(
    executors={"default": ThreadPoolExecutor(SchedulerConfig.SCHEDULER_EXECUTORS["default"]["max_workers"])})
scheduler.start()

print("[SCHEDULER ENGINE] Background task runner initialized successfully.")


def automated_sap_sync_job():
    print("\n" + "═" * 60)
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] AUTOMATED 10-MINUTE DELETE-AND-INSERT SYNC TRIGERRED")
    print("═" * 60)

    with nullcontext():
        sap_url = "https://kecapi.kirloskar-electric.com/sap/opu/odata/sap/ZFAGLFLEXA_SRV/ZFAGLFLEXASet?$format=xlsx"
        username = "corpodata"
        password = "Kec12345"

        try:
            # --- STEP 1: NETWORK REQUEST TO SAP ODATA ---
            print("[AUTO-SYNC STEP 1/4] Establishing secure connection link to live SAP systems...")
            response = requests.get(sap_url, auth=HTTPBasicAuth(username, password), timeout=360)
            print(f"--> Server response received: {response.status_code}")

            if response.status_code != 200:
                print("[CRITICAL AUTOMATED JOB ERROR] Live SAP stream network access drop. Job interrupted.")
                return

            # --- STEP 2: WORKBOOK STREAM PARSING ---
            print("[AUTO-SYNC STEP 2/4] Parsing streaming binary sheet data directly out of RAM...")
            excel_file_bytes = io.BytesIO(response.content)
            wb = openpyxl.load_workbook(excel_file_bytes, data_only=True)
            ws = wb.active

            total_rows_found = ws.max_row - 1
            print(f"--> Stream processed. {total_rows_found} live rows found inside the current stream pool.")

            if total_rows_found <= 0:
                print("[INFO] Incoming data grid array is empty. Closing interval.")
                return

            # --- STEP 3: DETECT CURRENT YEAR/MONTH AND PURGE THE TARGET BLOCK ---
            print("[AUTO-SYNC STEP 3/4] Initializing database transaction context...")
            conn = get_db_connection()
            cursor = conn.cursor()

            # Inspect the first data row (row 2) to dynamically identify the year and month to clear
            first_row = next(ws.iter_rows(min_row=2, max_row=2, values_only=True), None)

            if first_row and first_row[0] is not None and first_row[9] is not None:
                target_fiscal_year = str(first_row[0]).strip()

                # Parse out the month integer safely
                raw_pdate = first_row[9]
                if isinstance(raw_pdate, (datetime, date)):
                    target_month = raw_pdate.month
                else:
                    try:
                        target_month = datetime.strptime(str(raw_pdate).strip()[:10], '%Y-%m-%d').month
                    except ValueError:
                        target_month = None

                if target_fiscal_year and target_month:
                    print(f"--> Target scope detected: Fiscal Year {target_fiscal_year}, Month Integer {target_month}")
                    print(
                        f"--> Clearing existing database entries for the current target scope to prevent rewriting errors...")

                    # Wipes out only the current month's active records; historical entries are safe
                    delete_query = "DELETE FROM erm_data WHERE Fiscal_Year = %s AND MONTH(Posting_Date) = %s"
                    cursor.execute(delete_query, (target_fiscal_year, target_month))
                    print(f"--> Target scope purge complete. Removed {cursor.rowcount} temporary rows.")
                else:
                    print(
                        "[WARNING] Could not parse date bounds safely from row data. Proceeding without safety purge.")
            else:
                print("[WARNING] Spreadsheet data row structure invalid. Proceeding without safety purge.")

            # --- STEP 4: RECORD PROCESSING ENGINE LOOP ---
            print("[AUTO-SYNC STEP 4/4] Executing fresh insertion batch loop...")
            insert_query = """
                INSERT INTO erm_data (
                    Fiscal_Year, Document_Number, Company_Code, Line_Item, Currency, 
                    Account_Number, Profit_Center, LC_Amount, Fiscal_Year1, Posting_Date
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """

            inserted_counter = 0

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or all(v is None for v in row):
                    continue

                fy = str(row[0]).strip() if row[0] is not None else ""
                doc_num = str(row[1]).strip() if row[1] is not None else ""
                comp_code = str(row[2]).strip() if row[2] is not None else ""
                line_item = str(row[3]).strip() if row[3] is not None else ""
                currency = str(row[4]).strip() if row[4] is not None else ""

                # Strip leading zeros from Account Number strings
                raw_acct = str(row[5]).strip() if row[5] is not None else ""
                acct_num = raw_acct.lstrip('0')
                if not acct_num and raw_acct:
                    acct_num = "0"

                profit_ctr = str(row[6]).strip() if row[6] is not None else ""
                fy1 = str(row[8]).strip() if row[8] is not None else ""

                raw_lc_amount = row[7]
                lc_amount_decimal = Decimal('0.00')
                if raw_lc_amount is not None:
                    try:
                        clean_amount = str(raw_lc_amount).strip().replace(',', '')
                        lc_amount_decimal = Decimal(clean_amount)
                    except (InvalidOperation, ValueError):
                        lc_amount_decimal = Decimal('0.00')

                raw_posting_date = row[9]
                posting_date_str = ""
                if raw_posting_date is not None:
                    if isinstance(raw_posting_date, (datetime, date)):
                        posting_date_str = raw_posting_date.strftime('%Y-%m-%d')
                    else:
                        try:
                            parsed_date = datetime.strptime(str(raw_posting_date).strip()[:10], '%Y-%m-%d')
                            posting_date_str = parsed_date.strftime('%Y-%m-%d')
                        except ValueError:
                            posting_date_str = ""

                cursor.execute(insert_query, (
                    fy if fy else None, doc_num if doc_num else None, comp_code if comp_code else None,
                    line_item if line_item else None, currency if currency else None, acct_num if acct_num else None,
                    profit_ctr if profit_ctr else None, lc_amount_decimal, fy1 if fy1 else None,
                    posting_date_str if posting_date_str else None
                ))
                inserted_counter += 1

            conn.commit()

            print("═" * 60)
            print(f"BACKGROUND SCHEDULE JOB PIPELINE METRICS RECONCILED:")
            print(f"--> Total Fresh Rows Overwritten and Saved: {inserted_counter}")
            print("═" * 60 + "\n")

        except Exception as e:
            print(f"\n[BACKGROUND SCHEDULER SYSTEM EXCEPTION CRASH]: {e}")
            if 'conn' in locals(): conn.rollback()
        finally:
            if 'cursor' in locals(): cursor.close()
            if 'conn' in locals(): conn.close()


scheduler.add_job(automated_sap_sync_job, 'interval', id='sap_sync_10min_job', minutes=90, misfire_grace_time=900)

PROFIT_CENTER_MAP_1 = {
    'CFD CORP': ['CORPEXP', 'DUBAIEXP'],
    'Bran O_S': ['KOLEXP', 'FAREXP', 'LUDEXP', 'HYDEXP', 'MUMEXP', 'AHDEXP', 'PUNEXP', 'RPREXP', 'INDEXP', 'CHEEXP',
                 'CTREXP', 'KOCEXP', 'NPREXP', 'DELEXP', 'BBHEXP', 'BPLEXP', 'BRANAHD',
                 'BRANCHE', 'BRANDEL', 'BRANHYD', 'BRANKOL', 'BRANLUK', 'BRANPUN', 'DEHEXP', 'DUREXP', 'GTYEXP',
                 'JABEXP', 'JAMEXP', 'JPREXP', 'JSREXP', 'LUCEXP', 'MADEXP', 'NASEXP', 'PATEXP', 'RUREXP', 'SRTEXP',
                 'VADEXP', 'VIZEXP'],
    'Unit-1 Govanahalli': ['UN01ACM', 'UN01DCM', 'UN01ACG', 'UN01COM', 'UN15MCS', 'UN01MBS', 'UN01TRN'],
    'Unit-2 HUBLI-Special': ['UN02ACM', 'UN02ACG', 'UN02COM', ],
    'Unit-25ACM ACM-STD Motors': ['UN25ACM'],
    'Unit-25ACG ACG&EVM': ['UN25ACG', 'UN25EVM', 'UN25COM'],
    'Unit-7 TUMKUR': ['UN07ACM', 'UN07STG', 'UN07DIE', 'UN07WST', 'UN07COM', 'UN07DCM'],
    'UN05 Mysore Unit': [ 'UN05OFT', 'UN04ELE', ],
    'UN10 SWG': ['UN10SWG',],
    'Unit-16 PUNE': ['UN16CRT', 'UN16OFT', 'UN17CRT', ],
    'DG Set HUBLI': ['UN06DGS'],
    'Unit-12-20 Spares&Service': ['UN12SPS', 'UN20SPA', 'UN12SSD']
}

@app.get('/mis_unit_report')
def mis_unit_report(request: Request):
    conn = get_db_connection()
    years = []
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Fiscal_Year FROM erm_data ORDER BY Fiscal_Year DESC")
            years = [row['Fiscal_Year'] for row in cursor.fetchall()]
        except Exception:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Fiscal_Year FROM erm_data ORDER BY Fiscal_Year DESC")
            years = [row[0] for row in cursor.fetchall()]
        finally:
            cursor.close()
            conn.close()

    selected_year = request.query_params.get('fin_year', years[0] if years else "")

    return render_template(request,
        'MIS_UNIT.html',
        years=years,
        selected_year=selected_year,
        username=request.session.get('username', 'User')
    )


@app.get('/api/mis/unit-monthly-sales')
def get_unit_monthly_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        net_sales_accounts = [
            700001, 700004, 700002, 700003, 700006, 700011, 700009,
            700021, 700018, 700006, 700014, 700007, 700034, 700033,
            700008, 700010, 700013, 700012, 700015, 700016, 700017, 700019,
            700022, 700020, 700023, 700024, 700025, 700026, 700027, 700028,
            700029, 700030, 700031, 7000434
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(net_sales_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + net_sales_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "Net Sales",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })



    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/unit-scrap-sales')
def get_unit_scrap_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        scrap_sales_accounts = [
            700005
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(scrap_sales_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + scrap_sales_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "Scrap Sales",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })



    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/other-income-monthly-sales')
def get_unit_other_income_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        other_income_accounts = [459001, 711001, 712001, 712002, 713001, 714001, 715001, 715003, 715004, 715005, 717001,
                                 719001, 720009, 720006, 715006, 446105, 712006]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(other_income_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + other_income_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "Other income",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })



    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/iut-out-monthly-sales')
def get_unit_iut_out_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        iut_out_accounts = [900001]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(iut_out_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + iut_out_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "iut_out",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })



    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/rm-consp-monthly-sales')
def get_unit_rm_consp_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        rm_consp_accounts = [
            410001, 410002, 410003, 410004, 410005, 410006, 410007, 410008, 410009,
            410010, 410011, 410014, 410015, 410016, 410017, 410018, 410019, 410020,
            410021, 410022, 410023, 410024, 410101, 410102, 410103, 410104, 410105,
            410106, 410107, 410108, 410109, 410110, 410111, 410112, 410113, 410114,
            410115, 410116, 410117, 410118, 410119, 410120, 410121, 410122, 410123,
            410124, 410125, 410126, 410127, 410128, 410129, 410201, 410202, 410203,
            410204, 410205, 410207, 410208, 410209, 410211, 410212, 410213, 410214,
            410215, 410216, 410217, 410218, 410219, 410301, 410302, 410303, 410304,
            410305, 410306, 410307, 410308, 410309, 410310, 410311, 410312, 410313,
            410314, 410315, 410316, 410317, 410401, 410402, 410403, 410404, 410405,
            410406, 410407, 410408, 410409, 410410, 410411, 410412, 410413, 410414,
            410415, 410416, 410417, 410418, 410419, 410501, 410502, 410503, 410504,
            410505, 410506, 410507, 410508, 410509, 410510, 410511, 410512, 410513,
            410514, 410515, 410601, 410602, 410603, 410604, 410605, 410606, 410607,
            410608, 410609, 410610, 410611, 410612, 410701, 410702, 410703, 410704,
            410705, 410706, 410707, 410708, 410709, 410710, 410711, 410712, 410713,
            410714, 410715, 410716, 410717, 410718, 410719, 410720, 410721, 410722,
            410723, 410724, 410725, 410726, 410727, 410728, 410729, 410730, 410731,
            410732, 410733, 410734, 410735, 410736, 410737, 410738, 410739, 410740,
            438004, 448001, 720001, 410210
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(rm_consp_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + rm_consp_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "rm_consumption",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })


    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/employee-monthly-sales')
def get_unit_employee_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        employee_accounts = [
            420000, 420001, 420002, 420003, 420004, 420005, 420006, 420007, 420008, 420009,
            420010, 420011, 420012, 420013, 420016, 421214, 457002, 431006, 441001, 441008, 441021,
            421000, 421001, 421002, 421003, 421004, 421005, 421006, 421007, 421008, 421009,
            422001, 422002, 422003, 422004, 422005, 422006, 420101, 420102, 420103, 420104,
            420105, 420106, 420107, 420108, 420109, 420110, 420111, 420112, 420113, 420114,
            420115, 420116, 420117, 420118, 420119, 420120, 420121, 420122, 420123, 420124,
            420125, 420126, 420127, 420128, 420201, 420202, 420203, 420204, 420205, 420206,
            420207, 420208, 420209, 420210, 420211, 420212, 420213, 420214, 420215, 420216,
            420217, 420218, 420219, 420220, 420221, 420222, 420223, 420224, 420225, 420226,
            420227, 420228, 420229, 420230, 420234, 421102, 421103, 421104, 421105, 421106,
            421107, 421108, 421109, 421110, 421111, 421112, 421113, 421201, 421202, 421204,
            421205, 421206, 421207, 421208, 421209, 421210, 421211, 421212, 422101, 422102,
            422103, 422104, 422105, 422106, 422107, 422108, 422201, 422202, 422203, 422204,
            422205, 422206, 422207, 422208, 422209, 423101, 423102, 423103, 423106, 423108,
            423109, 423110, 423111, 423112, 423113, 423114, 420129, 420231, 420232, 421213,
            453001,
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(employee_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + employee_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "employee related expenses",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })

    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/iut-in-monthly-sales')
def get_unit_iut_in_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        iut_in_accounts = [900101]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(iut_in_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + iut_in_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "iut_in",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })

    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/transport-monthly-sales')
def get_unit_transport_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        transport_accounts = [
            432001, 432012, 432013, 432701, 432702,
            432703, 432704, 432705, 432706, 432707,
            432708, 432709
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(transport_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + transport_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "transposrt_charges",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })

    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/finance-cahrges-monthly-sales')
def get_unit_finance_charges_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        finance_charges_accounts = [
            446001, 446002, 446003, 446004, 454006, 454009,
            454001, 454002, 454003, 454004, 454005, 454017,
            454011, 454012, 454013, 454014, 454015, 454016,
            454022, 446104, 453101, 453102, 453104,
            454101, 454102, 454103, 454104, 454105, 454106,
            454107, 454108, 454109, 454110, 454111, 454112,
            454113, 454114, 454115, 454116, 453103
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(finance_charges_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + finance_charges_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "finance_charges",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })

    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/depreciation-monthly-sales')
def get_unit_depreciation_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        depreciation_accounts = [
            458001, 458002, 458003, 458004, 458005, 458006,
            458007, 458009, 458010, 458011, 458013,
            458014, 457101, 458008
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(depreciation_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + depreciation_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "depreciation",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })
    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/revenue-expenses-monthly-sales')
def get_unit_revenue_expenses_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        revenue_expenses_accounts = [
            430001, 420014, 420015, 430002, 430003, 430004, 430005, 431001, 431002, 431003,
            431004, 431005, 431008, 431009, 432002, 432004, 432005, 432006, 432007, 432008,
            432009, 432010, 432011, 432018, 433001, 434001, 435001, 436001, 437001, 437002,
            437004, 437005, 437006, 437007, 438001, 438003, 438005, 439002, 439003, 439005,
            439102, 440001, 440002, 440003, 441005, 441002, 441003, 441004, 441006, 441007,
            441009, 441010, 441011, 441012, 441013, 441014, 441015, 441016, 441017, 441018,
            441019, 441020, 441022, 441023, 442001, 442002, 442003, 443001, 443002, 444001,
            444002, 444004, 444005, 445001, 446009, 447001, 448002, 448003, 448005, 448006,
            448007, 450001, 456001, 454008, 446005, 446006, 446008, 446007, 720002, 720003,
            449003, 448008, 448009, 457003, 452001, 454020, 454021, 718001, 451001, 426101,
            426102, 426103, 426201, 426202, 426104, 426105, 426106, 426203, 426204, 427101,
            427102, 427201, 427202, 428101, 428102, 428103, 428104, 428105, 428106, 428107,
            428108, 430101, 430102, 430201, 430202, 430301, 430401, 430402, 430403, 430501,
            430502, 431101, 431201, 431301, 431401, 431501, 431601, 431701, 432101, 432102,
            432103, 432201, 432301, 432401, 432402, 432403, 432404, 432405, 432501, 432502,
            432504, 432601, 432602, 432604, 432801, 432802, 432803, 433101, 434101, 435101,
            435102, 435103, 435104, 436101, 437101, 437102, 437201, 437202, 437203, 437204,
            437205, 437206, 438101, 438102, 438103, 438106, 438107, 438108, 438109, 439101,
            439103, 440101, 440102, 440103, 440104, 441101, 441102, 441103, 441104, 441105,
            441106, 441107, 441108, 441109, 441110, 441111, 441112, 441113, 441114, 441115,
            441116, 441117, 441118, 441119, 441120, 441121, 441122, 441123, 442101, 442102,
            442103, 442104, 442105, 442201, 442202, 443101, 443102, 443103, 444101, 444102,
            444103, 444104, 444201, 444202, 444203, 444204, 444205, 444206, 444207, 444208,
            445101, 445102, 445103, 445104, 445105, 445106, 445107, 445201, 445202, 445203,
            445204, 445301, 445302, 446101, 446102, 446103, 447101, 447200, 448101, 449001,
            449101, 450101, 451101, 451102, 452101, 454118, 455101, 456101, 456102, 456103,
            456104, 456105, 456106, 459002, 459003, 460001, 461001, 430203, 432406
        ]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(revenue_expenses_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + revenue_expenses_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "revenue_expenses",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })

    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/ere-allocation-monthly-sales')
def get_unit_ere_allocation_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        ere_allocation_accounts = [424901, 424902, 424903, 424904]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(ere_allocation_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + ere_allocation_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "ere_allocation",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })
    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/ore-allocation-monthly-sales')
def get_unit_ore_allocation_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        ore_allocation_accounts = [425901, 425902, 425903, 425904]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(ore_allocation_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + ore_allocation_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "ore_allocation",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })
    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/intrest-allocation-monthly-sales')
def get_unit_intrest_allocation_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        intrest_allocation_accounts = [454121]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(intrest_allocation_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + intrest_allocation_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "interest_allocation",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })
    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/dep-allocation-monthly-sales')
def get_unit_dep_allocation_sales(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year')
    selected_unit = request.query_params.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}, status_code=400)

    profit_centers = PROFIT_CENTER_MAP_1.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}, status_code=400)

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        dep_allocation_accounts = [458012]

        # Dynamic query string parameters composition array
        pc_placeholders = ", ".join(["%s"] * len(profit_centers))
        account_placeholders = ", ".join(["%s"] * len(dep_allocation_accounts))

        # Core Query: Extract transaction totals grouped by chronological calendar month index
        query = f"""
            SELECT 
                MONTH(Posting_Date) AS month_idx,
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s
              AND Profit_Center IN ({pc_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY MONTH(Posting_Date), Account_Number
        """

        query_params = [selected_year] + profit_centers + dep_allocation_accounts
        cursor.execute(query, tuple(query_params))
        rows = cursor.fetchall()
        fiscal_months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]

        # Initialize flat dictionary keys mapping tracking slots cleanly
        monthly_totals = {m: 0.0 for m in fiscal_months}
        account_breakdown = {}

        for row in rows:
            if is_dict:
                m_idx = int(row['month_idx'])
                acct = int(row['Account_Number'])
                amt = float(row['total_amount'] or 0.0)
            else:
                m_idx = int(row[0])
                acct = int(row[1])
                amt = float(row[2] or 0.0)

            # Apply absolute inversion sign process as requested in step 2
            pos_amt = amt

            if m_idx in monthly_totals:
                monthly_totals[m_idx] += pos_amt

                if acct not in account_breakdown:
                    account_breakdown[acct] = {m: 0.0 for m in fiscal_months}
                    account_breakdown[acct]['GRAND TOTAL'] = 0.0

                account_breakdown[acct][m_idx] += pos_amt
                account_breakdown[acct]['GRAND TOTAL'] += pos_amt

        # Final map reconstruction to map indices cleanly to readable key names
        month_name_map = {
            4: "April", 5: "May", 6: "June", 7: "July", 8: "August", 9: "September",
            10: "October", 11: "November", 12: "December", 1: "January", 2: "February", 3: "March"
        }

        formatted_monthly_totals = {}
        grand_total = 0.0
        for m in fiscal_months:
            name = month_name_map[m]
            val = monthly_totals[m]
            formatted_monthly_totals[name] = val
            grand_total += val

        formatted_monthly_totals['GRAND TOTAL'] = grand_total

        # Re-map ledger sub-accounts breakdowns structure similarly
        formatted_account_breakdown = {}
        for acct, m_map in account_breakdown.items():
            formatted_account_breakdown[acct] = {}
            for m in fiscal_months:
                name = month_name_map[m]
                formatted_account_breakdown[acct][name] = m_map[m]
            formatted_account_breakdown[acct]['GRAND TOTAL'] = m_map['GRAND TOTAL']

        return jsonify({
            "row_title": "interest_allocation",
            "monthly_totals": formatted_monthly_totals,
            "dropdown_accounts": formatted_account_breakdown
        })
    except Exception as e:
        print(f"[UNIT MONTHLY MIS API CRASH]: {e}")
        return jsonify({"error": "Internal Processing Error"}, status_code=500)
    finally:
        cursor.close()
        conn.close()


@app.get('/api/mis/download-erm-excel')
def download_erm_excel(request: Request):
    selected_year = request.query_params.get('fin_year') or request.query_params.get('fiscal_year') or request.query_params.get('fis_year')
    selected_months = request.query_params.getlist('months[]') or request.query_params.getlist('month')

    if not selected_months and request.query_params.get('month'):
        selected_months = [request.query_params.get('month')]

    if not selected_year or not selected_months:
        return HTMLResponse("Fiscal Year and Month filters are required", status_code=400)

    month_nums = []
    for m in selected_months:
        try:
            month_nums.append(datetime.strptime(m.strip(), "%B").month)
        except ValueError:
            return HTMLResponse(f"Invalid Month name structure: {m}", status_code=400)

    # 1. DEFINE MATRIX STRUCTURE MATCHING FRONTEND COLS
    business_units = [
        'CFD CORP', 'Bran O_S','Unit-1 Govanahalli','Unit-2 HUBLI-Special',
        'Unit-25ACM ACM-STD Motors','Unit-25ACG ACG&EVM','Unit-7 TUMKUR',
        'UN05 Mysore Unit','Unit-16 PUNE','DG Set HUBLI','Unit-12-20 Spares&Service'
    ],
    business_units_2 = [
        'Corporate', 'Branches', 'Unit-1 HVM, HVG, DCM, RRM', 'Unit-2 LVM-Special','Unit-25 LVM-STD',
        'Unit-25 LVG & EVM','Unit-7 STG' , 'Unit-5 TRF(Mysore)','Unit16 TRF(Pune)','Unit-6 DG Sets',
        'Unit-12-20 Spares&Service'
    ]
    # Define General Ledger Account groupings exactly as configured in previous steps
    account_groups = {
        'net_sales': [ 700001, 700004, 700002, 700003, 700006, 700011, 700009,
            700021, 700018, 700006, 700014, 700007, 700034,
            700008, 700010, 700013, 700012, 700015, 700016, 700017, 700019,
            700022, 700020, 700023, 700024, 700025, 700026, 700027, 700028,
            700029, 700030, 700031, 700033,
            ],
        'scrap_sales': [700005],
        'iut_out': [900001],
        'other_income': [459001, 711001, 712001, 712002, 713001, 714001, 715001, 715003, 715004, 715005, 717001,
                             719001, 720009, 720006, 715006, 446105, 712006,],
        'rm_consumption': [
            410001, 410002, 410003, 410004, 410005, 410006, 410007, 410008, 410009,
            410010, 410011, 410014, 410015, 410016, 410017, 410018, 410019, 410020,
            410021, 410022, 410023, 410024, 410101, 410102, 410103, 410104, 410105,
            410106, 410107, 410108, 410109, 410110, 410111, 410112, 410113, 410114,
            410115, 410116, 410117, 410118, 410119, 410120, 410121, 410122, 410123,
            410124, 410125, 410126, 410127, 410128, 410129, 410201, 410202, 410203,
            410204, 410205, 410207, 410208, 410209, 410211, 410212, 410213, 410214,
            410215, 410216, 410217, 410218, 410219, 410301, 410302, 410303, 410304,
            410305, 410306, 410307, 410308, 410309, 410310, 410311, 410312, 410313,
            410314, 410315, 410316, 410317, 410401, 410402, 410403, 410404, 410405,
            410406, 410407, 410408, 410409, 410410, 410411, 410412, 410413, 410414,
            410415, 410416, 410417, 410418, 410419, 410501, 410502, 410503, 410504,
            410505, 410506, 410507, 410508, 410509, 410510, 410511, 410512, 410513,
            410514, 410515, 410601, 410602, 410603, 410604, 410605, 410606, 410607,
            410608, 410609, 410610, 410611, 410612, 410701, 410702, 410703, 410704,
            410705, 410706, 410707, 410708, 410709, 410710, 410711, 410712, 410713,
            410714, 410715, 410716, 410717, 410718, 410719, 410720, 410721, 410722,
            410723, 410724, 410725, 410726, 410727, 410728, 410729, 410730, 410731,
            410732, 410733, 410734, 410735, 410736, 410737, 410738, 410739, 410740,
            438004, 448001, 720001, 410210
        ],
        'iut_in': [900101],
        'employee_expenses': [
            420000, 420001, 420002, 420003, 420004, 420005, 420006, 420007, 420008, 420009,
            420010, 420011, 420012, 420013, 420016, 421214, 457002, 431006, 441001, 441008, 441021,
            421000, 421001, 421002, 421003, 421004, 421005, 421006, 421007, 421008, 421009,
            422001, 422002, 422003, 422004, 422005, 422006, 420101, 420102, 420103, 420104,
            420105, 420106, 420107, 420108, 420109, 420110, 420111, 420112, 420113, 420114,
            420115, 420116, 420117, 420118, 420119, 420120, 420121, 420122, 420123, 420124,
            420125, 420126, 420127, 420128, 420201, 420202, 420203, 420204, 420205, 420206,
            420207, 420208, 420209, 420210, 420211, 420212, 420213, 420214, 420215, 420216,
            420217, 420218, 420219, 420220, 420221, 420222, 420223, 420224, 420225, 420226,
            420227, 420228, 420229, 420230, 420234, 421102, 421103, 421104, 421105, 421106,
            421107, 421108, 421109, 421110, 421111, 421112, 421113, 421201, 421202, 421204,
            421205, 421206, 421207, 421208, 421209, 421210, 421211, 421212, 422101, 422102,
            422103, 422104, 422105, 422106, 422107, 422108, 422201, 422202, 422203, 422204,
            422205, 422206, 422207, 422208, 422209, 423101, 423102, 423103, 423106, 423108,
            423109, 423110, 423111, 423112, 423113, 423114, 420129, 420231, 420232, 421213,
            453001,
        ],
        'revenue_expenses': [
            430001, 420014, 420015, 430002, 430003, 430004, 430005, 431001, 431002, 431003,
            431004, 431005, 431008, 431009, 432002, 432004, 432005, 432006, 432007, 432008,
            432009, 432010, 432011, 432018, 433001, 434001, 435001, 436001, 437001, 437002,
            437004, 437005, 437006, 437007, 438001, 438003, 438005, 439002, 439003, 439005,
            439102, 440001, 440002, 440003, 441005, 441002, 441003, 441004, 441006, 441007,
            441009, 441010, 441011, 441012, 441013, 441014, 441015, 441016, 441017, 441018,
            441019, 441020, 441022, 441023, 442001, 442002, 442003, 443001, 443002, 444001,
            444002, 444004, 444005, 445001, 446009, 447001, 448002, 448003, 448005, 448006,
            448007, 450001, 456001, 454008, 446005, 446006, 446008, 446007, 720002, 720003,
            449003, 448008, 448009, 457003, 452001, 454020, 454021, 718001, 451001, 426101,
            426102, 426103, 426201, 426202, 426104, 426105, 426106, 426203, 426204, 427101,
            427102, 427201, 427202, 428101, 428102, 428103, 428104, 428105, 428106, 428107,
            428108, 430101, 430102, 430201, 430202, 430301, 430401, 430402, 430403, 430501,
            430502, 431101, 431201, 431301, 431401, 431501, 431601, 431701, 432101, 432102,
            432103, 432201, 432301, 432401, 432402, 432403, 432404, 432405, 432501, 432502,
            432504, 432601, 432602, 432604, 432801, 432802, 432803, 433101, 434101, 435101,
            435102, 435103, 435104, 436101, 437101, 437102, 437201, 437202, 437203, 437204,
            437205, 437206, 438101, 438102, 438103, 438106, 438107, 438108, 438109, 439101,
            439103, 440101, 440102, 440103, 440104, 441101, 441102, 441103, 441104, 441105,
            441106, 441107, 441108, 441109, 441110, 441111, 441112, 441113, 441114, 441115,
            441116, 441117, 441118, 441119, 441120, 441121, 441122, 441123, 442101, 442102,
            442103, 442104, 442105, 442201, 442202, 443101, 443102, 443103, 444101, 444102,
            444103, 444104, 444201, 444202, 444203, 444204, 444205, 444206, 444207, 444208,
            445101, 445102, 445103, 445104, 445105, 445106, 445107, 445201, 445202, 445203,
            445204, 445301, 445302, 446101, 446102, 446103, 447101, 447200, 448101, 449001,
            449101, 450101, 451101, 451102, 452101, 454118, 455101, 456101, 456102, 456103,
            456104, 456105, 456106, 459002, 459003, 460001, 461001, 430203, 432406
        ],
        'transport_charges': [432001, 432012, 432013, 432701, 432702,432703, 432704, 432705, 432706, 432707,
            432708, 432709],
        'finance_charges': [ 446001, 446002, 446003, 446004,
            454001, 454002, 454003, 454004, 454005, 454006, 454009,
            454011, 454012, 454013, 454014, 454015, 454016, 454017,
            454022, 446104, 453101, 453102, 453104,
            454101, 454102, 454103, 454104, 454105, 454106,
            454107, 454108, 454109, 454110, 454111, 454112,
            454113, 454114, 454115, 454116, 453103],
        'depreciation': [458001, 458002, 458003, 458004, 458005, 458006,
            458007, 458009, 458010, 458011, 458013, 458014, 457101, 458008],
        'ere_allocation': [424901, 424902, 424903, 424904],
        'ore_allocation': [425901, 425902, 425903, 425904],
        'interest_allocation': [454121],
        'depreciation_allocation': [458012]
    }

    # Extract single linear listing of all accounts to push to database criteria placeholder
    flattened_accounts = []
    for group, accts in account_groups.items():
        flattened_accounts.extend(accts)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Build SQL dynamic profit center CASE segments
        case_segments = []
        all_pcs = []
        for unit in business_units:
            pcs = PROFIT_CENTER_MAP.get(unit, [])
            if pcs:
                pc_placeholders = ", ".join(["%s"] * len(pcs))
                case_segments.append(f"WHEN Profit_Center IN ({pc_placeholders}) THEN '{unit}'")
                all_pcs.extend(pcs)

        case_statement = "CASE " + " ".join(case_segments) + " ELSE 'OTHERS' END AS business_unit"

        month_placeholders = ", ".join(["%s"] * len(month_nums))
        account_placeholders = ", ".join(["%s"] * len(flattened_accounts))

        query = f"""
            SELECT 
                {case_statement},
                Account_Number,
                SUM(LC_Amount)/100000 AS total_amount
            FROM erm_data
            WHERE Fiscal_Year = %s 
              AND MONTH(Posting_Date) IN ({month_placeholders})
              AND Account_Number IN ({account_placeholders})
            GROUP BY business_unit, Account_Number
        """

        query_params = all_pcs + [selected_year] + month_nums + flattened_accounts
        cursor.execute(query, tuple(query_params))
        db_rows = cursor.fetchall()

        # Organize pulled rows into direct cross matrix state
        matrix_data = {group: {u: 0.0 for u in business_units} for group in account_groups.keys()}

        for r in db_rows:
            bu = r['business_unit']
            acct = int(r['Account_Number'])
            amt = float(r['total_amount'] or 0.0)

            if bu in business_units:
                for group, accts in account_groups.items():
                    if acct in accts:
                        matrix_data[group][bu] += amt

        # 2. GENERATE NATIVE EXCEL WORKBOOK VIA OPENPYXL
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ERM Report"
        ws.views.sheetView[0].showGridLines = True

        # Layout Styling Palettes
        font_family = "Segoe UI"
        navy_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        subtotal_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
        pbt_blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        accent_yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        title_font = Font(name=font_family, size=16, bold=True, color="1F4E78")
        meta_font = Font(name=font_family, size=10, italic=True)
        header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
        bold_font = Font(name=font_family, size=10, bold=True)
        regular_font = Font(name=font_family, size=10)

        thin_border_side = Side(border_style="thin", color="D3D3D3")
        thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side,
                             bottom=thin_border_side)
        double_bottom_border = Border(top=thin_border_side, bottom=Side(border_style="double", color="000000"),
                                      left=thin_border_side, right=thin_border_side)

        # Title block headers
        ws['A1'] = "ERM OF MONTH"
        ws['A3'].font = meta_font
        ws['A1'].font = title_font
        ws['A3'] = f"Fiscal Year: {selected_year} | Filtered Months: {', '.join(selected_months)}"
        ws['A3'].font = meta_font
        ws.row_dimensions[1].height = 20
        ws.row_dimensions[2].height = 16

        # Build column grid header names
        headers = ["Particulars"] + business_units + ["TOTAL"]
        for col_idx, h_text in enumerate(headers, start=1):
            cell = ws.cell(row=4, column=col_idx, value=h_text)
            cell.font = header_font
            cell.fill = navy_header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[4].height = 28

        row_definitions = [
            {'title': 'A. Income', 'type': 'header'},
            {'title': 'Net Sales', 'type': 'data', 'key': 'net_sales','transform': 'abs'},
            {'title': 'Scrap Sales', 'type': 'data', 'key': 'scrap_sales','transform': 'abs'},
            {'title': 'IUT - Out', 'type': 'data', 'key': 'iut_out','transform': 'abs'},
            {'title': 'Other Income', 'type': 'data', 'key': 'other_income','transform': 'abs'},
            {'title': 'Total Income - A', 'type': 'formula', 'formula': '= ABS({col}6+{col}7+{col}8+{col}9)',
             'fill': subtotal_fill},
            {'title': 'B. Expenditure', 'type': 'header'},
            {'title': 'RM Consumption', 'type': 'data', 'key': 'rm_consumption'},
            {'title': 'IUT- In', 'type': 'data', 'key': 'iut_in'},
            {'title': 'Employee related expenses', 'type': 'data', 'key': 'employee_expenses'},
            {'title': 'Revenue expenses', 'type': 'data', 'key': 'revenue_expenses'},
            {'title': 'Transport charges', 'type': 'data', 'key': 'transport_charges'},
            {'title': 'Interest & finance charges', 'type': 'data', 'key': 'finance_charges'},
            {'title': 'Depreciation', 'type': 'data', 'key': 'depreciation'},
            {'title': 'Total Expenditure - B', 'type': 'formula', 'formula': '={col}12+{col}13+{col}14+{col}15+{col}16+{col}17+{col}18',
             'fill': subtotal_fill},
            {'title': 'Gross Profit/(Loss)', 'type': 'formula', 'formula': '= ({col}6+{col}7+{col}8)-({col}12+{col}13)'},
            {'title': 'EBIDTA from operation (Bfor allocn)', 'type': 'formula',
             'formula': '={col}22+{col}18+{col}17-{col}9'},
            {'title': 'PBT before allocation [C=A-B]', 'type': 'formula', 'formula': '={col}10-{col}19',
             'fill': pbt_blue_fill, 'border': double_bottom_border},
            {'title': 'D. Corp & branch allocation', 'type': 'header'},
            {'title': 'ERE - allocation', 'type': 'data', 'key': 'ere_allocation'},
            {'title': 'ORE - allocation', 'type': 'data', 'key': 'ore_allocation'},
            {'title': 'Interest - allocation', 'type': 'data', 'key': 'interest_allocation'},
            {'title': 'Depreciation - allocation', 'type': 'data', 'key': 'depreciation_allocation'},
            {'title': 'Total allocation - D', 'type': 'formula', 'formula': '=SUM({col}24+{col}25+{col}26+{col}27)',
             'fill': subtotal_fill},
            {'title': 'EBIDTA from operation (after allocn)', 'type': 'formula',
             'formula': '={col}30+{col}27+{col}26+{col}18+{col}17-{col}9'},
            {'title': 'PBT after allocation [ E=C-D ]', 'type': 'formula', 'formula': '={col}22-{col}28',
             'fill': pbt_blue_fill, 'border': double_bottom_border},
            {'title': '% to total income before allocation', 'type': 'header'},
            {'title': 'RM Consumption %', 'type': 'ratio', 'ratio_type': 'rm_pct'},
            {'title': 'RM Consumption including freight %', 'type': 'ratio', 'ratio_type': 'rmf_pct'},
            {'title': 'Employee related expenses %', 'type': 'ratio', 'ratio_type': 'ere_pct'},
            {'title': 'Revenue expenses %', 'type': 'ratio', 'ratio_type': 'rev_pct'},
            {'title': 'Transport charges %', 'type': 'ratio', 'ratio_type': 'tc_pct'},
            {'title': 'Interest & finance charges %', 'type': 'ratio', 'ratio_type': 'fc_pct'},
            {'title': 'Depreciation %', 'type': 'ratio', 'ratio_type': 'dep_pct'},
            {'title': 'PBT before allocation %', 'type': 'ratio', 'ratio_type': 'pbt_pct'}
        ]

        # Financial Number Formatting Pattern (Accounting comma separation with parenthetical credits)
        currency_format = '#,##0.00;(#,##0.00);"0.00";@'
        percent_format = '0.00"%";(0.00"%");"0.00%"'

        current_row = 5
        for r_def in row_definitions:
            # Type A: Section Main Category Headers
            if r_def['type'] == 'header':
                cell = ws.cell(row=current_row, column=1, value=r_def['title'])
                cell.font = bold_font
                ws.row_dimensions[current_row].height = 20
                # Border outline fill across the row length
                for c in range(1, len(business_units) + 3):
                    ws.cell(row=current_row, column=c).border = thin_border
                current_row += 1
                continue

            # Title Particular column allocation
            title_cell = ws.cell(row=current_row, column=1, value=r_def['title'])
            title_cell.font = bold_font if r_def['type'] in ['formula', 'ratio'] else regular_font
            title_cell.border = r_def.get('border', thin_border)
            if 'fill' in r_def:
                title_cell.fill = r_def['fill']

            # Populate business unit columns sequentially
            for u_idx, unit in enumerate(business_units, start=2):
                col_letter = get_column_letter(u_idx)
                cell = ws.cell(row=current_row, column=u_idx)
                cell.border = r_def.get('border', thin_border)
                if 'fill' in r_def:
                    cell.fill = r_def['fill']

                # Type B: Standard raw transactional account metrics
                # Type B: Standard raw transactional account metrics
                if r_def['type'] == 'data':
                    val = matrix_data[r_def['key']][unit]
                    if r_def.get('transform') == 'abs' and val is not None:
                        val = abs(val)
                    cell.value = val
                    cell.number_format = currency_format
                    cell.font = regular_font

                # Type C: Cross horizontal summary math row strings
                elif r_def['type'] == 'formula':
                    cell.value = r_def['formula'].format(col=col_letter)
                    cell.number_format = currency_format
                    cell.font = bold_font

                # Type D: Advanced operational efficiency management ratio percentages
                elif r_def['type'] == 'ratio':
                    cell.font = bold_font if r_def['title'] == 'PBT before allocation %' else regular_font
                    if unit in ['CFD CORP', 'Bran O_S']:
                        cell.value = " "
                        cell.alignment = Alignment(horizontal="center")
                    else:
                        rtype = r_def['ratio_type']
                        if rtype == 'rm_pct':
                            cell.value = f"=(({col_letter}12+{col_letter}13-{col_letter}8)/({col_letter}6+{col_letter}7))*100"
                        elif rtype == 'rmf_pct':
                            cell.value = f"=(({col_letter}12+{col_letter}13-{col_letter}8+{col_letter}16)/({col_letter}6+{col_letter}7))*100"
                        elif rtype == 'ere_pct':
                            cell.value = f"=({col_letter}14/{col_letter}10)*100"
                        elif rtype == 'rev_pct':
                            cell.value = f"=({col_letter}15/{col_letter}10)*100"
                        elif rtype == 'tc_pct':
                            cell.value = f"=({col_letter}16/({col_letter}6+{col_letter}8))*100"
                        elif rtype == 'fc_pct':
                            cell.value = f"=({col_letter}17/{col_letter}10)*100"
                        elif rtype == 'dep_pct':
                            cell.value = f"=({col_letter}18/{col_letter}10)*100"
                        elif rtype == 'pbt_pct':
                            cell.value = f"=({col_letter}22/{col_letter}10)*100"
                            if 'fill' not in r_def:
                                cell.fill = pbt_blue_fill

                        # Strip standard percentage format multipliers inside OpenPyXL calculations
                        # cell.value = cell.value.replace('*100', '')
                        cell.number_format = percent_format

            # 3. BUILD MATRIX GRAND TOTAL SUMMARY COL (Last Right Column Block)
            tot_col_idx = len(business_units) + 2
            tot_col_letter = get_column_letter(tot_col_idx)
            tot_cell = ws.cell(row=current_row, column=tot_col_idx)
            tot_cell.font = bold_font
            tot_cell.border = r_def.get('border', thin_border)
            tot_cell.fill = accent_yellow_fill

            if r_def['type'] == 'data':
                tot_cell.value = f"=SUM(B{current_row}:{get_column_letter(tot_col_idx - 1)}{current_row})"
                tot_cell.number_format = currency_format
            elif r_def['type'] == 'formula':
                tot_cell.value = r_def['formula'].format(col=tot_col_letter)
                tot_cell.number_format = currency_format

            elif r_def['type'] == 'ratio':
                rtype = r_def['ratio_type']
                # Grand total calculations matching the rules we designed previously
                if rtype == 'rm_pct':
                    tot_cell.value = f"=(SUM(D12:M12)/SUM(N6+N7))*100"
                elif rtype == 'rmf_pct':
                    tot_cell.value = f"=((SUM(D12:M12)+SUM(D16:M16))/SUM(N6+N7))*100"
                elif rtype == 'ere_pct':
                    tot_cell.value = f"=({tot_col_letter}14/{tot_col_letter}10)*100"
                elif rtype == 'rev_pct':
                    tot_cell.value = f"=({tot_col_letter}15/{tot_col_letter}10)*100"
                elif rtype == 'tc_pct':
                    tot_cell.value = f"=(SUM(D16:L16)/(SUM(D6:L6)+SUM(D8:L8)))*100"
                elif rtype == 'fc_pct':
                    tot_cell.value = f"=({tot_col_letter}17/{tot_col_letter}10)*100"
                elif rtype == 'dep_pct':
                    tot_cell.value = f"=({tot_col_letter}18/{tot_col_letter}10)*100"
                elif rtype == 'pbt_pct':
                    tot_cell.value = f"=({tot_col_letter}22/{tot_col_letter}10)*100"
                tot_cell.number_format = percent_format

            ws.row_dimensions[current_row].height = 20
            current_row += 1

        # Adjust tracking grid column auto width constraints to clear formatting overflows
        ws.freeze_panes = "B5"  # Freezes Particulars column and headers row
        ws.column_dimensions['A'].width = 38
        for col in range(2, tot_col_idx + 1):
            col_letter = get_column_letter(col)
            ws.column_dimensions[col_letter].width = 16

        # Pipe generated bytes cleanly down file stream transmission
        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)

        filename = f"ERM_Report_{selected_year}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        return send_file(
            excel_stream,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Excel Export System Dropped: {e}")
        return HTMLResponse("Internal Export Failure", status_code=500)
    finally:
        cursor.close()
        conn.close()

# --- CONFIGURATION ---
DB_CONFIG = {'host': 'localhost', 'user': 'root', 'password': '', 'database': 'dashboard'}
GMAIL_USER = "hsmadhan86@gmail.com"
GMAIL_TO = 'b20madanhs@gmail.com'
GMAIL_APP_PASS = "cxay hwxn ocxi ptrt"

# Paths
NETWORK_USER = "SAP"
NETWORK_PASS = "sap@kec"
NETWORK_IP = "192.7.200.4"
SHARE_NAME = "SapDownloads"
LOCAL_DIR = r'D:\Dashboard Material'
BILLING_FILENAME = 'BILLING_DATA_NEW.xls'
ORDER_FILENAME = 'ORDER_BOOK_NEW.xls'
PENDING_ORDERS = 'PENDING_ORDER_NEW.xls'
STOCK_FG = 'STOCK_CARD_FG.xls'
CREDIT_FG = 'STOCK_ANLYS_FG.xls'
OUTSTANDING = 'ZCM_AR_DATEWISE.xls'
COLLECTION = 'ZCOLL_DOWNLOAD.xls'
FILES_TO_FETCH = ['BILLING_DATA_NEW.xls', 'ORDER_BOOK_NEW.xls', 'PENDING_ORDER_NEW.xls', 'STOCK_CARD_FG.xls',
                  'STOCK_ANLYS_FG.xls', 'ZCM_AR_DATEWISE.xls', 'ZCOLL_DOWNLOAD.xls']
BILLING_PATH = os.path.join(LOCAL_DIR, BILLING_FILENAME)
ORDER_PATH = os.path.join(LOCAL_DIR, ORDER_FILENAME)
PENDING_ORDERS_PATH = os.path.join(LOCAL_DIR, PENDING_ORDERS)
STOCK_FG_PATH = os.path.join(LOCAL_DIR, STOCK_FG)
CREDIT_FG_PATH = os.path.join(LOCAL_DIR, CREDIT_FG)
OUTSTANDING_PATH = os.path.join(LOCAL_DIR, OUTSTANDING)
COLLECTION_PATH = os.path.join(LOCAL_DIR, COLLECTION)

# DB_URI = f"mysql+pymysql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
DB_URI = f"mysql+pymysql://root:@localhost:3306/dashboard"
engine = create_engine(DB_URI, pool_size=10, max_overflow=20)


def send_gmail_notification(subject, body):
    msg = EmailMessage()
    msg.set_content(body)
    msg['Subject'] = subject
    msg['From'] = GMAIL_USER
    msg['To'] = GMAIL_TO
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            smtp.starttls()
            smtp.login(GMAIL_USER, GMAIL_APP_PASS)
            smtp.send_message(msg)
        print("Email notification sent.")
    except Exception as e:
        print(f"Email failed: {e}")


def copy_network_files():
    """
    Authenticates with the protected network share and copies
    the specific  files.
    """
    try:

        if not os.path.exists(LOCAL_DIR):
            os.makedirs(LOCAL_DIR)
            print(f"Created local directory: {LOCAL_DIR}")

        # Configure SMB credentials
        smbclient.reset_connection_cache()
        smbclient.ClientConfig(username=NETWORK_USER, password=NETWORK_PASS)

        print(f"Connecting to \\\\{NETWORK_IP}\\{SHARE_NAME}...")

        for filename in FILES_TO_FETCH:
            # Source path on Network: \\192.7.200.4\SapDownloads\
            src_path = f"\\\\{NETWORK_IP}\\{SHARE_NAME}\\{filename}"
            # Destination path on Local
            dst_path = os.path.join(LOCAL_DIR, filename)

            print(f"Attempting to copy: {filename}...")

            with smbclient.open_file(src_path, mode="rb") as remote_file:
                with open(dst_path, "wb") as local_file:
                    shutil.copyfileobj(remote_file, local_file)

            print(f"Successfully saved to: {dst_path}")

        print("--- All files fetched successfully ---")
        return True

    except Exception as e:
        print(f"Error during network fetch: {str(e)}")

        return False


def read_excel_safely(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File not found: {path}")

    # Try 1: Modern Excel (XLSX)
    try:
        return pd.read_excel(path, engine='openpyxl')
    except:
        pass

    # Try 2: Legacy Excel (XLS)
    try:
        return pd.read_excel(path, engine='xlrd')
    except:
        pass

    # Try 3: Tab-Delimited (TSV)
    for enc in ['latin1', 'utf-16', 'utf-8']:
        try:
            return pd.read_csv(path, sep='\t', encoding=enc, low_memory=False)
        except:
            continue

    # Try 4: HTML Table
    try:
        dfs = pd.read_html(path)
        if dfs: return dfs[0]
    except:
        pass

    raise ValueError(f"Failed to read {path}.")


def fix_trailing_minus(x):
    if isinstance(x, str) and x.endswith('-'):
        return f"-{x[:-1]}"
    return x


def sync_data_safe():
    start_time = datetime.now()
    print(f"\n--- Full Database Refresh Started at: {start_time.strftime('%H:%M:%S')} ---")

    if not copy_network_files():
        send_gmail_notification("Sync Failed", "Network path unreachable. Could not fetch latest excels.")
        return

    report = []
    try:
        # 1. BILLING DATA ---
        print("\nReading Billing Excel...")
        df_bill = read_excel_safely(BILLING_PATH)
        df_bill.columns = [str(c).strip() for c in df_bill.columns]

        # ... (Your existing Billing rename logic remains same) ...
        df_bill = df_bill.rename(columns={
            'Billing Doc': 'Billing_Doc',
            'Billing Type': 'Billing_Type',
            'Sales Org': 'Sales_Org',
            'Sales Org Text': 'Sales_Org_Text',
            'Dist Channel': 'Dist_Channel',
            'Dist Channel TEXT': 'Dist_Channel_TEXT',
            'SHIP COND': 'SHIP_COND',
            'Payment Terms': 'Payment_Terms',
            'Created By': 'Created_By',
            'Billing Date': 'Billing_Date',
            'Billing Time': 'Billing_Time',
            'Division Name': 'Division_Name',
            'Sales office Cd': 'Sales_office_Cd',
            'Sales office': 'Sales_office',
            'Sales Region': 'Sales_Region',
            'Sales Region Name': 'Sales_Region_Name',
            'Purchase Order': 'Purchase_Order',
            'Billing Qty': 'Billing_Qty',
            'Cond Num': 'Cond_Num',
            'Tax of the item': 'Tax_of_the_item',
            'Net Value': 'Net_Value',
            'Gross Value': 'Gross_Value',
            'Order Number': 'Order_Number',
            'Material Code': 'Material_Code',
            'Desc': 'Description',
            'Prod hierarchy': 'Prod_hierarchy',
            'Shipping poin': 'Shipping_point',
            'Stor.Loc': 'Store_Loc',
            'Profict Center': 'Profit_Center',
            'Exchange Rate': 'Exchange_Rate',
            'Fin_Year': 'fin_Year',
            'eWay_Billno': 'eWay_Billno',
            'eWay_Billdate': 'eWay_Billdate',
            'LR_No': 'LR_No',
            'LR_Date': 'LR_Date',
            'Transporter_Name': 'Transporter_Name',
            'Vehicle No': 'Vehicle_No',
        })
        #  Date Conversion
        df_bill['Billing_Date'] = pd.to_datetime(df_bill['Billing_Date'], errors='coerce')
        df_bill['eWay_Billdate'] = pd.to_datetime(df_bill['eWay_Billdate'], errors='coerce')
        df_bill['LR_Date'] = pd.to_datetime(df_bill['LR_Date'], errors='coerce')

        sql_bill_columns = [
            'Billing_Doc', 'Billing_Type', 'Curr', 'Sales_Org', 'Sales_Org_Text', 'Dist_Channel',
            'Dist_Channel_TEXT', 'SHIP_COND', 'Inco1', 'Inco2', 'Payment_Terms', 'Unit', 'Created_By',
            'Billing_Date', 'Billing_Time', 'Customer', 'Name', 'Division', 'Division_Name',
            'Sales_office_Cd', 'Sales_office', 'Sales_Region', 'Sales_Region_Name', 'Purchase_Order',
            'Billing_Qty', 'UOM', 'Cond_Num', 'Tax_of_the_item', 'Net_Value', 'Gross_Value', 'Order_Number',
            'Item', 'Material_Code', 'Description', 'Prod_hierarchy', 'VTEXT1', 'VTEXT2', 'VTEXT3', 'VTEXT4',
            'Shipping_point', 'Store_Loc', 'Plant', 'Profit_Center', 'Prod', 'SRLNO', 'Exchange_Rate', 'fin_Year',
            'eWay_Billno', 'eWay_Billdate', 'LR_No', 'LR_Date', 'Transporter_Name', 'Vehicle_No'
        ]

        final_bill_df = df_bill[df_bill.columns.intersection(sql_bill_columns)]

        with engine.connect() as conn:
            print("Truncating billing_data table...")
            conn.execute(text("TRUNCATE TABLE billing_data"))
            conn.commit()

            print(f"Inserting {len(final_bill_df)} rows into billing_data...")
            final_bill_df.to_sql('billing_data', con=conn, if_exists='append', index=False, chunksize=10000)
            conn.commit()
            report.append(f"Billing: Full Refresh completed. {len(final_bill_df)} rows rewritten.")

            # --- 2. ORDER DATA ---
            print("\nReading Order Excel...")
            df_order = read_excel_safely(ORDER_PATH)
            df_order.columns = [str(c).strip() for c in df_order.columns]

            # Standard Renames new columns
            df_order = df_order.rename(columns={
                'Sales Org': 'Sales_Org',
                'Sales Region': 'Sales_Region',
                'Sales office Cd': 'Sales_office_Cd',
                'Sales office': 'Sales_office',
                'Order Number': 'Order_Number',
                'Order Typ': 'Order_Typ',
                'Dist Channel': 'Dist_Channel',
                'Created Date': 'Created_Date',
                'Released date': 'Released_date',
                'PO Delivery Date': 'PO_Delivery_Date',
                'Created By': 'Created_By',
                'Released By': 'Released_By',
                'Sales Engg': 'Sales_Engg',
                'Sales Engg Name': 'Sales_Engg_Name',
                'Material Code': 'Material_Code',
                'Desc': 'Description',
                'Group': 'Groups',
                'Frame Size': 'Frame_Size',
                'Frame Series': 'Frame_Series',
                'Order Qty': 'Order_Qty',
                'Net Value': 'Net_Value',
                'Fin_Year': 'fin_Year',
                'LD Clause Applicable': 'LD_Clause_Applicable',
                'LD Clause Min percentage': 'LD_Clause_Min_percentage',
                'LD Clause Max percentage': 'LD_Clause_Max_percentage',
                'LD max period': 'LD_max_period',
                'ABG Applicable': 'ABG_Applicable',
                'ABG % age': 'ABG_percent_age',
                'ABG Value': 'ABG_Value',
                'ABG Validity': 'ABG_Validity',
                'PBG Applicable': 'PBG_Applicable',
                'PBG % age': 'PBG_percent_age',
                'PBG Value': 'PBG_Value',
                'PBG Validity': 'PBG_Validity',
                'ACG Applicable': 'ACG_Applicable',
                'CG for Advance %age': 'CG_for_Advance_percent_age',
                'CG for Advance Value': 'CG_for_Advance_Value',
                'CG for Advance Validity': 'CG_for_Advance_Validity',
                'PCG Applicable': 'PCG_Applicable',
                'CG for Performance %age': 'CG_for_Performance_Percentage',
                'CG for Performance Value': 'CG_for_Performance_Value',
                'CG for Performance Validity': 'CG_for_Performance_Validity',
                'LD as per Week / Month': 'LD_as_ per_ Week _Month',
                'Email': 'Email',
                'PO No.': 'PO_No',
                'PO Date': 'PO_Date',
                'Total SO Amount': 'Total_SO_Value',
                'Calculated Advance': 'Advance_Value',
                'Collected Advance': 'Advance_Recevied',
                'Adv.To be Collect': 'Advance_To_be_received',
                'Payment Terms HdrText': 'Payment_Terms',
            })
            order_columns_schema = [
                'Location', 'Unit', 'Product1', 'Product2', 'Plant', 'Sales_Org',
                'Sales_Region', 'Sales_office_Cd', 'Sales_office', 'Customer', 'Name', 'Order_Number', 'Item',
                'Order_Typ', 'Dist_Channel', 'Division', 'Created_Date', 'Released_date', 'PO_Delivery_Date',
                'Created_By',
                'Released_By',
                'Sales_Engg', 'Sales_Engg_Name', 'Material_Code', 'Description', 'UOM',
                'Groups', 'Product', 'Type', 'Frame_Size', 'Frame_Series', 'Order_Qty', 'Net_Value', 'fin_Year',
                'LD_Clause_Applicable', 'LD_Clause_Min_percentage', 'LD_Clause_Max_percentage', 'LD_max_period',
                'ABG_Applicable',
                'ABG_%_age', 'PBG_Value', 'PBG_Validity', 'ACG_Applicable', 'CG_for_Advance_%_age',
                'CG_for_Advance_Value', 'CG_for_Advance_Validity',
                'PCG_Applicable', 'CG_ for _Performance _%age', 'CG_for_Performance_Value',
                'CG_for_Performance_Validity', 'LD_as_ per_ Week _Month',
                'Email', 'PO_No', 'PO_Date', 'Total_SO_Value', 'Advance_Value', 'Advance_Recevied',
                'Advance_To_be_received', 'Payment_Terms'
            ]

            # Convert date columns
            for d_col in ['Created_Date', 'Released_date', 'PO_Date']:
                if d_col in df_order.columns:
                    df_order[d_col] = pd.to_datetime(df_order[d_col], errors='coerce')

            numeric_cols_to_fix = ['Advance_Value', 'Advance_Recevied', 'Advance_To_be_received']
            for col in numeric_cols_to_fix:
                if col in df_order.columns:
                    df_order[col] = df_order[col].apply(fix_trailing_minus)
                    df_order[col] = pd.to_numeric(df_order[col], errors='coerce')

            final_order_df = df_order[df_order.columns.intersection(order_columns_schema)]

            with engine.connect() as conn:
                print("Truncating order_data table...")
                conn.execute(text("TRUNCATE TABLE order_data"))
                conn.commit()

                print(f"Inserting {len(final_order_df)} rows into order_data...")
                final_order_df.to_sql('order_data', con=conn, if_exists='append', index=False, chunksize=10000)
                conn.commit()
                report.append(f"Orders: Full Refresh completed. {len(final_order_df)} rows rewritten.")

                # PENDING_ORDERS

                print("\nReading Pending Order Excel...")
                df_pending_order = read_excel_safely(PENDING_ORDERS_PATH)
                df_pending_order.columns = [str(c).strip() for c in df_pending_order.columns]

                df_pending_order = df_pending_order.rename(columns={
                    'As on Date': 'As_on_Date',
                    'Location': 'Location',
                    'Unit': 'Unit',
                    'Product1': 'Product1',
                    'Product2': 'Product2',
                    'Plant': 'Plant',
                    'Sale Order No.': 'Sale_Order_No',
                    'Item No.': 'Item_No',
                    'SO. Type': 'SO_Type',
                    'Sale Ord. Dt.': 'Sale_Ord_Dt',
                    'Delv.Date.': 'Delv_Date',
                    'FIRST DATE': 'FIRST_DATE',
                    'Plant': 'Plant2',
                    'Sales Org': 'Sales_Org',
                    'Distr. Channel': 'Distr_Channel',
                    'Division': 'Division',
                    'Sales Group': 'Sales_Group',
                    'Branch': 'Branch',
                    'Branch Name': 'Branch_Name',
                    'Payment Terms': 'Payment_Terms',
                    'Payment Terms Descp.': 'Payment_Terms_Descp.',
                    'Material Type': 'Material_Type',
                    'Material Group': 'Material_Group',
                    'Material Code': 'Material_Code',
                    'Material Desc.': 'Material_Desc',
                    'Prod. Hierarchy': 'Prod_Hierarchy',
                    'Level-1': 'Level_1',
                    'Level-2': 'Level-2',
                    'Level-3': 'Level-3',
                    'Level-4': 'Level-4',
                    'Cust. Code': 'Cust_Code',
                    'Cust. Name': 'Cust_Name',
                    'PO No.': 'PO_No',
                    'PO Date': 'PO_Date',
                    'OA No.': 'OA_No',
                    'OA Dt.': 'OA_Dt',
                    'Order Reason.': 'Order_Reason',
                    'Created On': 'Created_On',
                    'Order Qty.': 'Order_Qty',
                    'Delivered Qty.': 'Delivered_Qty',
                    'Pending Qty.': 'Pending_Qty',
                    'Net Price': 'Net_Price',
                    'Net Value': 'Net_Value',
                    'Reason for rej.': 'Reason_for_rej.',
                    'Ship To Party': 'Ship_To_Party',
                    'Ship To Party name': 'Ship_To_Party_name',
                    'LD Clause Applicable': 'LD_Clause_Applicable',
                    'LD Clause Min percentage': 'LD_Clause_Min_percentage',
                    'LD Clause Max percentage': 'LD_Clause_Max_percentage',
                    'LD max period': 'LD_max_period',
                    'ABG Applicable': 'ABG_Applicable',
                    'ABG % age': 'ABG_percent_age',
                    'ABG Value': 'ABG_Value',
                    'ABG Validity': 'ABG_Validity',
                    'PBG Applicable': 'PBG_Applicable',
                    'PBG % age': 'PBG_percent_age',
                    'PBG Value': 'PBG_Value',
                    'PBG Validity': 'PBG_Validity',
                    'ACG Applicable': 'ACG_Applicable',
                    'CG for Advance %age': 'CG_for_Advance_percent_age',
                    'CG for Advance Value': 'CG_for_Advance_Value',
                    'CG for Advance Validity': 'CG_for_Advance_Validity',
                    'PCG Applicable': 'PCG_Applicable',
                    'CG for Performance %age': 'CG_for_Performance_Percentage',
                    'CG for Performance Value': 'CG_for_Performance_Value',
                    'CG for Performance Validity': 'CG_for_Performance_Validity',
                    'LD as per Week / Month': 'LD_as_ per_ Week _Month',
                    'Email': 'Email'

                })

                pending_order_columns_schema = [
                    'As_on_Date', 'Location', 'Unit', 'Product1', 'Product2', 'Plant',
                    'Sale_Order_No', 'Item_No', 'SO_Type', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'Plant2',
                    'Sales_Org', 'Distr_Channel', 'Division', 'Sales_Group', 'Branch', 'Branch_Name',
                    'Payment_Terms', 'Payment_Terms_Descp.',
                    'Material_Type', 'Material_Group', 'Material_Code', 'Material_Desc', 'Prod_Hierarchy',
                    'Level_1', 'Level-2', 'Level-3', 'Level-4', 'Cust_Code', 'Cust_Name', 'PO_No', 'PO_Date',
                    'OA_No', 'OA_Dt', '	Order_Reason', 'Created_On', 'Order_Qty', 'Delivered_Qty', 'Pending_Qty',
                    'Net_Price',
                    'Net_Value', 'Reason_for_rej.', 'Ship_To_Party', 'Ship_To_Party_name',
                    'LD_Clause_Applicable', 'LD_Clause_Min_percentage', 'LD_Clause_Max_percentage', 'LD_max_period',
                    'ABG_Applicable',
                    'ABG_%_age', 'PBG_Value', 'PBG_Validity', 'ACG_Applicable', 'CG_for_Advance_%_age',
                    'CG_for_Advance_Value', 'CG_for_Advance_Validity',
                    'PCG_Applicable', 'CG_ for _Performance _%age', 'CG_for_Performance_Value',
                    'CG_for_Performance_Validity', 'LD_as_ per_ Week _Month', 'Email'
                ]
                final_pending_order_df = df_pending_order[
                    df_pending_order.columns.intersection(pending_order_columns_schema)]
                for pending_order_col in ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date',
                                          'Created_On']:
                    if pending_order_col in df_pending_order.columns:
                        df_pending_order[pending_order_col] = pd.to_datetime(df_pending_order[pending_order_col],
                                                                             errors='coerce')

                with engine.connect() as conn:
                    print("Deleting pending_order_data table of today's...")
                    conn.execute(text("DELETE FROM pending_order WHERE As_on_Date = CURRENT_DATE;"))
                    conn.commit()

                    print(f"Inserting {len(final_pending_order_df)} rows into pending_order...")
                    final_pending_order_df.to_sql('pending_order', con=conn, if_exists='append', index=False,
                                                  chunksize=10000)
                    conn.commit()
                    report.append(
                        f"Pending Orders: Full Refresh completed. {len(final_pending_order_df)} rows rewritten.")

            # STOCK_DATA
            print("\nReading Stock_FG Excel...")

            df_stock_fg = read_excel_safely(STOCK_FG_PATH)
            df_stock_fg.columns = [str(c).strip() for c in df_stock_fg.columns]
            df_stock_fg = df_stock_fg.rename(columns={
                'AS ON Date': 'AS_ON_Date',
                'Location': 'Location',
                'Product1': 'Product1',
                'Product2': 'Product2',
                'Unit': 'Unit',
                'Plant': 'Plant',
                'Material': 'Material',
                'Desc': 'Description',
                'UOM.': 'UOM',
                'Mat_type': 'Mat_type',
                'Mat Group': 'Mat_Group',
                'Proc Type': 'Proc_Type',
                'Spl Proc': 'Spl_Proc',
                'Pur Grp': 'Pur_Grp',
                'Profict Center': 'Profit_Center',
                'Stock Qty': 'Stock_Qty',
                'Stock Value': 'Stock_Value',
                'Block Qty': 'Block_Qty',
                'Block Value': 'Block_Value',
                '0-90 Days': '0_90_Days',
                '0-90 Val': '0_90_Val',
                '91-180 Days': '91_180_Days',
                '91-180 Val': '91_180_Val',
                'Above 180 Days': 'Above_180_Days',
                'Above 180 Val': 'Above_180_Val',
                'Return Stock': 'Return_Stock'
            })
        stock_fg_columns_schema = [
            'AS_ON_Date', 'Location', 'Product1', 'Product2', 'Unit', 'Plant', 'Material', 'Description', 'UOM',
            'Mat_type', 'Mat_Group', 'Proc_Type', 'Spl_Proc',
            'Pur_Grp', 'Profit_Center', 'Stock_Qty', 'Stock_Value', 'Block_Qty', 'Block_Value', '0_90_Days', '0_90_Val',
            '91_180_Days', '91_180_Val', 'Above_180_Days', 'Above_180_Val', 'Return_Stock'
        ]
        final_stock_fg_df = df_stock_fg[df_stock_fg.columns.intersection(stock_fg_columns_schema)]
        for stock_fg_col in ['AS_ON_Date']:
            if stock_fg_col in df_stock_fg.columns:
                df_stock_fg[stock_fg_col] = pd.to_datetime(df_stock_fg[stock_fg_col], errors='coerce')

        with engine.connect() as conn:
            print("Truncating stock_data table...")
            conn.execute(text("TRUNCATE TABLE stock_data"))
            conn.commit()
            print(f"Inserting {len(final_stock_fg_df)} rows into stock_data...")
            final_stock_fg_df.to_sql('stock_data', con=conn, if_exists='append', index=False, chunksize=10000)
            conn.commit()
            report.append(f"Stock_data: Full Refresh completed. {len(final_stock_fg_df)} rows rewritten.")

        # CREDIT_FG
        print("\nReading Stock_Analysis Excel...")

        df_credit_fg = read_excel_safely(CREDIT_FG_PATH)
        df_credit_fg.columns = [str(c).strip() for c in df_credit_fg.columns]

        df_credit_fg = df_credit_fg.rename(columns={
            'AS ON Date': 'As_on_date',
            'Location': 'Location',
            'Product1': 'Product1',
            'Product2': 'Product2',
            'Unit': 'Unit',
            'Plant': 'Plant',
            'Material': 'Material',
            'Desc': 'Material_Description',
            'UOM.': 'UOM',
            'Material Type': 'Material_Type',
            'Mat Group': 'Mat_Group',
            'Proc Type': 'Procurement_type',
            'Spl Proc': 'Spl_procurement',
            'Pur Grp': 'Purchasing_Group',
            'Profict Center': 'Profit_Center',
            'Posting Date': 'Posting_Date',
            'Opening Qty': 'Opening_Qty',
            'Opening Val': 'Opening_Val',
            'Receipt Qty': 'Receipt_Qty',
            'Receipt Val': 'Receipt_Val',
            'Issue Qty': 'Issue_Qty',
            'Issue Val': 'Issue_Val',
            'Closing Qty': 'Closing_Qty',
            'Closing Val': 'Closing_Val',
            'Vtext3': 'Vtext3',
            'Product hierarchy': 'Product_hierarchy',
            'Currency': 'Currency'
        })

        credit_fg_column_schema = [
            'As_on_date', 'Location', 'Product1', 'Product2', 'Unit', 'Plant', 'Material', 'Material_Description',
            'UOM', 'Material_Type',
            'Mat_Group', 'Procurement_type', 'Spl_procurement', 'Purchasing_Group', 'Profit_Center', 'Posting_Date',
            'Opening_Qty', 'Opening_Val',
            'Receipt_Qty', 'Receipt_Val', 'Issue_Qty', 'Issue_Val', 'Closing_Qty', 'Closing_Val', 'Vtext3',
            'Product_hierarchy', 'Currency'
        ]
        final_credit_fg_df = df_credit_fg[df_credit_fg.columns.intersection(credit_fg_column_schema)]
        for credit_fg_col in ['As_on_date', 'Posting_Date']:
            if credit_fg_col in df_credit_fg.columns:
                df_credit_fg[credit_fg_col] = pd.to_datetime(
                    df_credit_fg[credit_fg_col],
                    format='%m/%d/%Y',
                    errors='coerce'
                )

        with engine.connect() as conn:
            print("Truncating stock_analysis table...")
            conn.execute(text("TRUNCATE TABLE stock_analysis"))
            conn.commit()
            print(f"Inserting {len(final_credit_fg_df)} rows into stock_analysis...")
            final_credit_fg_df.to_sql('stock_analysis', con=conn, if_exists='append', index=False, chunksize=10000)
            conn.commit()
            report.append(f"stock_analysis: Full Refresh completed. {len(final_credit_fg_df)} rows rewritten.")

        # --- FINAL REPORTING ---
        end_time = datetime.now()
        duration = end_time - start_time
        body = "\n".join(report) + f"\n\nTotal Time Taken: {duration}"
        send_gmail_notification("Full Dashboard Sync Completed", body)
        print(f"Sync Finished Successfully in {duration}")

    except Exception as e:
        print(f"Sync Error: {e}")
        send_gmail_notification("Full Dashboard Sync FAILED", str(e))




scheduler = BackgroundScheduler()
# Slot 1: 09:00 AM
scheduler.add_job(
    func=sync_data_safe,
    trigger="cron",
    hour=9,
    minute=20,
    id="sync_morning"
)
scheduler.add_job(
    func=sync_data_safe,
    trigger="cron",
    hour=10,
    minute=37,
    id="sync_morning_2"
)
# Slot 2: 02:00 PM (14:00)
scheduler.add_job(
    func=sync_data_safe,
    trigger="cron",
    hour=14,
    minute=30,
    id="sync_afternoon"
)

# Slot 3: 05:00 PM (17:00)
scheduler.add_job(
    func=sync_data_safe,
    trigger="cron",
    hour=17,
    minute=22,
    id="sync_evening"
)

scheduler.start()
print(" Scheduler initialized: Task set for  daily.")


@app.get('/setup')
def index():
    return HTMLResponse("Data Sync Service is Running...")


if __name__ == '__main__':
    uvicorn.run(app, host='192.7.200.48', port=8081)
