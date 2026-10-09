from fastapi import APIRouter, Request
from fastapi.responses import HTMLResponse
from fastapi_compat import render_template, redirect, url_for, jsonify, send_file
from db_config import get_db_connection
import pandas as pd
import io, time, threading, calendar
from collections import defaultdict
from io import BytesIO
import shutil
import smbclient
import calendar
from smbprotocol.connection import Connection
import xlsxwriter
from sqlalchemy import create_engine, text, engine
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.message import EmailMessage
from datetime import datetime
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (
    XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
)
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io

prod_app = APIRouter()

engine = create_engine("mysql+pymysql://root:root@192.7.200.4:3306/dashboard")


@prod_app.get('/prod_home')
def product_home(request: Request):
    if request.session.get('role_type') != 'product_head':
        return redirect(url_for('login'))
    return render_template(request,'product_home.html', username=request.session.get('username'))


@prod_app.get('/dashboard/dcm')
def dashboard_dcm(request: Request):
    return render_template(request,'Product1/ytd.html', username=request.session['username'])


def build_filter_query_dcm(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected.
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

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if Kolkata is selected (case-insensitive)
    kolkata_selected = [o for o in offices if o.strip().lower() == 'kolkata']

    if kolkata_selected:

        offices = kolkata_selected
    else:
        # Default behavior: Restrict to ONLY 'DC MACHINES' and 'RRM'
        allowed_products = ['DC MACHINES', 'RRM']
        product_placeholders = ','.join(['%s'] * len(allowed_products))

        if table_type == "data":
            filters.append(f"Prod IN ({product_placeholders})")
        else:
            filters.append(f"Product IN ({product_placeholders})")

        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/dcm')
def get_filters_dcm(request: Request):
    """Fetches distinct filter options, applying the dynamic Kolkata / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_kolkata = any(o.strip().lower() == 'kolkata' for o in offices_selected)

    if is_kolkata:
        # If Kolkata is active, allow all products to populate the dropdowns, but bound to Kolkata
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'Kolkata'"
    else:
        # Otherwise, restrict dropdowns to DC MACHINES and RRM.
        # We explicitly add `OR Sales_office = 'Kolkata'` so Kolkata always remains a visible option to click.
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND (Prod IN ('DC MACHINES', 'RRM') OR Sales_office = 'Kolkata')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"""
                SELECT DISTINCT Sales_Region_Name 
                FROM billing_data 
                WHERE {base_where} AND Sales_Region_Name IS NOT NULL
            """)
            regions = [r['Sales_Region_Name'] for r in cursor.fetchall() if r['Sales_Region_Name']]

            cursor.execute(f"""
                SELECT DISTINCT fin_year 
                FROM billing_data 
                WHERE {base_where} AND fin_year IS NOT NULL
            """)
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM billing_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"""
                SELECT DISTINCT Sales_office 
                FROM billing_data 
                WHERE {base_where} AND Sales_office IS NOT NULL
            """
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit)
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod 
                    END as mapped_prod
                FROM billing_data
                WHERE {base_where} AND Prod IS NOT NULL
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


@prod_app.get('/api/kpis/dcm')
def get_kpis_dcm(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_dcm(request, "data")
    target_where, target_params = build_filter_query_dcm(request, "target")

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


@prod_app.get('/api/sales-trend/dcm')
def get_sales_trend_dcm(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_dcm(request, "data")

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

@prod_app.get('/api/products-revenue/dcm')
def get_products_revenue_dcm(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_dcm(request, "data")

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


@prod_app.get('/api/sales-by-office/dcm')
def get_sales_by_office_dcm(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_dcm(request, "data")

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


@prod_app.get('/api/sales-by-region/dcm')
def get_sales_by_region_dcm(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_dcm(request, "data")

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



@prod_app.get('/api/aop-achievement/dcm')
def get_aop_achievement_dcm(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_dcm(request, "data")
    target_where, target_params = build_filter_query_dcm(request, "target")

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

@prod_app.get('/api/export/sales_excel/dcm')
def export_sales_excel_dcm(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_dcm(request, "data")
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

@prod_app.get('/order_dashboard/dcm')
def order_dashboard_dcm(request: Request):
    return render_template(request,'Product1/order_ytd.html', username=request.session['username'])

def build_filter_orders_query_dcm(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters specifically for Orders based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected (Kolkata exception).
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    kolkata_selected = [o for o in order_offices if o.strip().lower() == 'kolkata']

    if kolkata_selected:
        # If Kolkata is selected: Allow ALL products, but FORCE the data to only be Kolkata.
        order_offices = kolkata_selected
    else:
        # Default behavior: Restrict to ONLY 'DC MACHINES' and 'RRM'
        allowed_products = ['DC Machines', 'RRM']
        product_placeholders = ','.join(['%s'] * len(allowed_products))
        filters.append(f"Product1 IN ({product_placeholders})")
        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/orders/dcm')
def get_orders_filters_dcm(request: Request):
    """Fetches distinct filter options for Orders, applying the dynamic Kolkata / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_kolkata = any(o.strip().lower() == 'kolkata' for o in offices_selected)

    if is_kolkata:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'Kolkata'"
    else:
        base_where = "Product1 IS NOT NULL AND (Product1 IN ('DC Machines', 'RRM') OR Sales_office = 'Kolkata')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r['Sales_Region'] for r in cursor.fetchall() if r['Sales_Region']]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM order_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1 
                    END as mapped_prod
                FROM order_data
                WHERE {base_where}
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

@prod_app.get('/api/order/kpis/dcm')
def get_orders_kpis_dcm(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_dcm(request, "order_data")
    target_where, target_params = build_filter_orders_query_dcm(request, "order_target")

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


@prod_app.get('/api/orders-trend/dcm')
def get_orders_trend_dcm(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_dcm(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/dcm')
def get_orders_products_revenue_dcm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_dcm(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/dcm')
def get_orders_sales_by_office_dcm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_dcm(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/dcm')
def get_order_sales_by_region_dcm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_dcm(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/dcm')
def get_orders_sales_engineers_dcm(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_dcm(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/dcm')
def get_orders_aop_achievement_dcm(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_dcm(request, "order_data")
    target_where, target_params = build_filter_orders_query_dcm(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/dcm')
def export_orders_excel_dcm(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_dcm(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

#---PENDING_ORDERS(PH1)---

@prod_app.get('/ph1/pending_orders')
def ph1_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product1/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


PH1_PENDING_PRODS = ('HV Motors', 'HV Generators')

# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph1/pending/total')
def get_ph1_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
    """, (order_date, PH1_PENDING_PRODS))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph1/pending/products')
def get_ph1_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    product_abbr = [('HV Motors', 'HVM'), ('HV Generators', 'HVG')]

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT Product2, SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s AND `Product2` IN %s  AND Branch_Name != 'Bangalore - IMD'
            GROUP BY Product2
        """, (order_date, PH1_PENDING_PRODS))

        rows = cursor.fetchall()
        actual_dict = {r['Product2']: float(r['value'] or 0) for r in rows}

        return jsonify({
            "labels": [p[1] for p in product_abbr],
            "values": [actual_dict.get(p[0], 0) for p in product_abbr]
        })
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph1/pending/branches')
def get_ph1_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2` IN %s
          AND Branch_Name != '' AND  Branch_Name != 'Bangalore - IMD'
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date, PH1_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph1/pending/sectors')
def get_ph1_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    products = request.query_params.getlist('product') or list(PH1_PENDING_PRODS)

    if not order_date:
        return jsonify({"error": "date is required"}, status_code=400)

    conn = get_db_connection()
    cursor = conn.cursor()

    prod_placeholders = ', '.join(['%s'] * len(products))

    query = f"""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2` IN ({prod_placeholders}) 
          AND Branch_Name != 'Bangalore - IMD'
        GROUP BY `Distr_Channel`
    """

    try:
        cursor.execute(query, (order_date, *products))
        rows = cursor.fetchall()

        processed_data = {}
        for r in rows:
            label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
            processed_data[label] = processed_data.get(label, 0) + float(r['value'] or 0)

        return jsonify({
            "labels": list(processed_data.keys()),
            "values": list(processed_data.values())
        })

    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()

# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph1/pending/units')
def get_ph1_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product2 IN %s AND `As_on_Date` = %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Unit
    """, (PH1_PENDING_PRODS, order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph1/pending/customers')
def get_ph1_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` IN %s and  Branch_Name != 'Bangalore - IMD'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, PH1_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Names
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})

@prod_app.get('/collection_fin_report_dcm')
def finance_dcm_dashboard(request: Request):
    return render_template(request,'Product1/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_dcm(request):
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries dynamically."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    kolkata_selected = [o for o in sales_offices if o.strip().lower() == 'kolkata']

    if kolkata_selected:
        # If Kolkata is selected: Allow ALL profit centers, but FORCE data to only be Kolkata.
        # This overrides other multi-selected offices to keep it restricted.
        sales_offices = kolkata_selected
    else:
        # Default behavior: Restrict to ONLY 'UN01DCM' and 'UN01ACM'
        allowed_pcs = ['UN01DCM', 'UN01ACM']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        # standard IN clause (secured by the baseline rule above if not Kolkata)
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/dcm')
def get_collections_filters_dcm(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic UN1/Kolkata boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_kolkata = any(o.strip().lower() == 'kolkata' for o in offices_selected)

    if is_kolkata:
        base_where = "WHERE Coll_BR_DESC = 'Kolkata'"
    else:
        base_where = "WHERE (Profit_Centre IN ('UN01DCM', 'UN01ACM') OR Coll_BR_DESC = 'Kolkata')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:  # Added dictionary=True to prevent TypeErrors

            # Independent Filters scoped to boundary
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC"
            )
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL"
            )
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL"
            )
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL"
            )
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
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

@prod_app.get('/api/collections/dashboard/dcm')
def get_collections_dashboard_data_dcm(request: Request):
    where, params = build_collections_filter_dcm(request)

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

            # 2. Spend Trend
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


@prod_app.get('/api/export/collections_excel_dcm')
def export_collections_excel_dcm(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_dcm(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

@prod_app.get('/dashboard/hvm')
def dashboard_hvm(request: Request):
    return render_template(request,'Product2/ytd.html', username=request.session['username'])

def build_filter_query_hvm(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected.
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

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if Raipur is selected (case-insensitive)
    Raipur_selected = [o for o in offices if o.strip().lower() == 'raipur']

    if Raipur_selected:

        offices = Raipur_selected
    else:
        # Default behavior: Restrict to ONLY 'HV MOTORS' and 'HV GENERATORS'
        allowed_products = ['HV MOTORS', 'HV GENERATORS']
        product_placeholders = ','.join(['%s'] * len(allowed_products))

        if table_type == "data":
            filters.append(f"Prod IN ({product_placeholders})")
        else:
            filters.append(f"Product IN ({product_placeholders})")

        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/hvm')
def get_filters_hvm(request: Request):
    """Fetches distinct filter options, applying the dynamic Raipur / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_Raipur = any(o.strip().lower() == 'raipur'.lower() for o in offices_selected)

    if is_Raipur:
        # If Raipur is active, allow all products to populate the dropdowns, but bound to Raipur
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'raipur'"
    else:
        # Otherwise, restrict dropdowns to HV MOTORS and HV GENERATORS.
        # We explicitly add `OR Sales_office = 'Raipur'` so Raipur always remains a visible option to click.
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND (Prod IN ('HV MOTORS', 'HV GENERATORS') OR Sales_office = 'raipur')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"""
                SELECT DISTINCT Sales_Region_Name 
                FROM billing_data 
                WHERE {base_where} AND Sales_Region_Name IS NOT NULL
            """)
            regions = [r['Sales_Region_Name'] for r in cursor.fetchall() if r['Sales_Region_Name']]

            cursor.execute(f"""
                SELECT DISTINCT fin_year 
                FROM billing_data 
                WHERE {base_where} AND fin_year IS NOT NULL
            """)
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM billing_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"""
                SELECT DISTINCT Sales_office 
                FROM billing_data 
                WHERE {base_where} AND Sales_office IS NOT NULL
            """
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit)
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod 
                    END as mapped_prod
                FROM billing_data
                WHERE {base_where} AND Prod IS NOT NULL
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


@prod_app.get('/api/kpis/hvm')
def get_kpis_hvm(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_hvm(request, "data")
    target_where, target_params = build_filter_query_hvm(request, "target")

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


@prod_app.get('/api/sales-trend/hvm')
def get_sales_trend_hvm(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_hvm(request, "data")

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

@prod_app.get('/api/products-revenue/hvm')
def get_products_revenue_hvm(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_hvm(request, "data")

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


@prod_app.get('/api/sales-by-office/hvm')
def get_sales_by_office_hvm(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_hvm(request, "data")

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


@prod_app.get('/api/sales-by-region/hvm')
def get_sales_by_region_hvm(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_hvm(request, "data")

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



@prod_app.get('/api/aop-achievement/hvm')
def get_aop_achievement_hvm(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_hvm(request, "data")
    target_where, target_params = build_filter_query_hvm(request, "target")

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

@prod_app.get('/api/export/sales_excel/hvm')
def export_sales_excel_hvm(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_hvm(request, "data")
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

@prod_app.get('/order_dashboard/hvm')
def order_dashboard_hvm(request: Request):
    return render_template(request,'Product2/order_ytd.html', username=request.session['username'])

def build_filter_orders_query_hvm(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters specifically for Orders based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected (Raipur exception).
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    Raipur_selected = [o for o in order_offices if o.strip().lower() == 'raipur']

    if Raipur_selected:
        # If Raipur is selected: Allow ALL products, but FORCE the data to only be Raipur.
        order_offices = Raipur_selected
    else:
        # Default behavior: Restrict to ONLY 'DC MACHINES' and 'RRM'
        allowed_products = ['HV Motors', 'HV Generators']
        product_placeholders = ','.join(['%s'] * len(allowed_products))
        filters.append(f"Product1 IN ({product_placeholders})")
        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/orders/hvm')
def get_orders_filters_hvm(request: Request):
    """Fetches distinct filter options for Orders, applying the dynamic Raipur / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_Raipur = any(o.strip().lower() == 'raipur'.lower() for o in offices_selected)

    if is_Raipur:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'Raipur'"
    else:
        base_where = "Product1 IS NOT NULL AND (Product1 IN ('HV Generators', 'HV Generators') OR Sales_office = 'Raipur')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r['Sales_Region'] for r in cursor.fetchall() if r['Sales_Region']]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM order_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1 
                    END as mapped_prod
                FROM order_data
                WHERE {base_where}
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

@prod_app.get('/api/order/kpis/hvm')
def get_orders_kpis_hvm(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_hvm(request, "order_data")
    target_where, target_params = build_filter_orders_query_hvm(request, "order_target")

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


@prod_app.get('/api/orders-trend/hvm')
def get_orders_trend_hvm(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_hvm(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/hvm')
def get_orders_products_revenue_hvm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_hvm(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/hvm')
def get_orders_sales_by_office_hvm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_hvm(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/hvm')
def get_order_sales_by_region_hvm(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_hvm(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/hvm')
def get_orders_sales_engineers_hvm(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_hvm(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/hvm')
def get_orders_aop_achievement_hvm(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_hvm(request, "order_data")
    target_where, target_params = build_filter_orders_query_hvm(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/hvm')
def export_orders_excel_hvm(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_hvm(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

#---PENDING_ORDERS(PH1)---

@prod_app.get('/ph2/pending_orders')
def ph2_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product2/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


PH2_PENDING_PRODS = ('DC Machines', 'RRM')

# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph2/pending/total')
def get_ph2_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
    """, (order_date, PH2_PENDING_PRODS))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph2/pending/products')
def get_ph2_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    product_abbr = [('DC Machines', 'DCM'), ('RRM', 'RRM')]

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT Product2, SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
            GROUP BY Product2
        """, (order_date, PH2_PENDING_PRODS))

        rows = cursor.fetchall()
        actual_dict = {r['Product2']: float(r['value'] or 0) for r in rows}

        return jsonify({
            "labels": [p[1] for p in product_abbr],
            "values": [actual_dict.get(p[0], 0) for p in product_abbr]
        })
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph2/pending/branches')
def get_ph2_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
          AND Branch_Name != '' 
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date, PH2_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph2/pending/sectors')
def get_ph2_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    products = request.query_params.getlist('product') or list(PH2_PENDING_PRODS)

    conn = get_db_connection()
    cursor = conn.cursor()
    prod_placeholders = ', '.join(['%s'] * len(products))

    cursor.execute(f"""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` IN ({prod_placeholders}) AND Branch_Name != 'Bangalore - IMD'
        GROUP BY `Distr_Channel`
    """, (order_date, *products))

    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['value'] or 0)

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph2/pending/units')
def get_ph2_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product2 IN %s AND `As_on_Date` = %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Unit
    """, (PH2_PENDING_PRODS, order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph2/pending/customers')
def get_ph2_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, PH2_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})

@prod_app.get('/collection_fin_report_hvm')
def finance_hvm_dashboard(request: Request):
    return render_template(request,'Product2/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_hvm(request):
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries dynamically."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    raipur_selected = [o for o in sales_offices if o.strip().lower() == 'bhilai / durg']

    if raipur_selected:
        # If raipur is selected: Allow ALL profit centers, but FORCE data to only be raipur.
        # This overrides other multi-selected offices to keep it restricted.
        sales_offices = raipur_selected
    else:
        # Default behavior: Restrict to ONLY 'UN01DCM' and 'UN01ACM'
        allowed_pcs = ['UN01ACG', 'UN01ACM']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        # standard IN clause (secured by the baseline rule above if not raipur)
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/hvm')
def get_collections_filters_hvm(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic UN1/raipur boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_bhilai_durg = any(o.strip().lower() == 'bhilai / durg'.lower() for o in offices_selected)

    if is_bhilai_durg:
        base_where = "WHERE Coll_BR_DESC = 'bhilai / durg'"
    else:
        base_where = "WHERE (Profit_Centre IN ('UN01ACG', 'UN01ACM') OR Coll_BR_DESC = 'raipur')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:  # Added dictionary=True to prevent TypeErrors

            # Independent Filters scoped to boundary
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC"
            )
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL"
            )
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL"
            )
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL"
            )
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
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

@prod_app.get('/api/collections/dashboard/hvm')
def get_collections_dashboard_data_hvm(request: Request):
    where, params = build_collections_filter_hvm(request)

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

            # 2. Spend Trend
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

@prod_app.get('/api/export/collections_excel_hvm')
def export_collections_excel_hvm(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_hvm(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# ---oft&SWGR----

@prod_app.get('/dashboard/trf')
def dashboard_trf(request: Request):
    return render_template(request,'Product3/ytd.html', username=request.session['username'])

def build_filter_query_oft(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected.
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

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if bangalore is selected (case-insensitive)
    bangalore_selected = [o for o in offices if o.strip().lower() == 'bangalore']

    if bangalore_selected:

        offices = bangalore_selected
    else:
        # Default behavior: Restrict to ONLY 'TRANSFORMERS MYSORE'
        allowed_products = ['TRANSFORMERS MYSORE',]
        product_placeholders = ','.join(['%s'] * len(allowed_products))

        if table_type == "data":
            filters.append(f"Prod IN ({product_placeholders})")
        else:
            filters.append(f"Product IN ({product_placeholders})")

        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/oft')
def get_filters_oft(request: Request):
    """Fetches distinct filter options, applying the dynamic bangalore / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_bangalore = any(o.strip().lower() == 'bangalore' for o in offices_selected)

    if is_bangalore:
        # If bangalore is active, allow all products to populate the dropdowns, but bound to bangalore
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'bangalore'"
    else:
        # Otherwise, restrict dropdowns to TRANSFORMERS MYSORE.
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND (Prod IN ('TRANSFORMERS MYSORE') OR Sales_office = 'bangalore')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"""
                SELECT DISTINCT Sales_Region_Name 
                FROM billing_data 
                WHERE {base_where} AND Sales_Region_Name IS NOT NULL
            """)
            regions = [r['Sales_Region_Name'] for r in cursor.fetchall() if r['Sales_Region_Name']]

            cursor.execute(f"""
                SELECT DISTINCT fin_year 
                FROM billing_data 
                WHERE {base_where} AND fin_year IS NOT NULL
            """)
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM billing_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"""
                SELECT DISTINCT Sales_office 
                FROM billing_data 
                WHERE {base_where} AND Sales_office IS NOT NULL
            """
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit)
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod 
                    END as mapped_prod
                FROM billing_data
                WHERE {base_where} AND Prod IS NOT NULL
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


@prod_app.get('/api/kpis/oft')
def get_kpis_oft(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_oft(request, "data")
    target_where, target_params = build_filter_query_oft(request, "target")

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


@prod_app.get('/api/sales-trend/oft')
def get_sales_trend_oft(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_oft(request, "data")

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

@prod_app.get('/api/products-revenue/oft')
def get_products_revenue_oft(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_oft(request, "data")

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


@prod_app.get('/api/sales-by-office/oft')
def get_sales_by_office_oft(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_oft(request, "data")

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


@prod_app.get('/api/sales-by-region/oft')
def get_sales_by_region_oft(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_oft(request, "data")

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



@prod_app.get('/api/aop-achievement/oft')
def get_aop_achievement_oft(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_oft(request, "data")
    target_where, target_params = build_filter_query_oft(request, "target")

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

@prod_app.get('/api/export/sales_excel/oft')
def export_sales_excel_oft(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_oft(request, "data")
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
        filename = f"Sales_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


@prod_app.get('/order_dashboard/trf')
def order_dashboard_trf(request: Request):
    return render_template(request,'Product3/order_ytd.html', username=request.session['username'])


def build_filter_orders_query_oft(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters specifically for Orders based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected (bangalore exception).
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    bangalore_selected = [o for o in order_offices if o.strip().lower() == 'bangalore']

    if bangalore_selected:
        # If bangalore is selected: Allow ALL products, but FORCE the data to only be bangalore.
        order_offices = bangalore_selected
    else:
        # Default behavior: Restrict to ONLY 'Transformer Mysore'
        allowed_products = ['Transformer Mysore']
        product_placeholders = ','.join(['%s'] * len(allowed_products))
        filters.append(f"Product1 IN ({product_placeholders})")
        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/orders/oft')
def get_orders_filters_oft(request: Request):
    """Fetches distinct filter options for Orders, applying the dynamic bangalore / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_bangalore = any(o.strip().lower() == 'bangalore' for o in offices_selected)

    if is_bangalore:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'bangalore'"
    else:
        base_where = "Product1 IS NOT NULL AND (Product1 IN ('Transformer Mysore') OR Sales_office = 'bangalore')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r['Sales_Region'] for r in cursor.fetchall() if r['Sales_Region']]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM order_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1 
                    END as mapped_prod
                FROM order_data
                WHERE {base_where}
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

@prod_app.get('/api/order/kpis/oft')
def get_orders_kpis_oft(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_oft(request, "order_data")
    target_where, target_params = build_filter_orders_query_oft(request, "order_target")

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


@prod_app.get('/api/orders-trend/oft')
def get_orders_trend_oft(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_oft(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/oft')
def get_orders_products_revenue_oft(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_oft(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/oft')
def get_orders_sales_by_office_oft(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_oft(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/oft')
def get_order_sales_by_region_oft(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_oft(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/oft')
def get_orders_sales_engineers_oft(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_oft(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/oft')
def get_orders_aop_achievement_oft(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_oft(request, "order_data")
    target_where, target_params = build_filter_orders_query_oft(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/oft')
def export_orders_excel_oft(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_oft(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# ---PENDING_ORDERS(PH1)---

@prod_app.get('/ph5/pending_orders')
def ph5_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product3/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


PH3_PENDING_PRODS = ('Transformer Mysore',)


# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph5/pending/total')
def get_ph5_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s 
          AND Product1 IN('Transformer Mysore') 
          AND Level_1 != 'Switchgear' 
          AND Branch_Name != 'Bangalore - IMD'
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph5/pending/products')
def get_ph5_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('Transformer Mysore', 'TRF(M)'),
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
                WHERE `As_on_Date` = %s AND Branch_Name != 'Bangalore - IMD'
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


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph5/pending/branches')
def get_ph5_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2` IN %s
          AND Branch_Name != '' AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date, PH3_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph5/pending/sectors')
def get_ph5_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    products = request.query_params.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT `Distr_Channel` AS channel, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Branch_Name` != 'Bangalore - IMD'
    """
    params = [order_date]

    if products:
        product_clauses = []
        for p in products:
            if p == 'SWITCHGEARS':
                product_clauses.append("(`Product2` = 'Transformer Mysore' AND `Level_1` = 'Switchgear')")
            else:

                product_clauses.append("`Product2` = %s")
                params.append(p)
        query += " AND (" + " OR ".join(product_clauses) + ")"
    else:
        query += " AND `Product2` IN %s"
        params.append(tuple(PH3_PENDING_PRODS))

    query += " GROUP BY `Distr_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()

    # Data Processing / Label Mapping
    processed_data = {}
    for r in rows:
        chan = r['channel']
        val = float(r['value'] or 0)
        label = 'DD' if chan in ['SZ', 'DD'] else chan
        processed_data[label] = processed_data.get(label, 0) + val

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": [float(v) for v in processed_data.values()]
    })

# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph5/pending/units')
def get_ph5_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product2 IN %s AND `As_on_Date` = %s 
        GROUP BY Unit
    """, (PH3_PENDING_PRODS, order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph5/pending/customers')
def get_ph5_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` IN %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, PH3_PENDING_PRODS))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})

@prod_app.get('/collection_fin_report_oft')
def finance_oft_dashboard(request: Request):
    return render_template(request,'Product3/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_oft(request):
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries dynamically."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    bangalore_selected = [o for o in sales_offices if o.strip().lower() == 'bangalore']

    if bangalore_selected:
        # If raipur is selected: Allow ALL profit centers, but FORCE data to only be raipur.
        # This overrides other multi-selected offices to keep it restricted.
        sales_offices = bangalore_selected
    else:
        # Default behavior: Restrict to ONLY 'UN01DCM' and 'UN01ACM'
        allowed_pcs = ['UN05OFT']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        # standard IN clause (secured by the baseline rule above if not raipur)
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/oft')
def get_collections_filters_oft(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic UN1/raipur boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_bangalore = any(o.strip().lower() == 'bangalore'.lower() for o in offices_selected)

    if is_bangalore:
        base_where = "WHERE Coll_BR_DESC = 'bangalore'"
    else:
        base_where = "WHERE (Profit_Centre IN ('UN05OFT') OR Coll_BR_DESC = 'bangalore')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:  # Added dictionary=True to prevent TypeErrors

            # Independent Filters scoped to boundary
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC"
            )
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL"
            )
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL"
            )
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL"
            )
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
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

@prod_app.get('/api/collections/dashboard/oft')
def get_collections_dashboard_data_oft(request: Request):
    where, params = build_collections_filter_oft(request)

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

            # 2. Spend Trend
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

@prod_app.get('/api/export/collections_excel_oft')
def export_collections_excel_oft(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_oft(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


# ---LV Motors(S)----

SW_ALLOWED_REGIONS = ['South India', 'West India']
SW_SPECIAL_OFFICE = 'Hyderabad'
SW_REGION_SQL = "'South India','West India'"


def build_filter_query_sw(request, table_type="data"):
    is_data = table_type == "data"
    filters, params = [], []

    regions   = request.query_params.getlist('region')
    offices   = request.query_params.getlist('sales_office')
    products  = request.query_params.getlist('product')
    units     = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months    = request.query_params.getlist('month')

    special = [o for o in offices if o.strip().lower() == SW_SPECIAL_OFFICE.lower()]
    if special:
        offices = special
    else:
        filters.append("Prod = 'LV MOTORS'" if is_data
                       else "Product IN ('LV MOTORS(S)', 'LV MOTORS(NS)')")
        filters.append(_in("Sales_Region_Name", SW_ALLOWED_REGIONS, params))

    if regions:   filters.append(_in("Sales_Region_Name", regions, params))
    if offices:   filters.append(_in("Sales_office", offices, params))
    if fin_years: filters.append(_in("fin_year", fin_years, params))

    if products:
        if is_data:
            conds, std = [], []
            for p in products:
                if p == 'LV MOTORS (NS)':  conds.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
                elif p == 'LV MOTORS (S)': conds.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
                else: std.append(p)
            if std: conds.append(_in("Prod", std, params))
            filters.append("(" + " OR ".join(conds) + ")")
        else:
            filters.append(_in("Product", products, params))

    if units:
        conds, std = [], []
        for u in units:
            if u == 'UN16':
                conds.append("(Unit = 'UN16' OR Prod = 'TRANSFORMERS PUNE')" if is_data
                             else "(Unit = 'UN16' OR Product = 'TRANSFORMERS PUNE')")
            elif u == 'UN25' and is_data:
                conds.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
            else:
                std.append(u)
        if std: conds.append(_in("Unit", std, params))
        filters.append("(" + " OR ".join(conds) + ")")

    if months and is_data:
        nums = month_numbers(months)
        if nums:
            filters.append(f"MONTH(Billing_Date) IN ({','.join(['%s'] * len(nums))})")
            params.extend(nums)

    if is_data:
        filters.append("Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod IS NOT NULL")

    return "WHERE " + " AND ".join(filters), params


# =====================================================================
# SHARED CACHED DATA LAYER (3 queries per filter combination)
# =====================================================================
def get_base_agg_sw(request):
    def compute():
        where, params = build_filter_query_sw(request, "data")
        sql = f"""
            SELECT Sales_Region_Name AS region, Sales_office AS office,
                   Prod, Plant, Unit,
                   SUM(Net_Value) AS revenue, SUM(Billing_Qty) AS qty
            FROM billing_data
            {where}
            GROUP BY Sales_Region_Name, Sales_office, Prod, Plant, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("sw_agg|" + filter_key(request), 300, compute)


def get_daily_agg_sw(request):
    def compute():
        where, params = build_filter_query_sw(request, "data")
        sql = f"""
            SELECT Billing_Date AS d, SUM(Net_Value) AS revenue
            FROM billing_data
            {where}
            GROUP BY Billing_Date
            ORDER BY Billing_Date
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("sw_daily|" + filter_key(request), 300, compute)


def get_target_rows_sw(request):
    def compute():
        where, params = build_filter_query_sw(request, "target")
        sql = f"""
            SELECT Product, Unit, SUM(Target) * 100000 AS target_revenue
            FROM billing_target1
            {where}
            GROUP BY Product, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("sw_target|" + filter_key(request), 300, compute)


def proration_factor_sw(request):
    months = request.query_params.getlist('month')
    return len(months) / 12.0 if months else 1.0


# =====================================================================
# ROUTES
# =====================================================================
@prod_app.get('/dashboard/lvm_sw')
def dashboard_lvm_sw(request: Request):
    return render_template(request,'Product4/ytd.html', username=request.session.get('username'))


@prod_app.get('/api/filters/sw')
def get_filters_sw(request: Request):
    regions_selected = request.query_params.getlist('region')
    is_special = any(o.strip().lower() == SW_SPECIAL_OFFICE.lower()
                     for o in request.query_params.getlist('sales_office'))

    def compute():
        if is_special:
            base_where = ("Dist_Channel_TEXT != 'Inter Unit Transfer' "
                          f"AND Sales_office = '{SW_SPECIAL_OFFICE}'")
        else:
            base_where = ("Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod = 'LV MOTORS' "
                          f"AND Sales_Region_Name IN ({SW_REGION_SQL})")

        sql = f"""
            SELECT Sales_Region_Name, Sales_office, fin_year, Prod, Plant, Unit
            FROM billing_data
            WHERE {base_where}
            GROUP BY Sales_Region_Name, Sales_office, fin_year, Prod, Plant, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()

        regions, offices, fin_years, units, products = set(), set(), set(), set(), set()
        for r in rows:
            reg, off = r.get('Sales_Region_Name'), r.get('Sales_office')
            if reg:
                regions.add(reg)
                if off and (not regions_selected or reg in regions_selected):
                    offices.add(off)
            if r.get('fin_year'):
                fin_years.add(r['fin_year'])
            if r.get('Prod'):
                products.add(map_prod(r['Prod'], r.get('Plant')))
            u = r.get('Unit')
            if r.get('Plant') in ('AC25', 'AG25'):
                units.add('UN25')
            elif u:
                units.add(u)

        if not is_special:
            offices.add(SW_SPECIAL_OFFICE)

        return {
            "regions":   sorted(regions),
            "offices":   sorted(offices),
            "fin_years": sorted(fin_years),
            "products":  sorted(products),
            "units":     sorted(units - {'UN07', 'UN15'}),
        }

    key = f"sw_filters|{is_special}|{'|'.join(sorted(regions_selected))}"
    return jsonify(cached(key, 600, compute))


@prod_app.get('/api/kpis/sw')
def get_kpis_sw(request: Request):
    factor = proration_factor_sw(request)
    overall = sum(_f(r['revenue']) for r in get_base_agg_sw(request))
    target = sum(_f(t['target_revenue']) for t in get_target_rows_sw(request)) * factor
    return jsonify({"overall_sales": overall, "total_target": target})



@prod_app.get('/api/sales-trend/sw')
def get_sales_trend_sw(request: Request):
    daily = get_daily_agg_sw(request)
    monthly = defaultdict(float)
    weekly = {}

    for r in daily:
        d, rev = r['d'], _f(r['revenue'])
        if d is None:
            continue
        monthly[d.strftime('%Y-%m')] += rev
        iso = d.isocalendar()
        k = (iso[0], iso[1])
        if k not in weekly:
            weekly[k] = [d, 0.0]
        weekly[k][1] += rev

    return jsonify({
        "monthly": [{"time_period": k, "total_revenue": v} for k, v in sorted(monthly.items())],
        "weekly":  [{"time_period": v[0].strftime('%d %b %y'), "total_revenue": v[1]}
                    for _, v in sorted(weekly.items(), key=lambda kv: kv[1][0])],
    })


@prod_app.get('/api/products-revenue/sw')
def get_products_revenue_sw(request: Request):
    agg = defaultdict(lambda: [0.0, 0.0])
    for r in get_base_agg_sw(request):
        p = map_prod(r['Prod'], r['Plant'])
        agg[p][0] += _f(r['revenue'])
        agg[p][1] += _f(r['qty'])
    out = [{"mapped_product": k, "revenue": v[0], "quantity": v[1]} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out)


@prod_app.get('/api/sales-by-office/sw')
def get_sales_by_office_sw(request: Request):
    agg = defaultdict(float)
    for r in get_base_agg_sw(request):
        agg[r['office']] += _f(r['revenue'])
    out = [{"Sales_office": k, "revenue": v} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out[:12])


@prod_app.get('/api/sales-by-region/sw')
def get_sales_by_region_sw(request: Request):
    agg = defaultdict(float)
    for r in get_base_agg_sw(request):
        agg[r['region']] += _f(r['revenue'])
    out = [{"Sales_Region_Name": k, "revenue": v} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out)


@prod_app.get('/api/aop-achievement/sw')
def get_aop_achievement_sw(request: Request):
    factor = proration_factor_sw(request)
    merged = {}

    for t in get_target_rows_sw(request):
        product = t['Product']
        unit = 'UN16' if product == 'TRANSFORMERS PUNE' else t['Unit']
        k = (product, unit)
        if k not in merged:
            merged[k] = {"product": product, "unit": unit, "target": 0.0, "actual": 0.0}
        merged[k]["target"] += _f(t['target_revenue']) * factor

    for r in get_base_agg_sw(request):
        product = map_prod(r['Prod'], r['Plant'])
        unit = map_unit(r['Prod'], r['Plant'], r['Unit'])
        k = (product, unit)
        if k not in merged:
            merged[k] = {"product": product, "unit": unit, "target": 0.0, "actual": 0.0}
        merged[k]["actual"] += _f(r['revenue'])

    results = []
    for d in sorted(merged.values(), key=lambda d: (d['unit'] or '')):
        if d["target"] > 0:
            d["achievement_percentage"] = round(d["actual"] / d["target"] * 100, 2)
        else:
            d["achievement_percentage"] = 100.0 if d["actual"] > 0 else 0.0
        results.append(d)
    return jsonify(results)


@prod_app.get('/api/export/sales_excel/sw')
def export_sales_excel_sw(request: Request):
    where_clause, params = build_filter_query_sw(request, "data")
    query = f"SELECT * FROM billing_data {where_clause}"
    try:
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(query, params)
            results = cur.fetchall()
        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        df = pd.DataFrame(results)
        for col in ['Created_Date', 'Released_date', 'Billing_Date',
                    'eWay_Billdate', 'LR_Date', 'PO_Date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sales_Export')
        output.seek(0)

        fy = "_".join(request.query_params.getlist('fin_year')) or "All_Years"
        return send_file(output,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=f"Sales_Export_SW_{fy}.xlsx")
    except Exception as e:
        print(f"Export Error (SW): {e}")
        return HTMLResponse(f"Error exporting data: {e}", status_code=500)

@prod_app.get('/order_dashboard/lvm_wsc')
def order_dashboard_lvm_wsc(request: Request):
    return render_template(request,'Product4/order_ytd.html', username=request.session['username'])


def build_filter_orders_query_sw(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause for Orders.
    Optimized for performance by mapping products/units in Python to avoid slow SQL CASE statements.
    Includes dynamic restrictions for LV Motors, specific Regions, and hyderabad.
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    hyderabad_selected = [o for o in order_offices if o.strip().lower() == 'hyderabad']

    if hyderabad_selected:
        # If hyderabad is selected: Allow ALL products and regions, but FORCE data to only be hyderabad.
        order_offices = hyderabad_selected
    else:
        # Default behavior: Restrict to ONLY 'LV Motors' and specific Regions
        filters.append("Product1 IN('LV Motors','LV MOTORS(S)') ")

        allowed_regions = ['South India', 'West India']
        region_placeholders = ','.join(['%s'] * len(allowed_regions))
        filters.append(f"Sales_Region IN ({region_placeholders})")
        params.extend(allowed_regions)
    # ------------------------------------

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

    # --- OPTIMIZED PRODUCT HANDLING ---
    if order_products:
        if table_type == "order_data":
            prod_conditions = []
            standard_prods = []
            for p in order_products:
                if p == 'LV MOTORS(S)':
                    prod_conditions.append("(Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)')")
                else:
                    standard_prods.append(p)

            if standard_prods:
                placeholders = ','.join(['%s'] * len(standard_prods))
                prod_conditions.append(f"Product1 IN ({placeholders})")
                params.extend(standard_prods)

            if prod_conditions:
                filters.append("(" + " OR ".join(prod_conditions) + ")")
        else:
            placeholders = ','.join(['%s'] * len(order_products))
            filters.append(f"Product1 IN ({placeholders})")
            params.extend(order_products)

    # --- OPTIMIZED UNIT HANDLING ---
    if order_units:
        unit_conditions = []
        standard_units = []

        if table_type == "order_data":
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                elif u == 'UN25':
                    unit_conditions.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
                else:
                    standard_units.append(u)
        else:
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                else:
                    standard_units.append(u)

        if standard_units:
            placeholders = ','.join(['%s'] * len(standard_units))
            unit_conditions.append(f"Unit IN ({placeholders})")
            params.extend(standard_units)

        if unit_conditions:
            filters.append("(" + " OR ".join(unit_conditions) + ")")

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


@prod_app.get('/api/filters/orders/sw')
def get_orders_filters_sw(request: Request):
    """Fetches distinct filter options for Orders, utilizing high-speed Python mapping."""
    regions_selected = request.query_params.getlist('region')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_hyderabad = any(o.strip().lower() == 'hyderabad' for o in offices_selected)

    if is_hyderabad:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'hyderabad'"
    else:
        base_where = "Product1 IS NOT NULL AND ((Product1 IN('LV Motors','LV MOTORS(S)') AND Sales_Region IN ('South India', 'West India')) OR Sales_office = 'hyderabad')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions & Fin Years
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Optimized Unit Mapping
            cursor.execute(f"SELECT DISTINCT Unit, Product1, Plant FROM order_data WHERE {base_where}")
            raw_units = cursor.fetchall()
            unit_set = set()
            for r in raw_units:
                u = r.get('Unit')
                p1 = r.get('Product1')
                pl = r.get('Plant')

                if p1 == 'Transformer Pune':
                    unit_set.add('UN16')
                elif pl in ('AC25', 'AG25'):
                    unit_set.add('UN25')
                elif u:
                    unit_set.add(u)

            excluded_units = {'UN07', 'UN15'}
            units = [u for u in unit_set if u not in excluded_units]

            # 3. Dependent Office Filter
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Optimized Product Mapping
            cursor.execute(f"SELECT DISTINCT Product1, Product2 FROM order_data WHERE {base_where}")
            raw_prods = cursor.fetchall()
            prod_set = set()
            for r in raw_prods:
                p1 = r.get('Product1')
                p2 = r.get('Product2')

                if p1 == 'LV Motors' and p2 == 'LV Motors (Standard)':
                    prod_set.add('LV MOTORS(S)')
                elif p1:
                    prod_set.add(p1)

            products = list(prod_set)

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@prod_app.get('/api/order/kpis/sw')
def get_orders_kpis_sw(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_sw(request, "order_data")
    target_where, target_params = build_filter_orders_query_sw(request, "order_target")

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


@prod_app.get('/api/orders-trend/sw')
def get_orders_trend_sw(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_sw(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/sw')
def get_orders_products_revenue_sw(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_sw(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/sw')
def get_orders_sales_by_office_sw(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_sw(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/sw')
def get_order_sales_by_region_sw(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_sw(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/sw')
def get_orders_sales_engineers_sw(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_sw(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/sw')
def get_orders_aop_achievement_sw(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_sw(request, "order_data")
    target_where, target_params = build_filter_orders_query_sw(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/sw')
def export_orders_excel_sw(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_sw(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


@prod_app.get('/ph3/pending_orders')
def ph3_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product4/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph3/pending/total')
def get_ph3_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product1` = 'LV Motors' and Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph3/pending/products')
def get_ph3_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('LV Motors (Standard)', 'LVM(S)'),
         ('LV Motors (Non-Standard)','LVM(NS)'),
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
                WHERE `As_on_Date` = %s AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')
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


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph3/pending/branches')
def get_ph3_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product1` = 'LV Motors'
          AND Branch_Name != '' 
         AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph3/pending/sectors')
def get_ph3_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product1` ='LV Motors'
         AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')
        GROUP BY `Distr_Channel`
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['value'] or 0)

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph3/pending/units')
def get_ph3_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product1 = 'LV Motors' AND `As_on_Date` = %s 
         AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur') 
        GROUP BY Unit
    """, (order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph3/pending/customers')
def get_ph3_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product1`='LV Motors'
         AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})


@prod_app.get('/collection_fin_report_lvm_sw')
def finance_lvm_sw_dashboard(request: Request):
    return render_template(request,'Product4/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_sw(request):
    """Builds SQL WHERE clause for Collections, enforcing dynamic LV/hyderabad boundaries."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if hyderabad is selected (case-insensitive)
    hyderabad_selected = [o for o in sales_offices if o.strip().lower() == 'hyderabad']

    if hyderabad_selected:
        # If hyderabad is selected: Allow ALL profit centers, but FORCE data to only be hyderabad.
        sales_offices = hyderabad_selected
    else:
        # Default behavior: Restrict to ONLY specific Profit Centers AND Sales Offices
        allowed_pcs = ['UN02ACM', 'UN25ACM']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)

        allowed_offices = ['Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur']
        office_placeholders = ','.join(['%s'] * len(allowed_offices))
        filters.append(f"Coll_BR_DESC IN ({office_placeholders})")
        params.extend(allowed_offices)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/sw')
def get_collections_filters_sw(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic LV/hyderabad boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_hyderabad = any(o.strip().lower() == 'hyderabad' for o in offices_selected)

    if is_hyderabad:
        # If hyderabad is active, allow all profit centers to populate the dropdowns, but bound to hyderabad
        base_where = "WHERE Coll_BR_DESC = 'hyderabad'"
    else:
        # Fast query without OR: Restricts to specific PCs and Offices.
        # hyderabad is in the allowed list, so it will naturally appear in the dropdown.
        base_where = "WHERE Profit_Centre IN ('UN02ACM', 'UN25ACM') AND Coll_BR_DESC IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:

            # Independent Filters scoped to boundary
            cursor.execute(f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC")
            years = [str(r.get('yr')) for r in cursor.fetchall() if r.get('yr')]

            cursor.execute(f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL")
            months = [r.get('mth') for r in cursor.fetchall() if r.get('mth')]

            cursor.execute(f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL")
            units = [r.get('Unit_Code') for r in cursor.fetchall() if r.get('Unit_Code')]

            cursor.execute(f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL")
            sales_offices = [r.get('Coll_BR_DESC') for r in cursor.fetchall() if r.get('Coll_BR_DESC')]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
                pc_params.extend(unit_selected)

            cursor.execute(pc_query, pc_params)
            profit_centers = [r.get('Profit_Centre') for r in cursor.fetchall() if r.get('Profit_Centre')]

    return jsonify({
        "years": years,
        "months": months,
        "units": units,
        "profit_centers": profit_centers,
        "sales_offices": sales_offices
    })

@prod_app.get('/api/collections/dashboard/sw')
def get_collections_dashboard_data_sw(request: Request):
    where, params = build_collections_filter_sw(request)

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

            # 2. Spend Trend
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


@prod_app.get('/api/export/collections_excel_sw')
def export_collections_excel_sw(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_sw(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


# ---LV Motors(NS)----
_cache, _cache_lock = {}, threading.Lock()
_key_locks = defaultdict(threading.Lock)

def _cache_get(key):
    with _cache_lock:
        hit = _cache.get(key)
        if hit and hit[0] > time.time():
            return hit[1]
        _cache.pop(key, None)
    return None

def _cache_set(key, value, ttl):
    with _cache_lock:
        _cache[key] = (time.time() + ttl, value)
        if len(_cache) > 500:                      # simple size guard
            for k in sorted(_cache, key=lambda k: _cache[k][0])[:100]:
                _cache.pop(k, None)

def cached(key, ttl, compute):
    v = _cache_get(key)
    if v is not None:
        return v
    with _cache_lock:
        lock = _key_locks[key]
    with lock:                                     # others wait here, then hit cache
        v = _cache_get(key)
        if v is not None:
            return v
        v = compute()
        _cache_set(key, v, ttl)
        return v

def filter_key(request):
    items = sorted((k, tuple(sorted(request.query_params.getlist(k))))
                   for k in request.query_params.keys() if k != '_')   # ignore cache-buster
    return repr(items)


# =====================================================================
# PYTHON MAPPERS (replace SQL CASE statements)
# =====================================================================
def map_prod(prod, plant):
    if prod == 'LV MOTORS' and plant == 'AC02': return 'LV MOTORS (NS)'
    if prod == 'LV MOTORS' and plant == 'AC25': return 'LV MOTORS (S)'
    return prod

def map_unit(prod, plant, unit):
    if prod == 'TRANSFORMERS PUNE': return 'UN16'
    if plant in ('AC25', 'AG25'):   return 'UN25'
    return unit

def _f(x):
    return float(x or 0)

_MONTHS = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}

def month_numbers(months):
    out = []
    for m in months:
        m = str(m).strip()
        if m.lower() in _MONTHS:
            out.append(_MONTHS[m.lower()])
        elif m.isdigit() and 1 <= int(m) <= 12:
            out.append(int(m))
    return out


# =====================================================================
# FILTER BUILDER
# =====================================================================
def _in(col, values, params):
    params.extend(values)
    return f"{col} IN ({','.join(['%s'] * len(values))})"

def build_filter_query_lv(request, table_type="data"):
    is_data = table_type == "data"
    filters, params = [], []

    regions   = request.query_params.getlist('region')
    offices   = request.query_params.getlist('sales_office')
    products  = request.query_params.getlist('product')
    units     = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months    = request.query_params.getlist('month')

    ludhiana = [o for o in offices if o.strip().lower() == 'ludhiana']
    if ludhiana:
        offices = ludhiana
    else:
        filters.append("Prod = 'LV MOTORS'" if is_data
                       else "Product IN ('LV MOTORS(S)', 'LV MOTORS(NS)')")
        filters.append(_in("Sales_Region_Name",
                           ['North India', 'East India', 'Central India'], params))

    if regions:   filters.append(_in("Sales_Region_Name", regions, params))
    if offices:   filters.append(_in("Sales_office", offices, params))
    if fin_years: filters.append(_in("fin_year", fin_years, params))

    if products:
        if is_data:
            conds, std = [], []
            for p in products:
                if p == 'LV MOTORS (NS)':  conds.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
                elif p == 'LV MOTORS (S)': conds.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
                else: std.append(p)
            if std: conds.append(_in("Prod", std, params))
            filters.append("(" + " OR ".join(conds) + ")")
        else:
            filters.append(_in("Product", products, params))

    if units:
        conds, std = [], []
        for u in units:
            if u == 'UN16':
                conds.append("(Unit = 'UN16' OR Prod = 'TRANSFORMERS PUNE')" if is_data
                             else "(Unit = 'UN16' OR Product = 'TRANSFORMERS PUNE')")
            elif u == 'UN25' and is_data:
                conds.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
            else:
                std.append(u)
        if std: conds.append(_in("Unit", std, params))
        filters.append("(" + " OR ".join(conds) + ")")

    if months and is_data:
        nums = month_numbers(months)
        if nums:   # integer compare, much cheaper than MONTHNAME() string compare
            filters.append(f"MONTH(Billing_Date) IN ({','.join(['%s'] * len(nums))})")
            params.extend(nums)

    if is_data:
        filters.append("Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod IS NOT NULL")

    return "WHERE " + " AND ".join(filters), params


# =====================================================================
# SHARED CACHED DATA LAYER (3 queries total per filter combination)
# =====================================================================
def get_base_agg(request):
    """Revenue/qty grouped by raw columns only (index friendly, small result)."""
    def compute():
        where, params = build_filter_query_lv(request, "data")
        sql = f"""
            SELECT Sales_Region_Name AS region, Sales_office AS office,
                   Prod, Plant, Unit,
                   SUM(Net_Value) AS revenue, SUM(Billing_Qty) AS qty
            FROM billing_data
            {where}
            GROUP BY Sales_Region_Name, Sales_office, Prod, Plant, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("agg|" + filter_key(request), 300, compute)


def get_daily_agg(request):
    """Daily revenue; monthly and weekly trends are derived from this in Python."""
    def compute():
        where, params = build_filter_query_lv(request, "data")
        sql = f"""
            SELECT Billing_Date AS d, SUM(Net_Value) AS revenue
            FROM billing_data
            {where}
            GROUP BY Billing_Date
            ORDER BY Billing_Date
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("daily|" + filter_key(request), 300, compute)


def get_target_rows(request):
    """Un-prorated targets grouped by Product/Unit. Proration is applied in Python."""
    def compute():
        where, params = build_filter_query_lv(request, "target")
        sql = f"""
            SELECT Product, Unit, SUM(Target) * 100000 AS target_revenue
            FROM billing_target1
            {where}
            GROUP BY Product, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchall()
    return cached("target|" + filter_key(request), 300, compute)


def proration_factor(request):
    months = request.query_params.getlist('month')
    return len(months) / 12.0 if months else 1.0


# =====================================================================
# ROUTES
# =====================================================================
@prod_app.get('/dashboard/lvm_ne')
def dashboard_lvm_nec(request: Request):
    return render_template(request,'Product5/ytd.html', username=request.session.get('username'))


@prod_app.get('/api/filters/lv')
def get_filters_lv(request: Request):
    regions_selected = request.query_params.getlist('region')
    is_ludhiana = any(o.strip().lower() == 'ludhiana'
                      for o in request.query_params.getlist('sales_office'))

    def compute():
        if is_ludhiana:
            base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'Ludhiana'"
        else:
            base_where = ("Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod = 'LV MOTORS' "
                          "AND Sales_Region_Name IN ('North India','East India','Central India')")

        # ONE query replaces five DISTINCT scans
        sql = f"""
            SELECT Sales_Region_Name, Sales_office, fin_year, Prod, Plant, Unit
            FROM billing_data
            WHERE {base_where}
            GROUP BY Sales_Region_Name, Sales_office, fin_year, Prod, Plant, Unit
        """
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(sql)
            rows = cur.fetchall()

        regions, offices, fin_years, units, products = set(), set(), set(), set(), set()
        for r in rows:
            reg, off = r.get('Sales_Region_Name'), r.get('Sales_office')
            if reg:
                regions.add(reg)
                if off and (not regions_selected or reg in regions_selected):
                    offices.add(off)
            if r.get('fin_year'):
                fin_years.add(r['fin_year'])
            if r.get('Prod'):
                products.add(map_prod(r['Prod'], r.get('Plant')))
            u = r.get('Unit')
            if r.get('Plant') in ('AC25', 'AG25'):
                units.add('UN25')
            elif u:
                units.add(u)

        if not is_ludhiana:
            offices.add('Ludhiana')

        return {
            "regions":   sorted(regions),
            "offices":   sorted(offices),
            "fin_years": sorted(fin_years),
            "products":  sorted(products),
            "units":     sorted(units - {'UN07', 'UN15'}),
        }

    key = f"filters|{is_ludhiana}|{'|'.join(sorted(regions_selected))}"
    return jsonify(cached(key, 600, compute))


@prod_app.get('/api/kpis/lv')
def get_kpis_lv(request: Request):
    factor = proration_factor(request)
    overall = sum(_f(r['revenue']) for r in get_base_agg(request))
    target = sum(_f(t['target_revenue']) for t in get_target_rows(request)) * factor
    return jsonify({"overall_sales": overall, "total_target": target})


@prod_app.get('/api/sales-trend/lv')
def get_sales_trend_lv(request: Request):
    daily = get_daily_agg(request)
    monthly = defaultdict(float)
    weekly = {}                       # (iso_year, iso_week) -> [first_date, revenue]

    for r in daily:
        d, rev = r['d'], _f(r['revenue'])
        if d is None:
            continue
        monthly[d.strftime('%Y-%m')] += rev
        iso = d.isocalendar()
        k = (iso[0], iso[1])
        if k not in weekly:
            weekly[k] = [d, 0.0]
        weekly[k][1] += rev           # rows are date-ordered, first date = week start in data

    return jsonify({
        "monthly": [{"time_period": k, "total_revenue": v} for k, v in sorted(monthly.items())],
        "weekly":  [{"time_period": v[0].strftime('%d %b %y'), "total_revenue": v[1]}
                    for _, v in sorted(weekly.items(), key=lambda kv: kv[1][0])],
    })


@prod_app.get('/api/products-revenue/lv')
def get_products_revenue_lv(request: Request):
    agg = defaultdict(lambda: [0.0, 0.0])
    for r in get_base_agg(request):
        p = map_prod(r['Prod'], r['Plant'])
        agg[p][0] += _f(r['revenue'])
        agg[p][1] += _f(r['qty'])
    out = [{"mapped_product": k, "revenue": v[0], "quantity": v[1]} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out)


@prod_app.get('/api/sales-by-office/lv')
def get_sales_by_office_lv(request: Request):
    agg = defaultdict(float)
    for r in get_base_agg(request):
        agg[r['office']] += _f(r['revenue'])
    out = [{"Sales_office": k, "revenue": v} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out[:12])


@prod_app.get('/api/sales-by-region/lv')
def get_sales_by_region_lv(request: Request):
    agg = defaultdict(float)
    for r in get_base_agg(request):
        agg[r['region']] += _f(r['revenue'])
    out = [{"Sales_Region_Name": k, "revenue": v} for k, v in agg.items()]
    out.sort(key=lambda x: -x['revenue'])
    return jsonify(out)


@prod_app.get('/api/aop-achievement/lv')
def get_aop_achievement_lv(request: Request):
    factor = proration_factor(request)
    merged = {}

    for t in get_target_rows(request):
        product = t['Product']
        unit = 'UN16' if product == 'TRANSFORMERS PUNE' else t['Unit']
        k = (product, unit)
        if k not in merged:
            merged[k] = {"product": product, "unit": unit, "target": 0.0, "actual": 0.0}
        merged[k]["target"] += _f(t['target_revenue']) * factor

    for r in get_base_agg(request):
        product = map_prod(r['Prod'], r['Plant'])
        unit = map_unit(r['Prod'], r['Plant'], r['Unit'])
        k = (product, unit)
        if k not in merged:
            merged[k] = {"product": product, "unit": unit, "target": 0.0, "actual": 0.0}
        merged[k]["actual"] += _f(r['revenue'])

    results = []
    for d in sorted(merged.values(), key=lambda d: (d['unit'] or '')):
        if d["target"] > 0:
            d["achievement_percentage"] = round(d["actual"] / d["target"] * 100, 2)
        else:
            d["achievement_percentage"] = 100.0 if d["actual"] > 0 else 0.0
        results.append(d)
    return jsonify(results)


@prod_app.get('/api/export/sales_excel/lv')
def export_sales_excel_lv(request: Request):
    where_clause, params = build_filter_query_lv(request, "data")
    query = f"SELECT * FROM billing_data {where_clause}"
    try:
        with get_db_connection() as conn, conn.cursor() as cur:
            cur.execute(query, params)
            results = cur.fetchall()
        if not results:
            return HTMLResponse("No data found for the selected filters.", status_code=404)

        df = pd.DataFrame(results)
        for col in ['Created_Date', 'Released_date', 'Billing_Date',
                    'eWay_Billdate', 'LR_Date', 'PO_Date']:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d-%m-%Y')

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:  # pip install xlsxwriter
            df.to_excel(writer, index=False, sheet_name='Sales_Export')
        output.seek(0)

        fy = "_".join(request.query_params.getlist('fin_year')) or "All_Years"
        return send_file(output,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=f"Sales_Export_{fy}.xlsx")
    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {e}", status_code=500)
# PH4 ORDER

@prod_app.get('/order_dashboard/lvm_ne')
def order_dashboard_lvm_ne(request: Request):
    return render_template(request,'Product5/order_ytd.html', username=request.session['username'])


def build_filter_orders_query_lv(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause for Orders.
    Optimized for performance by mapping products/units in Python to avoid slow SQL CASE statements.
    Includes dynamic restrictions for LV Motors, specific Regions, and Ludhiana.
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    ludhiana_selected = [o for o in order_offices if o.strip().lower() == 'ludhiana']

    if ludhiana_selected:
        # If Ludhiana is selected: Allow ALL products and regions, but FORCE data to only be Ludhiana.
        order_offices = ludhiana_selected
    else:
        # Default behavior: Restrict to ONLY 'LV Motors' and specific Regions
        filters.append("Product1 IN('LV Motors','LV MOTORS(S)') ")

        allowed_regions = ['North India', 'East India', 'Central India']
        region_placeholders = ','.join(['%s'] * len(allowed_regions))
        filters.append(f"Sales_Region IN ({region_placeholders})")
        params.extend(allowed_regions)
    # ------------------------------------

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

    # --- OPTIMIZED PRODUCT HANDLING ---
    if order_products:
        if table_type == "order_data":
            prod_conditions = []
            standard_prods = []
            for p in order_products:
                if p == 'LV MOTORS(S)':
                    prod_conditions.append("(Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)')")
                else:
                    standard_prods.append(p)

            if standard_prods:
                placeholders = ','.join(['%s'] * len(standard_prods))
                prod_conditions.append(f"Product1 IN ({placeholders})")
                params.extend(standard_prods)

            if prod_conditions:
                filters.append("(" + " OR ".join(prod_conditions) + ")")
        else:
            placeholders = ','.join(['%s'] * len(order_products))
            filters.append(f"Product1 IN ({placeholders})")
            params.extend(order_products)

    # --- OPTIMIZED UNIT HANDLING ---
    if order_units:
        unit_conditions = []
        standard_units = []

        if table_type == "order_data":
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                elif u == 'UN25':
                    unit_conditions.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
                else:
                    standard_units.append(u)
        else:
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                else:
                    standard_units.append(u)

        if standard_units:
            placeholders = ','.join(['%s'] * len(standard_units))
            unit_conditions.append(f"Unit IN ({placeholders})")
            params.extend(standard_units)

        if unit_conditions:
            filters.append("(" + " OR ".join(unit_conditions) + ")")

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


@prod_app.get('/api/filters/orders/lv')
def get_orders_filters_lv(request: Request):
    """Fetches distinct filter options for Orders, utilizing high-speed Python mapping."""
    regions_selected = request.query_params.getlist('region')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_ludhiana = any(o.strip().lower() == 'ludhiana' for o in offices_selected)

    if is_ludhiana:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'Ludhiana'"
    else:
        base_where = "Product1 IS NOT NULL AND ((Product1 IN('LV Motors','LV MOTORS(S)') AND Sales_Region IN ('North India', 'East India', 'Central India')) OR Sales_office = 'Ludhiana')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions & Fin Years
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Optimized Unit Mapping
            cursor.execute(f"SELECT DISTINCT Unit, Product1, Plant FROM order_data WHERE {base_where}")
            raw_units = cursor.fetchall()
            unit_set = set()
            for r in raw_units:
                u = r.get('Unit')
                p1 = r.get('Product1')
                pl = r.get('Plant')

                if p1 == 'Transformer Pune':
                    unit_set.add('UN16')
                elif pl in ('AC25', 'AG25'):
                    unit_set.add('UN25')
                elif u:
                    unit_set.add(u)

            excluded_units = {'UN07', 'UN15'}
            units = [u for u in unit_set if u not in excluded_units]

            # 3. Dependent Office Filter
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Optimized Product Mapping
            cursor.execute(f"SELECT DISTINCT Product1, Product2 FROM order_data WHERE {base_where}")
            raw_prods = cursor.fetchall()
            prod_set = set()
            for r in raw_prods:
                p1 = r.get('Product1')
                p2 = r.get('Product2')

                if p1 == 'LV Motors' and p2 == 'LV Motors (Standard)':
                    prod_set.add('LV MOTORS(S)')
                elif p1:
                    prod_set.add(p1)

            products = list(prod_set)

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@prod_app.get('/api/order/kpis/lv')
def get_orders_kpis_lv(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_lv(request, "order_data")
    target_where, target_params = build_filter_orders_query_lv(request, "order_target")

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


@prod_app.get('/api/orders-trend/lv')
def get_orders_trend_lv(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_lv(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/lv')
def get_orders_products_revenue_lv(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_lv(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/lv')
def get_orders_sales_by_office_lv(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_lv(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/lv')
def get_order_sales_by_region_lv(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_lv(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/lv')
def get_orders_sales_engineers_lv(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_lv(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/lv')
def get_orders_aop_achievement_lv(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_lv(request, "order_data")
    target_where, target_params = build_filter_orders_query_lv(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/lv')
def export_orders_excel_lv(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_lv(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


# ---PENDING_ORDERS(LVM(S))---

@prod_app.get('/ph4/pending_orders')
def ph4_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product5/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph4/pending/total')
def get_ph4_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product1` = 'LV Motors'
        AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph4/pending/products')
def get_ph4_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('LV Motors (Standard)', 'LVM(S)'),
        ('LV Motors (Non-Standard)', 'LVM(NS)'),
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
                WHERE `As_on_Date` = %s AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
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


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph4/pending/branches')
def get_ph4_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product1` = 'LV Motors'
          AND Branch_Name != '' 
          AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph4/pending/sectors')
def get_ph4_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product1` ='LV Motors'
        AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
        GROUP BY `Distr_Channel`
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['value'] or 0)

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph4/pending/units')
def get_ph4_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product1 = 'LV Motors' AND `As_on_Date` = %s 
        AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
        GROUP BY Unit
    """, (order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph4/pending/customers')
def get_ph4_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product1`='LV Motors'
        AND Branch_Name IN ('Delhi','Ludhiana','Kolkata','Faridabad','Lucknow','Raipur')
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})

@prod_app.get('/collection_fin_report_lvm_ne')
def finance_lvm_ne_dashboard(request: Request):
    return render_template(request,'Product5/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_lv(request):
    """Builds SQL WHERE clause for Collections, enforcing dynamic LV/Ludhiana boundaries."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if Ludhiana is selected (case-insensitive)
    ludhiana_selected = [o for o in sales_offices if o.strip().lower() == 'ludhiana']

    if ludhiana_selected:
        # If Ludhiana is selected: Allow ALL profit centers, but FORCE data to only be Ludhiana.
        sales_offices = ludhiana_selected
    else:
        # Default behavior: Restrict to ONLY specific Profit Centers AND Sales Offices
        allowed_pcs = ['UN02ACM', 'UN25ACM']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)

        allowed_offices = ['Delhi', 'Ludhiana', 'Kolkata', 'Faridabad', 'Lucknow', 'bhilai/durg']
        office_placeholders = ','.join(['%s'] * len(allowed_offices))
        filters.append(f"Coll_BR_DESC IN ({office_placeholders})")
        params.extend(allowed_offices)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/lv')
def get_collections_filters_lv(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic LV/Ludhiana boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_ludhiana = any(o.strip().lower() == 'ludhiana' for o in offices_selected)

    if is_ludhiana:
        # If Ludhiana is active, allow all profit centers to populate the dropdowns, but bound to Ludhiana
        base_where = "WHERE Coll_BR_DESC = 'Ludhiana'"
    else:
        # Fast query without OR: Restricts to specific PCs and Offices.
        # Ludhiana is in the allowed list, so it will naturally appear in the dropdown.
        base_where = "WHERE Profit_Centre IN ('UN02ACM', 'UN25ACM') AND Coll_BR_DESC IN ('Delhi', 'Ludhiana', 'Kolkata', 'Faridabad', 'Lucknow', 'bhilai/durg')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:

            # Independent Filters scoped to boundary
            cursor.execute(f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC")
            years = [str(r.get('yr')) for r in cursor.fetchall() if r.get('yr')]

            cursor.execute(f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL")
            months = [r.get('mth') for r in cursor.fetchall() if r.get('mth')]

            cursor.execute(f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL")
            units = [r.get('Unit_Code') for r in cursor.fetchall() if r.get('Unit_Code')]

            cursor.execute(f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL")
            sales_offices = [r.get('Coll_BR_DESC') for r in cursor.fetchall() if r.get('Coll_BR_DESC')]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
                pc_params.extend(unit_selected)

            cursor.execute(pc_query, pc_params)
            profit_centers = [r.get('Profit_Centre') for r in cursor.fetchall() if r.get('Profit_Centre')]

    return jsonify({
        "years": years,
        "months": months,
        "units": units,
        "profit_centers": profit_centers,
        "sales_offices": sales_offices
    })

@prod_app.get('/api/collections/dashboard/lv')
def get_collections_dashboard_data_lv(request: Request):
    where, params = build_collections_filter_lv(request)

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

            # 2. Spend Trend
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


@prod_app.get('/api/export/collections_excel_lv')
def export_collections_excel_lv(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_lv(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# --- CR16----------

@prod_app.get('/dashboard/crt')
def dashboard_crt(request: Request):
    return render_template(request,'Product6/ytd.html', username=request.session['username'])


def build_filter_query_crt(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected.
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

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    # Check if chennai is selected (case-insensitive)
    chennai_selected = [o for o in offices if o.strip().lower() == 'chennai']

    if chennai_selected:

        offices = chennai_selected
    else:
        # Default behavior: Restrict to ONLY 'DC MACHINES' and 'RRM'
        allowed_products = ['SWITCHGEAR MYSORE','TRANSFORMERS PUNE']
        product_placeholders = ','.join(['%s'] * len(allowed_products))

        if table_type == "data":
            filters.append(f"Prod IN ({product_placeholders})")
        else:
            filters.append(f"Product IN ({product_placeholders})")

        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/crt')
def get_filters_crt(request: Request):
    """Fetches distinct filter options, applying the dynamic chennai / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_chennai = any(o.strip().lower() == 'chennai' for o in offices_selected)

    if is_chennai:
        # If chennai is active, allow all products to populate the dropdowns, but bound to chennai
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'chennai'"
    else:
        # Otherwise, restrict dropdowns to DC MACHINES and RRM.
        # We explicitly add `OR Sales_office = 'chennai'` so chennai always remains a visible option to click.
        base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND (Prod IN ('SWITCHGEAR MYSORE','TRANSFORMERS PUNE') OR Sales_office = 'chennai')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"""
                SELECT DISTINCT Sales_Region_Name 
                FROM billing_data 
                WHERE {base_where} AND Sales_Region_Name IS NOT NULL
            """)
            regions = [r['Sales_Region_Name'] for r in cursor.fetchall() if r['Sales_Region_Name']]

            cursor.execute(f"""
                SELECT DISTINCT fin_year 
                FROM billing_data 
                WHERE {base_where} AND fin_year IS NOT NULL
            """)
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'TRANSFORMERS PUNE' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM billing_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"""
                SELECT DISTINCT Sales_office 
                FROM billing_data 
                WHERE {base_where} AND Sales_office IS NOT NULL
            """
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products (Filtered by Unit)
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        ELSE Prod 
                    END as mapped_prod
                FROM billing_data
                WHERE {base_where} AND Prod IS NOT NULL
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


@prod_app.get('/api/kpis/crt')
def get_kpis_crt(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_crt(request, "data")
    target_where, target_params = build_filter_query_crt(request, "target")

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


@prod_app.get('/api/sales-trend/crt')
def get_sales_trend_crt(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_crt(request, "data")

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

@prod_app.get('/api/products-revenue/crt')
def get_products_revenue_crt(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_crt(request, "data")

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


@prod_app.get('/api/sales-by-office/crt')
def get_sales_by_office_crt(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_crt(request, "data")

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


@prod_app.get('/api/sales-by-region/crt')
def get_sales_by_region_crt(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_crt(request, "data")

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



@prod_app.get('/api/aop-achievement/crt')
def get_aop_achievement_crt(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_crt(request, "data")
    target_where, target_params = build_filter_query_crt(request, "target")

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

@prod_app.get('/api/export/sales_excel/crt')
def export_sales_excel_crt(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_crt(request, "data")
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


@prod_app.get('/order_dashboard/crt')
def order_dashboard_crt(request: Request):
    return render_template(request,'Product6/order_ytd.html', username=request.session['username'])


def build_filter_orders_query_crt(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters specifically for Orders based on frontend filters.
    Includes dynamic product restrictions based on the Sales Office selected (chennai exception).
    """
    filters = []
    params = []

    order_regions = request.query_params.getlist('region')
    order_offices = request.query_params.getlist('sales_office')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    chennai_selected = [o for o in order_offices if o.strip().lower() == 'chennai']

    if chennai_selected:
        # If chennai is selected: Allow ALL products, but FORCE the data to only be chennai.
        order_offices = chennai_selected
    else:

        allowed_products = ['Switchgear','Transformer Pune']
        product_placeholders = ','.join(['%s'] * len(allowed_products))
        filters.append(f"Product1 IN ({product_placeholders})")
        params.extend(allowed_products)
    # ------------------------------------

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

    # Unit Handling
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


@prod_app.get('/api/filters/orders/crt')
def get_orders_filters_crt(request: Request):
    """Fetches distinct filter options for Orders, applying the dynamic chennai / Product boundaries."""
    regions_selected = request.query_params.getlist('region')
    units_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_chennai = any(o.strip().lower() == 'chennai' for o in offices_selected)

    if is_chennai:
        base_where = "Product1 IS NOT NULL AND Sales_office = 'chennai'"
    else:
        base_where = "Product1 IS NOT NULL AND (Product1 IN ('Switchgear','Transformer Pune') OR Sales_office = 'chennai')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters
            cursor.execute(f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r['Sales_Region'] for r in cursor.fetchall() if r['Sales_Region']]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r['fin_year'] for r in cursor.fetchall() if r['fin_year']]

            unit_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'Transformer Pune' THEN 'UN16'
                        WHEN Plant IN ('AC25', 'AG25') THEN 'UN25'
                        ELSE Unit 
                    END as unit
                FROM order_data 
                WHERE {base_where} AND Unit IS NOT NULL
            """
            cursor.execute(unit_query)
            excluded_units = {'UN07', 'UN15'}
            units = [r['unit'] for r in cursor.fetchall() if r['unit'] and r['unit'] not in excluded_units]

            # 2. Dependent Filter: Sales Office
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r['Sales_office'] for r in cursor.fetchall() if r['Sales_office']]

            # 3. Dependent Filter: Products
            prod_query = f"""
                SELECT DISTINCT 
                    CASE 
                        WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                        ELSE Product1 
                    END as mapped_prod
                FROM order_data
                WHERE {base_where}
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

@prod_app.get('/api/order/kpis/crt')
def get_orders_kpis_crt(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_crt(request, "order_data")
    target_where, target_params = build_filter_orders_query_crt(request, "order_target")

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


@prod_app.get('/api/orders-trend/crt')
def get_orders_trend_crt(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_crt(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/crt')
def get_orders_products_revenue_crt(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_crt(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/crt')
def get_orders_sales_by_office_crt(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_crt(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/crt')
def get_order_sales_by_region_crt(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_crt(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/crt')
def get_orders_sales_engineers_crt(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_crt(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/crt')
def get_orders_aop_achievement_crt(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_crt(request, "order_data")
    target_where, target_params = build_filter_orders_query_crt(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/crt')
def export_orders_excel_crt(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_crt(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)


# ---PENDING_ORDERS(CRT)---

@prod_app.get('/ph6/pending_orders')
def ph6_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'Product6/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/ph6/pending/total')
def get_ph6_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
       SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s 
          AND Branch_Name != 'Bangalore - IMD'
          AND (
                `Product2` = 'Transformer Pune'
             OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear')
          )
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/ph6/pending/products')
def get_ph6_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('TRF(P)', 'TRF(P)'),
        ('SWITCHGEARS', 'SWGR'),
    ]

    try:
        cursor.execute("""
            SELECT 
                CASE
                    WHEN Product2 = 'Transformer Mysore' AND Level_1 = 'Switchgear' THEN 'SWITCHGEARS'
                    WHEN Product2 = 'Transformer Pune' THEN 'TRF(P)'
                    ELSE Product2
                END AS display_product, 
                SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s 
              AND Branch_Name != 'Bangalore - IMD'
              AND (
                  Product2 = 'Transformer Pune'
                  OR (Product2 = 'Transformer Mysore' AND Level_1 = 'Switchgear')
              )
            GROUP BY display_product
        """, (order_date,))

        rows = cursor.fetchall()
        actual_dict = {r['display_product']: float(r['value'] or 0) for r in rows}

        final_labels = []
        final_values = []

        for key_name, display_label in product_abbr:
            if key_name in actual_dict:
                final_labels.append(display_label)
                final_values.append(actual_dict[key_name])

        return jsonify({
            "labels": final_labels,
            "values": final_values
        })

    except Exception as e:
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/ph6/pending/branches')
def get_ph6_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
       WHERE `As_on_Date` = %s 
          AND Branch_Name != 'Bangalore - IMD'
          AND (
                `Product2` = 'Transformer Pune'
             OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear')
          )
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/ph6/pending/sectors')
def get_ph6_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
       WHERE `As_on_Date` = %s 
          AND Branch_Name != 'Bangalore - IMD'
          AND (
                `Product2` = 'Transformer Pune'
             OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear')
          )
        GROUP BY `Distr_Channel`
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['value'] or 0)

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": list(processed_data.values())
    })


# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/ph6/pending/units')
def get_ph6_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
       WHERE `As_on_Date` = %s 
          AND Branch_Name != 'Bangalore - IMD'
          AND (
                `Product2` = 'Transformer Pune'
             OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear')
          )
        GROUP BY Unit
    """, (order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/ph6/pending/customers')
def get_ph6_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND Branch_Name != 'Bangalore - IMD'
          AND (
                `Product2` = 'Transformer Pune'
             OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear')
          )
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})


@prod_app.get('/collection_fin_report_crt')
def finance_crt_dashboard(request: Request):
    return render_template(request,'Product6/collection_report.html', username=request.session.get('username', 'User'),)

def build_collections_filter_crt(request):
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries dynamically."""
    filters = []
    params = []

    # Extract UI Filters
    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    ui_pcs = request.query_params.getlist('profit_center')
    sales_offices = request.query_params.getlist('sales_office')

    # --- DYNAMIC AUTHORIZATION LOGIC ---
    chennai_selected = [o for o in sales_offices if o.strip().lower() == 'chennai']

    if chennai_selected:
        # If chennai is selected: Allow ALL profit centers, but FORCE data to only be chennai.
        # This overrides other multi-selected offices to keep it restricted.
        sales_offices = chennai_selected
    else:
        # Default behavior: Restrict to ONLY 'UN01crt' and 'UN01ACM'
        allowed_pcs = ['UN16CRT']
        pc_placeholders = ','.join(['%s'] * len(allowed_pcs))
        filters.append(f"Profit_Centre IN ({pc_placeholders})")
        params.extend(allowed_pcs)
    # ------------------------------------

    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    if ui_pcs:
        # standard IN clause (secured by the baseline rule above if not chennai)
        placeholders = ','.join(['%s'] * len(ui_pcs))
        filters.append(f"Profit_Centre IN ({placeholders})")
        params.extend(ui_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters) if filters else ""
    return where_clause, params


@prod_app.get('/api/filters/collections/crt')
def get_collections_filters_crt(request: Request):
    """Fetches unique values for filters, restricted strictly to dynamic UN1/chennai boundaries."""

    unit_selected = request.query_params.getlist('unit')
    offices_selected = request.query_params.getlist('sales_office')

    # --- DYNAMIC DROPDOWN BOUNDARY LOGIC ---
    is_chennai = any(o.strip().lower() == 'chennai' for o in offices_selected)

    if is_chennai:
        base_where = "WHERE Coll_BR_DESC = 'chennai'"
    else:
        base_where = "WHERE (Profit_Centre IN ('UN16CRT') OR Coll_BR_DESC = 'chennai')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:  # Added dictionary=True to prevent TypeErrors

            # Independent Filters scoped to boundary
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC"
            )
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL"
            )
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL"
            )
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL"
            )
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
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

@prod_app.get('/api/collections/dashboard/crt')
def get_collections_dashboard_data_crt(request: Request):
    where, params = build_collections_filter_crt(request)

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

            # 2. Spend Trend
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


@prod_app.get('/api/export/collections_excel_crt')
def export_collections_excel_crt(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_crt(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# ---- International Market------

@prod_app.get('/dashboard/imd')
def dashboard_imd(request: Request):
    return render_template(request,'imd/ytd.html', username=request.session['username'])

def build_filter_query_imd(request, table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the IMD dashboard.
    Optimized for MySQL using Python mapping and hardcoded Bangalore - IMD restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Bangalore - IMD for all queries on this dashboard
    filters.append("Sales_office = 'Bangalore - IMD'")
    # ----------------------------

    # Extract UI selections
    regions = request.query_params.getlist('region')
    products = request.query_params.getlist('product')
    units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # Region Handling
    if regions:
        placeholders = ','.join(['%s'] * len(regions))
        filters.append(f"Sales_Region_Name IN ({placeholders})")
        params.extend(regions)

    # --- PYTHON MAPPING FOR PRODUCTS ---
    if products:
        if table_type == "data":
            prod_conditions = []
            standard_prods = []
            for p in products:
                if p == 'LV MOTORS (NS)':
                    prod_conditions.append("(Prod = 'LV MOTORS' AND Plant = 'AC02')")
                elif p == 'LV MOTORS (S)':
                    prod_conditions.append("(Prod = 'LV MOTORS' AND Plant = 'AC25')")
                else:
                    standard_prods.append(p)

            if standard_prods:
                placeholders = ','.join(['%s'] * len(standard_prods))
                prod_conditions.append(f"Prod IN ({placeholders})")
                params.extend(standard_prods)

            if prod_conditions:
                filters.append("(" + " OR ".join(prod_conditions) + ")")
        else:
            placeholders = ','.join(['%s'] * len(products))
            filters.append(f"Product IN ({placeholders})")
            params.extend(products)

    # --- PYTHON MAPPING FOR UNITS ---
    if units:
        unit_conditions = []
        standard_units = []

        if table_type == "data":
            for u in units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Prod = 'TRANSFORMERS PUNE')")
                elif u == 'UN25':
                    unit_conditions.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
                else:
                    standard_units.append(u)
        else:
            for u in units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product = 'TRANSFORMERS PUNE')")
                else:
                    standard_units.append(u)

        if standard_units:
            placeholders = ','.join(['%s'] * len(standard_units))
            unit_conditions.append(f"Unit IN ({placeholders})")
            params.extend(standard_units)

        if unit_conditions:
            filters.append("(" + " OR ".join(unit_conditions) + ")")

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


@prod_app.get('/api/filters/imd')
def get_filters_imd():
    """Fetches distinct filter options, restricted strictly to Bangalore - IMD."""

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Sales_office = 'Bangalore - IMD'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions
            cursor.execute(
                f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units (Mapped entirely in Python to avoid SQL CASE delays)
            cursor.execute(f"SELECT DISTINCT Unit, Prod, Plant FROM billing_data WHERE {base_where}")
            raw_units = cursor.fetchall()
            unit_set = set()
            for r in raw_units:
                u = r.get('Unit')
                p = r.get('Prod')
                pl = r.get('Plant')

                if p == 'TRANSFORMERS PUNE':
                    unit_set.add('UN16')
                elif pl in ('AC25', 'AG25'):
                    unit_set.add('UN25')
                elif u:
                    unit_set.add(u)

            excluded_units = {'UN07', 'UN15'}
            units = [u for u in unit_set if u not in excluded_units]

            # 4. Sales Office
            offices = ['Bangalore - IMD']

            # 5. Products (Mapped in Python)
            cursor.execute(f"SELECT DISTINCT Prod, Plant FROM billing_data WHERE {base_where} AND Prod IS NOT NULL")
            raw_prods = cursor.fetchall()
            prod_set = set()
            for r in raw_prods:
                p = r.get('Prod')
                pl = r.get('Plant')

                if p == 'LV MOTORS' and pl == 'AC02':
                    prod_set.add('LV MOTORS (NS)')
                elif p == 'LV MOTORS' and pl == 'AC25':
                    prod_set.add('LV MOTORS (S)')
                elif p:
                    prod_set.add(p)

            products = list(prod_set)

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@prod_app.get('/api/kpis/imd')
def get_kpis_imd(request: Request):
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_imd(request, "data")
    target_where, target_params = build_filter_query_imd(request, "target")

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


@prod_app.get('/api/sales-trend/imd')
def get_sales_trend_imd(request: Request):
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_imd(request, "data")

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

@prod_app.get('/api/products-revenue/imd')
def get_products_revenue_imd(request: Request):
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_imd(request, "data")

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


@prod_app.get('/api/sales-by-office/imd')
def get_sales_by_office_imd(request: Request):
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_imd(request, "data")

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


@prod_app.get('/api/sales-by-region/imd')
def get_sales_by_region_imd(request: Request):
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_imd(request, "data")

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



@prod_app.get('/api/aop-achievement/imd')
def get_aop_achievement_imd(request: Request):
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_imd(request, "data")
    target_where, target_params = build_filter_query_imd(request, "target")

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

@prod_app.get('/api/export/sales_excel/imd')
def export_sales_excel_imd(request: Request):
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_imd(request, "data")
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



# --- IMD Orders -----

@prod_app.get('/order_dashboard/imd')
def order_dashboard_imd(request: Request):
    return render_template(request,'imd/order_ytd.html', username=request.session['username'])


def build_filter_orders_query_imd(request, table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the International Market Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded InternationalMarket restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce InternationalMarket for all queries on this dashboard
    filters.append("Sales_office = 'InternationalMarket'")
    # ----------------------------

    order_regions = request.query_params.getlist('region')
    order_products = request.query_params.getlist('product')
    order_units = request.query_params.getlist('unit')
    fin_years = request.query_params.getlist('fin_year')
    months = request.query_params.getlist('month')

    # Region
    if order_regions:
        placeholders = ','.join(['%s'] * len(order_regions))
        filters.append(f"Sales_Region IN ({placeholders})")
        params.extend(order_regions)

    # --- PYTHON MAPPING FOR PRODUCTS ---
    if order_products:
        if table_type == "order_data":
            prod_conditions = []
            standard_prods = []
            for p in order_products:
                if p == 'LV MOTORS(S)':
                    prod_conditions.append("(Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)')")
                else:
                    standard_prods.append(p)

            if standard_prods:
                placeholders = ','.join(['%s'] * len(standard_prods))
                prod_conditions.append(f"Product1 IN ({placeholders})")
                params.extend(standard_prods)

            if prod_conditions:
                filters.append("(" + " OR ".join(prod_conditions) + ")")
        else:
            placeholders = ','.join(['%s'] * len(order_products))
            filters.append(f"Product1 IN ({placeholders})")
            params.extend(order_products)

    # --- PYTHON MAPPING FOR UNITS ---
    if order_units:
        unit_conditions = []
        standard_units = []

        if table_type == "order_data":
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                elif u == 'UN25':
                    unit_conditions.append("(Unit = 'UN25' OR Plant IN ('AC25', 'AG25'))")
                else:
                    standard_units.append(u)
        else:
            for u in order_units:
                if u == 'UN16':
                    unit_conditions.append("(Unit = 'UN16' OR Product1 = 'Transformer Pune')")
                else:
                    standard_units.append(u)

        if standard_units:
            placeholders = ','.join(['%s'] * len(standard_units))
            unit_conditions.append(f"Unit IN ({placeholders})")
            params.extend(standard_units)

        if unit_conditions:
            filters.append("(" + " OR ".join(unit_conditions) + ")")

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


@prod_app.get('/api/filters/orders/imd')
def get_orders_filters_imd():
    """Fetches distinct filter options for Orders, restricted strictly to InternationalMarket."""

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 IS NOT NULL AND Sales_office = 'InternationalMarket'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units (Mapped entirely in Python to avoid SQL CASE delays)
            cursor.execute(f"SELECT DISTINCT Unit, Product1, Plant FROM order_data WHERE {base_where}")
            raw_units = cursor.fetchall()
            unit_set = set()
            for r in raw_units:
                u = r.get('Unit')
                p1 = r.get('Product1')
                pl = r.get('Plant')

                if p1 == 'Transformer Pune':
                    unit_set.add('UN16')
                elif pl in ('AC25', 'AG25'):
                    unit_set.add('UN25')
                elif u:
                    unit_set.add(u)

            excluded_units = {'UN07', 'UN15'}
            units = [u for u in unit_set if u not in excluded_units]

            # 4. Sales Office
            # We don't need to query the database for this since it's statically bound
            offices = ['InternationalMarket']

            # 5. Products (Mapped entirely in Python to avoid SQL CASE delays)
            cursor.execute(f"SELECT DISTINCT Product1, Product2 FROM order_data WHERE {base_where}")
            raw_prods = cursor.fetchall()
            prod_set = set()
            for r in raw_prods:
                p1 = r.get('Product1')
                p2 = r.get('Product2')

                if p1 == 'LV Motors' and p2 == 'LV Motors (Standard)':
                    prod_set.add('LV MOTORS(S)')
                elif p1:
                    prod_set.add(p1)

            products = list(prod_set)

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@prod_app.get('/api/order/kpis/imd')
def get_orders_kpis_imd(request: Request):
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_imd(request, "order_data")
    target_where, target_params = build_filter_orders_query_imd(request, "order_target")

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


@prod_app.get('/api/orders-trend/imd')
def get_orders_trend_imd(request: Request):
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_imd(request, "order_data")

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


@prod_app.get('/api/orders/products-revenue/imd')
def get_orders_products_revenue_imd(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_imd(request, "order_data")

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


@prod_app.get('/api/orders/sales-by-office/imd')
def get_orders_sales_by_office_imd(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_imd(request, "order_data")

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


@prod_app.get('/api/order/sales-by-region/imd')
def get_order_sales_by_region_imd(request: Request):
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_imd(request, "order_data")

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


@prod_app.get('/api/orders/sales-engineers/imd')
def get_orders_sales_engineers_imd(request: Request):
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_imd(request, "order_data")

    query = f"""
        SELECT 
            Sales_Engg_Name, 
            Sales_office, 
            SUM(Net_Value) AS actual
        FROM order_data
        {data_where}
        GROUP BY Sales_Engg_Name, Sales_office
        ORDER BY actual DESC
    """

    table_data = []
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, data_params)
            results = cursor.fetchall()

            for idx, row in enumerate(results, 1):
                table_data.append({
                    "sl_no": idx,
                    "sales_engineer": row.get('Sales_Engg_Name') or 'Unknown',
                    "sales_office": row.get('Sales_office') or 'Unknown',
                    "actual": float(row.get('actual') or 0),
                    "target": 0,
                    "achievement_percentage": 0
                })

    return jsonify(table_data)


@prod_app.get('/api/orders/aop-achievement/imd')
def get_orders_aop_achievement_imd(request: Request):
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_imd(request, "order_data")
    target_where, target_params = build_filter_orders_query_imd(request, "order_target")

    months = request.query_params.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # Actuals Query (Grouped only by Product)
    actuals_query = f"""
        SELECT 
            CASE 
               WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                ELSE Product1
            END AS mapped_product,
            SUM(Net_Value) as actual_revenue
        FROM order_data
        {data_where}
        GROUP BY mapped_product
        ORDER BY mapped_product ASC
    """

    # Targets Query (Grouped only by Product)
    targets_query = f"""
        SELECT 
            Product1, 
            (SUM(Target) * 100000 * {proration_factor}) as target_revenue 
        FROM order_target1
        {target_where}
        GROUP BY Product1
        ORDER BY Product1 ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(targets_query, target_params)
            targets_raw = cursor.fetchall()

            cursor.execute(actuals_query, data_params)
            actuals_raw = cursor.fetchall()

    merged_data = {}

    # Map Targets
    for row in targets_raw:
        product = row['Product1']

        merged_data[product] = {
            "product": product,
            "order_target": float(row['target_revenue'] or 0),
            "actual": 0.0,
            "achievement_percentage": 0.0
        }

    # Map Actuals
    for row in actuals_raw:
        product = row['mapped_product']

        if product not in merged_data:
            merged_data[product] = {
                "product": product,
                "order_target": 0.0,
                "actual": 0.0,
                "achievement_percentage": 0.0
            }

        merged_data[product]["actual"] += float(row['actual_revenue'] or 0)

    # Calculate final percentages
    final_results = []
    for data in merged_data.values():
        if data["order_target"] > 0:
            data["achievement_percentage"] = round((data["actual"] / data["order_target"]) * 100, 2)
        else:
            data["achievement_percentage"] = 100.0 if data["actual"] > 0 else 0.0

        final_results.append(data)

    return jsonify(final_results)

@prod_app.get('/api/export/orders_excel/imd')
def export_orders_excel_imd(request: Request):
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_imd(request, "order_data")

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
        filename = f"Orders_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)

# ---PENDING_ORDERS(CRT)---

@prod_app.get('/imd/pending_orders')
def imd_pending_orders_page(request: Request):
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template(request,'imd/Pending_Orders.html', username=request.session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@prod_app.get('/api/imd/pending/total')
def get_imd_total_pending(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Branch_Name` = 'Bangalore - IMD' 
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@prod_app.get('/api/imd/pending/products')
def get_imd_pending_products(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('DC Machines', 'DCM'), ('HV Motors', 'HVM'), ('HV Generators', 'HVG'),
        ('RRM', 'RRM'), ('LV Generators', 'LVG'),
        ('LV Motors (Non-Standard)', 'LVM(NS)'), ('LV Motors (Standard)', 'LVM(S)'), ('EVM', 'EVM'), ('LV Generators', 'LVG'),
        ('DG Sets', 'DGS'),
        ('Transformer Pune', 'TRF(P)'), ('Transformer Mysore', 'TRF(M)'), ('Switchgear', 'SWG'),
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
                WHERE `As_on_Date` = %s AND `Branch_Name` = 'Bangalore - IMD'
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


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@prod_app.get('/api/imd/pending/branches')
def get_imd_pending_branches(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Branch_Name` = 'Bangalore - IMD'
        GROUP BY Branch_Name
        ORDER BY value DESC
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    # Manual consolidation for Faridabad/Delhi consistency
    processed_dict = {}
    for r in rows:
        name = r['label'].strip()
        val = float(r['value'] or 0)
        if name in ['Delhi', 'Faridabad']:
            processed_dict['Faridabad'] = processed_dict.get('Faridabad', 0) + val
        else:
            processed_dict[name] = val

    sorted_branches = sorted(processed_dict.items(), key=lambda x: x[1], reverse=True)

    return jsonify({
        "labels": [item[0] for item in sorted_branches],
        "values": [float(item[1]) for item in sorted_branches]
    })


# --- 4. PH1 PENDING SECTORS ---
@prod_app.get('/api/imd/pending/sectors')
def get_imd_pending_sectors(request: Request):
    order_date = request.query_params.get('date')
    products = request.query_params.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()


    query = """
        SELECT `Distr_Channel` AS channel, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Branch_Name` = 'Bangalore - IMD'
    """
    params = [order_date]

    # 3. Dynamic Product Filtering Logic
    if products:
        product_clauses = []
        for p in products:
            if p == 'SWITCHGEARS':
                # Special condition for Switchgears
                product_clauses.append("(`Product2` = 'Transformer Mysore' AND `Level_1` = 'Switchgear')")
            else:
                # Standard condition for other products
                product_clauses.append("`Product2` = %s")
                params.append(p)


        query += " AND (" + " OR ".join(product_clauses) + ")"

    query += " GROUP BY `Distr_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        print(f"Database Error: {e}")
        return jsonify({"error": str(e)}, status_code=500)
    finally:
        conn.close()

    processed_data = {}
    for r in rows:
        chan = r['channel']
        val = float(r['value'] or 0)
        label = 'DD' if chan in ['SZ', 'DD'] else chan
        processed_data[label] = processed_data.get(label, 0) + val

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": [float(v) for v in processed_data.values()]
    })
# --- 5. PH1 PENDING UNITS ---
@prod_app.get('/api/imd/pending/units')
def get_imd_pending_units(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE `Branch_Name` = 'Bangalore - IMD' AND `As_on_Date` = %s 
        GROUP BY Unit
    """, (order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@prod_app.get('/api/imd/pending/customers')
def get_imd_pending_customers(request: Request):
    order_date = request.query_params.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Branch_Name` = 'Bangalore - IMD'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
        # Standardize KEC Name
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['value'] or 0))

    return jsonify({"labels": labels, "values": values})

@prod_app.get('/collection_fin_report_imd')
def finance_imd_dashboard(request: Request):
    return render_template(request,'imd/collection_report.html', username=request.session.get('username', 'User'),)


def build_collections_filter_imd(request):
    """Builds SQL WHERE clause for Collections queries, strictly restricted to Exports (IMD)."""
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Exports (IMD) for all queries on this dashboard
    filters.append("Coll_BR_DESC = 'Exports (IMD)'")
    # ----------------------------

    years = request.query_params.getlist('year')
    months = request.query_params.getlist('month')
    exact_date = request.query_params.get('date')
    units = request.query_params.getlist('unit')
    profit_centers = request.query_params.getlist('profit_center')


    if years:
        placeholders = ','.join(['%s'] * len(years))
        filters.append(f"YEAR(Posting_Date) IN ({placeholders})")
        params.extend(years)

    if months:
        month_map = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6,
                     "July": 7, "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
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

    where_clause = "WHERE " + " AND ".join(filters)
    return where_clause, params


@prod_app.get('/api/filters/collections/imd')
def get_collections_filters_imd(request: Request):
    """Fetches unique values for the Collections Dashboard filters, restricted to Exports (IMD)."""
    unit_selected = request.query_params.getlist('unit')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "WHERE Coll_BR_DESC = 'Exports (IMD)'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:  # Added dictionary=True for safety

            # Independent Filters
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC")
            years = [str(r.get('yr')) for r in cursor.fetchall() if r.get('yr')]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL")
            months = [r.get('mth') for r in cursor.fetchall() if r.get('mth')]

            cursor.execute(f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL")
            units = [r.get('Unit_Code') for r in cursor.fetchall() if r.get('Unit_Code')]

            # Sales Office is statically bound, no need to query the database
            sales_offices = ['Exports (IMD)']

            # Dependent Filter: Profit Center (Filtered by Unit)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params = []
            if unit_selected:
                placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({placeholders})"
                pc_params.extend(unit_selected)

            cursor.execute(pc_query, pc_params)
            profit_centers = [r.get('Profit_Centre') for r in cursor.fetchall() if r.get('Profit_Centre')]

    return jsonify({
        "years": years,
        "months": months,
        "units": units,
        "profit_centers": profit_centers,
        "sales_offices": sales_offices
    })


@prod_app.get('/api/collections/dashboard/imd')
def get_collections_dashboard_data_imd(request: Request):
    where, params = build_collections_filter_imd(request)

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

            # 2. Spend Trend
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


@prod_app.get('/api/export/collections_excel_imd')
def export_collections_excel_imd(request: Request):
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_imd(request)
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
        filename = f"Collections_Export.xlsx"

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f"Export Error: {e}")
        return HTMLResponse(f"Error exporting data: {str(e)}", status_code=500)



@prod_app.get('/api/ph1/pending_order_excel')
def ph1_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('HV Motors','HV Generators') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)



@prod_app.get('/api/ph2/pending_order_excel')
def ph2_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('DC Machines','RRM') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)

@prod_app.get('/api/ph6/pending_order_excel')
def ph6_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(
                "SELECT * FROM pending_order "
                "WHERE `As_on_Date` = :dt "
                "AND Branch_Name != 'Bangalore - IMD' "
                "AND ( "
                "    `Product2` = 'Transformer Pune' "
                "    OR (`Product2` = 'Transformer Mysore' AND Level_1 = 'Switchgear') "
                ")"
            )
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)



ph5_PENDING_PRODS = ('Transformer Mysore', 'Switchgear')

@prod_app.get('/export_ph5_pending_report')
def export_ph5_pending_report(request: Request):
    order_date = request.query_params.get('date')
    if not order_date:
        return HTMLResponse("As-on-Date is required", status_code=400)

    def set_slide_title(slide, text):
        title_shape = slide.shapes.title
        title_shape.text = text
        for paragraph in title_shape.text_frame.paragraphs:
            paragraph.font.size = Pt(24)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor(44, 62, 80)

    conn = get_db_connection()
    cursor = conn.cursor()
    prs = Presentation()

    # --- 1. KPI SUMMARY SLIDE ---
    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS total 
        FROM pending_order 
        WHERE As_on_Date = %s AND Product2 IN %s AND Branch_Name != 'Bangalore - IMD'
    """, (order_date, ph5_PENDING_PRODS))

    total_val = cursor.fetchone()['total'] or 0

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_title(slide, f"ph5 PENDING SUMMARY - As on {order_date}")

    table = slide.shapes.add_table(2, 2, Inches(1.5), Inches(1.5), Inches(7), Inches(2)).table
    kpi_rows = [
        ("Metric", "Value (Lakhs)"),
        ("ph5 Total Pending", f"{total_val:,.2f} L")
    ]

    for i, (label, val) in enumerate(kpi_rows):
        for col in range(2):
            cell = table.cell(i, col)
            cell.text = label if col == 0 else str(val)
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            para.font.size = Pt(18)
            if i == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(192, 57, 43)  # Distinct color for ph5
                para.font.color.rgb = RGBColor(255, 255, 255)

    # --- 2. PRODUCT-WISE SPLIT (HVM vs HVG) ---
    product_abbr = [('Transformer Mysore', 'TRF(M)'), ('SWITCHGEARS', 'SWGR')]
    cursor.execute("""
        SELECT 
                    CASE
                        WHEN Product2 = 'Transformer Mysore' AND Level_1 = 'Switchgear' THEN 'SWITCHGEARS'
                        ELSE Product2
                    END AS display_product, 
                    SUM(`Net_Value`)/100000 AS value
                FROM pending_order
        WHERE As_on_Date = %s AND Product2 IN %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY display_product
    """, (order_date, ph5_PENDING_PRODS))

    p_rows = cursor.fetchall()
    p_map = {r['display_product']: float(r['value'] or 0) for r in p_rows}

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_title(slide, f"ph5 Product-wise Pending ({order_date})")

    p_data = CategoryChartData()
    p_data.categories = [p[1] for p in product_abbr]
    p_data.add_series('Value (L)', [p_map.get(p[0], 0) for p in product_abbr])

    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(4.5),
                                   p_data).chart
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.number_format = '#,##0'

    # --- 3. BRANCH-WISE PENDING ---
    cursor.execute("""
        SELECT Branch_Name, SUM(Net_Value)/100000 AS value
        FROM pending_order
        WHERE As_on_Date = %s AND Product2 IN %s AND Branch_Name NOT IN ('', 'Bangalore - IMD')
        GROUP BY Branch_Name
    """, (order_date, ph5_PENDING_PRODS))

    b_rows = cursor.fetchall()
    processed_branches = {}
    for r in b_rows:
        name = 'Faridabad' if r['Branch_Name'].strip() in ['Delhi', 'Faridabad'] else r['Branch_Name'].strip()
        processed_branches[name] = processed_branches.get(name, 0) + float(r['value'] or 0)

    sorted_b = dict(sorted(processed_branches.items(), key=lambda x: x[1], reverse=True))

    slide = prs.slides.add_slide(prs.slide_layouts[5])
    set_slide_title(slide, f"ph5 Branch-wise Pending ({order_date})")

    b_data = CategoryChartData()
    b_data.categories = list(sorted_b.keys())
    b_data.add_series('Value (L)', list(sorted_b.values()))

    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.5), Inches(1.2), Inches(9), Inches(5.5),
                                   b_data).chart
    chart.series[0].has_data_labels = True
    chart.series[0].data_labels.font.size = Pt(8)

    # --- 4. SECTOR-WISE (PIE) ---
    cursor.execute("""
        SELECT Distr_Channel, SUM(Net_Value)/100000 AS value
        FROM pending_order
        WHERE As_on_Date = %s AND Product2 IN %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Distr_Channel
    """, (order_date, ph5_PENDING_PRODS))

    s_rows = cursor.fetchall()
    sec_dict = {}
    for r in s_rows:
        lbl = 'DD' if r['Distr_Channel'] in ['SZ', 'DD'] else r['Distr_Channel']
        sec_dict[lbl] = sec_dict.get(lbl, 0) + float(r['value'] or 0)

    if sec_dict:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        set_slide_title(slide, f"ph5 Sector-wise Distribution ({order_date})")
        s_data = CategoryChartData()
        s_data.categories = list(sec_dict.keys())
        s_data.add_series('Value (L)', list(sec_dict.values()))
        chart = slide.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(1.5), Inches(1.5), Inches(7), Inches(4.5),
                                       s_data).chart
        chart.has_legend = True
        chart.plots[0].has_data_labels = True
        chart.plots[0].data_labels.number_format = '#,##0 "L"'

    # --- 5. TOP 10 CUSTOMERS (BAR) ---
    cursor.execute("""
        SELECT Cust_Name, SUM(Net_Value)/100000 AS value
        FROM pending_order
        WHERE As_on_Date = %s AND Product2 IN %s AND Branch_Name != 'Bangalore - IMD'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, ph5_PENDING_PRODS))

    c_rows = cursor.fetchall()
    if c_rows:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        set_slide_title(slide, f"ph5 Top 10 Customers ({order_date})")
        c_data = CategoryChartData()

        categories = []
        for r in c_rows:
            name = r['Cust_Name'].strip() if r['Cust_Name'] else "Unknown"
            if "KIRLOSKAR ELECTRIC" in name.upper(): name = "KEC, Ajman"
            categories.append(name[:20])  # Truncate for chart

        c_data.categories = categories
        c_data.add_series('Value (L)', [float(r['value']) for r in c_rows])
        chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.5), Inches(1.2), Inches(9), Inches(5.5),
                                       c_data).chart
        chart.series[0].has_data_labels = True

    # --- FINALIZE ---
    conn.close()
    target_stream = io.BytesIO()
    prs.save(target_stream)
    target_stream.seek(0)

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=f"ph5_Pending_Report_{order_date}.pptx",
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )

ph5_BILLING_PRODS = ('TRANSFORMERS MYSORE',)
ph5_ORDER_PRODS = ('Transformer Mysore', 'Switchgear')

engine = create_engine("mysql+pymysql://root:@localhost/dashboard")
def get_ph5_condition(page_type):
    if page_type == "billing":
        return "Prod", ph5_BILLING_PRODS
    else:
        return "Product1", ph5_ORDER_PRODS


def format_date_columns(df):
    date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df


def serve_excel(df, filename):
    if df.empty:
        return HTMLResponse("No data found for the selected period", status_code=404)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Export_Data')
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )





@prod_app.get('/api/ph5/pending_order_excel')
def ph5_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query  = text(
    "SELECT * FROM pending_order "
    "WHERE Product1 IN('Transformer Mysore') "
    "AND Level_1 != 'Switchgear' "
    "AND As_on_Date = :dt"
)
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)


@prod_app.get('/api/ph3/pending_order_excel')
def ph3_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('LV Motors') AND As_on_Date = :dt AND Branch_Name IN ('Bangalore','Mumbai','Coimbatore','Chennai','Cochin','Hyderabad','Ahmedabad','Nagpur') ")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)



@prod_app.get('/api/ph4/pending_order_excel')
def ph4_pending_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('LV Motors') AND As_on_Date = :dt AND Branch_Name NOT IN ('Ludhiana','Delhi','Kolkata','Faridabad','Lucknow','Raipur')")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)


@prod_app.get('/api/export/pending_order_excel')
def pending_imd_order_excel(request: Request):
    order_date = request.query_params.get('date')
    page_type = request.query_params.get('type', 'pending_orders')

    if not order_date:
        return HTMLResponse("Please Select Date", status_code=400)
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE As_on_Date = :dt AND Branch_Name = 'Bangalore - IMD' ")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return HTMLResponse(f"No records found for {order_date} in pending_orders", status_code=404)

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return HTMLResponse(f"Database Error: {str(e)}", status_code=500)
