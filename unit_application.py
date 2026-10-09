
from flask import Blueprint, render_template, request, redirect, url_for, session, jsonify, flash,send_file, request
from db_config import get_db_connection
import io
from sqlalchemy import create_engine,text
import pandas as pd
from io import BytesIO
import shutil
import smbclient
import calendar
from smbprotocol.connection import Connection
import xlsxwriter
from apscheduler.schedulers.background import BackgroundScheduler
import smtplib
from email.message import EmailMessage
from datetime import datetime,timedelta
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import (
    XL_CHART_TYPE, XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
)
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

unit_app = Blueprint('unit_head', __name__)

@unit_app.route('/unit_home')
def unit_home():
    if session.get('role_type') != 'unit_head':
        return redirect(url_for('login'))
    return render_template('unit_home.html', username=session.get('username'))

@unit_app.route('/dashboard/un01')
def unit_1_dashboard():
    return render_template('Unit1/ytd.html', username=session['username'])

def build_filter_query_unit1(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 1 dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN01 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit = 'UN01'")
    # ----------------------------

    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
    units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/unit1', methods=['GET'])
def get_filters_unit1():
    """Fetches distinct filter options, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit = 'UN01'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions
            cursor.execute(f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units
            units = ['UN01']

            # 4. Dependent Office Filter (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM billing_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

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

@unit_app.route('/api/kpis/unit1', methods=['GET'])
def get_kpis_unit1():
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_unit1("data")
    target_where, target_params = build_filter_query_unit1("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
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


@unit_app.route('/api/sales-trend/unit1', methods=['GET'])
def get_sales_trend_unit1():
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_unit1("data")

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

@unit_app.route('/api/products-revenue/unit1', methods=['GET'])
def get_products_revenue_unit1():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_unit1("data")

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


@unit_app.route('/api/sales-by-office/unit1', methods=['GET'])
def get_sales_by_office_unit1():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_unit1("data")

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


@unit_app.route('/api/sales-by-region/unit1', methods=['GET'])
def get_sales_by_region_unit1():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_unit1("data")

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



@unit_app.route('/api/aop-achievement/unit1', methods=['GET'])
def get_aop_achievement_unit1():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_unit1("data")
    target_where, target_params = build_filter_query_unit1("target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.args.getlist('month')
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

@unit_app.route('/api/export/sales_excel/unit1', methods=['GET'])
def export_sales_excel_unit1():
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_unit1("data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500

@unit_app.route('/order_dashboard/unit1')
def unit_1_order_dashboard():
    return render_template('Unit1/order_ytd.html', username=session['username'])


def build_filter_orders_query_unit1(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 1 Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN01 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit = 'UN01'")
    # ----------------------------

    order_regions = request.args.getlist('region')
    order_offices = request.args.getlist('sales_office')
    order_products = request.args.getlist('product')
    order_units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/orders/unit1', methods=['GET'])
def get_orders_filters_unit1():
    """Fetches distinct filter options for Orders, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 IS NOT NULL AND Unit = 'UN01'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters (Regions, Fin Years)
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Units
            # We don't need to query the database for this since it's statically bound to UN01
            units = ['UN01']

            # 3. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Products (Mapped in Python to avoid SQL CASE delays)
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

@unit_app.route('/api/order/kpis/unit1', methods=['GET'])
def get_orders_kpis_unit1():
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_unit1("order_data")
    target_where, target_params = build_filter_orders_query_unit1("order_target")

    months = request.args.getlist('month')
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


@unit_app.route('/api/orders-trend/unit1', methods=['GET'])
def get_orders_trend_unit1():
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_unit1("order_data")

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


@unit_app.route('/api/orders/products-revenue/unit1', methods=['GET'])
def get_orders_products_revenue_unit1():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit1("order_data")

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


@unit_app.route('/api/orders/sales-by-office/unit1', methods=['GET'])
def get_orders_sales_by_office_unit1():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit1("order_data")

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


@unit_app.route('/api/order/sales-by-region/unit1', methods=['GET'])
def get_order_sales_by_region_unit1():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit1("order_data")

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


@unit_app.route('/api/orders/sales-engineers/unit1', methods=['GET'])
def get_orders_sales_engineers_unit1():
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_unit1("order_data")

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


@unit_app.route('/api/orders/aop-achievement/unit1', methods=['GET'])
def get_orders_aop_achievement_unit1():
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_unit1("order_data")
    target_where, target_params = build_filter_orders_query_unit1("order_target")

    months = request.args.getlist('month')
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

@unit_app.route('/api/export/orders_excel/unit1', methods=['GET'])
def export_orders_excel_unit1():
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_unit1("order_data")

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
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500


#---PENDING_ORDERS(un1)---
@unit_app.route('/un1/pending_orders')
def un1_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit1/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)

# --- 1. un1 TOTAL PENDING KPI ---
@unit_app.route('/api/un1/pending/total')
def get_un1_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND Unit = 'UN01'
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. un1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un1/pending/products')
def get_un1_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    product_abbr = [('HV Motors', 'HVM'), ('HV Generators', 'HVG'),('DC Machines','DCM'),('RRM','RRM'),('Traction','TRC')]
    un1_PENDING_PRODS = ('HV Motors','HV Generators','DC Machines','RRM','Traction',)
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT Product2, SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s AND `Product2` IN %s
            GROUP BY Product2
        """, (order_date, un1_PENDING_PRODS))

        rows = cursor.fetchall()
        actual_dict = {r['Product2']: float(r['value'] or 0) for r in rows}

        return jsonify({
            "labels": [p[1] for p in product_abbr],
            "values": [actual_dict.get(p[0], 0) for p in product_abbr]
        })
    finally:
        conn.close()


# --- 3. un1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un1/pending/branches')
def get_un1_pending_branches():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # un1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND Unit = 'UN01'
          AND Branch_Name != '' 
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


# --- 4. un1 PENDING SECTORS ---
@unit_app.route('/api/un1/pending/sectors')
def get_un1_pending_sectors():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND Unit = 'UN01'
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


# --- 5. un1 PENDING UNITS ---
@unit_app.route('/api/un1/pending/units')
def get_un1_pending_units():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Unit = 'UN01' AND `As_on_Date` = %s 
        GROUP BY Unit
    """, ( order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. un1 PENDING TOP 10 CUSTOMERS ---
@unit_app.route('/api/un1/pending/customers')
def get_un1_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND Unit = 'UN01'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date,))

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


# ---Unit 2----
@unit_app.route('/dashboard/unit2')
def unit_2_dashboard():
    return render_template('Unit2/ytd.html', username=session['username'])

def build_filter_query_unit2(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 2 dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN02 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit IN ('UN02','UN06')")
    # ----------------------------

    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
    units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/unit2', methods=['GET'])
def get_filters_unit2():
    """Fetches distinct filter options, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit IN ('UN02','UN06')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions
            cursor.execute(f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units
            units = ['UN02','UN06']

            # 4. Dependent Office Filter (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM billing_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

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

@unit_app.route('/api/kpis/unit2', methods=['GET'])
def get_kpis_unit2():
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_unit2("data")
    target_where, target_params = build_filter_query_unit2("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
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


@unit_app.route('/api/sales-trend/unit2', methods=['GET'])
def get_sales_trend_unit2():
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_unit2("data")

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

@unit_app.route('/api/products-revenue/unit2', methods=['GET'])
def get_products_revenue_unit2():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_unit2("data")

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


@unit_app.route('/api/sales-by-office/unit2', methods=['GET'])
def get_sales_by_office_unit2():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_unit2("data")

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


@unit_app.route('/api/sales-by-region/unit2', methods=['GET'])
def get_sales_by_region_unit2():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_unit2("data")

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



@unit_app.route('/api/aop-achievement/unit2', methods=['GET'])
def get_aop_achievement_unit2():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_unit2("data")
    target_where, target_params = build_filter_query_unit2("target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.args.getlist('month')
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

@unit_app.route('/api/export/sales_excel/unit2', methods=['GET'])
def export_sales_excel_unit2():
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_unit2("data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500


# un2 ORDER
@unit_app.route('/order_dashboard/unit_2')
def order_dashboard_unit_2():
    return render_template('Unit2/order_ytd.html', username=session['username'])


def build_filter_orders_query_unit2(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 1 Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN02 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit IN ('UN02','UN06')")
    # ----------------------------

    order_regions = request.args.getlist('region')
    order_offices = request.args.getlist('sales_office')
    order_products = request.args.getlist('product')
    order_units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/orders/unit2', methods=['GET'])
def get_orders_filters_unit2():
    """Fetches distinct filter options for Orders, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 IS NOT NULL AND Unit IN ( 'UN02','UN06')"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters (Regions, Fin Years)
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Units
            # We don't need to query the database for this since it's statically bound to UN01
            units = ['UN02','UN06']

            # 3. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Products (Mapped in Python to avoid SQL CASE delays)
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

@unit_app.route('/api/order/kpis/unit2', methods=['GET'])
def get_orders_kpis_unit2():
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_unit2("order_data")
    target_where, target_params = build_filter_orders_query_unit2("order_target")

    months = request.args.getlist('month')
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


@unit_app.route('/api/orders-trend/unit2', methods=['GET'])
def get_orders_trend_unit2():
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_unit2("order_data")

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


@unit_app.route('/api/orders/products-revenue/unit2', methods=['GET'])
def get_orders_products_revenue_unit2():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit2("order_data")

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


@unit_app.route('/api/orders/sales-by-office/unit2', methods=['GET'])
def get_orders_sales_by_office_unit2():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit2("order_data")

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


@unit_app.route('/api/order/sales-by-region/unit2', methods=['GET'])
def get_order_sales_by_region_unit2():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit2("order_data")

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


@unit_app.route('/api/orders/sales-engineers/unit2', methods=['GET'])
def get_orders_sales_engineers_unit2():
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_unit2("order_data")

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


@unit_app.route('/api/orders/aop-achievement/unit2', methods=['GET'])
def get_orders_aop_achievement_unit2():
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_unit2("order_data")
    target_where, target_params = build_filter_orders_query_unit2("order_target")

    months = request.args.getlist('month')
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

@unit_app.route('/api/export/orders_excel/unit2', methods=['GET'])
def export_orders_excel_unit2():
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_unit2("order_data")

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
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500
# ---PENDING_ORDERS(LVM(S))---

@unit_app.route('/un2/pending_orders')
def un2_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit2/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@unit_app.route('/api/un2/pending/total')
def get_un2_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND Unit IN ('UN02','UN06')
    """, (order_date,))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un2/pending/products')
def get_un2_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
        ('LV Motors (Non-Standard)', 'LVM(NS)'), ('LV Motors (Standard)', 'LVM(S)'),
        ('EVM', 'EVM'), ('LV Generators', 'LVG'), ('DG Sets','DGS'),
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
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un2/pending/branches')
def get_un2_pending_branches():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND Unit IN ('UN02','UN06')
          AND Branch_Name != '' 
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

UN2_PENDING_PRODS = ('LV Motors (Non-Standard)', 'LV Motors (Standard)', 'LV Generators', 'EVM', 'DG Sets')
# --- 4. PH1 PENDING SECTORS ---
@unit_app.route('/api/un2/pending/sectors')
def get_un2_pending_sectors():
    order_date = request.args.get('date')
    products = request.args.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
            SELECT `Distr_Channel` AS channel, SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s AND `Branch_Name` != '' AND  Unit IN ('UN02','UN06')
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
        params.append(tuple(UN2_PENDING_PRODS))

    query += " GROUP BY `Distr_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        return jsonify({"error": str(e)}), 500
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
@unit_app.route('/api/un2/pending/units')
def get_un2_pending_units():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Unit IN ('UN02','UN06') AND `As_on_Date` = %s 
        GROUP BY Unit
    """, (order_date))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@unit_app.route('/api/un2/pending/customers')
def get_un2_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND Unit IN ('UN02','UN06')
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

# ---TRF(M)&SWGR----

@unit_app.route('/unit_5/dashboard')
def unit_5_dashboard():
    return render_template('Unit5/ytd.html', username=session['username'])

def build_filter_query_unit5(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 16 dashboard.
    Optimized for MySQL using Python mapping and hardcoded TRANSFORMERS PUNE restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Product = 'TRANSFORMERS PUNE' for all queries on this dashboard
    if table_type == "data":
        filters.append("Prod IN ('TRANSFORMERS MYSORE','SWITCHGEAR MYSORE')")
    else:
        filters.append("Product IN ('TRANSFORMERS MYSORE','SWITCHGEAR MYSORE')")
    # ----------------------------

    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
    units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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


@unit_app.route('/api/filters/unit5', methods=['GET'])
def get_filters_unit5():
    """Fetches distinct filter options, restricted strictly to TRANSFORMERS PUNE."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod = 'TRANSFORMERS PUNE'"

    with get_db_connection() as conn:
        with conn.cursor(dictionary=True) as cursor: # Added dictionary=True to prevent .get() TypeErrors
            # 1. Regions
            cursor.execute(f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units
            # Since the data is bound to Transformers Pune, the mapped unit is always UN16
            units = ['UN05']

            # 4. Dependent Office Filter (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM billing_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 5. Products
            # We don't need to query the database for this since it's statically bound to TRANSFORMERS PUNE
            products = ['TRANSFORMERS MYSORE','SWITCHGEAR MYSORE']

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })


@unit_app.route('/api/kpis/unit5', methods=['GET'])
def get_kpis_unit5():
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_unit5("data")
    target_where, target_params = build_filter_query_unit5("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
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


@unit_app.route('/api/sales-trend/unit5', methods=['GET'])
def get_sales_trend_unit5():
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_unit5("data")

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

@unit_app.route('/api/products-revenue/unit5', methods=['GET'])
def get_products_revenue_unit5():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_unit5("data")

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


@unit_app.route('/api/sales-by-office/unit5', methods=['GET'])
def get_sales_by_office_unit5():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_unit5("data")

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


@unit_app.route('/api/sales-by-region/unit5', methods=['GET'])
def get_sales_by_region_unit5():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_unit5("data")

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



@unit_app.route('/api/aop-achievement/unit5', methods=['GET'])
def get_aop_achievement_unit5():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_unit5("data")
    target_where, target_params = build_filter_query_unit5("target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.args.getlist('month')
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

@unit_app.route('/api/export/sales_excel/unit5', methods=['GET'])
def export_sales_excel_unit5():
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_unit5("data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500


@unit_app.route('/order_dashboard/unit5')
def order_dashboard_unit5():
    return render_template('Unit5/order_ytd.html', username=session['username'])

def build_filter_orders_query_unit5(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 16 Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded Transformer Pune restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Transformer Pune for all queries on this dashboard
    filters.append("Product1 IN ( 'Transformer Mysore','Switchgear')")
    # ----------------------------

    order_regions = request.args.getlist('region')
    order_offices = request.args.getlist('sales_office')
    order_products = request.args.getlist('product')
    order_units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    where_clause = " AND ".join(filters) if filters else "1=1"
    return "WHERE " + where_clause, params


@unit_app.route('/api/filters/orders/unit5', methods=['GET'])
def get_orders_filters_unit5():
    """Fetches distinct filter options for Orders, restricted strictly to Transformer Pune."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 IN ( 'Transformer Mysore','Switchgear') "

    with get_db_connection() as conn:
        with conn.cursor(dictionary=True) as cursor:
            # 1. Independent Filters (Regions, Fin Years)
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL"
            )
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(
                f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL"
            )
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Units
            units = ['UN05']

            # 3. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Products
            products = [ 'Transformer Mysore','Switchgear']

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@unit_app.route('/api/order/kpis/unit5', methods=['GET'])
def get_orders_kpis_unit5():
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_unit5("order_data")
    target_where, target_params = build_filter_orders_query_unit5("order_target")

    months = request.args.getlist('month')
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


@unit_app.route('/api/orders-trend/unit5', methods=['GET'])
def get_orders_trend_unit5():
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_unit5("order_data")

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


@unit_app.route('/api/orders/products-revenue/unit5', methods=['GET'])
def get_orders_products_revenue_unit5():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit5("order_data")

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


@unit_app.route('/api/orders/sales-by-office/unit5', methods=['GET'])
def get_orders_sales_by_office_unit5():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit5("order_data")

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


@unit_app.route('/api/order/sales-by-region/unit5', methods=['GET'])
def get_order_sales_by_region_unit5():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit5("order_data")

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


@unit_app.route('/api/orders/sales-engineers/unit5', methods=['GET'])
def get_orders_sales_engineers_unit5():
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_unit5("order_data")

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


@unit_app.route('/api/orders/aop-achievement/unit5', methods=['GET'])
def get_orders_aop_achievement_unit5():
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_unit5("order_data")
    target_where, target_params = build_filter_orders_query_unit5("order_target")

    months = request.args.getlist('month')
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

@unit_app.route('/api/export/orders_excel/unit5', methods=['GET'])
def export_orders_excel_unit5():
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_unit5("order_data")

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
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500

# ---PENDING_ORDERS(PH1)---

@unit_app.route('/un5/pending_orders')
def un5_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit5/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)

PH3_PENDING_PRODS = ('Transformer Mysore', 'Switchgear')

# --- 1. PH1 TOTAL PENDING KPI ---
@unit_app.route('/api/un5/pending/total')
def get_un5_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2` IN %s
    """, (order_date, PH3_PENDING_PRODS))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un5/pending/products')
def get_un5_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('Transformer Mysore', 'TRF(M)'), ('SWITCHGEARS', 'SWG'),
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
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un5/pending/branches')
def get_un5_pending_branches():
    order_date = request.args.get('date')
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
          AND Branch_Name != '' 
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
@unit_app.route('/api/un5/pending/sectors')
def get_un5_pending_sectors():
    order_date = request.args.get('date')
    products = request.args.getlist('product')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
            SELECT `Distr_Channel` AS channel, SUM(`Net_Value`)/100000 AS value
            FROM pending_order
            WHERE `As_on_Date` = %s AND `Branch_Name` != ''
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
                # Standard product logic
                product_clauses.append("`Product2` = %s")
                params.append(p)

        query += " AND (" + " OR ".join(product_clauses) + ")"
    else:
        # Default behavior if no products are passed
        query += " AND `Product2` IN %s"
        params.append(tuple(PH3_PENDING_PRODS))

    query += " GROUP BY `Distr_Channel`"

    try:
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

    processed_data = {}
    for r in rows:
        chan = r['channel']
        val = float(r['value'] or 0)
        # Mapping SZ and DD to a single 'DD' label
        label = 'DD' if chan in ['SZ', 'DD'] else chan
        processed_data[label] = processed_data.get(label, 0) + val

    return jsonify({
        "labels": list(processed_data.keys()),
        "values": [round(float(v), 2) for v in processed_data.values()]
    })


# --- 5. PH1 PENDING UNITS ---
@unit_app.route('/api/un5/pending/units')
def get_un5_pending_units():
    order_date = request.args.get('date')
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
@unit_app.route('/api/un5/pending/customers')
def get_un5_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` IN %s
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

# ---TRF(M)&SWGR----

@unit_app.route('/unit_6/dashboard')
def unit_6_dashboard():
    return render_template('Unit6/ytd.html', username=session['username'])

@unit_app.route('/api/un6/ytd/kpis')
def get_un6_ytd_kpi_data():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS total FROM billing_data WHERE fin_year=%s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'",
        (year, ))
    actual = cursor.fetchone()['total'] or 0
    cursor.execute("SELECT SUM(Target) AS TARGET FROM billing_target where fin_year=%s AND Product = 'DG SETS' ",
                   (year, ))
    target = cursor.fetchone()['TARGET'] or 0
    prev_y_start = int(year.split('-')[0]) - 1
    prev_y_end = int(year.split('-')[0])
    prev_year_str = f"{prev_y_start}-{prev_y_end}"
    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS total FROM billing_data WHERE fin_year=%s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'",
        (prev_year_str, ))
    last_year_actual = cursor.fetchone()['total'] or 0

    conn.close()

    percent_val = (actual / target * 100) if target > 0 else 0
    return jsonify({
        "actual": f"{float(actual):,.5f}L",
        "target": f"{float(target):,.2f}L",
        "percent": percent_val,
        "raw_percent": percent_val,
        "last_year": f"{float(last_year_actual):,.5f}L"
    })

@unit_app.route('/api/un6/ytd/products')
def get_un6_ytd_product_data():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_ORDER = [
        ('DG SETS','DGS')

    ]

    DATABASE_VALID_PRODS = (
       'DG SETS',
    )

    query = """
        SELECT 
            CASE 
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                WHEN Prod = 'TRANSFORMERS MYSORE' AND VTEXT1 = 'Switchgear' THEN 'SWITCHGEARS'
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



@unit_app.route('/api/un6/ytd/regions')
def get_un6_ytd_region_data():
    year = request.args.get('fin_year')
    if not year:
        return jsonify({"error": "fin_year required"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')


    query = """
        SELECT 
            CASE 
                WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket'
                ELSE Sales_Region_Name
            END AS display_region, 
            SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_region
        HAVING display_region IN %s
    """
    cursor.execute(query, (year, valid_regions))
    rows = cursor.fetchall()


    target_query = """
        SELECT Sales_Region_Name AS region, SUM(Target) AS annual_target
        FROM billing_target
        WHERE Sales_Region_Name IN %s AND fin_year = %s AND Product = 'DG SETS'
        GROUP BY Sales_Region_Name
    """
    cursor.execute(target_query, (valid_regions, year,))
    target_rows = cursor.fetchall()
    conn.close()

    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['region']: float(r['annual_target'] or 0) for r in target_rows}

    region_labels, region_actuals, region_targets = [], [], []
    for region in valid_regions:
        region_labels.append(region)
        region_actuals.append(actual_dict.get(region, 0))
        region_targets.append(target_dict.get(region, 0))

    return jsonify({"labels": region_labels, "actuals": region_actuals, "targets": region_targets})


@unit_app.route('/api/un6/ytd/sector')
def get_un6_ytd_sector_data():
    year = request.args.get('fin_year')
    if not year:
        return jsonify({"error": "fin_year required"}), 400

    valid_sectors = ('Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ')

    query = """
        SELECT Dist_Channel_TEXT AS channel,
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s 
          AND Prod = 'DG SETS'
          AND Dist_Channel_TEXT IN %s
          AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Dist_Channel_TEXT
    """

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(query, (year, valid_sectors))
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

        return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})
    finally:
        conn.close()


@unit_app.route('/api/un6/ytd/Unit')
def get_un6_ytd_Unit_data():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT Sales_office AS display_unit, 
               SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE Prod =  'DG SETS' AND  fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_unit
    """
    cursor.execute(query, (year,))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })



@unit_app.route('/api/un6/get_ytd_name_data')
def get_un6_ytd_name():
    fin_year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT Name, SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND  Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Name
        ORDER BY total_value DESC
        LIMIT 10
    """
    cursor.execute(query, (fin_year,))
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


@unit_app.route('/mtd/un_6')
def mtd_unit_6():
    conn = get_db_connection()
    years = []
    if conn:
        cursor = conn.cursor()

        cursor.execute("SELECT DISTINCT fin_year FROM billing_data ORDER BY fin_year DESC")
        years = [row['fin_year'] for row in cursor.fetchall()]
        conn.close()

    selected_year = request.args.get('fin_year', years[0] if years else "")
    return render_template('Unit6/mtd.html', years=years, selected_year=selected_year, username=session['username'])


@unit_app.route('/api/un6/mtd/billing_kpi')
def get_un6_mtd_billing_kpi():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')

    if not months or not fin_year:
        return jsonify({"error": "Missing month or fin_year"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. ACTUAL MTD VALUE (un6 Products Only)
    actual_sql = """
        SELECT SUM(`Net_Value`)/100000 AS total_val 
        FROM billing_data 
        WHERE fin_year=%s 
        AND Prod = 'DG SETS'
        AND MONTHNAME(`Billing_Date`) IN (%s)
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
    """

    placeholders = ', '.join(['%s'] * len(months))
    actual_sql = actual_sql % ('%s', placeholders)

    cursor.execute(actual_sql, (fin_year, *months))
    res = cursor.fetchone()
    actual_val = res['total_val'] if res and res['total_val'] else 0

    target_sql = """
        SELECT SUM(`Target`) AS total_target 
        FROM billing_target 
        WHERE fin_year=%s AND Product = 'DG SETS'
    """
    cursor.execute(target_sql, (fin_year,))
    t_res = cursor.fetchone()
    annual_target = t_res['total_target'] if t_res and t_res['total_target'] else 0

    num_months = len(months)
    pro_rata_target = (annual_target / 12) * num_months

    try:
        current_month_name = months[0]
        current_year_start = int(fin_year.split('-')[0])

        month_dt = datetime.strptime(current_month_name, "%B")
        month_number = month_dt.month

        if month_number == 4:
            last_month_num = 3
            lm_y_start, lm_y_end = current_year_start - 1, current_year_start
        else:
            last_month_num = 12 if month_number == 1 else month_number - 1
            lm_y_start, lm_y_end = current_year_start, current_year_start + 1

        last_month_name = datetime(2000, last_month_num, 1).strftime('%B')
        last_month_fin_year = f"{lm_y_start}-{lm_y_end}"

        lm_sql = """
            SELECT SUM(`Net_Value`)/100000 AS total 
            FROM billing_data 
            WHERE fin_year=%s 
            AND Prod = 'DG SETS'
            AND MONTHNAME(`Billing_Date`) = %s
            AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        """
        cursor.execute(lm_sql, (last_month_fin_year, last_month_name))
        lm_res = cursor.fetchone()
        last_month_actual = lm_res['total'] if lm_res and lm_res['total'] else 0
    except Exception as e:
        print(f"Error PH2 Last Month: {e}")
        last_month_actual = 0

    conn.close()

    achievement_pct = (actual_val / pro_rata_target * 100) if pro_rata_target > 0 else 0

    return jsonify({
        "actual_value": f"{float(actual_val):,.2f}L",
        "target_value": f"{float(pro_rata_target):,.2f}L",
        "achievement": f"{achievement_pct:.1f}%",
        "last_month_actual": f"{float(last_month_actual):,.2f}L"
    })


@unit_app.route('/api/un6/mtd/product_breakdown')
def get_un6_mtd_product_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    if not fin_year:
        return jsonify({"error": "Missing fin_year"}), 400

    PRODUCT_ORDER = [
      ('DG SETS','DGS'),
    ]

    # Valid product names as they appear in the billing_data 'Prod' column
    DATABASE_VALID_PRODS = (
      'DG SETS',
    )

    conn = None
    cursor = None
    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        actual_query = """
                SELECT 
                    CASE 
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                        WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                        WHEN Prod = 'TRANSFORMERS MYSORE' AND VTEXT1='Switchgear' THEN 'SWITCHGEARS'
                        ELSE Prod
                    END AS display_prod, 
                    SUM(Net_Value)/100000 as actual
                FROM billing_data 
                WHERE Prod IN %s AND fin_year = %s AND Dist_Channel_TEXT != 'Inter Unit Transfer' 
            """
        actual_params = [DATABASE_VALID_PRODS, fin_year]

        if months:
            placeholders = ', '.join(['%s'] * len(months))
            actual_query += f" AND MONTHNAME(Billing_Date) IN ({placeholders})"
            actual_params.extend(months)

        actual_query += " GROUP BY display_prod"
        cursor.execute(actual_query, tuple(actual_params))
        actual_rows = cursor.fetchall()
        actual_dict = {row['display_prod']: float(row['actual'] or 0) for row in actual_rows}
        target_query = """
                SELECT Product, SUM(Target) as annual_target 
                FROM billing_target 
                WHERE fin_year = %s 
                GROUP BY Product
            """
        cursor.execute(target_query, (fin_year,))
        target_rows = cursor.fetchall()

        num_months = len(months) if months else 12
        month_factor = num_months / 12.0

        target_dict = {row['Product']: float(row['annual_target'] or 0) * month_factor for row in target_rows}

        final_labels = []
        final_actuals = []
        final_targets = []

        for full_name, abbr in PRODUCT_ORDER:
            final_labels.append(abbr)
            final_actuals.append(actual_dict.get(full_name, 0))
            final_targets.append(target_dict.get(full_name, 0))

        return jsonify({
            "labels": final_labels,
            "actuals": final_actuals,
            "targets": final_targets,
            "fin_year": fin_year,
            "months": months
        })
    except Exception as e:
        return jsonify({"error": f"Database error: {str(e)}"}), 500
    finally:
        if cursor: cursor.close()
        if conn: conn.close()

@unit_app.route('/api/un6/mtd/regions')
def get_un6_mtd_region_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')

    conn = get_db_connection()
    cursor = conn.cursor()
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    # 1. Actuals for PH1 only
    actual_query = """
        SELECT 
            CASE WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket' ELSE Sales_Region_Name END AS display_region, 
            SUM(Net_Value)/100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        AND MONTHNAME(Billing_Date) IN (%s)
        GROUP BY display_region
    """
    placeholders = ', '.join(['%s'] * len(months))
    actual_query = actual_query % ('%s', placeholders)
    cursor.execute(actual_query, (fin_year,  *months))
    actual_rows = cursor.fetchall()

    # 2. Targets for PH1
    target_query = """
        SELECT Sales_Region_Name AS region, SUM(Target) AS annual_target
        FROM billing_target
        WHERE Sales_Region_Name IN %s AND fin_year = %s AND Product = 'DG SETS'
        GROUP BY Sales_Region_Name
    """
    cursor.execute(target_query, (valid_regions, fin_year, ))
    target_rows = cursor.fetchall()
    conn.close()

    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in actual_rows}
    month_factor = len(months) / 12.0
    target_dict = {r['region']: float(r['annual_target'] or 0) * month_factor for r in target_rows}

    labels, actuals, targets = [], [], []
    for region in valid_regions:
        labels.append(region)
        actuals.append(actual_dict.get(region, 0))
        targets.append(target_dict.get(region, 0))

    return jsonify({"labels": labels, "actuals": actuals, "targets": targets})


@unit_app.route('/api/un6/mtd/sectors')
def get_un6_mtd_sector_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    valid_sectors = ('Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ')
    query = """
        SELECT Dist_Channel_TEXT as channel, SUM(Net_Value)/100000 as actual
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND Dist_Channel_TEXT IN %s
        AND MONTHNAME(Billing_Date) IN (%s)
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Dist_Channel_TEXT
    """
    placeholders = ', '.join(['%s'] * len(months))
    query = query % ('%s',  '%s', placeholders)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, (fin_year,  valid_sectors, *months))
    rows = cursor.fetchall()
    conn.close()

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

    return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})


@unit_app.route('/api/un6/mtd/units')
def get_un6_mtd_unit_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    query = """
        SELECT Sales_office as display_unit, SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE Prod = 'DG SETS'  AND fin_year = %s AND MONTHNAME(Billing_Date) IN (%s)
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_unit
    """
    placeholders = ', '.join(['%s'] * len(months))
    query = query % ( '%s', placeholders)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, (fin_year, *months))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })


@unit_app.route('/api/un6/mtd/top_customers')
def get_un6_mtd_top_customers():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    query = """
        SELECT Name, SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND MONTHNAME(Billing_Date) IN (%s)
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Name
        ORDER BY total_value DESC LIMIT 10
    """
    placeholders = ', '.join(['%s'] * len(months))
    query = query % ('%s', placeholders)

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, (fin_year, *months))
    rows = cursor.fetchall()
    conn.close()

    # Apply your KEC replacement logic
    processed_labels = []
    for r in rows:
        name = r['Name'].strip() if r['Name'] else "Unknown"
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        processed_labels.append(name)

    return jsonify({"labels": processed_labels, "values": [float(r['total_value'] or 0) for r in rows]})



@unit_app.route('/un6/qtd')
def qtd_unit_6():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT fin_year FROM billing_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    available_years = [row['fin_year'] for row in cursor.fetchall()]

    default_year = available_years[0] if available_years else "2025-2026"

    conn.close()

    return render_template(
        'Unit6/qtd.html',
        username=session.get('username', 'User'),
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


@unit_app.route('/api/un6/quarterly/kpis')
def get_un6_quarterly_kpis():
    year = request.args.get('fin_year')
    quarters_list = request.args.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({"actual": "0L", "target": "0L", "achievement": 0, "last_year": "0L"})

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT SUM(Net_Value)/100000 AS total 
        FROM billing_data 
        WHERE fin_year=%s AND Prod = 'DG SETS' AND MONTH(Billing_Date) IN %s
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
    """, (year, months))
    actual = cursor.fetchone()['total'] or 0

    cursor.execute("""
        SELECT SUM(Target) AS annual FROM billing_target 
        WHERE fin_year=%s AND Product = 'DG SETS'
    """, (year,))
    annual_target = cursor.fetchone()['annual'] or 0

    target = (annual_target / 4) * len(quarters_list)

    try:
        start_year = int(year.split('-')[0])
        prev_year_str = f"{start_year - 1}-{start_year}"
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total 
            FROM billing_data 
            WHERE fin_year=%s AND Prod = 'DG SETS' AND MONTH(Billing_Date) IN %s
            AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        """, (prev_year_str, months))
        last_year_actual = cursor.fetchone()['total'] or 0
    except:
        last_year_actual = 0

    conn.close()
    percent_val = (actual / target * 100) if target > 0 else 0

    return jsonify({
        "actual": f"{float(actual):,.5f}L",
        "target": f"{float(target):,.0f}L",
        "achievement": round(percent_val, 1),
        "raw_percent": percent_val,
        "last_year": f"{float(last_year_actual):,.5f}L"
    })


@unit_app.route('/api/un6/quarterly/products')
def get_un6_quarterly_product_data():
    year = request.args.get('fin_year')
    quarters = request.args.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_ORDER = [
        ('DG SETS', 'DGS'),
    ]

    query = """
            SELECT 
                CASE 
                    WHEN Prod = 'LV MOTORS' AND Plant = 'AC02' THEN 'LV MOTORS (NS)'
                    WHEN Prod = 'LV MOTORS' AND Plant = 'AC25' THEN 'LV MOTORS (S)'
                    WHEN Prod = 'TRANSFORMERS MYSORE' AND VTEXT1 = 'Switchgear' THEN 'SWITCHGEARS'
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

@unit_app.route('/api/un6/quarterly/regions')
def get_un6_quarterly_region_data():
    year = request.args.get('fin_year')
    quarters = request.args.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    query = """
        SELECT CASE WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket' ELSE Sales_Region_Name END AS display_region, 
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_region HAVING display_region IN %s
    """
    cursor.execute(query, (year, months, valid_regions))
    rows = cursor.fetchall()

    target_query = """
        SELECT Sales_Region_Name AS region, (SUM(Target)/4)*%s AS q_target 
        FROM billing_target 
        WHERE fin_year = %s AND Product = 'DG SETS'
        GROUP BY region
    """
    cursor.execute(target_query, (len(quarters), year, ))
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


@unit_app.route('/api/un6/quarterly/sector')
def get_un6_quarterly_sector_data():
    year = request.args.get('fin_year')
    quarters = request.args.getlist('quarters')
    months = get_quarter_months(quarters)

    if not year or not months:
        return jsonify({"labels": [], "values": []})

    valid_sectors = ('Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ')

    query = """
        SELECT Dist_Channel_TEXT AS channel,
               SUM(Net_Value) / 100000 AS actual
        FROM billing_data
        WHERE fin_year = %s 
          AND Prod = 'DG SETS'
          AND MONTH(Billing_Date) IN %s
          AND Dist_Channel_TEXT IN %s 
          AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Dist_Channel_TEXT
    """

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, (year,  months, valid_sectors))
    rows = cursor.fetchall()
    conn.close()

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

    return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})


@unit_app.route('/api/un6/quarterly/unit')
def get_un6_quarterly_unit_data():
    year = request.args.get('fin_year')
    quarters = request.args.getlist('quarters')
    months = get_quarter_months(quarters)
    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT Sales_office AS display_unit, 
               SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_unit
    """
    cursor.execute(query, (year, months))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({"labels": [r['display_unit'] for r in rows], "values": [float(r['total'] or 0) for r in rows]})


@unit_app.route('/api/un6/quarterly/top_customers')
def get_un6_quarterly_customers():
    year = request.args.get('fin_year')
    quarters = request.args.getlist('quarters')
    months = get_quarter_months(quarters)

    conn = get_db_connection()
    cursor = conn.cursor()
    query = """
        SELECT Name, SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE fin_year = %s AND Prod = 'DG SETS' AND MONTH(Billing_Date) IN %s AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Name ORDER BY total_value DESC LIMIT 10
    """
    cursor.execute(query, (year,  months))
    rows = cursor.fetchall()
    conn.close()

    processed_labels = []
    for r in rows:
        name = r['Name'].strip() if r['Name'] else ""
        if "KIRLOSKAR ELECTRIC COMPANY LIMITED" in name.upper():
            name = "KEC, Ajman"
        processed_labels.append(name)

    return jsonify({"labels": processed_labels, "values": [float(r['total_value'] or 0) for r in rows]})


@unit_app.route('/un6/daily')
def un6_daily():
    today_date = datetime.now().strftime('%Y-%m-%d')
    selected_date = request.args.get('date', today_date)
    return render_template('Unit6/daily_sales.html', selected_date=selected_date, username=session['username'])


@unit_app.route('/api/un6/daily/summary')
def get_un6_daily_summary():
    selected_date_str = request.args.get('date')
    if not selected_date_str:
        return jsonify({"error": "No date selected"}), 400

    selected_date = datetime.strptime(selected_date_str, '%Y-%m-%d')
    day_of_month = selected_date.day

    month = selected_date.month
    year = selected_date.year
    fin_year = f"{year}-{year + 1}" if month >= 4 else f"{year - 1}-{year}"
    mtd_start = selected_date.replace(day=1).strftime('%Y-%m-%d')

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # 1. PH1 Today Actual
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
            WHERE Billing_Date = %s AND Prod = 'DG SETS'
            AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        """, (selected_date_str, ))
        today_actual = cursor.fetchone()['total'] or 0

        # 2. PH1 MTD Actual (Cumulative to selected date)
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS total FROM billing_data 
            WHERE Billing_Date >= %s AND Billing_Date <= %s 
            AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        """, (mtd_start, selected_date_str, ))
        mtd_actual = cursor.fetchone()['total'] or 0

        # 3. PH1 Targets (Annual / 365)
        cursor.execute("""
            SELECT SUM(Target) as annual FROM billing_target 
            WHERE fin_year=%s AND Product = 'DG SETS'
        """, (fin_year,))
        annual_target = cursor.fetchone()['annual'] or 0

        daily_target = annual_target / 365
        achievement = (today_actual / daily_target * 100) if daily_target > 0 else 0

        return jsonify({
            "today_actual": f"{float(today_actual):,.5f}",
            "today_target": f"{float(daily_target):,.2f}",
            "achievement": round(achievement, 1),
            "mtd_actual": f"{float(mtd_actual):,.5f} L",
        })
    finally:
        conn.close()


@unit_app.route('/api/un6/daily/products')
def get_un6_daily_products():
    selected_date = request.args.get('date')
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
                    WHEN Prod = 'TRANSFORMERS MYSORE' AND VTEXT1='Switchgear' THEN 'SWITCHGEARS'
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
         ('DG SETS', 'DGS'),
    ]

    return jsonify({
        "labels": [p[1] for p in PRODUCT_MAP],
        "actuals": [actual_map.get(p[0], 0) for p in PRODUCT_MAP],
        "targets": [target_map.get(p[0], 0) for p in PRODUCT_MAP]
    })


@unit_app.route('/api/un6/daily/regions')
def get_un6_daily_regions():
    selected_date = request.args.get('date')
    date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{date_obj.year}-{date_obj.year + 1}" if date_obj.month >= 4 else f"{date_obj.year - 1}-{date_obj.year}"
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
            CASE WHEN Sales_office = 'Bangalore - IMD' THEN 'InternationalMarket' ELSE Sales_Region_Name END AS display_region, 
            SUM(Net_Value)/100000 AS actual
        FROM billing_data
        WHERE Billing_Date = %s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY display_region
    """, (selected_date,))
    actual_dict = {r['display_region']: float(r['actual'] or 0) for r in cursor.fetchall()}

    cursor.execute("""
        SELECT Sales_Region_Name as region, SUM(Target)/365 as daily_target 
        FROM billing_target 
        WHERE fin_year=%s AND Product = 'DG SETS'
        GROUP BY region
    """, (fin_year,))
    target_dict = {r['region']: float(r['daily_target'] or 0) for r in cursor.fetchall()}
    conn.close()

    return jsonify({
        "labels": valid_regions,
        "actuals": [actual_dict.get(r, 0) for r in valid_regions],
        "targets": [target_dict.get(r, 0) for r in valid_regions]
    })


@unit_app.route('/api/un6/daily/sectors')
def get_un6_daily_sectors():
    selected_date = request.args.get('date')
    valid_sectors = ['Direct Export', 'Domestic Direct', 'Domestic Dealers', 'Deemed Export', 'SEZ']

    query = """
        SELECT Dist_Channel_TEXT as channel, SUM(Net_Value)/100000 as actual
        FROM billing_data
        WHERE Dist_Channel_TEXT IN %s 
        AND Billing_Date = %s 
        AND Prod = 'DG SETS'
        AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Dist_Channel_TEXT
    """

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, (valid_sectors, selected_date,))
    rows = cursor.fetchall()
    conn.close()

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


@unit_app.route('/api/un6/daily/units')
def get_un6_daily_units():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Sales_office AS display_unit, SUM(Net_Value)/100000 as total 
        FROM billing_data
        WHERE Billing_Date = %s AND Prod = 'DG SETS' AND  Dist_Channel_TEXT != 'Inter Unit Transfer'
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


@unit_app.route('/api/un6/daily/customers')
def get_un6_daily_customers():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Name, SUM(Net_Value) / 100000 AS total_value
        FROM billing_data
        WHERE Billing_Date = %s AND Prod = 'DG SETS' AND Dist_Channel_TEXT != 'Inter Unit Transfer'
        GROUP BY Name ORDER BY total_value DESC LIMIT 10
    """
    cursor.execute(query, (selected_date,))
    rows = cursor.fetchall()
    conn.close()

    processed_labels = []
    processed_values = []
    for r in rows:
        name = r['Name'].strip() if r['Name'] else "Unknown"
        if "KIRLOSKAR ELECTRIC" in name.upper():
            name = "KEC, Ajman"
        processed_labels.append(name)
        processed_values.append(float(r['total_value'] or 0))

    return jsonify({"labels": processed_labels, "values": processed_values})


@unit_app.route('/order_dashboard/Unit6')
def order_dashboard_Unit6():
    return render_template('Unit6/order_ytd.html', username=session['username'])


@unit_app.route('/api/un6/order_ytd/kpis')
def get_un6_order_ytd_kpi_data():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()


    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS total FROM order_data WHERE fin_year=%s AND Product1 = 'DG Sets'",
        (year, ))
    actual = cursor.fetchone()['total'] or 0
    cursor.execute("SELECT SUM(Target) AS TARGET FROM order_target where fin_year=%s AND Product1 = 'DG Sets' ",
                   (year, ))
    target = cursor.fetchone()['TARGET'] or 0
    prev_y_start = int(year.split('-')[0]) - 1
    prev_y_end = int(year.split('-')[0])
    prev_year_str = f"{prev_y_start}-{prev_y_end}"
    cursor.execute(
        "SELECT SUM(Net_Value)/100000 AS total FROM order_data WHERE fin_year=%s AND Product1  = 'DG Sets' ",
        (prev_year_str, ))
    last_year_actual = cursor.fetchone()['total'] or 0

    conn.close()
    achievement_pct = (actual / target * 100) if target > 0 else 0
    return jsonify({
        "actuals": f"{float(actual):,.5f}L",
        "last_year": f"{float(last_year_actual):,.5f}L",
        "targets": f"{float(target):,.0f}L",
        "achievement": f"{achievement_pct:.1f}%",
    })


@unit_app.route('/api/un6/order_ytd/products')
def get_un6_order_ytd_product_data():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    PRODUCT_ORDER = [
        ('DG Sets','DGS'),

    ]

    DATABASE_VALID_PRODS = (
        'DG Sets',
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
    cursor.execute(query, (year, DATABASE_VALID_PRODS))
    rows = cursor.fetchall()

    target_query = "SELECT Product1, SUM(Target) as annual_target FROM order_target WHERE fin_year = %s GROUP BY Product1"
    cursor.execute(target_query, (year,))
    target_rows = cursor.fetchall()
    conn.close()

    results_dict = {r['display_product']: float(r['actual'] or 0) for r in rows}
    target_dict = {r['Product1']: float(r['annual_target'] or 0) for r in target_rows}

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


@unit_app.route('/api/un6/ytd/order_regions')
def get_un6_order_ytd_region_data():
    year = request.args.get('fin_year')
    if not year:
        return jsonify({"error": "Financial year required"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    query = """
        SELECT Sales_Region AS region,
               SUM(Net_Value)/100000 AS actual
        FROM order_data
        WHERE fin_year = %s 
          AND Sales_Region IN %s 
          AND Product1  = 'DG Sets'
        GROUP BY Sales_Region
    """
    try:
        cursor.execute(query, (year, valid_regions, ))
        rows = cursor.fetchall()
        actual_dict = {r['region']: float(r['actual'] or 0) for r in rows}

        target_query = """
            SELECT Sales_Region AS region, 
                   SUM(Target) AS target
            FROM order_target 
            WHERE Sales_Region IN %s 
              AND fin_year = %s 
              AND Product1  = 'DG Sets'
            GROUP BY Sales_Region
        """
        cursor.execute(target_query, (valid_regions, year, ))
        target_rows = cursor.fetchall()
        target_dict = {r['region']: float(r['target'] or 0) for r in target_rows}

    except Exception as e:
        print(f"Database Error: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

    final_actuals = []
    final_targets = []

    for region in valid_regions:
        final_actuals.append(actual_dict.get(region, 0))
        final_targets.append(target_dict.get(region, 0))

    return jsonify({
        "labels": valid_regions,  # Use hardcoded labels for chart consistency
        "actuals": final_actuals,
        "targets": final_targets
    })



@unit_app.route('/api/un6/order_ytd/sector')
def get_un6_order_ytd_sector_data():
    year = request.args.get('fin_year')
    if not year:
        return jsonify({"error": "fin_year is required"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
            SELECT `Dist_Channel` AS channel, 
                   SUM(`Net_Value`)/100000 as actual 
            FROM order_data 
            WHERE fin_year = %s AND `Product1`  = 'DG Sets'
            GROUP BY `Dist_Channel`
        """

    try:
        cursor.execute(query, (year, ))
        rows = cursor.fetchall()
        conn.close()

        processed_data = {}
        for r in rows:
            label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
            processed_data[label] = processed_data.get(label, 0) + float(r['actual'] or 0)

        return jsonify({
            "labels": list(processed_data.keys()),
            "values": list(processed_data.values())
        })
    except Exception as e:
        if conn: conn.close()
        return jsonify({"error": str(e)}), 500


@unit_app.route('/api/un6/order_ytd/unit')
def get_un6_order_ytd_unit_data():
    year = request.args.get('fin_year')
    if not year:
        return jsonify({"error": "fin_year is required"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Sales_office AS display_unit, 
               SUM(Net_Value)/100000 as total 
        FROM order_data
        WHERE Product1 = 'DG Sets' AND fin_year = %s 
        GROUP BY display_unit
    """

    cursor.execute(query, (year,))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })



@unit_app.route('/api/un6/order_ytd/top_customers')
def get_un6_order_ytd_customers():
    fin_year = request.args.get('fin_year')
    months = request.args.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        query = """
            SELECT Name, SUM(Net_Value) / 100000 AS total_value
            FROM order_data
            WHERE fin_year = %s AND Product1  = 'DG Sets'
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
            name = r['Name'].strip() if r['Name'] else "Unknown"
            for long_name, alias in NAME_REPLACEMENTS.items():
                if long_name in name.upper():
                    name = alias
                    break
            processed_labels.append(name)

        return jsonify({
            "labels": processed_labels,
            "values": [float(r['total_value'] or 0) for r in rows]
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

def process_sales_table(actual_rows):
    table_data = []
    for i, row in enumerate(actual_rows, 1):
        name = row['Sales_Engg_Name']
        actual = float(row['actual'] or 0)

        achievement = 0

        table_data.append({
            "si": i,
            "name": name,
            "office": row['Sales_office'],
            "net_value": round(actual, 2),
            "achievement": round(achievement, 1)
        })

    sorted_data = sorted(table_data, key=lambda x: x['achievement'], reverse=True)

    for index, item in enumerate(sorted_data, 1):
        item['si'] = index

    return sorted_data


@unit_app.route('/api/un6/order_ytd/sales_eng_table')
def get_un6_order_ytd_sales_eng():
    year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    query_actuals = """
        SELECT Sales_Engg_Name, Sales_office, SUM(Net_Value) AS actual
        FROM order_data
        WHERE fin_year = %s AND Product1  = 'DG Sets'
        GROUP BY Sales_Engg_Name
        ORDER BY actual DESC   
    """
    cursor.execute(query_actuals, (year,))
    actual_rows = cursor.fetchall()

    return jsonify(process_sales_table(actual_rows))


@unit_app.route('/un6/order_mtd')
def un6_order_mtd():
    conn = get_db_connection()
    years = []
    if conn:
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT fin_year FROM order_data ORDER BY fin_year DESC")
        years = [row['fin_year'] for row in cursor.fetchall()]
        conn.close()
    selected_year = request.args.get('fin_year', years[0] if years else "")
    return render_template('Unit6/order_mtd.html', years=years, selected_year=selected_year,
                           username=session['username'])


@unit_app.route('/api/un6/orders/kpi')
def get_un6_orders_kpi():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Current Actuals for PH1
    actual_sql = "SELECT SUM(Net_Value)/100000 AS total_val FROM order_data WHERE fin_year=%s AND Product1  = 'DG Sets'"
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        actual_sql += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    cursor.execute(actual_sql, tuple(params))
    res = cursor.fetchone()
    actual_val = res['total_val'] if res and res['total_val'] else 0

    # 2. Target for PH1 (Pro-rata)
    target_sql = "SELECT SUM(`Target`) AS total_target FROM order_target WHERE fin_year=%s AND Product1  = 'DG Sets'"
    cursor.execute(target_sql, (fin_year,))
    t_res = cursor.fetchone()
    annual_target = t_res['total_target'] if t_res and t_res['total_target'] else 0

    num_months = len(months) if months else 12
    pro_rata_target = (annual_target / 12) * num_months

    # 3. Last Month Performance for PH1
    last_month_actual = 0
    if months:
        try:
            current_month_name = months[0]
            current_year_start = int(fin_year.split('-')[0])
            month_number = datetime.strptime(current_month_name, "%B").month

            # Determine previous month and its financial year
            if month_number == 4:  # If April, go to March of previous FY
                last_month_num, lm_y_start = 3, current_year_start - 1
            else:
                last_month_num, lm_y_start = (12 if month_number == 1 else month_number - 1), current_year_start

            last_month_name = datetime(2000, last_month_num, 1).strftime('%B')
            last_month_fin_year = f"{lm_y_start}-{lm_y_start + 1}"

            lm_sql = """
                SELECT SUM(`Net_Value`)/100000 AS total 
                FROM order_data 
                WHERE fin_year=%s AND Product1  = 'DG Sets' AND MONTHNAME(`Created_Date`) = %s
            """
            cursor.execute(lm_sql, (last_month_fin_year, last_month_name))
            lm_res = cursor.fetchone()
            last_month_actual = lm_res['total'] if lm_res and lm_res['total'] else 0
        except Exception as e:
            print(f"Error: {e}")

    conn.close()
    achievement_pct = (actual_val / pro_rata_target * 100) if pro_rata_target > 0 else 0

    return jsonify({
        "actual_value": f"{float(actual_val):,.2f}L",
        "target_value": f"{float(pro_rata_target):,.2f}L",
        "achievement": f"{achievement_pct:.1f}%",
        "last_year": f"{float(last_month_actual):,.2f}L"
    })


@unit_app.route('/api/un6/orders/product_data')
def get_un6_orders_product_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    ph3_abbr = [('DG Sets','DGS')]

    # Actuals
    actual_query = """
         SELECT Product1, SUM(`Net_Value`)/100000 AS actual 
         FROM order_data 
         WHERE fin_year = %s AND Product1 = 'DG Sets'
    """
    params = [fin_year,]
    if months:
        placeholders = ', '.join(['%s'] * len(months))
        actual_query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    actual_query += " GROUP BY Product1"
    cursor.execute(actual_query, tuple(params))
    actual_dict = {r['Product1']: float(r['actual'] or 0) for r in cursor.fetchall()}

    # Targets
    target_query = "SELECT Product1, SUM(Target) as annual_target FROM order_target WHERE fin_year = %s AND Product1 = 'DG Sets' GROUP BY Product1"
    cursor.execute(target_query, (fin_year,))
    month_factor = len(months) / 12.0 if months else 1.0
    target_dict = {r['Product1']: float(r['annual_target'] or 0) * month_factor for r in cursor.fetchall()}

    conn.close()

    return jsonify({
        "labels": [p[1] for p in ph3_abbr],
        "actuals": [actual_dict.get(p[0], 0) for p in ph3_abbr],
        "targets": [target_dict.get(p[0], 0) for p in ph3_abbr]
    })


@unit_app.route('/api/un6/orders/unit_data')
def get_un6_orders_unit_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Sales_office AS display_unit, SUM(Net_Value)/100000 as actual 
        FROM order_data
        WHERE Product1 = 'DG Sets' AND fin_year = %s 
    """
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    query += " GROUP BY display_unit"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['actual'] or 0) for r in rows]
    })



@unit_app.route('/api/un6/orders/region_data')
def get_un6_orders_region_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    conn = get_db_connection()
    cursor = conn.cursor()

    # Actuals
    query = "SELECT Sales_Region as region, SUM(Net_Value)/100000 as actual FROM order_data WHERE fin_year = %s AND Product1  = 'DG Sets' "
    params = [fin_year]
    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)
    query += " GROUP BY Sales_Region"
    cursor.execute(query, tuple(params))
    actual_dict = {r['region']: float(r['actual'] or 0) for r in cursor.fetchall()}

    # Targets
    target_query = "SELECT Sales_Region AS region, SUM(Target) AS target FROM order_target WHERE Sales_Region IN %s AND fin_year = %s AND Product1  = 'DG Sets' GROUP BY Sales_Region"
    cursor.execute(target_query, (valid_regions, fin_year,))
    month_factor = len(months) / 12.0 if months else 1.0
    target_dict = {r['region']: float(r['target'] or 0) * month_factor for r in cursor.fetchall()}

    conn.close()

    return jsonify({
        "labels": valid_regions,
        "actuals": [actual_dict.get(r, 0) for r in valid_regions],
        "targets": [target_dict.get(r, 0) for r in valid_regions]
    })


@unit_app.route('/api/un6/orders/sector_data')
def get_un6_orders_sector_data():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT Dist_Channel AS channel, SUM(Net_Value)/100000 as actual FROM order_data WHERE fin_year = %s AND Product1  = 'DG Sets' "
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    query += " GROUP BY Dist_Channel"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['actual'] or 0)

    return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})


@unit_app.route('/api/un6/orders/name')
def get_un6_order_name():
    months = request.args.getlist('month')
    fin_year = request.args.get('fin_year')

    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT Name, SUM(Net_Value)/100000 AS total_value FROM order_data WHERE fin_year = %s AND Product1  = 'DG Sets' "
    params = [fin_year]

    if months:
        placeholders = ', '.join(['%s'] * len(months))
        query += f" AND MONTHNAME(Created_Date) IN ({placeholders})"
        params.extend(months)

    query += " GROUP BY Name ORDER BY total_value DESC LIMIT 10"
    cursor.execute(query, tuple(params))
    rows = cursor.fetchall()
    conn.close()

    processed_labels = []
    for r in rows:
        name = r['Name'].strip() if r['Name'] else "Unknown"
        if "KIRLOSKAR ELECTRIC COMPANY" in name.upper():
            name = "KEC, Ajman"
        processed_labels.append(name)

    return jsonify({"labels": processed_labels, "values": [float(r['total_value'] or 0) for r in rows]})

@unit_app.route('/api/un6/order_mtd/sales_eng_table')
def get_un6_order_mtd_sales_eng():
    year = request.args.get('fin_year')
    months = request.args.getlist('month')

    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. MTD Actuals
    query_actuals = """
        SELECT Sales_Engg_Name, Sales_office, SUM(Net_Value) AS actual
        FROM order_data
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTHNAME(Created_Date) IN %s
        GROUP BY Sales_Engg_Name
        ORDER BY actual DESC
    """

    placeholders = ', '.join(['%s'] * len(months))
    query_actuals = query_actuals % ('%s', f"({placeholders})")

    cursor.execute(query_actuals, (year,  *months))
    actual_rows = cursor.fetchall()
    conn.close()
    month_factor = len(months) / 12.0 if months else 1.0
    return jsonify(process_sales_table(actual_rows))


@unit_app.route('/un6/quarterly_order')
def qtd_un6_order_dashboard():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT DISTINCT fin_year FROM order_data WHERE fin_year IS NOT NULL ORDER BY fin_year DESC")
    available_years = [row['fin_year'] for row in cursor.fetchall()]
    default_year = available_years[0] if available_years else "2025-2026"
    conn.close()

    return render_template(
        'Unit6/quarterly_orders.html',
        username=session.get('username', 'User'),
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


# --- 1. PH1 QUARTERLY KPI ---
@unit_app.route('/api/un6/order/quarterly_kpis')
def get_un6_order_quarterly_kpis():
    year = request.args.get('fin_year')
    quarters_list = request.args.getlist('quarters')
    months = get_quarter_months(quarters_list)

    if not year or not months:
        return jsonify({"actuals": "0L", "targets": "0L", "achievement": "0%", "raw_percent": 0})

    conn = get_db_connection()
    cursor = conn.cursor()

    m_placeholders = ', '.join(['%s'] * len(months))

    cursor.execute(f"""
        SELECT SUM(Net_Value)/100000 AS total 
        FROM order_data 
        WHERE fin_year=%s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
    """, (year,  *months))
    actual = cursor.fetchone()['total'] or 0

    # Pro-rated Target
    cursor.execute("SELECT SUM(Target) AS annual FROM order_target WHERE fin_year=%s AND Product1  = 'DG Sets' ",
                   (year,))
    annual_target = cursor.fetchone()['annual'] or 0
    target = (annual_target / 4) * len(quarters_list)

    # Last Year Comparison
    try:
        prev_y_start = int(year.split('-')[0]) - 1
        prev_year_str = f"{prev_y_start}-{prev_y_start + 1}"
        cursor.execute(f"""
            SELECT SUM(Net_Value)/100000 AS total 
            FROM order_data 
            WHERE fin_year=%s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
        """, (prev_year_str,  *months))
        last_year_actual = cursor.fetchone()['total'] or 0
    except:
        last_year_actual = 0

    conn.close()
    achievement_pct = (actual / target * 100) if target > 0 else 0

    return jsonify({
        "actuals": f"{float(actual):,.2f}L",
        "last_year": f"{float(last_year_actual):,.2f}L",
        "targets": f"{float(target):,.0f}L",
        "achievement": f"{achievement_pct:.1f}%",
        "raw_percent": achievement_pct
    })


@unit_app.route('/api/un6/order/quarterly_products')
def get_un6_order_quarterly_products():
    year = request.args.get('fin_year')
    quarters_list = request.args.getlist('quarters')
    months = get_quarter_months(quarters_list)

    m_placeholders = ', '.join(['%s'] * len(months))
    product_abbr = [('DG Sets','DGS')]

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT Product1, SUM(Net_Value)/100000 AS actual 
        FROM order_data 
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
        GROUP BY Product1
    """, (year,  *months))
    actual_dict = {r['Product1']: float(r['actual'] or 0) for r in cursor.fetchall()}

    cursor.execute(
        "SELECT Product1, SUM(Target) as annual FROM order_target WHERE fin_year=%s AND Product1  = 'DG Sets' GROUP BY Product1",
        (year,))
    target_dict = {r['Product1']: (float(r['annual'] or 0) / 4) * len(quarters_list) for r in cursor.fetchall()}

    conn.close()

    return jsonify({
        "labels": [p[1] for p in product_abbr],
        "actuals": [actual_dict.get(p[0], 0) for p in product_abbr],
        "targets": [target_dict.get(p[0], 0) for p in product_abbr]
    })


# --- 3. PH1 QUARTERLY REGIONS ---
@unit_app.route('/api/un6/order/quarterly_regions')
def get_un6_order_quarterly_regions():
    year = request.args.get('fin_year')
    quarters_list = request.args.getlist('quarters')
    months = get_quarter_months(quarters_list)
    m_placeholders = ', '.join(['%s'] * len(months))
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute(f"""
        SELECT Sales_Region AS region, SUM(Net_Value)/100000 AS actual
        FROM order_data
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
        GROUP BY Sales_Region
    """, (year,  *months))
    actual_dict = {r['region']: float(r['actual'] or 0) for r in cursor.fetchall()}

    cursor.execute(
        "SELECT Sales_Region AS region, SUM(Target) AS annual FROM order_target WHERE fin_year=%s AND Product1  = 'DG Sets' GROUP BY Sales_Region",
        (year, ))
    target_dict = {r['region']: (float(r['annual'] or 0) / 4) * len(quarters_list) for r in cursor.fetchall()}

    conn.close()
    return jsonify({
        "labels": valid_regions,
        "actuals": [actual_dict.get(r, 0) for r in valid_regions],
        "targets": [target_dict.get(r, 0) for r in valid_regions]
    })


# --- 4. PH1 QUARTERLY SECTOR ---
@unit_app.route('/api/un6/order/quarterly_sector')
def get_un6_order_quarterly_sector():
    year = request.args.get('fin_year')
    months = get_quarter_months(request.args.getlist('quarters'))
    m_placeholders = ', '.join(['%s'] * len(months))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT Dist_Channel AS channel, SUM(Net_Value)/100000 as actual 
        FROM order_data 
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
        GROUP BY Dist_Channel
    """, (year,  *months))
    rows = cursor.fetchall()
    conn.close()

    processed_data = {}
    for r in rows:
        label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
        processed_data[label] = processed_data.get(label, 0) + float(r['actual'] or 0)

    return jsonify({"labels": list(processed_data.keys()), "values": list(processed_data.values())})


# --- 5. PH1 QUARTERLY UNIT ---
@unit_app.route('/api/un6/order/quarterly_unit')
def get_un6_order_quarterly_unit():
    year = request.args.get('fin_year')
    months = get_quarter_months(request.args.getlist('quarters'))
    m_placeholders = ', '.join(['%s'] * len(months))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT Sales_office, SUM(Net_Value)/100000 as total 
        FROM order_data
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders}) 
        GROUP BY Sales_office
    """, (year,  *months))
    rows = cursor.fetchall()
    conn.close()

    return jsonify({
        "labels": [r['Sales_office'] for r in rows],
        "values": [float(r['total'] or 0) for r in rows]
    })



# --- 6. PH1 QUARTERLY TOP CUSTOMERS ---
@unit_app.route('/api/un6/order/quarterly_customers')
def get_un6_order_quarterly_customers_swgr():
    year = request.args.get('fin_year')
    months = get_quarter_months(request.args.getlist('quarters'))
    m_placeholders = ', '.join(['%s'] * len(months))

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(f"""
        SELECT Name, SUM(Net_Value)/100000 AS total_value
        FROM order_data
        WHERE fin_year = %s AND Product1  = 'DG Sets' AND MONTH(Created_Date) IN ({m_placeholders})
        GROUP BY Name ORDER BY total_value DESC LIMIT 10
    """, (year,  *months,))
    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = r['Name'].strip().upper() if r['Name'] else "UNKNOWN"
        if "KIRLOSKAR ELECTRIC" in name: name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['total_value'] or 0))

    return jsonify({"labels": labels, "values": values})



@unit_app.route('/api/un6/orders/quarterly_sales_eng')
def get_un6_quarterly_sales_eng():
    year = request.args.get('fin_year')
    months = get_quarter_months(request.args.getlist('quarters'))
    conn = get_db_connection()
    cursor = conn.cursor()

    try:

        m_placeholders = ', '.join(['%s'] * len(months))

        query_actuals = f"""
            SELECT Sales_Engg_Name, Sales_office, SUM(Net_Value) AS actual
            FROM order_data
            WHERE fin_year = %s 
             AND Product1  = 'DG Sets'
              AND MONTH(Created_Date) IN ({m_placeholders})
            GROUP BY Sales_Engg_Name
            ORDER BY actual DESC
        """

        cursor.execute(query_actuals, (year, *months))
        actual_rows = cursor.fetchall()
        return jsonify(process_sales_table(actual_rows))
    finally:
        conn.close()


# ----DAILY PH1 ORDERS-----

@unit_app.route('/un6/order_daily')
def un6_order_daily():
    today_date = datetime.now().strftime('%Y-%m-%d')
    selected_date = request.args.get('date', today_date)
    return render_template('Unit6/daily_orders.html', selected_date=selected_date, username=session['username'])


# --- 1. DAILY SUMMARY KPI ---
@unit_app.route('/api/un6/daily/order_summary')
def get_un6_daily_order_summary():
    selected_date = request.args.get('date')
    if not selected_date:
        return jsonify({"error": "No date selected"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Today's Actual for PH1
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS actual 
            FROM order_data 
            WHERE Created_Date = %s AND Product1  = 'DG Sets'
        """, (selected_date,))
        res = cursor.fetchone()
        today_actual = float(res['actual'] or 0)

        # PH1 Daily Target (PH1 Annual / 365)
        dt = datetime.strptime(selected_date, '%Y-%m-%d')
        fin_year = f"{dt.year}-{dt.year + 1}" if dt.month >= 4 else f"{dt.year - 1}-{dt.year}"

        cursor.execute("""
            SELECT SUM(Target) AS annual_target 
            FROM order_target 
            WHERE fin_year=%s AND Product1  = 'DG Sets'
        """, (fin_year,))
        t_res = cursor.fetchone()
        annual_target = float(t_res['annual_target'] or 0)
        today_target = annual_target / 365

        # MTD Actual for PH1
        cursor.execute("""
            SELECT SUM(Net_Value)/100000 AS mtd 
            FROM order_data 
            WHERE fin_year = %s 
            AND MONTH(Created_Date) = %s 
            AND Created_Date <= %s
            AND Product1  = 'DG Sets'
        """, (fin_year, dt.month, selected_date,))
        mtd_actual = float(cursor.fetchone()['mtd'] or 0)

        achievement = (today_actual / today_target * 100) if today_target > 0 else 0

        return jsonify({
            "today_actual": today_actual,
            "today_target": today_target,
            "achievement": achievement,
            "mtd_actual": f"{mtd_actual:,.2f}"
        })
    finally:
        conn.close()


# --- 2. DAILY PRODUCT SPLIT ---
@unit_app.route('/api/un6/daily/order_products')
def get_un6_daily_order_products():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    ph2_abbr = [ ('DG Sets','DGS')]

    # Actuals for the day
    cursor.execute("""
        SELECT Product1, SUM(Net_Value)/100000 AS actual 
        FROM order_data 
        WHERE Created_Date = %s AND Product1  = 'DG Sets'
        GROUP BY Product1
    """, (selected_date,))
    actual_dict = {r['Product1']: float(r['actual'] or 0) for r in cursor.fetchall()}

    # Daily Target
    dt = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{dt.year}-{dt.year + 1}" if dt.month >= 4 else f"{dt.year - 1}-{dt.year}"

    cursor.execute("""
        SELECT Product1, SUM(Target)/365 as daily_target 
        FROM order_target 
        WHERE fin_year = %s AND Product1  = 'DG Sets'
        GROUP BY Product1
    """, (fin_year,))
    target_dict = {r['Product1']: float(r['daily_target'] or 0) for r in cursor.fetchall()}

    conn.close()
    return jsonify({
        "labels": [p[1] for p in ph2_abbr],
        "actuals": [actual_dict.get(p[0], 0) for p in ph2_abbr],
        "targets": [target_dict.get(p[0], 0) for p in ph2_abbr]
    })


# --- 3. DAILY REGION PERFORMANCE ---
@unit_app.route('/api/un6/daily/order_regions')
def get_un6_daily_order_regions():
    selected_date = request.args.get('date')
    valid_regions = ('Central India', 'East India', 'North India', 'South India', 'West India')

    conn = get_db_connection()
    cursor = conn.cursor()

    # Actuals
    cursor.execute("""
        SELECT Sales_Region as region, SUM(Net_Value)/100000 as actual 
        FROM order_data 
        WHERE Created_Date = %s AND Product1  = 'DG Sets'
        GROUP BY Sales_Region
    """, (selected_date,))
    actual_dict = {r['region']: float(r['actual'] or 0) for r in cursor.fetchall()}

    # Daily Target
    dt = datetime.strptime(selected_date, '%Y-%m-%d')
    fin_year = f"{dt.year}-{dt.year + 1}" if dt.month >= 4 else f"{dt.year - 1}-{dt.year}"

    cursor.execute("""
        SELECT Sales_Region AS region, SUM(Target)/365 AS daily_target
        FROM order_target 
        WHERE Sales_Region IN %s AND fin_year = %s AND Product1  = 'DG Sets'
        GROUP BY Sales_Region
    """, (valid_regions, fin_year,))
    target_dict = {r['region']: float(r['daily_target'] or 0) for r in cursor.fetchall()}

    conn.close()
    return jsonify({
        "labels": valid_regions,
        "actuals": [actual_dict.get(r, 0) for r in valid_regions],
        "targets": [target_dict.get(r, 0) for r in valid_regions]
    })


# --- 4. DAILY TOP 10 CUSTOMERS ---
@unit_app.route('/api/un6/daily/order_customers')
def get_un6_daily_order_customers():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Name, SUM(Net_Value)/100000 AS total_value
        FROM order_data
        WHERE Created_Date = %s AND Product1  = 'DG Sets'
        GROUP BY Name ORDER BY total_value DESC LIMIT 10
    """, (selected_date,))
    rows = cursor.fetchall()
    conn.close()

    labels, values = [], []
    for r in rows:
        name = (r['Name'].strip() if r['Name'] else "Unknown")
        if "KIRLOSKAR ELECTRIC" in name.upper(): name = "KEC, Ajman"
        labels.append(name)
        values.append(float(r['total_value'] or 0))

    return jsonify({"labels": labels, "values": values})



# --- DAILY UNIT WISE ORDERS (PH1) ---
@unit_app.route('/api/un6/daily/order_units')
def get_un6_daily_order_units():
    selected_date = request.args.get('date')
    if not selected_date:
        return jsonify({"error": "No date provided"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
           SELECT 
               Sales_office AS display_unit, 
               SUM(Net_Value)/100000 as actual 
           FROM order_data
           WHERE Product1  = 'DG Sets'
             AND Created_Date = %s 

           GROUP BY display_unit
       """

    try:
        cursor.execute(query, (selected_date,))
        rows = cursor.fetchall()

        return jsonify({
            "labels": [r['display_unit'] for r in rows],
            "values": [float(r['actual'] or 0) for r in rows]
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

# --- DAILY SECTOR WISE ORDERS (PH1) ---
@unit_app.route('/api/un6/daily/order_sectors')
def get_un6_daily_order_sectors():
    selected_date = request.args.get('date')
    if not selected_date:
        return jsonify({"error": "No date provided"}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT Dist_Channel AS channel, 
               SUM(Net_Value)/100000 as actual 
        FROM order_data 
        WHERE Created_Date = %s 
          AND Product1  = 'DG Sets'
        GROUP BY Dist_Channel
    """

    try:
        cursor.execute(query, (selected_date, ))
        rows = cursor.fetchall()

        processed_data = {}
        for r in rows:
            label = 'DD' if r['channel'] in ['SZ', 'DD'] else r['channel']
            processed_data[label] = processed_data.get(label, 0) + float(r['actual'] or 0)

        return jsonify({
            "labels": list(processed_data.keys()),
            "values": list(processed_data.values())
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 5. DAILY DETAILED TABLE ---
@unit_app.route('/api/un6/daily/details')
def get_un6_daily_order_details():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
       SELECT Product2, Order_Number, Sales_office, Name, SUM(Order_Qty) as Order_Qty, SUM(Net_Value) as Net_Value 
       FROM order_data 
       WHERE Created_Date = %s AND Product1  = 'DG Sets'
       GROUP BY Order_Number, Product2, Sales_office, Name;
    """
    cursor.execute(query, (selected_date,))
    results = cursor.fetchall()
    conn.close()
    return jsonify(results)


@unit_app.route('/api/un6/order_daily/sales_eng_table')
def get_un6_order_daily_sales_eng():
    selected_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    query_actuals = """
        SELECT Sales_Engg_Name, Sales_office, SUM(Net_Value) AS actual
        FROM order_data
        WHERE Created_Date = %s  AND Product1  = 'DG Sets'
        GROUP BY Sales_Engg_Name
        ORDER BY actual DESC
    """
    cursor.execute(query_actuals, (selected_date,))
    actual_rows = cursor.fetchall()

    return jsonify(process_sales_table(actual_rows))


# ---PENDING_ORDERS(PH1)---

@unit_app.route('/un6/pending_orders')
def un6_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit6/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@unit_app.route('/api/un6/pending/total')
def get_un6_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2`  = 'DG Sets'
    """, (order_date, ))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un6/pending/products')
def get_un6_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('DG Sets','DGS')
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
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un6/pending/branches')
def get_un6_pending_branches():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2`  = 'DG Sets'
          AND Branch_Name != '' 
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
@unit_app.route('/api/un6/pending/sectors')
def get_un6_pending_sectors():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2`  = 'DG Sets'
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
@unit_app.route('/api/un6/pending/units')
def get_un6_pending_units():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product2  = 'DG Sets' AND `As_on_Date` = %s 
        GROUP BY Unit
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 6. PH1 PENDING TOP 10 CUSTOMERS ---
@unit_app.route('/api/un6/pending/customers')
def get_un6_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` = 'DG Sets'
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


# ---TRF(M)&SWGR----

@unit_app.route('/unit_16/dashboard')
def unit_16_dashboard():
    return render_template('Unit16/ytd.html', username=session['username'])

def build_filter_query_unit16(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 16 dashboard.
    Optimized for MySQL using Python mapping and hardcoded TRANSFORMERS PUNE restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    if table_type == "data":
        filters.append("Prod = 'TRANSFORMERS PUNE'")
    else:
        filters.append("Product = 'TRANSFORMERS PUNE'")
    # ----------------------------

    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
    units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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


@unit_app.route('/api/filters/unit16', methods=['GET'])
def get_filters_unit16():
    """Fetches distinct filter options, restricted strictly to TRANSFORMERS PUNE."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Prod = 'TRANSFORMERS PUNE'"

    with get_db_connection() as conn:
        with conn.cursor(dictionary=True) as cursor: # Added dictionary=True to prevent .get() TypeErrors
            # 1. Regions
            cursor.execute(f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units
            # Since the data is bound to Transformers Pune, the mapped unit is always UN16
            units = ['UN16']

            # 4. Dependent Office Filter (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM billing_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 5. Products
            # We don't need to query the database for this since it's statically bound to TRANSFORMERS PUNE
            products = ['TRANSFORMERS PUNE']

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })

@unit_app.route('/api/kpis/unit16', methods=['GET'])
def get_kpis_unit16():
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_unit16("data")
    target_where, target_params = build_filter_query_unit16("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
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


@unit_app.route('/api/sales-trend/unit16', methods=['GET'])
def get_sales_trend_unit16():
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_unit16("data")

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

@unit_app.route('/api/products-revenue/unit16', methods=['GET'])
def get_products_revenue_unit16():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_unit16("data")

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


@unit_app.route('/api/sales-by-office/unit16', methods=['GET'])
def get_sales_by_office_unit16():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_unit16("data")

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


@unit_app.route('/api/sales-by-region/unit16', methods=['GET'])
def get_sales_by_region_unit16():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_unit16("data")

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



@unit_app.route('/api/aop-achievement/unit16', methods=['GET'])
def get_aop_achievement_unit16():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_unit16("data")
    target_where, target_params = build_filter_query_unit16("target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.args.getlist('month')
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

@unit_app.route('/api/export/sales_excel/unit16', methods=['GET'])
def export_sales_excel_unit16():
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_unit16("data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500

@unit_app.route('/order_dashboard/Unit16')
def order_dashboard_unit16():
    return render_template('Unit16/order_ytd.html', username=session['username'])


def build_filter_orders_query_unit16(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 16 Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded Transformer Pune restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Transformer Pune for all queries on this dashboard
    filters.append("Product1 = 'Transformer Pune'")
    # ----------------------------

    order_regions = request.args.getlist('region')
    order_offices = request.args.getlist('sales_office')
    order_products = request.args.getlist('product')
    order_units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    where_clause = " AND ".join(filters) if filters else "1=1"
    return "WHERE " + where_clause, params


@unit_app.route('/api/filters/orders/unit16', methods=['GET'])
def get_orders_filters_unit16():
    """Fetches distinct filter options for Orders, restricted strictly to Transformer Pune."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 = 'Transformer Pune'"

    with get_db_connection() as conn:
        with conn.cursor(dictionary=True) as cursor:
            # 1. Independent Filters (Regions, Fin Years)
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL"
            )
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(
                f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL"
            )
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Units
            # Since data is locked to Transformer Pune, mapped unit is UN16
            units = ['UN16']

            # 3. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Products
            # Statically bound to Transformer Pune
            products = ['Transformer Pune']

    return jsonify({
        "regions": regions,
        "offices": offices,
        "fin_years": fin_years,
        "products": products,
        "units": units
    })
@unit_app.route('/api/order/kpis/unit16', methods=['GET'])
def get_orders_kpis_unit16():
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_unit16("order_data")
    target_where, target_params = build_filter_orders_query_unit16("order_target")

    months = request.args.getlist('month')
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


@unit_app.route('/api/orders-trend/unit16', methods=['GET'])
def get_orders_trend_unit16():
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_unit16("order_data")

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


@unit_app.route('/api/orders/products-revenue/unit16', methods=['GET'])
def get_orders_products_revenue_unit16():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit16("order_data")

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


@unit_app.route('/api/orders/sales-by-office/unit16', methods=['GET'])
def get_orders_sales_by_office_unit16():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit16("order_data")

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


@unit_app.route('/api/order/sales-by-region/unit16', methods=['GET'])
def get_order_sales_by_region_unit16():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit16("order_data")

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


@unit_app.route('/api/orders/sales-engineers/unit16', methods=['GET'])
def get_orders_sales_engineers_unit16():
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_unit16("order_data")

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


@unit_app.route('/api/orders/aop-achievement/unit16', methods=['GET'])
def get_orders_aop_achievement_unit16():
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_unit16("order_data")
    target_where, target_params = build_filter_orders_query_unit16("order_target")

    months = request.args.getlist('month')
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

@unit_app.route('/api/export/orders_excel/unit16', methods=['GET'])
def export_orders_excel_unit16():
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_unit16("order_data")

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
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500

# ---PENDING_ORDERS(PH1)---
@unit_app.route('/un16/pending_orders')
def un16_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit16/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@unit_app.route('/api/un16/pending/total')
def get_un16_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2`  = 'Transformer Pune'
    """, (order_date, ))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un16/pending/products')
def get_un16_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('Transformer Pune','TRF(P)')
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
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un16/pending/branches')
def get_un16_pending_branches():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2`  = 'Transformer Pune'
          AND Branch_Name != '' 
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
@unit_app.route('/api/un16/pending/sectors')
def get_un16_pending_sectors():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2`  = 'Transformer Pune'
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
@unit_app.route('/api/un16/pending/units')
def get_un16_pending_units():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT 
               CASE 
                   WHEN Product2 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END AS display_unit, 
               SUM(Net_Value)/100000 as value 
           FROM pending_order
           WHERE Product2 IN ('Transformer Pune') AND `As_on_Date` = %s AND Unit != ''
           GROUP BY 
               CASE 
                   WHEN Product2 = 'Transformer Pune' THEN 'UN16'
                   ELSE Unit 
               END
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['display_unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 16. PH1 PENDING TOP 10 CUSTOMERS ---
@unit_app.route('/api/un16/pending/customers')
def get_un16_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` = 'Transformer Pune'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, ))

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

@unit_app.route('/unit_20/dashboard')
def unit_20_dashboard():
    return render_template('Unit20/ytd.html', username=session['username'])

def build_filter_query_unit20(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 1 dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN16 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit = 'UN20'")
    # ----------------------------

    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
    units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/unit20', methods=['GET'])
def get_filters_unit20():
    """Fetches distinct filter options, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Dist_Channel_TEXT != 'Inter Unit Transfer' AND Unit = 'UN20'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Regions
            cursor.execute(f"SELECT DISTINCT Sales_Region_Name FROM billing_data WHERE {base_where} AND Sales_Region_Name IS NOT NULL")
            regions = [r.get('Sales_Region_Name') for r in cursor.fetchall() if r.get('Sales_Region_Name')]

            # 2. Fin Years
            cursor.execute(f"SELECT DISTINCT fin_year FROM billing_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 3. Units
            units = ['UN20']

            # 4. Dependent Office Filter (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM billing_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region_Name IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

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

@unit_app.route('/api/kpis/unit20', methods=['GET'])
def get_kpis_unit20():
    """Returns Overall Sales and Total Target for the KPI grid and Achievement Gauge."""
    data_where, data_params = build_filter_query_unit20("data")
    target_where, target_params = build_filter_query_unit20("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
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


@unit_app.route('/api/sales-trend/unit20', methods=['GET'])
def get_sales_trend_unit20():
    """Returns Net Revenue Trend over time (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_query_unit20("data")

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

@unit_app.route('/api/products-revenue/unit20', methods=['GET'])
def get_products_revenue_unit20():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query_unit20("data")

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


@unit_app.route('/api/sales-by-office/unit20', methods=['GET'])
def get_sales_by_office_unit20():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query_unit20("data")

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


@unit_app.route('/api/sales-by-region/unit20', methods=['GET'])
def get_sales_by_region_unit20():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query_unit20("data")

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



@unit_app.route('/api/aop-achievement/unit20', methods=['GET'])
def get_aop_achievement_unit20():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query_unit20("data")
    target_where, target_params = build_filter_query_unit20("target")

    # Determine proration factor based on selected months
    # If 1 month is selected, target is multiplied by (1/12). If none, it remains (1).
    months = request.args.getlist('month')
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

@unit_app.route('/api/export/sales_excel/unit20', methods=['GET'])
def export_sales_excel_unit20():
    """Exports the filtered billing_data to an Excel file."""
    where_clause, params = build_filter_query_unit20("data")
    query = f"SELECT * FROM billing_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500


@unit_app.route('/order_dashboard/Unit20')
def order_dashboard_unit20():
    return render_template('Unit20/order_ytd.html', username=session['username'])

def build_filter_orders_query_unit20(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters for the Unit 1 Orders dashboard.
    Optimized for MySQL using Python mapping and hardcoded UN20 restriction.
    """
    filters = []
    params = []

    # --- BASELINE RESTRICTION ---
    # Strictly enforce Unit 1 for all queries on this dashboard
    filters.append("Unit = 'UN20'")
    # ----------------------------

    order_regions = request.args.getlist('region')
    order_offices = request.args.getlist('sales_office')
    order_products = request.args.getlist('product')
    order_units = request.args.getlist('unit')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

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

    # --- PYTHON MAPPING FOR PRODUCTS ---
    # Avoids slow CASE WHEN table scans in MySQL
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


@unit_app.route('/api/filters/orders/unit20', methods=['GET'])
def get_orders_filters_unit20():
    """Fetches distinct filter options for Orders, restricted strictly to Unit 1."""
    regions_selected = request.args.getlist('region')

    # Hardcoded baseline to ensure high-speed querying without OR conditions
    base_where = "Product1 IS NOT NULL AND Unit = 'UN20'"

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Independent Filters (Regions, Fin Years)
            cursor.execute(
                f"SELECT DISTINCT Sales_Region FROM order_data WHERE {base_where} AND Sales_Region IS NOT NULL")
            regions = [r.get('Sales_Region') for r in cursor.fetchall() if r.get('Sales_Region')]

            cursor.execute(f"SELECT DISTINCT fin_year FROM order_data WHERE {base_where} AND fin_year IS NOT NULL")
            fin_years = [r.get('fin_year') for r in cursor.fetchall() if r.get('fin_year')]

            # 2. Units
            # We don't need to query the database for this since it's statically bound to UN20
            units = ['UN20']

            # 3. Dependent Filter: Sales Office (Filtered by Region)
            office_query = f"SELECT DISTINCT Sales_office FROM order_data WHERE {base_where} AND Sales_office IS NOT NULL"
            office_params = []
            if regions_selected:
                placeholders = ','.join(['%s'] * len(regions_selected))
                office_query += f" AND Sales_Region IN ({placeholders})"
                office_params.extend(regions_selected)

            cursor.execute(office_query, office_params)
            offices = [r.get('Sales_office') for r in cursor.fetchall() if r.get('Sales_office')]

            # 4. Products (Mapped in Python to avoid SQL CASE delays)
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

@unit_app.route('/api/order/kpis/unit20', methods=['GET'])
def get_orders_kpis_unit20():
    """Returns Overall Sales, Total Target, and Customers Billed for the Orders Dashboard."""
    data_where, data_params = build_filter_orders_query_unit20("order_data")
    target_where, target_params = build_filter_orders_query_unit20("order_target")

    months = request.args.getlist('month')
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


@unit_app.route('/api/orders-trend/unit20', methods=['GET'])
def get_orders_trend_unit20():
    """Returns Net Revenue Trend over time for Orders (Both Weekly and Monthly aggregation)."""
    where_clause, params = build_filter_orders_query_unit20("order_data")

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


@unit_app.route('/api/orders/products-revenue/unit20', methods=['GET'])
def get_orders_products_revenue_unit20():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit20("order_data")

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


@unit_app.route('/api/orders/sales-by-office/unit20', methods=['GET'])
def get_orders_sales_by_office_unit20():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit20("order_data")

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


@unit_app.route('/api/order/sales-by-region/unit20', methods=['GET'])
def get_order_sales_by_region_unit20():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query_unit20("order_data")

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


@unit_app.route('/api/orders/sales-engineers/unit20', methods=['GET'])
def get_orders_sales_engineers_unit20():
    """Replaces AOP Endpoint: Returns Sales Engineer table with hardcoded 0 targets."""
    data_where, data_params = build_filter_orders_query_unit20("order_data")

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


@unit_app.route('/api/orders/aop-achievement/unit20', methods=['GET'])
def get_orders_aop_achievement_unit20():
    """Calculates Sales Achieved VS AOP based strictly on Product, ignoring units."""
    data_where, data_params = build_filter_orders_query_unit20("order_data")
    target_where, target_params = build_filter_orders_query_unit20("order_target")

    months = request.args.getlist('month')
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

@unit_app.route('/api/export/orders_excel/unit20', methods=['GET'])
def export_orders_excel_unit20():
    """Exports the filtered order_data to an Excel file."""
    # 1. Dynamically fetch the WHERE clause specifically for Orders
    where_clause, params = build_filter_orders_query_unit20("order_data")

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
            return "No data found for the selected filters.", 404

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
        fin_years = request.args.getlist('fin_year')
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
        return f"Error exporting data: {str(e)}", 500

# ---PENDING_ORDERS(PH1)---

@unit_app.route('/un20/pending_orders')
def un20_pending_orders_page():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit20/Pending_Orders.html', username=session.get('username', 'User'), order_date=today)


# --- 1. PH1 TOTAL PENDING KPI ---
@unit_app.route('/api/un20/pending/total')
def get_un20_total_pending():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"total": 0})

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT SUM(`Net_Value`)/100000 AS total 
        FROM pending_order 
        WHERE `As_on_Date` = %s AND `Product2`  = 'SPARES & SERVICE'
    """, (order_date, ))

    result = cursor.fetchone()
    conn.close()
    return jsonify({"total": float(result['total'] or 0)})


# --- 2. PH1 PENDING PRODUCT SPLIT ---
@unit_app.route('/api/un20/pending/products')
def get_un20_pending_products():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()

    product_abbr = [
         ('SPARES & SERVICE','SP&S')
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
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()


# --- 3. PH1 PENDING BRANCHES/OFFICES ---
@unit_app.route('/api/un20/pending/branches')
def get_un20_pending_branches():
    order_date = request.args.get('date')
    if not order_date:
        return jsonify({"labels": [], "values": []})

    conn = get_db_connection()
    cursor = conn.cursor()
    # PH1 specific branch performance
    cursor.execute("""
        SELECT Branch_Name AS label, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s 
          AND `Product2`  = 'SPARES & SERVICE'
          AND Branch_Name != '' 
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
@unit_app.route('/api/un20/pending/sectors')
def get_un20_pending_sectors():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT `Distr_Channel` AS channel , SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2`  = 'SPARES & SERVICE'
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
@unit_app.route('/api/un20/pending/units')
def get_un20_pending_units():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT Unit, SUM(Net_Value)/100000 as value 
        FROM pending_order
        WHERE Product2  = 'SPARES & SERVICE' AND `As_on_Date` = %s 
        GROUP BY Unit
    """, (order_date,))

    rows = cursor.fetchall()
    conn.close()
    return jsonify({
        "labels": [r['Unit'] for r in rows],
        "values": [float(r['value']) for r in rows]
    })


# --- 20. PH1 PENDING TOP 10 CUSTOMERS ---
@unit_app.route('/api/un20/pending/customers')
def get_un20_pending_customers():
    order_date = request.args.get('date')
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT Cust_Name, SUM(`Net_Value`)/100000 AS value
        FROM pending_order
        WHERE `As_on_Date` = %s AND `Product2` = 'SPARES & SERVICE'
        GROUP BY Cust_Name
        ORDER BY value DESC LIMIT 10
    """, (order_date, ))

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



engine = create_engine("mysql+pymysql://root:@localhost/dashboard")

@unit_app.route('/api/pending_order_excel_un1')
def pending_order_excel_un1():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE As_on_Date = :dt AND Unit = 'UN01' ")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500



@unit_app.route('/api/pending_order_excel_un2')
def pending_order_excel_un2():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE As_on_Date = :dt AND Unit IN ('UN02','UN06') ")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500




engine = create_engine("mysql+pymysql://root:@localhost/dashboard")
un5_BILLING_PRODS = ('TRANSFORMERS MYSORE',)
un5_ORDER_PRODS = ('Transformer Mysore', 'Switchgear')


def get_un5_condition(page_type):
    if page_type == "billing":
        return "Prod", un5_BILLING_PRODS
    else:
        return "Product1", un5_ORDER_PRODS


def format_date_columns(df):
    date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date','PO_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df


def serve_excel(df, filename):
    if df.empty:
        return "No data found for the selected period", 404

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


@unit_app.route('/api/un5/pending_order_excel')
def un5_pending_order_excel():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('Transformer Mysore') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date', 'Created_On']

        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')

        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500



un16_BILLING_PRODS = ('TRANSFORMERS PUNE',)
un16_ORDER_PRODS = ('Transformer Pune',)

engine = create_engine("mysql+pymysql://root:@localhost/dashboard")
def get_un16_condition(page_type):
    if page_type == "billing":
        return "Prod", un16_BILLING_PRODS
    else:
        return "Product1", un16_ORDER_PRODS


def format_date_columns(df):
    date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date','PO_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df


def serve_excel(df, filename):
    if df.empty:
        return "No data found for the selected period", 404

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

@unit_app.route('/api/un16/pending_order_excel')
def un16_pending_order_excel():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('Transformer Pune') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500


un6_BILLING_PRODS = ('DG SETS',)
un6_ORDER_PRODS = ('DG Sets',)

engine = create_engine("mysql+pymysql://root:@localhost/dashboard")
def get_un6_condition(page_type):
    if page_type == "billing":
        return "Prod", un6_BILLING_PRODS
    else:
        return "Product1", un6_ORDER_PRODS


def format_date_columns(df):
    date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date','PO_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df


def serve_excel(df, filename):
    if df.empty:
        return "No data found for the selected period", 404

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



@unit_app.route('/api/un6/pending_order_excel')
def un6_pending_order_excel():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('DG Sets') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date','Created_On']

        for col in date_cols:
            if col in df.columns:

                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')


        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500


un20_BILLING_PRODS = ('SPARES & SERVICE',)
un20_ORDER_PRODS = ('SPARES & SERVICE',)

engine = create_engine("mysql+pymysql://root:@localhost/dashboard")


def get_un20_condition(page_type):
    if page_type == "billing":
        return "Prod", un20_BILLING_PRODS
    else:
        return "Product1", un20_ORDER_PRODS


def format_date_columns(df):
    date_cols = ['Created_Date', 'Released_date', 'Billing_Date', 'eWay_Billdate', 'LR_Date','PO_Date']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    return df


def serve_excel(df, filename):
    if df.empty:
        return "No data found for the selected period", 404

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

@unit_app.route('/api/un20/pending_order_excel')
def un20_pending_order_excel():
    order_date = request.args.get('date')
    page_type = request.args.get('type', 'pending_orders')

    if not order_date:
        return "Please Select Date", 400
    page_type == "pending_orders"

    try:
        with engine.connect() as conn:

            query = text(f"SELECT * FROM pending_order WHERE Product1 IN('SPARES & SERVICE') AND As_on_Date = :dt")
            df = pd.read_sql(query, conn, params={"dt": order_date})

        if df.empty:
            return f"No records found for {order_date} in pending_orders", 404

        date_cols = ['As_on_Date', 'Sale_Ord_Dt', 'Delv_Date', 'FIRST_DATE', 'PO_Date', 'Created_On']

        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
                df[col] = df[col].dt.strftime('%d-%m-%Y')

        filename = f"{page_type}_{order_date}.xlsx"
        return serve_excel(df, filename)

    except Exception as e:
        print(f"SQL Error: {e}")
        return f"Database Error: {str(e)}", 500


@unit_app.route('/collection_report_un20')
def collection_report_un20():
    today = datetime.now().strftime('%Y-%m-%d')
    return render_template('Unit20/collection_report.html', username=session.get('username', 'User'),
                           selected_date=today)

FIN20_PROFIT_CENTERS = [
    'UN12SPS', 'UN20SPA', 'UN12SSD'
]


@unit_app.route('/collection_fin_report_un20')
def finance_20_dashboard():
    return render_template('Unit20/collection_report.html', username=session.get('username', 'User'),)


# ==========================================
# FILTER BUILDER (Secured)
# ==========================================

def build_collections_filter_un20():
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries."""
    filters = []
    params = []

    # 1. Enforce Static RBAC Data Boundary
    pc_placeholders = ','.join(['%s'] * len(FIN20_PROFIT_CENTERS))
    filters.append(f"Profit_Centre IN ({pc_placeholders})")
    params.extend(FIN20_PROFIT_CENTERS)

    # 2. Extract UI Filters
    years = request.args.getlist('year')
    months = request.args.getlist('month')
    exact_date = request.args.get('date')
    units = request.args.getlist('unit')
    ui_pcs = request.args.getlist('profit_center')
    sales_offices = request.args.getlist('sales_office')

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
            filters.append("DATE(Posting_Date) BETWEEN %s AND %s")
            params.extend([start_date, end_date])
        else:
            filters.append("DATE(Posting_Date) = %s")
            params.append(exact_date)

    if units:
        placeholders = ','.join(['%s'] * len(units))
        filters.append(f"Unit_Code IN ({placeholders})")
        params.extend(units)

    if ui_pcs:
        # Prevent user from requesting a PC outside their allowed boundary
        valid_pcs = [pc for pc in ui_pcs if pc in FIN20_PROFIT_CENTERS]
        if valid_pcs:
            placeholders = ','.join(['%s'] * len(valid_pcs))
            filters.append(f"Profit_Centre IN ({placeholders})")
            params.extend(valid_pcs)

    if sales_offices:
        placeholders = ','.join(['%s'] * len(sales_offices))
        filters.append(f"Coll_BR_DESC IN ({placeholders})")
        params.extend(sales_offices)

    where_clause = "WHERE " + " AND ".join(filters)
    return where_clause, params


# ==========================================
# API ENDPOINTS
# ==========================================

@unit_app.route('/api/filters/collections/un20', methods=['GET'])
def get_collections_filters_un20():
    """Fetches unique values for filters, restricted strictly to Fin1 boundaries."""

    # Establish base security condition for dropdowns
    pc_placeholders = ','.join(['%s'] * len(FIN20_PROFIT_CENTERS))
    base_where = f"WHERE Profit_Centre IN ({pc_placeholders})"
    base_params = list(FIN20_PROFIT_CENTERS)

    unit_selected = request.args.getlist('unit')

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # Independent Filters scoped to boundary
            cursor.execute(
                f"SELECT DISTINCT YEAR(Posting_Date) as yr FROM collections_data {base_where} AND Posting_Date IS NOT NULL ORDER BY yr DESC",
                base_params)
            years = [str(r['yr']) for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT MONTHNAME(Posting_Date) as mth FROM collections_data {base_where} AND Posting_Date IS NOT NULL",
                base_params)
            months = [r['mth'] for r in cursor.fetchall()]

            cursor.execute(f"SELECT DISTINCT Unit_Code FROM collections_data {base_where} AND Unit_Code IS NOT NULL",
                           base_params)
            units = [r['Unit_Code'] for r in cursor.fetchall()]

            cursor.execute(
                f"SELECT DISTINCT Coll_BR_DESC FROM collections_data {base_where} AND Coll_BR_DESC IS NOT NULL",
                base_params)
            sales_offices = [r['Coll_BR_DESC'] for r in cursor.fetchall()]

            # Dependent Filter: Profit Center (Filtered by Unit AND Boundary)
            pc_query = f"SELECT DISTINCT Profit_Centre FROM collections_data {base_where} AND Profit_Centre IS NOT NULL"
            pc_params_list = list(base_params)

            if unit_selected:
                unit_placeholders = ','.join(['%s'] * len(unit_selected))
                pc_query += f" AND Unit_Code IN ({unit_placeholders})"
                pc_params_list.extend(unit_selected)

            cursor.execute(pc_query, pc_params_list)
            profit_centers = [r['Profit_Centre'] for r in cursor.fetchall()]

    return jsonify({
        "years": years,
        "months": months,
        "units": units,
        "profit_centers": profit_centers,
        "sales_offices": sales_offices
    })


@unit_app.route('/api/collections/dashboard/un20', methods=['GET'])
def get_collections_dashboard_data_un20():
    where, params = build_collections_filter_un20()

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
            months = request.args.getlist('month')
            exact_date = request.args.get('date')

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


@unit_app.route('/api/export/collections_excel_un20', methods=['GET'])
def export_collections_excel_un20():
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_un20()
    query = f"SELECT * FROM collections_data {where_clause}"

    try:
        with get_db_connection() as conn:
            with conn.cursor() as cursor:
                cursor.execute(query, params)
                results = cursor.fetchall()

        if not results:
            return "No data found for the selected filters.", 404

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
        years = request.args.getlist('year')
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
        return f"Error exporting data: {str(e)}", 500


PROFIT_CENTER_MAP_20 = {
'Unit-20': ['UN12SPS', 'UN20SPA','UN12SSD',],
}

@unit_app.route('/unit20_mis_report')
def unit_20_mis_report():
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

    selected_year = request.args.get('fin_year', years[0] if years else "")

    return render_template(
        'Unit20/MIS.html',
        years=years,
        selected_year=selected_year,
        username=session.get('username', 'User')
    )


@unit_app.route('/api/mis/unit20/unit-monthly-sales')
def get_unit20_monthly_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
            700021, 700018, 700006, 700014, 700007, 700034, 700033, 700032,
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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/unit-scrap-sales')
def get_unit20_scrap_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/other-income-monthly-sales')
def get_unit20_other_income_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        other_income_accounts = [459001,711001, 712001,712002,713001,714001,715001, 715003,715004, 715005, 717001, 719001,720009,720006,715006,]

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/iut-out-monthly-sales')
def get_unit20_iut_out_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/rm-consp-monthly-sales')
def get_unit20_rm_consp_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/employee-monthly-sales')
def get_unit20_employee_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/iut-in-monthly-sales')
def get_unit20_iut_in_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/transport-monthly-sales')
def get_unit20_transport_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/finance-cahrges-monthly-sales')
def get_unit20_finance_charges_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        is_dict = True
    except AttributeError:
        cursor = conn.cursor()
        is_dict = False

    try:
        finance_charges_accounts = [
            446001, 446002, 446003, 446004,454006, 454009,
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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/depreciation-monthly-sales')
def get_unit20_depreciation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/revenue-expenses-monthly-sales')
def get_unit20_revenue_expenses_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
            456104, 456105, 456106, 459002, 459003, 460001, 461001, 430203
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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/ere-allocation-monthly-sales')
def get_unit20_ere_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/ore-allocation-monthly-sales')
def get_unit20_ore_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

@unit_app.route('/api/mis/unit20/intrest-allocation-monthly-sales')
def get_unit20_intrest_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()


@unit_app.route('/api/mis/unit20/dep-allocation-monthly-sales')
def get_unit20_dep_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_20.get(selected_unit, [])
    if not profit_centers:
        return jsonify({"error": f"Selected Unit allocation matching code not found: {selected_unit}"}), 400

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
        return jsonify({"error": "Internal Processing Error"}), 500
    finally:
        cursor.close()
        conn.close()

