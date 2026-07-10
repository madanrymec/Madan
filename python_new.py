import os
import traceback
import io
from flask import Flask, render_template, request, redirect, url_for, session, jsonify, flash,send_file,Blueprint, request
import  pymysql
from datetime import datetime

DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'dashboard',
}
app = Flask(__name__)
app.secret_key = 'Kec@12345'

@app.route('/test')
def test():
    return "TEST OK"

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

def requires_login(f):
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)

    decorated_function.__name__ = f.__name__
    return decorated_function


@app.route('/', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        emp_id_input = request.form['username']
        password_input = request.form['password']

        conn = get_db_connection()
        cursor = conn.cursor()

        # Query remains the same, fetching the 'permission' column
        query = """
            SELECT emp_id, emp_name, permission 
            FROM user_mast 
            WHERE emp_id = %s AND BINARY password = %s AND active_status = 'active'
        """
        cursor.execute(query, (emp_id_input, password_input))
        user = cursor.fetchone()
        conn.close()

        if user:
            session['logged_in'] = True
            session['emp_id'] = user['emp_id'].strip()
            session['username'] = user['emp_name']

            perm = user['permission']


            if not perm or perm == 'NULL':
                return redirect(url_for('menu'))

            if perm.startswith('unit_'):
                target_route = f"unit_head.{perm}"
            elif perm.startswith('finance_'):
                target_route = f"finance_head.{perm}"
            else:
                target_route = f"product_head.{perm}"

            try:
                return redirect(url_for(target_route))
            except Exception as e:
                print(f"Routing Error for {target_route}: {e}")
                return redirect(url_for('menu'))
        else:
            error = "Invalid Employee ID or Password"

    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('username', None)
    return redirect(url_for('login'))



def build_filter_query(table_type="data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Optimized for MySQL using %s placeholders.
    """
    filters = []
    params = []

    # Extract multiple selections
    regions = request.args.getlist('region')
    offices = request.args.getlist('sales_office')
    products = request.args.getlist('product')
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

    # Product handling (Handles the LV MOTORS logic internally for billing_data)
    if products:
        placeholders = ','.join(['%s'] * len(products))
        if table_type == "data":
            # Map products dynamically using the Plant logic requested
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

    # Financial Year Handling (Now directly using the fin_year column for both tables)
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

@app.route('/menu')
def menu():
    return render_template('Menu.html')

@app.route('/home')
def home():
    # Assuming login logic is handled elsewhere, defaulting to a static render for now
    username = session.get('username', 'Guest')
    return render_template('sales_dashboard.html', username=username)


# ==========================================
# API ENDPOINTS
# ==========================================

@app.route('/api/filters', methods=['GET'])
def get_filters():
    """Fetches unique values to populate frontend dropdown menus in a single scan."""
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            # 1. Quick target fetch
            cursor.execute("SELECT DISTINCT Product FROM billing_target WHERE Product IS NOT NULL")
            products = [row['Product'] for row in cursor.fetchall()]

            # 2. Combined scan for billing_data filters
            combined_query = """
                SELECT 
                    GROUP_CONCAT(DISTINCT Sales_Region_Name) as regions,
                    GROUP_CONCAT(DISTINCT Sales_office) as offices,
                    GROUP_CONCAT(DISTINCT fin_year) as fin_years
                FROM billing_data 
                WHERE Dist_Channel_TEXT != 'Inter Unit Transfer'
            """
            cursor.execute(combined_query)
            row = cursor.fetchone()

            # Split comma-separated strings back into clean lists
            regions = [r for r in (row['regions'] or "").split(',') if r]
            offices = [o for o in (row['offices'] or "").split(',') if o]
            fin_years = [y for y in (row['fin_years'] or "").split(',') if y]

    return jsonify({"regions": regions, "offices": offices, "fin_years": fin_years, "products": products})


@app.route('/api/kpis', methods=['GET'])
def get_kpis():
    """Returns Overall Sales, Total Target, Invoices Raised, and Customers Billed."""
    data_where, data_params = build_filter_query("data")
    target_where, target_params = build_filter_query("target")

    # Determine proration factor based on selected months
    months = request.args.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    # 1. Query Actuals from billing_data
    data_query = f"""
        SELECT 
            SUM(Net_Value) as overall_sales,
            COUNT(DISTINCT Billing_Doc) as invoices_raised,
            COUNT(DISTINCT Customer) as customers_billed
        FROM billing_data
        {data_where}
    """

    # 2. Query Targets from billing_target
    target_query = f"""
        SELECT 
            (SUM(Target) * 100000 * {proration_factor}) as total_target
        FROM billing_target
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

    result['overall_sales'] = result.get('overall_sales') or 0
    result['total_target'] = target_result.get('total_target') or 0
    result['invoices_raised'] = result.get('invoices_raised') or 0
    result['customers_billed'] = result.get('customers_billed') or 0

    return jsonify(result)

@app.route('/api/sales-trend', methods=['GET'])
def get_sales_trend():
    """Returns Net Revenue Trend over time (Month-Year aggregation)."""
    where_clause, params = build_filter_query("data")

    # Double %% escapes the % for Python's format string
    query = f"""
        SELECT 
            DATE_FORMAT(Billing_Date, '%%Y-%%m') as month_year,
            SUM(Net_Value) as total_revenue
        FROM billing_data
        {where_clause}
        GROUP BY month_year
        ORDER BY month_year ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            results = cursor.fetchall()

    return jsonify(results)


@app.route('/api/products-revenue', methods=['GET'])
def get_products_revenue():
    """Returns Revenue Contribution by Product Line using the LV Motors plant logic."""
    where_clause, params = build_filter_query("data")

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


@app.route('/api/sales-by-office', methods=['GET'])
def get_sales_by_office():
    """Returns Top 12 Sales Offices by Revenue."""
    where_clause, params = build_filter_query("data")

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


@app.route('/api/sales-by-region', methods=['GET'])
def get_sales_by_region():
    """Returns Revenue by Region."""
    where_clause, params = build_filter_query("data")

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



@app.route('/api/aop-achievement', methods=['GET'])
def get_aop_achievement():
    """
    Calculates Sales Achieved VS AOP.
    Queries billing_data and billing_target separately and maps them via Python dictionary.
    Includes dynamic prorating for target revenue based on month selections.
    """
    data_where, data_params = build_filter_query("data")
    target_where, target_params = build_filter_query("target")

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
        FROM billing_target
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


# ==========================================
# QUERY BUILDER (Fixed & Corrected)
# ==========================================
def build_filter_orders_query(table_type="order_data"):
    """
    Dynamically builds the WHERE clause and parameters based on frontend filters.
    Optimized for MySQL using %s placeholders.
    """
    filters = []
    params = []

    order_regions = request.args.getlist('order_region')
    order_offices = request.args.getlist('order_sales_office')
    order_products = request.args.getlist('order_product')
    fin_years = request.args.getlist('fin_year')
    months = request.args.getlist('month')

    # Region Handling
    if order_regions:
        placeholders = ','.join(['%s'] * len(order_regions))
        filters.append(f"Sales_Region IN ({placeholders})")
        params.extend(order_regions)

    # Sales Office Handling
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
            # If target mapping needs to align with the frontend values
            filters.append(f"Product1 IN ({placeholders})")
        params.extend(order_products)

    # Financial Year Handling
    if fin_years:
        placeholders = ','.join(['%s'] * len(fin_years))
        filters.append(f"fin_year IN ({placeholders})")
        params.extend(fin_years)

    # Month Handling
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
@app.route('/orders')
def order():
    username = session.get('username', 'Guest')
    return render_template('orders_dashboard.html', username=username)


# ==========================================
# API ENDPOINTS (Points to build_filter_orders_query)
# ==========================================

@app.route('/api/filters/orders', methods=['GET'])
def get_orders_filters():
    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute("SELECT DISTINCT Product1 FROM order_target WHERE Product1 IS NOT NULL")
            products = [row['Product1'] for row in cursor.fetchall()]

            combined_query = """
                SELECT 
                    GROUP_CONCAT(DISTINCT Sales_Region) as regions,
                    GROUP_CONCAT(DISTINCT Sales_office) as offices,
                    GROUP_CONCAT(DISTINCT fin_year) as fin_years
                FROM order_data 
            """
            cursor.execute(combined_query)
            row = cursor.fetchone()

            regions = [r for r in (row['regions'] or "").split(',') if r]
            offices = [o for o in (row['offices'] or "").split(',') if o]
            fin_years = [y for y in (row['fin_years'] or "").split(',') if y]

    return jsonify({"regions": regions, "offices": offices, "fin_years": fin_years, "products": products})


@app.route('/api/order/kpis', methods=['GET'])
def get_orders_kpis():
    # FIXED FUNCTION CALLS HERE
    data_where, data_params = build_filter_orders_query("order_data")
    target_where, target_params = build_filter_orders_query("order_target")

    months = request.args.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    data_query = f"""
        SELECT 
            SUM(Net_Value) as overall_sales,
            COUNT(DISTINCT Order_Number) as invoices_raised,
            COUNT(DISTINCT Customer) as customers_billed
        FROM order_data
        {data_where}
    """

    target_query = f"""
        SELECT 
            (SUM(Target) * 100000 * {proration_factor}) as total_target
        FROM order_target
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
    result['invoices_raised'] = result.get('invoices_raised') or 0
    result['customers_billed'] = result.get('customers_billed') or 0

    return jsonify(result)


@app.route('/api/orders-trend', methods=['GET'])
def get_orders_trend():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query("order_data")

    query = f"""
        SELECT 
            DATE_FORMAT(Created_Date, '%%Y-%%m') as month_year,
            SUM(Net_Value) as total_revenue
        FROM order_data
        {where_clause}
        GROUP BY month_year
        ORDER BY month_year ASC
    """

    with get_db_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query, params)
            results = cursor.fetchall()

    return jsonify(results)


@app.route('/api/orders/products-revenue', methods=['GET'])
def get_orders_products_revenue():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query("order_data")

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


@app.route('/api/orders/sales-by-office', methods=['GET'])
def get_orders_sales_by_office():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query("order_data")

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


@app.route('/api/order/sales-by-region', methods=['GET'])
def get_order_sales_by_region():
    # FIXED FUNCTION CALL
    where_clause, params = build_filter_orders_query("order_data")

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


@app.route('/api/orders/aop-achievement', methods=['GET'])
def get_orders_aop_achievement():
    # FIXED FUNCTION CALLS HERE
    data_where, data_params = build_filter_orders_query("order_data")
    target_where, target_params = build_filter_orders_query("order_target")

    months = request.args.getlist('month')
    proration_factor = len(months) / 12.0 if months else 1.0

    actuals_query = f"""
            SELECT 
                CASE 
                   WHEN Product1 = 'LV Motors' AND Product2 = 'LV Motors (Standard)' THEN 'LV MOTORS(S)'
                    ELSE Product1
                END AS mapped_product,
                CASE 
                    WHEN Product1 = 'Transformer Pune' THEN 'UN16'
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
            FROM order_target
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

if __name__ == '__main__':
    app.run(debug=True, threaded=True, host='192.7.200.48', port=8082)
