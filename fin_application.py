
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

fin_app = Blueprint('finance_head', __name__)

engine = create_engine("mysql+pymysql://root:root@192.7.200.4:3306/dashboard")


@fin_app.route('/fin_home')
def finance_home():
    if session.get('role_type') != 'finance_head':
        return redirect(url_for('login'))

    return render_template('finance_home.html', username=session.get('username'))

 ###--- collection Data ----#####

# --- Static Data Boundaries for Finance Head 1 ---
FIN1_PROFIT_CENTERS = [
    'UN01ACM', 'UN01DCM', 'UN01ACG', 'UN01COM', 'UN01MBS', 'UN01TRN',
    'UN07ACM', 'UN07STG', 'UN07DIE', 'UN07WST', 'UN07COM', 'UN07DCM',
    'UN15MCS'
]


@fin_app.route('/collection_fin_report_un1')
def finance_1_dashboard():
    return render_template('Fin1/collection_report.html', username=session.get('username', 'User'),)


# ==========================================
# FILTER BUILDER (Secured)
# ==========================================

def build_collections_filter_un1():
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries."""
    filters = []
    params = []

    # 1. Enforce Static RBAC Data Boundary
    pc_placeholders = ','.join(['%s'] * len(FIN1_PROFIT_CENTERS))
    filters.append(f"Profit_Centre IN ({pc_placeholders})")
    params.extend(FIN1_PROFIT_CENTERS)

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
        valid_pcs = [pc for pc in ui_pcs if pc in FIN1_PROFIT_CENTERS]
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

@fin_app.route('/api/filters/collections/un1', methods=['GET'])
def get_collections_filters_un1():
    """Fetches unique values for filters, restricted strictly to Fin1 boundaries."""

    # Establish base security condition for dropdowns
    pc_placeholders = ','.join(['%s'] * len(FIN1_PROFIT_CENTERS))
    base_where = f"WHERE Profit_Centre IN ({pc_placeholders})"
    base_params = list(FIN1_PROFIT_CENTERS)

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


@fin_app.route('/api/collections/dashboard/un1', methods=['GET'])
def get_collections_dashboard_data_un1():
    where, params = build_collections_filter_un1()

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


@fin_app.route('/api/export/collections_excel_un1', methods=['GET'])
def export_collections_excel_un1():
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_un1()
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

PROFIT_CENTER_MAP = {
    'Unit-1 Govanahalli': ['UN01ACM', 'UN01DCM', 'UN01ACG', 'UN01COM','UN01MBS','UN01TRN'],
    'Unit-7 TUMKUR': ['UN07ACM','UN07STG','UN07DIE','UN07WST','UN07COM','UN07DCM'],
    'Unit-15':['UN15MCS']
}

@fin_app.route('/fin/unit1_mis_report')
def unit_1_fin_mis_report():
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
        'Fin1/MIS.html',
        years=years,
        selected_year=selected_year,
        username=session.get('username', 'User')
    )


@fin_app.route('/api/mis/fin/unit1/unit-monthly-sales')
def get_fin_unit1_monthly_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/unit-scrap-sales')
def get_fin_unit1_scrap_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/other-income-monthly-sales')
def get_fin_unit1_other_income_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/iut-out-monthly-sales')
def get_fin_unit1_iut_out_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/rm-consp-monthly-sales')
def get_fin_unit1_rm_consp_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/employee-monthly-sales')
def get_fin_unit1_employee_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/iut-in-monthly-sales')
def get_fin_unit1_iut_in_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/transport-monthly-sales')
def get_fin_unit1_transport_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/finance-cahrges-monthly-sales')
def get_fin_unit1_finance_charges_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/depreciation-monthly-sales')
def get_fin_unit1_depreciation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/revenue-expenses-monthly-sales')
def get_fin_unit1_revenue_expenses_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/ere-allocation-monthly-sales')
def get_fin_unit1_ere_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/ore-allocation-monthly-sales')
def get_fin_unit1_ore_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit1/intrest-allocation-monthly-sales')
def get_fin_unit1_intrest_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit1/dep-allocation-monthly-sales')
def get_fin_unit1_dep_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP.get(selected_unit, [])
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

# --- Static Data Boundaries for Finance Head 2 ---
FIN2_PROFIT_CENTERS = [
    'UN02ACM', 'UN02ACG', 'UN02COM',
    'UN25ACM', 'UN25ACG', 'UN25EVM', 'UN25COM',
    'UN06DGS'
]

@fin_app.route('/collection_fin_report_un2')
def finance_2_dashboard():
    return render_template('Fin2/collection_report.html', username=session.get('username', 'User'),)


# ==========================================
# FILTER BUILDER (Secured for Fin2)
# ==========================================

def build_collections_filter_fin2():
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries for Finance Head 2."""
    filters = []
    params = []

    # 1. Enforce Static RBAC Data Boundary
    pc_placeholders = ','.join(['%s'] * len(FIN2_PROFIT_CENTERS))
    filters.append(f"Profit_Centre IN ({pc_placeholders})")
    params.extend(FIN2_PROFIT_CENTERS)

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
        valid_pcs = [pc for pc in ui_pcs if pc in FIN2_PROFIT_CENTERS]
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
# API ENDPOINTS (Dedicated to Fin2 to prevent route collisions)
# ==========================================

@fin_app.route('/api/filters/collections_fin2', methods=['GET'])
def get_collections_filters_fin2():
    """Fetches unique values for filters, restricted strictly to Fin2 boundaries."""

    # Establish base security condition for dropdowns
    pc_placeholders = ','.join(['%s'] * len(FIN2_PROFIT_CENTERS))
    base_where = f"WHERE Profit_Centre IN ({pc_placeholders})"
    base_params = list(FIN2_PROFIT_CENTERS)

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


@fin_app.route('/api/collections/dashboard_fin2', methods=['GET'])
def get_collections_dashboard_data_fin2():
    where, params = build_collections_filter_fin2()

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


@fin_app.route('/api/export/collections_excel_un2', methods=['GET'])
def export_collections_excel_un2():
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_fin2()
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

PROFIT_CENTER_MAP_2 = {
    'Unit-2 HUBLI-Special': ['UN02ACM','UN02ACG','UN02COM',],
    'Unit-25ACM ACM-STD Motors': ['UN25ACM'],
    'Unit-25ACG ACG&EVM': ['UN25ACG', 'UN25EVM','UN25COM'],
    'DG Set HUBLI': ['UN06DGS'],
}

@fin_app.route('/fin/unit2_mis_report')
def unit_2_fin_mis_report():
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
        'Fin2/MIS.html',
        years=years,
        selected_year=selected_year,
        username=session.get('username', 'User')
    )


@fin_app.route('/api/mis/fin/unit2/unit-monthly-sales')
def get_fin_unit2_monthly_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/unit-scrap-sales')
def get_fin_unit2_scrap_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/other-income-monthly-sales')
def get_fin_unit2_other_income_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/iut-out-monthly-sales')
def get_fin_unit2_iut_out_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/rm-consp-monthly-sales')
def get_fin_unit2_rm_consp_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/employee-monthly-sales')
def get_fin_unit2_employee_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/iut-in-monthly-sales')
def get_fin_unit2_iut_in_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/transport-monthly-sales')
def get_fin_unit2_transport_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/finance-cahrges-monthly-sales')
def get_fin_unit2_finance_charges_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/depreciation-monthly-sales')
def get_fin_unit2_depreciation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/revenue-expenses-monthly-sales')
def get_fin_unit2_revenue_expenses_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/ere-allocation-monthly-sales')
def get_fin_unit2_ere_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/ore-allocation-monthly-sales')
def get_fin_unit2_ore_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit2/intrest-allocation-monthly-sales')
def get_fin_unit2_intrest_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit2/dep-allocation-monthly-sales')
def get_fin_unit2_dep_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_2.get(selected_unit, [])
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

@fin_app.route('/collection_fin_report_un5')
def finance_5_dashboard():
    return render_template('Fin5/collection_report.html', username=session.get('username', 'User'),)

# --- Static Data Boundaries for Finance Head 5 ---
FIN5_PROFIT_CENTERS = [
    'UN05OFT', 'UN04ELE', 'UN10SWG'
]


def build_collections_filter_fin5():
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries for Finance Head 5."""
    filters = []
    params = []

    # 1. Enforce Static RBAC Data Boundary
    pc_placeholders = ','.join(['%s'] * len(FIN5_PROFIT_CENTERS))
    filters.append(f"Profit_Centre IN ({pc_placeholders})")
    params.extend(FIN5_PROFIT_CENTERS)

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
        valid_pcs = [pc for pc in ui_pcs if pc in FIN5_PROFIT_CENTERS]
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
# API ENDPOINTS (Dedicated to Fin5 to prevent route collisions)
# ==========================================

@fin_app.route('/api/filters/collections_fin5', methods=['GET'])
def get_collections_filters_fin5():
    """Fetches unique values for filters, restricted strictly to Fin5 boundaries."""

    # Establish base security condition for dropdowns
    pc_placeholders = ','.join(['%s'] * len(FIN5_PROFIT_CENTERS))
    base_where = f"WHERE Profit_Centre IN ({pc_placeholders})"
    base_params = list(FIN5_PROFIT_CENTERS)

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


@fin_app.route('/api/collections/dashboard_fin5', methods=['GET'])
def get_collections_dashboard_data_fin5():
    where, params = build_collections_filter_fin5()

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

@fin_app.route('/api/export/collections_excel_un5', methods=['GET'])
def export_collections_excel_un5():
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_fin5()
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

PROFIT_CENTER_MAP_5 = {
 'UN05 Mysore Unit': ['UN05OFT'],
  'UN10 SWGR':['UN10SWG'],
}

@fin_app.route('/fin/unit5_mis_report')
def unit_5_fin_mis_report():
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
        'Fin5/MIS.html',
        years=years,
        selected_year=selected_year,
        username=session.get('username', 'User')
    )


@fin_app.route('/api/mis/fin/unit5/unit-monthly-sales')
def get_fin_unit5_monthly_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/unit-scrap-sales')
def get_fin_unit5_scrap_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/other-income-monthly-sales')
def get_fin_unit5_other_income_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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
        other_income_accounts = [459001,711001, 712001,712002,713001,714001,715001, 715003,715004, 715005, 717001, 719001,720009,720006,715006]

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


@fin_app.route('/api/mis/fin/unit5/iut-out-monthly-sales')
def get_fin_unit5_iut_out_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/rm-consp-monthly-sales')
def get_fin_unit5_rm_consp_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/employee-monthly-sales')
def get_fin_unit5_employee_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/iut-in-monthly-sales')
def get_fin_unit5_iut_in_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/transport-monthly-sales')
def get_fin_unit5_transport_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/finance-cahrges-monthly-sales')
def get_fin_unit5_finance_charges_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/depreciation-monthly-sales')
def get_fin_unit5_depreciation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/revenue-expenses-monthly-sales')
def get_fin_unit5_revenue_expenses_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/ere-allocation-monthly-sales')
def get_fin_unit5_ere_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/ore-allocation-monthly-sales')
def get_fin_unit5_ore_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit5/intrest-allocation-monthly-sales')
def get_fin_unit5_intrest_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit5/dep-allocation-monthly-sales')
def get_fin_unit5_dep_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_5.get(selected_unit, [])
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

@fin_app.route('/collection_fin_report_un16')
def finance_16_dashboard():
    return render_template('Fin16/collection_report.html', username=session.get('username', 'User'),)

FIN16_PROFIT_CENTERS = [
    'UN16CRT', 'UN16OFT', 'UN17CRT'
]


# ==========================================
# FILTER BUILDER (Secured for Fin16)
# ==========================================

def build_collections_filter_fin16():
    """Builds SQL WHERE clause, enforcing strict Profit Center boundaries for Finance Head 16."""
    filters = []
    params = []

    # 1. Enforce Static RBAC Data Boundary
    pc_placeholders = ','.join(['%s'] * len(FIN16_PROFIT_CENTERS))
    filters.append(f"Profit_Centre IN ({pc_placeholders})")
    params.extend(FIN16_PROFIT_CENTERS)

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
        valid_pcs = [pc for pc in ui_pcs if pc in FIN16_PROFIT_CENTERS]
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
# API ENDPOINTS (Dedicated to Fin16 to prevent route collisions)
# ==========================================

@fin_app.route('/api/filters/collections_fin16', methods=['GET'])
def get_collections_filters_fin16():
    """Fetches unique values for filters, restricted strictly to Fin16 boundaries."""

    # Establish base security condition for dropdowns
    pc_placeholders = ','.join(['%s'] * len(FIN16_PROFIT_CENTERS))
    base_where = f"WHERE Profit_Centre IN ({pc_placeholders})"
    base_params = list(FIN16_PROFIT_CENTERS)

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


@fin_app.route('/api/collections/dashboard_fin16', methods=['GET'])
def get_collections_dashboard_data_fin16():
    where, params = build_collections_filter_fin16()

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

@fin_app.route('/api/export/collections_excel_un16', methods=['GET'])
def export_collections_excel_un16():
    """Exports the filtered collections data to an Excel file."""
    # 1. Fetch the dynamic WHERE clause
    where_clause, params = build_collections_filter_fin16()
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


PROFIT_CENTER_MAP_16 = {
'Unit-16 PUNE': ['UN16CRT','UN16OFT','UN17CRT',],
}

@fin_app.route('/fin/unit16_mis_report')
def unit_16_fin_mis_report():
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
        'Fin16/MIS.html',
        years=years,
        selected_year=selected_year,
        username=session.get('username', 'User')
    )


@fin_app.route('/api/mis/fin/unit16/unit-monthly-sales')
def get_fin_unit16_monthly_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/unit-scrap-sales')
def get_fin_unit16_scrap_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/other-income-monthly-sales')
def get_fin_unit16_other_income_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/iut-out-monthly-sales')
def get_fin_unit16_iut_out_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/rm-consp-monthly-sales')
def get_fin_unit16_rm_consp_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/employee-monthly-sales')
def get_fin_unit16_employee_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/iut-in-monthly-sales')
def get_fin_unit16_iut_in_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/transport-monthly-sales')
def get_fin_unit16_transport_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/finance-cahrges-monthly-sales')
def get_fin_unit16_finance_charges_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/depreciation-monthly-sales')
def get_fin_unit16_depreciation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/revenue-expenses-monthly-sales')
def get_fin_unit16_revenue_expenses_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/ere-allocation-monthly-sales')
def get_fin_unit16_ere_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/ore-allocation-monthly-sales')
def get_fin_unit16_ore_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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

@fin_app.route('/api/mis/fin/unit16/intrest-allocation-monthly-sales')
def get_fin_unit16_intrest_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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


@fin_app.route('/api/mis/fin/unit16/dep-allocation-monthly-sales')
def get_fin_unit16_dep_allocation_sales():
    selected_year = request.args.get('fin_year') or request.args.get('fiscal_year')
    selected_unit = request.args.get('unit')

    if not selected_year or not selected_unit:
        return jsonify({"error": "Both Fiscal Year and Unit filters are required"}), 400

    profit_centers = PROFIT_CENTER_MAP_16.get(selected_unit, [])
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
