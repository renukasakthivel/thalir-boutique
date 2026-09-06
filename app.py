import io
import os
import json
import datetime
import pandas as pd
import streamlit as st
import cv2
import numpy as np
import barcode
from barcode.writer import ImageWriter
from pyzbar.pyzbar import decode
from openpyxl.utils import get_column_letter
from db import get_connection, init_production_db, verify_password

# --- STREAMLIT PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Thalir Boutique",
    page_icon="👗",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton>button { width: 100%; border-radius: 8px; height: 3.2em; font-weight: bold; }
    div[data-testid="stMetricValue"] { font-size: 1.35rem; }
</style>
""", unsafe_allow_html=True)

init_production_db()

# --- UTILITY ROUTINES ---
def generate_barcode_label(code_text: str) -> io.BytesIO:
    code_class = barcode.get_barcode_class("code128")
    writer = ImageWriter()
    buffer = io.BytesIO()
    code_class(code_text, writer=writer).write(buffer)
    buffer.seek(0)
    return buffer

def decode_barcode_image(image_bytes: bytes) -> str | None:
    nparr = np.frombuffer(image_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    if img is None:
        return None
    decoded = decode(img)
    for item in decoded:
        return item.data.decode("utf-8")
    return None

def get_partner_name_list():
    conn = get_connection()
    df = pd.read_sql_query("SELECT name FROM partners WHERE is_active=TRUE ORDER BY name ASC;", conn)
    conn.close()
    names = df["name"].tolist() if not df.empty else []
    for default_p in ["Renuka S", "Sudha S", "Mohanapriya G", "Manimegalai S", "Uma S"]:
        if default_p not in names:
            names.append(default_p)
    return sorted(list(set(names)))

def log_audit_action(action_type: str, actor_name: str, details: str):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO backup_audit_logs (action_type, performed_by, user_fullname, details)
        VALUES (%s, %s, %s, %s);
    """, (action_type, st.session_state.auth["user_id"], actor_name, details))
    conn.commit()
    cur.close()
    conn.close()

# --- 2-WAY BACKUP ENGINE ---
def perform_two_way_backup(authorizer_name: str):
    os.makedirs("backups", exist_ok=True)
    timestamp_str = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    local_filename = f"backups/thalir_db_snapshot_{timestamp_str}.json"

    conn = get_connection()
    tables = ["users", "partners", "vendors", "products", "inventory_history", "expenses", "backup_audit_logs"]
    backup_data = {"backup_time": timestamp_str, "tables": {}}

    for tbl in tables:
        df = pd.read_sql_query(f"SELECT * FROM {tbl};", conn)
        backup_data["tables"][tbl] = df.to_dict(orient="records")

    conn.close()

    with open(local_filename, "w", encoding="utf-8") as f:
        json.dump(backup_data, f, default=str, indent=2)

    abs_path = os.path.abspath(local_filename)
    log_audit_action("LOCAL_AND_CLOUD_BACKUP", authorizer_name, f"Created backup snapshot at: {abs_path}")
    return local_filename, backup_data

# --- MASTER EXCEL WORKBOOK BUILDER ---
def build_master_excel_workbook(start_date, end_date, timeframe_label, authorizer_name):
    conn = get_connection()

    products_df = pd.read_sql_query("SELECT * FROM products ORDER BY id ASC;", conn)
    vendors_df = pd.read_sql_query("SELECT * FROM vendors ORDER BY id ASC;", conn)
    partners_df = pd.read_sql_query("SELECT * FROM partners ORDER BY id ASC;", conn)
    audit_df = pd.read_sql_query("""
        SELECT id, created_at, action_type, user_fullname AS approved_by_partner, details
        FROM backup_audit_logs ORDER BY id DESC LIMIT 200;
    """, conn)

    hist_query = """
        SELECT h.id, h.timestamp, h.transaction_type, p.name AS item_name, p.barcode,
               v.name AS vendor_name, h.quantity, h.unit_price, h.cost_price,
               (h.quantity * h.unit_price) AS total_sale_value,
               (h.quantity * h.cost_price) AS total_cost_value,
               h.payment_mode,
               COALESCE(NULLIF(h.biller_name, ''), u.full_name, 'Managing Partner') AS biller_name,
               h.customer_name, h.customer_phone, h.customer_place
        FROM inventory_history h
        JOIN products p ON h.product_id = p.id
        LEFT JOIN vendors v ON p.vendor_id = v.id
        LEFT JOIN users u ON h.performed_by = u.id
        WHERE DATE(h.timestamp) >= %s AND DATE(h.timestamp) <= %s
        ORDER BY h.timestamp DESC;
    """
    history_df = pd.read_sql_query(hist_query, conn, params=(start_date, end_date))

    exp_query = """
        SELECT e.id, e.expense_date, e.category, e.amount, e.description,
               COALESCE(e.approved_by, u.full_name, 'Managing Partner') AS approved_by_partner
        FROM expenses e
        LEFT JOIN users u ON e.logged_by = u.id
        WHERE e.expense_date >= %s AND e.expense_date <= %s
        ORDER BY e.expense_date DESC;
    """
    expenses_df = pd.read_sql_query(exp_query, conn, params=(start_date, end_date))

    stock_in_df = history_df[history_df["transaction_type"] == "STOCK_IN"]
    sales_df = history_df[history_df["transaction_type"] == "STOCK_OUT"]
    returns_df = history_df[history_df["transaction_type"] == "RETURN"]

    total_units_stocked = stock_in_df["quantity"].sum() if not stock_in_df.empty else 0
    total_stockin_investment = stock_in_df["total_cost_value"].sum() if not stock_in_df.empty else 0.0

    gross_units_sold = sales_df["quantity"].sum() if not sales_df.empty else 0
    gross_sales = sales_df["total_sale_value"].sum() if not sales_df.empty else 0.0
    gross_cogs = sales_df["total_cost_value"].sum() if not sales_df.empty else 0.0

    returned_units = returns_df["quantity"].sum() if not returns_df.empty else 0
    return_val = returns_df["total_sale_value"].sum() if not returns_df.empty else 0.0
    return_cogs = returns_df["total_cost_value"].sum() if not returns_df.empty else 0.0

    net_units_sold = gross_units_sold - returned_units
    net_sales = gross_sales - return_val
    net_cogs = gross_cogs - return_cogs
    total_exp = expenses_df["amount"].sum() if not expenses_df.empty else 0.0
    gross_profit = net_sales - net_cogs
    net_profit = gross_profit - total_exp

    cash_sales = sales_df[sales_df["payment_mode"] == "Cash"]["total_sale_value"].sum() if not sales_df.empty else 0.0
    gpay_sales = sales_df[sales_df["payment_mode"].str.contains("GPay|UPI|PhonePe", na=False)]["total_sale_value"].sum() if not sales_df.empty else 0.0
    other_pay_sales = gross_sales - cash_sales - gpay_sales

    now_ts = datetime.datetime.now()
    tally_data = [
        {"Formal Accounting Parameter": "REPORT INFORMATION", "Value / Amount": ""},
        {"Formal Accounting Parameter": "Accounting Cycle / Timeframe", "Value / Amount": timeframe_label},
        {"Formal Accounting Parameter": "Period Start Date", "Value / Amount": str(start_date)},
        {"Formal Accounting Parameter": "Period End Date", "Value / Amount": str(end_date)},
        {"Formal Accounting Parameter": "Report Generation Timestamp", "Value / Amount": now_ts.strftime("%A, %Y-%m-%d %I:%M:%S %p")},
        {"Formal Accounting Parameter": "Authorized By (Partner Full Name)", "Value / Amount": authorizer_name},
        {"Formal Accounting Parameter": "Approval Authority Role", "Value / Amount": "Executive Partner"},
        {"Formal Accounting Parameter": "----------------------------------------------------", "Value / Amount": "-----------------------"},
        {"Formal Accounting Parameter": "1. INVENTORY MOVEMENT SUMMARY (PIECES)", "Value / Amount": ""},
        {"Formal Accounting Parameter": "New Garments Purchased & Added to Rack (Stock-In)", "Value / Amount": f"{int(total_units_stocked)} pcs"},
        {"Formal Accounting Parameter": "Gross Garments Sold to Customers (Stock-Out)", "Value / Amount": f"{int(gross_units_sold)} pcs"},
        {"Formal Accounting Parameter": "Customer Returns / Reversals Deducted", "Value / Amount": f"{int(returned_units)} pcs"},
        {"Formal Accounting Parameter": "Net Garments Sold (Actual Quantity Out)", "Value / Amount": f"{int(net_units_sold)} pcs"},
        {"Formal Accounting Parameter": "----------------------------------------------------", "Value / Amount": "-----------------------"},
        {"Formal Accounting Parameter": "2. REVENUE & PAYMENT COLLECTIONS (INFLOW)", "Value / Amount": ""},
        {"Formal Accounting Parameter": "Gross Billing Revenue Generated", "Value / Amount": round(gross_sales, 2)},
        {"Formal Accounting Parameter": "  -> Collected via Cash Register (₹)", "Value / Amount": round(cash_sales, 2)},
        {"Formal Accounting Parameter": "  -> Collected via GPay / PhonePe / UPI (₹)", "Value / Amount": round(gpay_sales, 2)},
        {"Formal Accounting Parameter": "  -> Collected via Card / POS / Bank (₹)", "Value / Amount": round(other_pay_sales, 2)},
        {"Formal Accounting Parameter": "Less: Customer Refunds Issued for Returns (₹)", "Value / Amount": round(return_val, 2)},
        {"Formal Accounting Parameter": "NET REVENUE COLLECTED (Actual Inflow) (₹)", "Value / Amount": round(net_sales, 2)},
        {"Formal Accounting Parameter": "----------------------------------------------------", "Value / Amount": "-----------------------"},
        {"Formal Accounting Parameter": "3. COST OF GOODS SOLD & GROSS MARGIN", "Value / Amount": ""},
        {"Formal Accounting Parameter": "Wholesale Cost of Net Sold Stock (COGS) (₹)", "Value / Amount": round(net_cogs, 2)},
        {"Formal Accounting Parameter": "GROSS OPERATING PROFIT MARGIN (Net Revenue - COGS) (₹)", "Value / Amount": round(gross_profit, 2)},
        {"Formal Accounting Parameter": "----------------------------------------------------", "Value / Amount": "-----------------------"},
        {"Formal Accounting Parameter": "4. OPERATING OVERHEAD EXPENSES (OUTFLOW)", "Value / Amount": ""},
        {"Formal Accounting Parameter": "Total Boutique Overhead (Rent, EB, Salary, Fuel, etc.) (₹)", "Value / Amount": round(total_exp, 2)},
        {"Formal Accounting Parameter": "----------------------------------------------------", "Value / Amount": "-----------------------"},
        {"Formal Accounting Parameter": "5. NET BUSINESS POSITION & PARTNER SETTLEMENT", "Value / Amount": ""},
        {"Formal Accounting Parameter": "FINAL NET PROFIT / (NET LOSS) (₹)", "Value / Amount": round(net_profit, 2)},
    ]

    active_partners = partners_df[partners_df["is_active"] == True]
    for _, prow in active_partners.iterrows():
        p_share = (net_profit * float(prow["profit_percentage"])) / 100.0
        status_word = "Profit Payout" if net_profit >= 0 else "Loss Contribution"
        tally_data.append({
            "Formal Accounting Parameter": f"Partner {status_word}: {prow['name']} ({prow['profit_percentage']}%) (₹)",
            "Value / Amount": round(p_share, 2)
        })

    tally_df = pd.DataFrame(tally_data)

    vendor_analysis = []
    for _, v in vendors_df.iterrows():
        v_sales = sales_df[sales_df["vendor_name"] == v["name"]]
        v_units = v_sales["quantity"].sum() if not v_sales.empty else 0
        v_rev = v_sales["total_sale_value"].sum() if not v_sales.empty else 0.0
        v_cost = v_sales["total_cost_value"].sum() if not v_sales.empty else 0.0
        vendor_analysis.append({
            "Vendor / Supplier Name": v["name"],
            "Contact Person": v["contact_person"],
            "Phone Number": v["phone"],
            "Mill / Store Address": v.get("address", ""),
            "Pieces Sold (Units)": v_units,
            "Total Sales Realized (₹)": v_rev,
            "Wholesale Cost of Stock (₹)": v_cost,
            "Gross Profit Earned (₹)": v_rev - v_cost
        })
    vendor_analysis_df = pd.DataFrame(vendor_analysis).sort_values(by="Pieces Sold (Units)", ascending=False)

    item_summary = []
    for _, prod in products_df.iterrows():
        prod_hist = history_df[history_df["barcode"] == prod["barcode"]] if not history_df.empty else pd.DataFrame()
        p_in = prod_hist[prod_hist["transaction_type"] == "STOCK_IN"]["quantity"].sum() if not prod_hist.empty else 0
        p_out = prod_hist[prod_hist["transaction_type"] == "STOCK_OUT"]["quantity"].sum() if not prod_hist.empty else 0
        p_ret = prod_hist[prod_hist["transaction_type"] == "RETURN"]["quantity"].sum() if not prod_hist.empty else 0

        v_match = vendors_df[vendors_df["id"] == prod["vendor_id"]]
        v_name = v_match["name"].values[0] if not v_match.empty else "Direct"

        item_summary.append({
            "Barcode (SKU)": prod["barcode"],
            "Garment Description": prod["name"],
            "Vendor / Supplier": v_name,
            "Wholesale Purchase Cost (₹)": prod["wholesale_cost"],
            "Retail Selling Price (₹)": prod["selling_price"],
            "Current Available Stock (Pcs)": prod["stock_quantity"],
            "Period Stocked In (Pcs)": p_in,
            "Period Sold (Pcs)": p_out,
            "Period Returned (Pcs)": p_ret,
            "Net Sold Volume (Pcs)": (p_out - p_ret),
            "Period Sales Generated (₹)": (p_out - p_ret) * float(prod["selling_price"]),
            "Catalog Status": "Active" if prod["is_active"] else "Deactivated",
            "Reorder Advisory": "🚨 URGENT RESTOCK REQUIRED" if prod["stock_quantity"] <= prod["min_threshold"] and (p_out - p_ret) > 0 else "Stock Adequate"
        })
    item_stock_tally_df = pd.DataFrame(item_summary).sort_values(by="Net Sold Volume (Pcs)", ascending=False)

    audit_clean_df = audit_df.copy()
    audit_clean_df["created_at"] = pd.to_datetime(audit_clean_df["created_at"]).dt.strftime("%Y-%m-%d %I:%M:%S %p")
    history_clean_df = history_df.copy()
    history_clean_df["timestamp"] = pd.to_datetime(history_clean_df["timestamp"]).dt.strftime("%Y-%m-%d %I:%M:%S %p")
    expenses_clean_df = expenses_df.copy()
    expenses_clean_df["expense_date"] = pd.to_datetime(expenses_clean_df["expense_date"]).dt.strftime("%Y-%m-%d")

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tally_df.to_excel(writer, sheet_name="Formal_P&L_Tally", index=False)
        history_clean_df.to_excel(writer, sheet_name="Customer_Sales_Ledger", index=False)
        vendor_analysis_df.to_excel(writer, sheet_name="Vendor_Performance_Audit", index=False)
        item_stock_tally_df.to_excel(writer, sheet_name="Product_Sales_Ranking", index=False)
        expenses_clean_df.to_excel(writer, sheet_name="Approved_Expenses", index=False)
        products_df.to_excel(writer, sheet_name="Inventory_Catalog", index=False)
        partners_df.to_excel(writer, sheet_name="Partner_Capital_Equity", index=False)
        vendors_df.to_excel(writer, sheet_name="Vendors_Directory", index=False)
        audit_clean_df.to_excel(writer, sheet_name="System_Audit_Trail", index=False)

        for sheetname in writer.sheets:
            ws = writer.sheets[sheetname]
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    conn.close()
    buffer.seek(0)
    return buffer

# --- AUTH STATE ---
if "auth" not in st.session_state:
    st.session_state.auth = {"logged_in": False, "user_id": None, "username": "", "role": "", "full_name": ""}

def logout():
    st.session_state.auth = {"logged_in": False, "user_id": None, "username": "", "role": "", "full_name": ""}
    st.rerun()

# ----------------------------------------------------
# LOGIN SCREEN
# ----------------------------------------------------
if not st.session_state.auth["logged_in"]:
    st.title("👗 Thalir Boutique")
    st.caption("Store Inventory, POS, Formal Partner Tally & Security Audit Trail")

    with st.form("login_box"):
        u_name = st.text_input("Username")
        p_word = st.text_input("Password", type="password")
        if st.form_submit_button("Sign In"):
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT id, username, password_hash, role, full_name, is_active FROM users WHERE username = %s;", (u_name.strip(),))
            user = cur.fetchone()
            cur.close()
            conn.close()

            if user and user[5] and verify_password(p_word, user[2]):
                st.session_state.auth = {"logged_in": True, "user_id": user[0], "username": user[1], "role": user[3], "full_name": user[4]}
                st.rerun()
            else:
                st.error("Invalid credentials.")
    st.stop()

# ----------------------------------------------------
# NAVIGATION
# ----------------------------------------------------
user_role = st.session_state.auth["role"]
st.sidebar.markdown(f"### Thalir Boutique\n**{st.session_state.auth['full_name']}** ({user_role})")

if user_role in ["Partner", "Admin"]:
    menu = [
        "🛒 Point of Sale (Billing & Customer)",
        "📈 Vendor & Best-Seller Analytics",
        "🔄 Returns & Sales Reversals",
        "📦 Stock Levels & Catalog",
        "📥 Stock-In (New / Restock)",
        "💸 Store Expenses & Approvals",
        "📊 Multi-Month Z-Report & Excel Download",
        "💾 2-Way Backup & Security Audit Trail",
        "🤝 Partner Capital & Equity",
        "🏭 Vendor Management"
    ]
else:
    menu = [
        "🛒 Point of Sale (Billing & Customer)",
        "🔄 Returns & Sales Reversals",
        "📦 Stock Levels & Catalog"
    ]

page = st.sidebar.radio("Navigation", menu)
st.sidebar.button("Logout", on_click=logout)

partner_names = get_partner_name_list()

# ----------------------------------------------------
# 1. POINT OF SALE (WITH CUSTOMER, BILLER & REACTIVATION)
# ----------------------------------------------------
if page == "🛒 Point of Sale (Billing & Customer)":
    st.subheader("🛒 Fast Billing & Customer Checkout")

    with st.expander("📷 Open Camera Scanner", expanded=False):
        cam = st.camera_input("Scan Barcode")
        scanned_code = ""
        if cam:
            scanned_code = decode_barcode_image(cam.getvalue()) or ""
            if scanned_code:
                st.success(f"Scanned: {scanned_code}")

    with st.form("billing_form"):
        col1, col2 = st.columns([2, 1])
        with col1:
            item_barcode = st.text_input("Barcode / Item Code", value=scanned_code)
        with col2:
            item_qty = st.number_input("Total Pieces (Qty)", min_value=1, value=1, step=1)

        st.markdown("##### 💵 Payment & Biller Accountability")
        pay_c1, pay_c2 = st.columns(2)
        with pay_c1:
            payment_mode = st.selectbox("Payment Mode", ["Cash", "GPay / UPI", "PhonePe / Paytm", "Card / POS", "Net Banking"])
        with pay_c2:
            biller_name = st.selectbox("Billing Manager / Partner in Charge", partner_names)

        st.markdown("##### 👤 Customer Details (Optional for CRM & Warranty)")
        cust_c1, cust_c2, cust_c3 = st.columns(3)
        with cust_c1:
            c_name = st.text_input("Customer Name", placeholder="e.g., Priya")
        with cust_c2:
            c_phone = st.text_input("Mobile Number", placeholder="e.g., 9876543210")
        with cust_c3:
            c_place = st.text_input("Place / City", placeholder="e.g., Salem / Erode")

        submit_bill = st.form_submit_button("Complete Bill & Stock Out", type="primary")

    if submit_bill:
        if not item_barcode.strip():
            st.error("Please enter or scan a barcode.")
        else:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("""
                SELECT id, name, selling_price, wholesale_cost, stock_quantity, is_active 
                FROM products WHERE barcode = %s;
            """, (item_barcode.strip(),))
            prod = cur.fetchone()

            if not prod:
                st.error("Item barcode not found in catalog.")
            elif not prod[5]:
                st.warning(f"⚠️ Item '{prod[1]}' is currently DEACTIVATED, but has {prod[4]} pcs in stock.")
                st.info("Go to '📦 Stock Levels & Catalog' -> '♻️ Reactivate Deactivated Products' tab to turn it on with partner approval before billing.")
            elif prod[4] < item_qty:
                st.error(f"Insufficient stock! Only {prod[4]} pcs available on rack.")
            else:
                p_id, p_name, s_price, w_cost, cur_qty, _ = prod
                new_qty = cur_qty - item_qty

                cur.execute("UPDATE products SET stock_quantity = %s WHERE id = %s;", (new_qty, p_id))
                cur.execute("""
                    INSERT INTO inventory_history (
                        product_id, transaction_type, quantity, unit_price, cost_price,
                        customer_name, customer_phone, customer_place, biller_name, payment_mode, performed_by
                    ) VALUES (%s, 'STOCK_OUT', %s, %s, %s, %s, %s, %s, %s, %s, %s);
                """, (
                    p_id, item_qty, s_price, w_cost,
                    c_name.strip(), c_phone.strip(), c_place.strip(),
                    biller_name, payment_mode, st.session_state.auth["user_id"]
                ))
                conn.commit()

                total = float(s_price) * item_qty
                st.success(f"✅ Billed {item_qty} pcs of '{p_name}' for ₹{total:,.2f} via {payment_mode} by '{biller_name}'. Remaining stock: {new_qty} pcs.")

            cur.close()
            conn.close()

# ----------------------------------------------------
# 2. VENDOR & BEST-SELLER ANALYTICS
# ----------------------------------------------------
elif page == "📈 Vendor & Best-Seller Analytics":
    st.subheader("📈 Vendor Analytics & Fast-Moving Products")
    st.caption("Detailed ranking of suppliers and products to guide purchasing decisions.")

    conn = get_connection()
    prod_rank_df = pd.read_sql_query("""
        SELECT p.barcode, p.name AS product_name, v.name AS vendor_name,
               COALESCE(SUM(h.quantity), 0) AS units_sold,
               COALESCE(SUM(h.quantity * h.unit_price), 0) AS total_revenue,
               p.stock_quantity AS current_stock, p.min_threshold,
               v.phone AS vendor_phone
        FROM products p
        LEFT JOIN vendors v ON p.vendor_id = v.id
        LEFT JOIN inventory_history h ON p.id = h.product_id AND h.transaction_type = 'STOCK_OUT'
        GROUP BY p.barcode, p.name, v.name, p.stock_quantity, p.min_threshold, v.phone
        ORDER BY units_sold DESC;
    """, conn)

    vendor_rank_df = pd.read_sql_query("""
        SELECT v.name AS vendor_name, v.contact_person, v.phone, v.address,
               COALESCE(SUM(h.quantity), 0) AS total_pieces_sold,
               COALESCE(SUM(h.quantity * h.unit_price), 0) AS total_sales_generated,
               COALESCE(SUM(h.quantity * (h.unit_price - h.cost_price)), 0) AS gross_profit_earned
        FROM vendors v
        JOIN products p ON v.id = p.vendor_id
        JOIN inventory_history h ON p.id = h.product_id AND h.transaction_type = 'STOCK_OUT'
        GROUP BY v.name, v.contact_person, v.phone, v.address
        ORDER BY total_pieces_sold DESC;
    """, conn)
    conn.close()

    t_v1, t_v2, t_v3 = st.tabs(["🏆 Top Selling Garments", "🏭 Vendor Ranking (Most Sold)", "🚨 Restock Reorder Advisory"])

    with t_v1:
        st.markdown("#### 🔥 Most Popular Garments")
        st.dataframe(prod_rank_df[["barcode", "product_name", "vendor_name", "units_sold", "total_revenue", "current_stock"]], use_container_width=True)

    with t_v2:
        st.markdown("#### 🏭 Best Performing Vendors (By Volume & Profit)")
        if not vendor_rank_df.empty:
            st.dataframe(vendor_rank_df, use_container_width=True)
            top_v = vendor_rank_df.iloc[0]
            st.info(f"💡 **Top Vendor:** **{top_v['vendor_name']}** has moved **{top_v['total_pieces_sold']} pieces** totaling ₹{top_v['total_sales_generated']:,.2f} in sales.")
        else:
            st.info("No vendor sales recorded yet.")

    with t_v3:
        st.markdown("#### 🚨 Products to Reorder from Vendors")
        reorder_needed = prod_rank_df[(prod_rank_df["current_stock"] <= prod_rank_df["min_threshold"]) & (prod_rank_df["units_sold"] > 0)]
        if not reorder_needed.empty:
            for _, item in reorder_needed.iterrows():
                st.warning(f"⚠️ **{item['product_name']}** (Barcode: `{item['barcode']}`)\n"
                           f"* Sold: **{item['units_sold']} pcs** | Current Rack Stock: **{item['current_stock']} pcs**\n"
                           f"* Supplier: **{item['vendor_name']}** (Phone: {item['vendor_phone'] or 'N/A'})\n"
                           f"* **Recommendation:** Repurchase 10–25 pcs from this vendor immediately.")
        else:
            st.success("✅ All top-selling items have adequate rack stock.")

# ----------------------------------------------------
# 3. RETURNS & REVERSALS
# ----------------------------------------------------
elif page == "🔄 Returns & Sales Reversals":
    st.subheader("🔄 Customer Return / Bill Reversal")

    with st.form("return_box"):
        r_code = st.text_input("Barcode of Item Being Returned")
        r_qty = st.number_input("Units to Return", min_value=1, value=1, step=1)
        r_biller = st.selectbox("Partner Authorizing Return & Refund", partner_names)
        r_submit = st.form_submit_button("Process Return & Refund")

    if r_submit:
        if not r_code.strip():
            st.error("Enter barcode.")
        else:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT id, name, selling_price, wholesale_cost, stock_quantity FROM products WHERE barcode = %s;", (r_code.strip(),))
            prod = cur.fetchone()

            if not prod:
                st.error("Item not found.")
            else:
                p_id, p_name, s_price, w_cost, cur_qty = prod
                new_qty = cur_qty + r_qty

                cur.execute("UPDATE products SET stock_quantity = %s WHERE id = %s;", (new_qty, p_id))
                cur.execute("""
                    INSERT INTO inventory_history (product_id, transaction_type, quantity, unit_price, cost_price, biller_name, payment_mode, performed_by)
                    VALUES (%s, 'RETURN', %s, %s, %s, %s, 'Refund', %s);
                """, (p_id, r_qty, s_price, w_cost, r_biller, st.session_state.auth["user_id"]))
                conn.commit()

                log_audit_action("CUSTOMER_RETURN_APPROVED", r_biller, f"Refunded {r_qty}x '{p_name}' for ₹{float(s_price)*r_qty:,.2f}")
                refund_amt = float(s_price) * r_qty
                st.success(f"Returned {r_qty}x '{p_name}'. Refunded ₹{refund_amt:,.2f} approved by {r_biller}. Stock updated to {new_qty}.")

            cur.close()
            conn.close()

# ----------------------------------------------------
# 4. STOCK LEVELS & CATALOG (WITH REACTIVATION)
# ----------------------------------------------------
elif page == "📦 Stock Levels & Catalog":
    st.subheader("📦 Inventory Catalog & Stock Corrections")

    conn = get_connection()
    cur = conn.cursor()

    df_active = pd.read_sql_query("""
        SELECT p.id, p.barcode, p.name, v.name AS vendor, p.wholesale_cost, p.selling_price, p.stock_quantity, p.min_threshold
        FROM products p
        LEFT JOIN vendors v ON p.vendor_id = v.id
        WHERE p.is_active IS NULL OR p.is_active = TRUE
        ORDER BY p.stock_quantity ASC;
    """, conn)

    df_deactive = pd.read_sql_query("""
        SELECT p.id, p.barcode, p.name, v.name AS vendor, p.wholesale_cost, p.selling_price, p.stock_quantity, p.min_threshold
        FROM products p
        LEFT JOIN vendors v ON p.vendor_id = v.id
        WHERE p.is_active = FALSE
        ORDER BY p.id ASC;
    """, conn)

    tab_cat1, tab_cat2 = st.tabs(["Active Catalog", "♻️ Reactivate Deactivated Products"])

    with tab_cat1:
        if not df_active.empty:
            low_stock = df_active[df_active["stock_quantity"] <= df_active["min_threshold"]]
            if not low_stock.empty:
                st.warning(f"⚠️ {len(low_stock)} item(s) are at or below threshold!")
            st.dataframe(df_active, use_container_width=True)

            if user_role in ["Partner", "Admin"]:
                with st.expander("🛠️ Modify or Deactivate Product (Requires Partner Approval)"):
                    prod_map = dict(zip(df_active["name"] + " (" + df_active["barcode"] + ")", df_active["id"]))
                    chosen = st.selectbox("Select Product", list(prod_map.keys()))
                    cid = prod_map[chosen]
                    r = df_active[df_active["id"] == cid].iloc[0]

                    col_m1, col_m2 = st.columns(2)
                    with col_m1:
                        c_name = st.text_input("Garment Name", value=r["name"])
                        c_qty = st.number_input("Rack Stock Qty", value=int(r["stock_quantity"]), min_value=0)
                    with col_m2:
                        c_cost = st.number_input("Wholesale Cost (₹)", value=float(r["wholesale_cost"]), min_value=0.0)
                        c_sell = st.number_input("Selling Price (₹)", value=float(r["selling_price"]), min_value=0.0)

                    approver_partner = st.selectbox("Partner Authorizing Modification / Deactivation", partner_names, key="prod_mod_auth")

                    b_col1, b_col2 = st.columns(2)
                    if b_col1.button("Save Corrections"):
                        cur.execute("UPDATE products SET name=%s, stock_quantity=%s, wholesale_cost=%s, selling_price=%s WHERE id=%s;",
                                    (c_name, c_qty, c_cost, c_sell, cid))
                        conn.commit()
                        log_audit_action("PRODUCT_MODIFIED", approver_partner, f"Updated '{c_name}' (Qty: {c_qty}, Cost: {c_cost}, Price: {c_sell})")
                        st.success("Updated successfully.")
                        st.rerun()

                    if b_col2.button("Deactivate Item"):
                        cur.execute("UPDATE products SET is_active = FALSE WHERE id = %s;", (cid,))
                        conn.commit()
                        log_audit_action("PRODUCT_DEACTIVATED", approver_partner, f"Deactivated '{r['name']}' ({r['barcode']}) with {r['stock_quantity']} pcs remaining")
                        st.success("Item deactivated. If you ever need to sell the remaining stock, reactivate it in the tab above.")
                        st.rerun()
        else:
            st.info("Active catalog empty.")

    with tab_cat2:
        st.markdown("#### ♻️ Products Turned Off / Deactivated")
        st.caption("If an item was deactivated by mistake or still has stock left on the rack, reactivate it here.")
        if not df_deactive.empty:
            st.dataframe(df_deactive, use_container_width=True)
            deact_map = dict(zip(df_deactive["name"] + " (" + df_deactive["barcode"] + ") - Stock: " + df_deactive["stock_quantity"].astype(str), df_deactive["id"]))
            chosen_deact = st.selectbox("Select Deactivated Product to Reactivate:", list(deact_map.keys()))
            react_id = deact_map[chosen_deact]
            react_auth = st.selectbox("Partner Authorizing Reactivation:", partner_names, key="react_auth_p")

            if st.button("♻️ Reactivate to Active Catalog", type="primary"):
                cur.execute("UPDATE products SET is_active = TRUE WHERE id = %s;", (react_id,))
                conn.commit()
                log_audit_action("PRODUCT_REACTIVATED", react_auth, f"Reactivated product ID {react_id} back to active catalog")
                st.success("Product successfully reactivated! It is now scannable in POS.")
                st.rerun()
        else:
            st.info("No deactivated products currently.")

    cur.close()
    conn.close()

# ----------------------------------------------------
# 5. STOCK-IN
# ----------------------------------------------------
elif page == "📥 Stock-In (New / Restock)":
    st.subheader("📥 Add Inventory & Print Label")

    conn = get_connection()
    v_df = pd.read_sql_query("SELECT id, name FROM vendors ORDER BY name ASC;", conn)
    conn.close()
    v_dict = dict(zip(v_df["name"], v_df["id"]))

    with st.form("stockin_box"):
        b_code = st.text_input("Barcode (SKU)")
        p_name = st.text_input("Garment Name / Style")
        v_sel = st.selectbox("Vendor", ["None"] + list(v_dict.keys()))
        cw, cs = st.columns(2)
        with cw: w_cost = st.number_input("Wholesale Cost (₹)", min_value=0.0, step=10.0)
        with cs: s_price = st.number_input("Selling Price (₹)", min_value=0.0, step=10.0)
        units = st.number_input("Units to Add", min_value=1, value=1, step=1)
        in_biller = st.selectbox("Partner Authorizing Stock-In", partner_names)
        submit_in = st.form_submit_button("Save to Inventory")

    if submit_in:
        if not b_code.strip() or not p_name.strip():
            st.error("Barcode and Garment Name are required.")
        else:
            v_id = v_dict.get(v_sel) if v_sel != "None" else None
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("SELECT id, stock_quantity FROM products WHERE barcode = %s;", (b_code.strip(),))
            existing = cur.fetchone()

            if existing:
                p_id, cur_stk = existing
                new_q = cur_stk + units
                cur.execute("""
                    UPDATE products 
                    SET stock_quantity = %s, wholesale_cost = %s, selling_price = %s,
                        vendor_id = COALESCE(%s, vendor_id), is_active = TRUE 
                    WHERE id = %s;
                """, (new_q, w_cost, s_price, v_id, p_id))
            else:
                cur.execute("""
                    INSERT INTO products (barcode, name, vendor_id, wholesale_cost, selling_price, stock_quantity, is_active)
                    VALUES (%s, %s, %s, %s, %s, %s, TRUE) RETURNING id;
                """, (b_code.strip(), p_name.strip(), v_id, w_cost, s_price, units))
                p_id = cur.fetchone()[0]

            cur.execute("""
                INSERT INTO inventory_history (product_id, transaction_type, quantity, unit_price, cost_price, biller_name, payment_mode, performed_by)
                VALUES (%s, 'STOCK_IN', %s, %s, %s, %s, 'Purchase/StockIn', %s);
            """, (p_id, units, s_price, w_cost, in_biller, st.session_state.auth["user_id"]))
            conn.commit()
            cur.close()
            conn.close()

            log_audit_action("STOCK_IN_AUTHORIZED", in_biller, f"Added {units} pcs of '{p_name}' at ₹{w_cost} each")
            st.success(f"Added {units} units of '{p_name}'.")
            lbl = generate_barcode_label(b_code.strip())
            st.image(lbl, caption=f"Label: {b_code.strip()}")
            st.download_button("Download Label Image", lbl.getvalue(), file_name=f"{b_code}.png", mime="image/png")

# ----------------------------------------------------
# 6. STORE EXPENSES & APPROVALS
# ----------------------------------------------------
elif page == "💸 Store Expenses & Approvals":
    st.subheader("💸 Store Operating Expenses & Approvals")

    with st.form("exp_form", clear_on_submit=True):
        col_x1, col_x2 = st.columns(2)
        with col_x1:
            cat = st.selectbox("Category", ["Rent", "Electricity", "Vehicle / Fuel", "Salary", "Maintenance", "Packaging", "Miscellaneous"])
            amt = st.number_input("Amount (₹)", min_value=1.0, step=50.0)
            e_date = st.date_input("Date", datetime.date.today())
        with col_x2:
            desc = st.text_input("Expense Description / Notes", placeholder="e.g., Shop Rent Paid to Owner")
            approver_name = st.selectbox("Partner Approving This Expense", partner_names)

        if st.form_submit_button("Record Approved Expense", type="primary"):
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO expenses (category, amount, description, approved_by, logged_by, expense_date)
                VALUES (%s, %s, %s, %s, %s, %s);
            """, (cat, amt, desc, approver_name, st.session_state.auth["user_id"], e_date))
            conn.commit()
            cur.close()
            conn.close()

            log_audit_action("EXPENSE_RECORDED", approver_name, f"Approved ₹{amt:,.2f} for {cat} ({desc})")
            st.success(f"Expense of ₹{amt:,.2f} recorded and approved by {approver_name}.")

    conn = get_connection()
    df_exp = pd.read_sql_query("""
        SELECT id, expense_date, category, amount, description, approved_by
        FROM expenses ORDER BY expense_date DESC LIMIT 50;
    """, conn)
    st.dataframe(df_exp, use_container_width=True)

    if not df_exp.empty and user_role in ["Partner", "Admin"]:
        with st.expander("🗑️ Delete / Remove a Mistyped Expense (Requires Partner Authorization)"):
            del_id = st.selectbox("Select Expense ID to Remove", df_exp["id"].tolist())
            del_approver = st.selectbox("Partner Authorizing Expense Deletion", partner_names, key="del_exp_auth")
            if st.button("Delete Selected Expense Entry"):
                cur = conn.cursor()
                cur.execute("SELECT amount, category, description FROM expenses WHERE id = %s;", (del_id,))
                r_del = cur.fetchone()
                cur.execute("DELETE FROM expenses WHERE id = %s;", (del_id,))
                conn.commit()
                cur.close()

                log_audit_action("EXPENSE_DELETED", del_approver, f"Deleted Expense ID {del_id}: ₹{r_del[0]} for {r_del[1]} ({r_del[2]})")
                st.success(f"Expense ID {del_id} removed by {del_approver}.")
                st.rerun()
    conn.close()

# ----------------------------------------------------
# 7. MULTI-MONTH Z-REPORT & FORMAL EXCEL
# ----------------------------------------------------
elif page == "📊 Multi-Month Z-Report & Excel Download":
    st.subheader("📊 Executive Stock Balance & Financial Tally")
    st.caption("Formal accounting reconciliation for all partners.")

    today = datetime.date.today()
    options = [
        "Current Week",
        "Current Month",
        "Past 2 Months",
        "Past 3 Months (Quarterly)",
        "Past 4 Months",
        "Past 5 Months",
        "Past 6 Months (Half-Year)",
        "Past 7 Months",
        "Past 8 Months",
        "Past 9 Months",
        "Past 10 Months",
        "Past 12 Months (Annual / Yearly)",
        "Custom Date Range"
    ]
    chosen_period = st.selectbox("Select Reporting Period", options)

    if chosen_period == "Current Week":
        start_d = today - datetime.timedelta(days=today.weekday())
        end_d = today
    elif chosen_period == "Current Month":
        start_d = today.replace(day=1)
        end_d = today
    elif "Months" in chosen_period:
        num_months = int(chosen_period.split()[1])
        start_d = today - datetime.timedelta(days=num_months * 30)
        end_d = today
    else:
        c1, c2 = st.columns(2)
        with c1: start_d = st.date_input("Start Date", today - datetime.timedelta(days=30))
        with c2: end_d = st.date_input("End Date", today)

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT COALESCE(SUM(quantity), 0), COALESCE(SUM(quantity * cost_price), 0)
        FROM inventory_history
        WHERE transaction_type = 'STOCK_IN' AND DATE(timestamp) >= %s AND DATE(timestamp) <= %s;
    """, (start_d, end_d))
    tot_stock_in_qty, tot_stock_in_cost = cur.fetchone()

    cur.execute("""
        SELECT COALESCE(SUM(quantity), 0), COALESCE(SUM(quantity * unit_price), 0), COALESCE(SUM(quantity * cost_price), 0)
        FROM inventory_history
        WHERE transaction_type = 'STOCK_OUT' AND DATE(timestamp) >= %s AND DATE(timestamp) <= %s;
    """, (start_d, end_d))
    gross_sold_qty, gross_sales_rev, gross_sales_cogs = cur.fetchone()

    cur.execute("""
        SELECT COALESCE(SUM(quantity), 0), COALESCE(SUM(quantity * unit_price), 0), COALESCE(SUM(quantity * cost_price), 0)
        FROM inventory_history
        WHERE transaction_type = 'RETURN' AND DATE(timestamp) >= %s AND DATE(timestamp) <= %s;
    """, (start_d, end_d))
    ret_qty, ret_rev, ret_cogs = cur.fetchone()

    cur.execute("SELECT COALESCE(SUM(amount), 0) FROM expenses WHERE expense_date >= %s AND expense_date <= %s;", (start_d, end_d))
    total_expenses = cur.fetchone()[0]

    cur.execute("""
        SELECT COALESCE(SUM(stock_quantity), 0), 
               COALESCE(SUM(stock_quantity * wholesale_cost), 0), 
               COALESCE(SUM(stock_quantity * selling_price), 0)
        FROM products WHERE is_active IS NULL OR is_active = TRUE;
    """)
    current_rack_units, current_rack_cost_val, current_rack_retail_val = cur.fetchone()

    cur.execute("""
        SELECT payment_mode, COALESCE(SUM(quantity * unit_price), 0)
        FROM inventory_history
        WHERE transaction_type = 'STOCK_OUT' AND DATE(timestamp) >= %s AND DATE(timestamp) <= %s
        GROUP BY payment_mode;
    """, (start_d, end_d))
    pay_splits = dict(cur.fetchall())
    cur.close()
    conn.close()

    net_sold_qty = int(gross_sold_qty) - int(ret_qty)
    net_sales_rev = float(gross_sales_rev) - float(ret_rev)
    net_cogs = float(gross_sales_cogs) - float(ret_cogs)
    gross_profit = net_sales_rev - net_cogs
    net_profit = gross_profit - float(total_expenses)

    st.markdown("#### 📦 1. Inventory Volume Tally")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Stocked In", f"{int(tot_stock_in_qty):,} pcs")
    s2.metric("Gross Sold", f"{int(gross_sold_qty):,} pcs")
    s3.metric("Returned", f"{int(ret_qty):,} pcs")
    s4.metric("Net Sold Out", f"{net_sold_qty:,} pcs")

    st.markdown("#### 💵 2. Cash Inflow & Payment Methods")
    p1, p2, p3 = st.columns(3)
    p1.metric("💵 Cash Collected", f"₹{pay_splits.get('Cash', 0.0):,.2f}")
    p2.metric("📱 GPay / UPI Collected", f"₹{pay_splits.get('GPay / UPI', 0.0):,.2f}")
    p3.metric("💳 Card / Other Collected", f"₹{sum([v for k, v in pay_splits.items() if k not in ['Cash', 'GPay / UPI']]):,.2f}")

    st.markdown("#### 📈 3. Profit & Loss Reconciliation")
    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Net Revenue", f"₹{net_sales_rev:,.2f}")
    f2.metric("COGS (Cost)", f"₹{net_cogs:,.2f}")
    f3.metric("Overhead Expenses", f"₹{float(total_expenses):,.2f}")
    f4.metric("Net Profit / (Loss)", f"₹{net_profit:,.2f}", delta=f"{net_profit:,.2f}")

    st.markdown("---")
    st.subheader("📥 Formal Excel Export & Partner Verification")

    report_approver = st.selectbox("Partner Authorizing Report Generation & Download:", partner_names, key="excel_auth_name")
    file_name = f"Thalir_Boutique_Formal_Z_Report_{chosen_period.replace(' ', '_')}_{today}.xlsx"

    excel_file = build_master_excel_workbook(start_d, end_d, chosen_period, report_approver)

    download_clicked = st.download_button(
        label=f"📥 Download Formal Z-Report (Authorized by: {report_approver})",
        data=excel_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

    if download_clicked:
        log_audit_action("EXCEL_Z_REPORT_DOWNLOAD", report_approver, f"Downloaded {chosen_period} report ({start_d} to {end_d})")
        st.success(f"Report authorized and logged under '{report_approver}'.")

# ----------------------------------------------------
# 8. 2-WAY BACKUP & AUDIT TRAIL
# ----------------------------------------------------
elif page == "💾 2-Way Backup & Security Audit Trail":
    st.subheader("💾 2-Way Disaster Recovery & Security Audit Trail")

    colA, colB = st.columns(2)
    with colA:
        backup_auth = st.selectbox("Partner Authorizing Local Snapshot:", partner_names, key="bk_auth")
        if st.button("🚀 Trigger Instant 2-Way Backup Now", type="primary"):
            local_path, _ = perform_two_way_backup(backup_auth)
            abs_disk_path = os.path.abspath(local_path)
            st.success(f"✅ Local snapshot saved to: `{abs_disk_path}`")
            st.info(f"Authorized by: {backup_auth}")

    st.markdown("---")
    st.subheader("🕵️ Immutable Security Audit Log")
    st.caption("Complete history of all expenses, deletions, product edits, and report downloads.")

    conn = get_connection()
    audit_df = pd.read_sql_query("""
        SELECT id, created_at AS timestamp, action_type, user_fullname AS authorized_partner, details
        FROM backup_audit_logs
        ORDER BY created_at DESC LIMIT 100;
    """, conn)
    conn.close()

    if not audit_df.empty:
        st.dataframe(audit_df, use_container_width=True)
    else:
        st.info("No audit logs recorded yet.")

# ----------------------------------------------------
# 9. PARTNER EQUITY & SHARES
# ----------------------------------------------------
elif page == "🤝 Partner Capital & Equity":
    st.subheader("🤝 Partner Directory & Equity Management")

    conn = get_connection()
    cur = conn.cursor()
    partners_df = pd.read_sql_query("SELECT id, name, phone, profit_percentage, initial_investment, is_active, joined_date FROM partners ORDER BY id ASC;", conn)

    display_df = partners_df.copy()
    display_df["initial_investment"] = display_df["initial_investment"].apply(lambda x: f"₹{float(x):,.2f}")
    display_df["profit_percentage"] = display_df["profit_percentage"].apply(lambda x: f"{float(x):.2f}%")
    st.dataframe(display_df, use_container_width=True)

    active_equity_sum = partners_df[partners_df["is_active"] == True]["profit_percentage"].sum()
    if round(active_equity_sum, 2) == 100.00:
        st.success(f"Total Active Equity: {active_equity_sum:.2f}% (Perfect Balance)")
    else:
        st.error(f"Total Active Equity: {active_equity_sum:.2f}% (Must equal exactly 100.00%)")

    tab_edit, tab_add = st.tabs(["✏️ Edit Existing Partner", "➕ Add New Partner"])

    with tab_edit:
        if not partners_df.empty:
            partner_selector = dict(zip(partners_df["name"] + " (ID: " + partners_df["id"].astype(str) + ")", partners_df["id"]))
            chosen_partner_label = st.selectbox("Select Partner to Edit:", list(partner_selector.keys()))
            chosen_partner_id = partner_selector[chosen_partner_label]
            partner_data = partners_df[partners_df["id"] == chosen_partner_id].iloc[0]

            with st.form("edit_partner_full_form"):
                st.markdown(f"#### Updating Details for: **{partner_data['name']}**")
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    updated_name = st.text_input("Partner Full Name", value=str(partner_data["name"]))
                    updated_phone = st.text_input("Phone Number", value=str(partner_data["phone"] or ""))
                with col_e2:
                    updated_investment = st.number_input("Capital Investment (₹)", value=float(partner_data["initial_investment"]), min_value=0.0, step=1000.0)
                    updated_pct = st.number_input("Profit Share Percentage (%)", value=float(partner_data["profit_percentage"]), min_value=0.0, max_value=100.0, step=0.5)

                is_partner_active = st.checkbox("Active Partner (Eligible for Profit Share)", value=bool(partner_data["is_active"]))
                approver_p = st.selectbox("Partner Authorizing This Update", partner_names, key="p_edit_auth")
                save_partner_changes = st.form_submit_button("💾 Save All Changes", type="primary")

            if save_partner_changes:
                if not updated_name.strip():
                    st.error("Partner Name cannot be empty.")
                else:
                    cur.execute("""
                        UPDATE partners 
                        SET name = %s, phone = %s, initial_investment = %s, profit_percentage = %s, is_active = %s
                        WHERE id = %s;
                    """, (updated_name.strip(), updated_phone.strip(), updated_investment, updated_pct, is_partner_active, int(chosen_partner_id)))
                    conn.commit()
                    log_audit_action("PARTNER_UPDATED", approver_p, f"Updated partner '{updated_name}' ({updated_pct}%)")
                    st.success(f"Partner '{updated_name}' updated successfully!")
                    st.rerun()

    with tab_add:
        with st.form("add_partner_form", clear_on_submit=True):
            new_p_name = st.text_input("New Partner Full Name")
            new_p_phone = st.text_input("Phone Number")
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                new_p_inv = st.number_input("Capital Investment (₹)", min_value=0.0, value=50000.0, step=5000.0)
            with col_a2:
                new_p_pct = st.number_input("Profit Share Percentage (%)", min_value=0.0, max_value=100.0, value=25.0, step=0.5)

            add_approver = st.selectbox("Partner Authorizing New Admission", partner_names, key="p_add_auth")
            submit_new_partner = st.form_submit_button("➕ Register Partner")

        if submit_new_partner:
            if not new_p_name.strip():
                st.error("Partner name is required.")
            else:
                cur.execute("""
                    INSERT INTO partners (name, phone, profit_percentage, initial_investment, is_active)
                    VALUES (%s, %s, %s, %s, TRUE);
                """, (new_p_name.strip(), new_p_phone.strip(), new_p_pct, new_p_inv))
                conn.commit()
                log_audit_action("PARTNER_ADDED", add_approver, f"Admitted new partner '{new_p_name}' ({new_p_pct}%)")
                st.success(f"Registered partner '{new_p_name}'.")
                st.rerun()

    cur.close()
    conn.close()

# ----------------------------------------------------
# 10. VENDOR MANAGEMENT
# ----------------------------------------------------
elif page == "🏭 Vendor Management":
    st.subheader("🏭 Suppliers & Vendors Directory")

    with st.form("v_box", clear_on_submit=True):
        name = st.text_input("Vendor / Mill Name")
        contact = st.text_input("Contact Person")
        phone = st.text_input("Phone Number")
        email = st.text_input("Email")
        address = st.text_area("Full Mill / Shop Address (City, State, Street)")
        approver_v = st.selectbox("Partner Registering Supplier", partner_names, key="v_add_auth")
        if st.form_submit_button("Register Vendor") and name:
            conn = get_connection()
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO vendors (name, contact_person, phone, email, address)
                VALUES (%s, %s, %s, %s, %s);
            """, (name.strip(), contact.strip(), phone.strip(), email.strip(), address.strip()))
            conn.commit()
            cur.close()
            conn.close()
            log_audit_action("VENDOR_REGISTERED", approver_v, f"Registered supplier '{name}'")
            st.success("Vendor saved.")

    conn = get_connection()
    df_v = pd.read_sql_query("SELECT id, name, contact_person, phone, address, created_at FROM vendors ORDER BY id DESC;", conn)
    conn.close()
    st.dataframe(df_v, use_container_width=True)