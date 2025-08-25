# Synthetic MISA-like enterprise dataset generator
# Creates an Excel workbook with many sheets covering core modules:
# GL, AR/AP, Inventory, Payroll, Fixed Assets, Tax, Budgets, etc.

import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import random
from dateutil.relativedelta import relativedelta

random.seed(42)
np.random.seed(42)

# -------------------- Helpers --------------------
def daterange(start_date, end_date):
    for n in range(int((end_date - start_date).days) + 1):
        yield start_date + timedelta(n)

def month_ends(start_date, end_date):
    cur = datetime(start_date.year, start_date.month, 1)
    while cur <= end_date:
        # month end = last day of month
        next_month = cur + relativedelta(months=1)
        month_end = next_month - timedelta(days=1)
        yield month_end if month_end <= end_date else end_date
        cur = next_month

def choose(seq):
    return seq[random.randrange(len(seq))]

def vn_date(d):
    return d.strftime("%Y-%m-%d")

# -------------------- Parameters --------------------
company_name = "CÔNG TY CP MISA (GIẢ LẬP)"
start = datetime(2024, 1, 1)
end   = datetime(2025, 8, 24)  # current convo date - 1
months = list(month_ends(start, end))

# -------------------- Dimensions --------------------
# Chart of Accounts (minimal, hierarchical for trading company)
coa = pd.DataFrame([
    # code, name, level, type, parent
    ("111", "Tiền mặt", 3, "Asset", "11"),
    ("112", "Tiền gửi ngân hàng", 3, "Asset", "11"),
    ("131", "Phải thu khách hàng", 3, "Asset", "13"),
    ("1331", "Thuế GTGT được khấu trừ", 4, "Asset", "13"),
    ("152", "Nguyên liệu, vật liệu", 3, "Asset", "15"),
    ("156", "Hàng hóa", 3, "Asset", "15"),
    ("211", "Tài sản cố định hữu hình", 3, "Asset", "21"),
    ("214", "Hao mòn lũy kế TSCĐ", 3, "ContraAsset", "21"),
    ("331", "Phải trả người bán", 3, "Liability", "33"),
    ("33311", "Thuế GTGT đầu ra", 5, "Liability", "33"),
    ("334", "Phải trả người lao động", 3, "Liability", "33"),
    ("338", "Phải trả, phải nộp khác", 3, "Liability", "33"),
    ("411", "Vốn đầu tư của chủ sở hữu", 3, "Equity", "41"),
    ("421", "Lợi nhuận sau thuế chưa phân phối", 3, "Equity", "42"),
    ("511", "Doanh thu bán hàng và CCDV", 3, "Revenue", "51"),
    ("515", "Doanh thu hoạt động tài chính", 3, "Revenue", "51"),
    ("521", "Các khoản giảm trừ doanh thu", 3, "ContraRevenue", "52"),
    ("632", "Giá vốn hàng bán", 3, "Expense", "63"),
    ("642", "Chi phí quản lý doanh nghiệp", 3, "Expense", "64"),
    ("635", "Chi phí tài chính", 3, "Expense", "63"),
    ("811", "Chi phí khác", 3, "Expense", "81"),
    ("711", "Thu nhập khác", 3, "Revenue", "71"),
], columns=["account_code","account_name","level","account_type","parent_code"])

# VAT codes
vat_codes = pd.DataFrame([
    ("VAT10", 0.10, "Thuế GTGT 10%"),
    ("VAT8",  0.08, "Thuế GTGT 8% (ưu đãi)"),
    ("VAT5",  0.05, "Thuế GTGT 5%"),
    ("VAT0",  0.00, "Thuế GTGT 0%"),
    ("NON",   0.00, "Không chịu thuế"),
], columns=["vat_code","vat_rate","description"])

# Warehouses
warehouses = pd.DataFrame([
    ("WH01","Kho Hà Nội"),
    ("WH02","Kho TP.HCM")
], columns=["warehouse_code","warehouse_name"])

# Cost Centers & Projects
cost_centers = pd.DataFrame([
    ("CC01","Kinh doanh miền Bắc"),
    ("CC02","Kinh doanh miền Nam"),
    ("CC03","Hành chính - Kế toán"),
], columns=["cost_center","cost_center_name"])

projects = pd.DataFrame([
    ("PRJ01","Triển khai khách hàng Enterprise"),
    ("PRJ02","Nâng cấp hạ tầng AMIS"),
    ("PRJ03","Chiến dịch Marketing Q2"),
], columns=["project_code","project_name"])

# Customers & Vendors
customer_names = [
    "CT TNHH Minh An", "CT CP Đông Á", "CT TNHH Hoa Sen",
    "CT CP ABC Group", "CT TNHH Bình Minh", "CT CP Sao Mai",
    "CT TNHH GreenTech", "CT CP Việt Phát", "CT TNHH Tân Long",
    "CT CP Đại Phú", "CT TNHH Blue Ocean", "CT CP Sunrise"
]
vendor_names = [
    "NCC Thiên Long", "NCC Phương Nam", "NCC KingStar",
    "NCC Việt Á", "NCC Hòa Bình", "NCC An Phú", "NCC Hi-Tech",
    "NCC Hồng Phát", "NCC Vạn Xuân", "NCC Đại Thành"
]
def make_party_df(names, prefix):
    rows = []
    for i, name in enumerate(names, start=1):
        code = f"{prefix}{i:03d}"
        tax = f"0{random.randint(100000000, 999999999)}"
        addr = choose(["Hà Nội","TP.HCM","Đà Nẵng","Hải Phòng","Cần Thơ"])
        rows.append((code, name, tax, addr, choose(["TM/CK","CK"]), choose(["30 ngày","45 ngày","60 ngày"])))
    return pd.DataFrame(rows, columns=[f"{prefix.lower()}_code", "name", "tax_code", "address", "payment_method","payment_terms"])

customers = make_party_df(customer_names, "CUS")
vendors   = make_party_df(vendor_names, "VEN")

# Items (SKUs)
items = []
for i in range(1, 41):  # 40 items
    sku = f"SKU{i:03d}"
    name = f"Phần mềm/Thiết bị {i:03d}"
    uom = choose(["cái","gói","bộ","license"])
    std_cost = round(random.uniform(300000, 5000000), -3)  # VND
    price = round(std_cost * random.uniform(1.15, 1.8), -3)
    vatc = choose(vat_codes.vat_code.tolist())
    items.append((sku, name, uom, std_cost, price, vatc))
items = pd.DataFrame(items, columns=["item_code","item_name","uom","standard_cost","list_price","default_vat_code"])

# Employees
emp_names = ["Nguyễn Văn An","Trần Thị Bình","Phạm Văn Cường","Lê Thị Dung","Hoàng Văn Em",
             "Đỗ Thị Giang","Phan Văn Huy","Vũ Thị Hạnh","Bùi Minh Khoa","Đặng Thị Loan"]
employees = []
for i, name in enumerate(emp_names, start=1):
    emp_code = f"EMP{i:03d}"
    dept = choose(["Kinh doanh","Kế toán","Hành chính","Kỹ thuật","Marketing"])
    base = random.randint(9, 25) * 1_000_000
    employees.append((emp_code, name, dept, base))
employees = pd.DataFrame(employees, columns=["emp_code","full_name","department","base_salary"])

# Opening balances (as of 2024-01-01)
opening_bal = pd.DataFrame([
    ("111", "Số dư tiền mặt", 50_000_000, 0),
    ("112", "Số dư tiền gửi", 2_000_000_000, 0),
    ("156", "Hàng hóa tồn đầu", 800_000_000, 0),
    ("211", "Nguyên giá TSCĐ", 600_000_000, 0),
    ("214", "HM lũy kế TSCĐ", 0, 120_000_000),
    ("331", "Phải trả NCC đầu kỳ", 0, 200_000_000),
    ("131", "Phải thu KH đầu kỳ", 300_000_000, 0),
    ("411", "Vốn chủ sở hữu", 0, 2_530_000_000),
], columns=["account_code","description","debit","credit"])

# -------------------- Transaction Generators --------------------
sales_headers = []
sales_lines = []
purch_headers = []
purch_lines = []
einvoices = []

journal = []  # GL journal entries
receipts = [] # AR receipts
payments = [] # AP payments
bank_tx = []  # Bank transactions
inv_moves = []# Inventory movements (qty & value)
doc_seq = 1000

def new_doc(prefix, dt):
    global doc_seq
    doc_seq += 1
    return f"{prefix}{dt.strftime('%y%m')}-{doc_seq}"

for m_end in months:
    m_start = datetime(m_end.year, m_end.month, 1)
    # volumes
    n_sales = random.randint(12, 22)
    n_purch = random.randint(8, 16)

    # Purchases first (to feed inventory)
    for _ in range(n_purch):
        dt = m_start + timedelta(days=random.randint(0, (m_end - m_start).days))
        ven = vendors.sample(1).iloc[0]
        wh = warehouses.sample(1).iloc[0]
        doc = new_doc("PN", dt)  # Purchase Note
        n_lines = random.randint(1, 3)
        total_net = 0
        total_vat = 0
        lines = []
        for li in range(1, n_lines+1):
            it = items.sample(1).iloc[0]
            qty = random.randint(5, 40)
            price = it["standard_cost"] * random.uniform(0.95, 1.05)  # near cost
            vatc = it["default_vat_code"]
            vat_rate = float(vat_codes.loc[vat_codes.vat_code==vatc,"vat_rate"].iloc[0])
            line_net = round(qty * price, 0)
            line_vat = round(line_net * vat_rate, 0)
            total_net += line_net
            total_vat += line_vat
            lines.append((doc, li, it["item_code"], qty, it["uom"], round(price,0), vatc, line_net, line_vat, wh["warehouse_code"]))
            # Inventory move (receipt)
            inv_moves.append((dt, doc, "IN", it["item_code"], qty, round(price,0), round(line_net,0), wh["warehouse_code"]))
        gross = total_net + total_vat

        purch_headers.append((doc, vn_date(dt), ven["ven_code"], ven["name"], total_net, total_vat, gross, wh["warehouse_code"], ven["payment_terms"]))
        for ln in lines:
            purch_lines.append(ln)

        # Journal: Dr 156 (net) + Dr 1331 (VAT) / Cr 331 (gross)
        journal.append((vn_date(dt), doc, "156", "Nợ", total_net, "Nhập mua hàng", "", "", "CC01", "", ""))
        if total_vat>0:
            journal.append((vn_date(dt), doc, "1331", "Nợ", total_vat, "Thuế GTGT được khấu trừ", "", "", "CC01", "", ""))
        journal.append((vn_date(dt), doc, "331", "Có", gross, f"Phải trả {ven['name']}", "", "", "CC01", "", ""))

        # Randomly pay some purchases
        if random.random() < 0.55:
            pay_date = dt + timedelta(days=random.randint(5, 25))
            if pay_date > end: 
                pay_date = end
            pay_doc = new_doc("PC", pay_date)
            payments.append((pay_doc, vn_date(pay_date), doc, ven["ven_code"], ven["name"], gross, "112"))
            bank_tx.append((vn_date(pay_date), pay_doc, "OUT", ven["name"], gross, f"Thanh toán HĐ {doc}"))
            # Journal: Dr 331 / Cr 112
            journal.append((vn_date(pay_date), pay_doc, "331", "Nợ", gross, f"TT NCC {ven['name']} hóa đơn {doc}", "", "", "CC01", "", ""))
            journal.append((vn_date(pay_date), pay_doc, "112", "Có", gross, "Chi tiền gửi NH", "", "", "CC01", "", ""))

    # Sales
    for _ in range(n_sales):
        dt = m_start + timedelta(days=random.randint(0, (m_end - m_start).days))
        cus = customers.sample(1).iloc[0]
        wh = warehouses.sample(1).iloc[0]
        doc = new_doc("SI", dt)  # Sales Invoice
        n_lines = random.randint(1, 3)
        total_net = 0
        total_vat = 0
        cogs_total = 0
        lines = []
        for li in range(1, n_lines+1):
            it = items.sample(1).iloc[0]
            qty = random.randint(2, 20)
            price = it["list_price"] * random.uniform(0.95, 1.05)
            vatc = it["default_vat_code"]
            vat_rate = float(vat_codes.loc[vat_codes.vat_code==vatc,"vat_rate"].iloc[0])
            line_net = round(qty * price, 0)
            line_vat = round(line_net * vat_rate, 0)
            total_net += line_net
            total_vat += line_vat
            # COGS from standard cost
            cogs = round(qty * it["standard_cost"] * random.uniform(0.98, 1.02), 0)
            cogs_total += cogs
            lines.append((doc, li, it["item_code"], qty, it["uom"], round(price,0), vatc, line_net, line_vat, wh["warehouse_code"], cogs))
            # Inventory move (issue)
            inv_moves.append((dt, doc, "OUT", it["item_code"], -qty, round(it["standard_cost"],0), -cogs, wh["warehouse_code"]))
        gross = total_net + total_vat

        sales_headers.append((doc, vn_date(dt), cus["cus_code"], cus["name"], total_net, total_vat, gross, wh["warehouse_code"], cus["payment_terms"]))
        for ln in lines:
            sales_lines.append(ln)

        # Journal: Dr 131 (gross) / Cr 511 (net), Cr 33311 (VAT)
        journal.append((vn_date(dt), doc, "131", "Nợ", gross, f"Phải thu {cus['name']}", "", "", "CC01", "", ""))
        journal.append((vn_date(dt), doc, "511", "Có", total_net, "Doanh thu bán hàng", "", "", "CC01", "", ""))
        if total_vat>0:
            journal.append((vn_date(dt), doc, "33311", "Có", total_vat, "Thuế GTGT đầu ra", "", "", "CC01", "", ""))
        # COGS: Dr 632 / Cr 156
        journal.append((vn_date(dt), doc, "632", "Nợ", cogs_total, "Giá vốn hàng bán", "", "", "CC01", "", ""))
        journal.append((vn_date(dt), doc, "156", "Có", cogs_total, "Xuất kho hàng bán", "", "", "CC01", "", ""))

        # Randomly receive payment
        if random.random() < 0.6:
            rc_date = dt + timedelta(days=random.randint(5, 30))
            if rc_date > end: 
                rc_date = end
            rc_doc = new_doc("PT", rc_date)
            receipts.append((rc_doc, vn_date(rc_date), doc, cus["cus_code"], cus["name"], gross, "112"))
            bank_tx.append((vn_date(rc_date), rc_doc, "IN", cus["name"], gross, f"Thu tiền HĐ {doc}"))
            # Journal: Dr 112 / Cr 131
            journal.append((vn_date(rc_date), rc_doc, "112", "Nợ", gross, f"Thu KH {cus['name']} hóa đơn {doc}", "", "", "CC01", "", ""))
            journal.append((vn_date(rc_date), rc_doc, "131", "Có", gross, "Giảm phải thu", "", "", "CC01", "", ""))

    # Monthly payroll & depreciation
    payroll_date = m_end
    total_payroll = 0
    for _, row in employees.iterrows():
        gross = round(row.base_salary * random.uniform(1.0, 1.15), -3)
        ins_tax = round(gross * 0.105, -3)  # BH + thuế tạm tính (giả lập)
        net = gross - ins_tax
        total_payroll += gross
    # Journal: Dr 642 / Cr 334 (gross), and payment Cr 112 when paid
    pay_doc = new_doc("PR", payroll_date)
    journal.append((vn_date(payroll_date), pay_doc, "642", "Nợ", total_payroll, "Chi phí lương tháng", "", "", "CC03", "", ""))
    journal.append((vn_date(payroll_date), pay_doc, "334", "Có", total_payroll, "Lương phải trả", "", "", "CC03", "", ""))
    if random.random() < 0.9:
        pay_date = payroll_date + timedelta(days=3)
        if pay_date > end: pay_date = end
        pay_doc2 = new_doc("PC", pay_date)
        payments.append((pay_doc2, vn_date(pay_date), pay_doc, "EMPALL", "Nhân viên", total_payroll, "112"))
        bank_tx.append((vn_date(pay_date), pay_doc2, "OUT", "Trả lương NV", total_payroll, f"Trả lương tháng {payroll_date.strftime('%m/%Y')}"))
        journal.append((vn_date(pay_date), pay_doc2, "334", "Nợ", total_payroll, "Chi trả lương", "", "", "CC03", "", ""))
        journal.append((vn_date(pay_date), pay_doc2, "112", "Có", total_payroll, "Chi tiền gửi NH", "", "", "CC03", "", ""))

    # Depreciation (simple straight-line)
    dep_base = 600_000_000
    useful_months = 60  # 5 years
    monthly_dep = round(dep_base / useful_months, 0)
    dep_doc = new_doc("KH", m_end)
    journal.append((vn_date(m_end), dep_doc, "642", "Nợ", monthly_dep, "Khấu hao TSCĐ tháng", "", "", "CC03", "", ""))
    journal.append((vn_date(m_end), dep_doc, "214", "Có", monthly_dep, "Hao mòn lũy kế", "", "", "CC03", "", ""))

# -------------------- Build DataFrames --------------------
sales_headers_df = pd.DataFrame(sales_headers, columns=[
    "invoice_no","invoice_date","customer_code","customer_name","net_amount","vat_amount","gross_amount","warehouse_code","payment_terms"
])
sales_lines_df = pd.DataFrame(sales_lines, columns=[
    "invoice_no","line_no","item_code","quantity","uom","unit_price","vat_code","line_net","line_vat","warehouse_code","estimated_cogs"
])

purch_headers_df = pd.DataFrame(purch_headers, columns=[
    "bill_no","bill_date","vendor_code","vendor_name","net_amount","vat_amount","gross_amount","warehouse_code","payment_terms"
])
purch_lines_df = pd.DataFrame(purch_lines, columns=[
    "bill_no","line_no","item_code","quantity","uom","unit_cost","vat_code","line_net","line_vat","warehouse_code"
])

receipts_df = pd.DataFrame(receipts, columns=[
    "receipt_no","receipt_date","ref_invoice","customer_code","customer_name","amount","bank_cash_account"
])
payments_df = pd.DataFrame(payments, columns=[
    "payment_no","payment_date","ref_bill_or_doc","vendor_code_or_emp","vendor_or_emp_name","amount","bank_cash_account"
])

inv_moves_df = pd.DataFrame(inv_moves, columns=[
    "move_date","ref_doc","direction","item_code","quantity","unit_cost","value","warehouse_code"
])

bank_tx_df = pd.DataFrame(bank_tx, columns=[
    "txn_date","txn_no","direction","counterparty","amount","description"
])

journal_df = pd.DataFrame(journal, columns=[
    "entry_date","document_no","account_code","drcr","amount","description","partner_code","item_code","cost_center","project_code","notes"
])

# Fixed assets register (simple)
fa_register = pd.DataFrame([
    ("FA001","Máy chủ ứng dụng","211", datetime(2023,1,1), 400_000_000, 60, "SL"),
    ("FA002","Thiết bị lưu trữ NAS","211", datetime(2023,6,1), 200_000_000, 60, "SL")
], columns=["asset_code","asset_name","account_code","acq_date","acq_cost","useful_life_months","method"])
fa_register["acq_date"] = fa_register["acq_date"].dt.strftime("%Y-%m-%d")

# -------------------- Derived Reports --------------------
# Trial Balance by month
journal_df["month"] = pd.to_datetime(journal_df["entry_date"]).dt.to_period('M').astype(str)
tb_rows = []
for (m, acc), g in journal_df.groupby(["month","account_code"]):
    debit = g.loc[g.drcr=="Nợ","amount"].sum()
    credit = g.loc[g.drcr=="Có","amount"].sum()
    tb_rows.append((m, acc, debit, credit))
trial_balance = pd.DataFrame(tb_rows, columns=["month","account_code","debit","credit"])
trial_balance = trial_balance.merge(coa[["account_code","account_name","account_type"]], on="account_code", how="left")

# AR/AP Aging as of end
def build_aging(ar=True):
    if ar:
        hdr = sales_headers_df.copy()
        hdr["partner_code"] = hdr["customer_code"]
        hdr["partner_name"] = hdr["customer_name"]
        hdr["doc_no"] = hdr["invoice_no"]
        hdr["doc_date"] = pd.to_datetime(hdr["invoice_date"])
        pay = receipts_df.copy()
        pay["doc_date"] = pd.to_datetime(pay["receipt_date"])
        pay = pay[["ref_invoice","amount","doc_date"]]
        hdr["paid"] = 0.0
        # Sum payments per invoice
        paid_map = pay.groupby("ref_invoice")["amount"].sum().to_dict()
        hdr["paid"] = hdr["invoice_no"].map(paid_map).fillna(0.0)
        hdr["outstanding"] = hdr["gross_amount"] - hdr["paid"]
    else:
        hdr = purch_headers_df.copy()
        hdr["partner_code"] = hdr["vendor_code"]
        hdr["partner_name"] = hdr["vendor_name"]
        hdr["doc_no"] = hdr["bill_no"]
        hdr["doc_date"] = pd.to_datetime(hdr["bill_date"])
        pay = payments_df.copy()
        # payments ref may include non-bill docs (e.g., payroll); filter those starting with PN (purchase)
        pay = pay[pay["ref_bill_or_doc"].str.startswith("PN", na=False)]
        pay["doc_date"] = pd.to_datetime(pay["payment_date"])
        pay = pay[["ref_bill_or_doc","amount","doc_date"]]
        hdr["paid"] = 0.0
        paid_map = pay.groupby("ref_bill_or_doc")["amount"].sum().to_dict()
        hdr["paid"] = hdr["bill_no"].map(paid_map).fillna(0.0)
        hdr["outstanding"] = hdr["gross_amount"] - hdr["paid"]

    today = pd.to_datetime(end.strftime("%Y-%m-%d"))
    hdr["days"] = (today - hdr["doc_date"]).dt.days
    bins = [-1, 30, 60, 90, 180, 365, 10_000]
    labels = ["0-30","31-60","61-90","91-180","181-365",">365"]
    hdr["aging_bucket"] = pd.cut(hdr["days"], bins=bins, labels=labels)
    aging = hdr.groupby(["partner_code","partner_name","aging_bucket"])["outstanding"].sum().reset_index()
    pivot = aging.pivot_table(index=["partner_code","partner_name"], columns="aging_bucket", values="outstanding", fill_value=0).reset_index()
    pivot.columns.name = None
    return pivot

ar_aging = build_aging(ar=True)
ap_aging = build_aging(ar=False)

# VAT monthly report
def vat_report():
    s = sales_lines_df.merge(sales_headers_df[["invoice_no","invoice_date"]], on="invoice_no", how="left")
    p = purch_lines_df.merge(purch_headers_df[["bill_no","bill_date"]], on="bill_no", how="left")
    s["month"] = pd.to_datetime(s["invoice_date"]).dt.to_period('M').astype(str)
    p["month"] = pd.to_datetime(p["bill_date"]).dt.to_period('M').astype(str)
    s_out = s.groupby("month")[["line_net","line_vat"]].sum().reset_index().rename(columns={"line_net":"sales_net","line_vat":"vat_output"})
    p_in  = p.groupby("month")[["line_net","line_vat"]].sum().reset_index().rename(columns={"line_net":"purch_net","line_vat":"vat_input"})
    v = s_out.merge(p_in, on="month", how="outer").fillna(0)
    v["vat_payable"] = v["vat_output"] - v["vat_input"]
    return v

vat_monthly = vat_report()

# Budgets (simple annual budgets per account type)
budget_rows = []
for year in [2024, 2025]:
    budget_rows.append((year, "511", "Doanh thu", 50_000_000_000 if year==2025 else 35_000_000_000))
    budget_rows.append((year, "632", "Giá vốn",   32_000_000_000 if year==2025 else 23_000_000_000))
    budget_rows.append((year, "642", "Chi phí QLDN", 6_000_000_000 if year==2025 else 4_000_000_000))
budgets = pd.DataFrame(budget_rows, columns=["year","account_code","label","amount_vnd"])

# README & Data Dictionary
readme_text = [
    f"Doanh nghiệp giả lập: {company_name}",
    f"Chu kỳ dữ liệu: {start.strftime('%Y-%m-%d')} đến {end.strftime('%Y-%m-%d')}",
    "Tất cả dữ liệu là NGẪU NHIÊN (synthetic) để học và thực hành phân tích. Không phản ánh dữ liệu thật.",
    "",
    "Các sheet chính:",
    "- COA: Hệ thống tài khoản",
    "- VAT_CODES: Danh mục thuế GTGT",
    "- CUSTOMERS, VENDORS, ITEMS, WAREHOUSES, COST_CENTERS, PROJECTS, EMPLOYEES",
    "- OPENING_BALANCES",
    "- SALES_HEADERS, SALES_LINES, PURCHASE_HEADERS, PURCHASE_LINES",
    "- RECEIPTS (thu tiền), PAYMENTS (chi tiền), BANK_TX (sao kê ngân hàng)",
    "- INVENTORY_MOVES (nhập/xuất kho)",
    "- JOURNAL (bút toán tổng hợp)",
    "- TRIAL_BALANCE (Bảng cân đối số phát sinh theo tháng)",
    "- AR_AGING, AP_AGING (tuổi nợ phải thu/phải trả)",
    "- VAT_MONTHLY (kê khai thuế GTGT theo tháng)",
    "- FA_REGISTER (danh mục TSCĐ)",
    "- BUDGETS (ngân sách năm)",
    "",
    "Quan hệ logic:",
    "* Hóa đơn bán (SALES_HEADERS/LINES) phát sinh: Dr 131 / Cr 511, Cr 33311; đồng thời Dr 632 / Cr 156.",
    "* Hóa đơn mua (PURCHASE_*) phát sinh: Dr 156, Dr 1331 / Cr 331.",
    "* Thu/Chi tiền cập nhật 112, đối ứng 131/331/334...",
    "* Khấu hao: Dr 642 / Cr 214.",
    "",
    "Gợi ý phân tích:",
    "- Doanh thu, biên lợi nhuận gộp theo tháng/sản phẩm/khách hàng.",
    "- Vòng quay hàng tồn kho, tồn kho an toàn theo SKU.",
    "- Tuổi nợ phải thu/phải trả và dòng tiền.",
    "- Đối chiếu VAT đầu ra/đầu vào; so ngân sách vs thực tế.",
]

readme_df = pd.DataFrame({"README": readme_text})

data_dict = pd.DataFrame([
    ("COA","account_code","Mã tài khoản","text"),
    ("COA","account_type","Loại (Asset/Liab/Revenue/Expense)","text"),
    ("SALES_HEADERS","invoice_no","Số hóa đơn bán","text"),
    ("SALES_HEADERS","gross_amount","Tổng tiền thanh toán (có VAT)","number"),
    ("SALES_LINES","item_code","Mã hàng hóa/dịch vụ","text"),
    ("PURCHASE_HEADERS","bill_no","Số hóa đơn mua","text"),
    ("PURCHASE_LINES","unit_cost","Đơn giá mua","number"),
    ("JOURNAL","drcr","Nợ/Có","text"),
    ("JOURNAL","amount","Số tiền bút toán","number"),
    ("INVENTORY_MOVES","direction","IN (nhập)/OUT (xuất)","text"),
    ("VAT_MONTHLY","vat_payable","VAT phải nộp (= đầu ra - đầu vào)","number"),
    ("AR_AGING","0-30..>365","Dư nợ theo nhóm tuổi","number"),
], columns=["table","column","definition_vn","type"])

# -------------------- Write Excel --------------------
path = "/mnt/data/misa_simulated_dataset_v1.xlsx"
with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
    coa.to_excel(writer, sheet_name="COA", index=False)
    vat_codes.to_excel(writer, sheet_name="VAT_CODES", index=False)
    warehouses.to_excel(writer, sheet_name="WAREHOUSES", index=False)
    cost_centers.to_excel(writer, sheet_name="COST_CENTERS", index=False)
    projects.to_excel(writer, sheet_name="PROJECTS", index=False)
    customers.to_excel(writer, sheet_name="CUSTOMERS", index=False)
    vendors.to_excel(writer, sheet_name="VENDORS", index=False)
    items.to_excel(writer, sheet_name="ITEMS", index=False)
    employees.to_excel(writer, sheet_name="EMPLOYEES", index=False)
    opening_bal.to_excel(writer, sheet_name="OPENING_BALANCES", index=False)

    sales_headers_df.to_excel(writer, sheet_name="SALES_HEADERS", index=False)
    sales_lines_df.to_excel(writer, sheet_name="SALES_LINES", index=False)
    purch_headers_df.to_excel(writer, sheet_name="PURCHASE_HEADERS", index=False)
    purch_lines_df.to_excel(writer, sheet_name="PURCHASE_LINES", index=False)

    receipts_df.to_excel(writer, sheet_name="RECEIPTS", index=False)
    payments_df.to_excel(writer, sheet_name="PAYMENTS", index=False)
    bank_tx_df.to_excel(writer, sheet_name="BANK_TX", index=False)

    inv_moves_df.to_excel(writer, sheet_name="INVENTORY_MOVES", index=False)
    journal_df.to_excel(writer, sheet_name="JOURNAL", index=False)

    trial_balance.to_excel(writer, sheet_name="TRIAL_BALANCE", index=False)
    ar_aging.to_excel(writer, sheet_name="AR_AGING", index=False)
    ap_aging.to_excel(writer, sheet_name="AP_AGING", index=False)
    vat_monthly.to_excel(writer, sheet_name="VAT_MONTHLY", index=False)

    fa_register.to_excel(writer, sheet_name="FA_REGISTER", index=False)
    budgets.to_excel(writer, sheet_name="BUDGETS", index=False)

    readme_df.to_excel(writer, sheet_name="README", index=False)
    data_dict.to_excel(writer, sheet_name="DATA_DICTIONARY", index=False)

# Show a small preview to the user
preview = pd.DataFrame({
    "Sheet": ["COA","CUSTOMERS","ITEMS","SALES_HEADERS","PURCHASE_HEADERS","JOURNAL","TRIAL_BALANCE","AR_AGING","VAT_MONTHLY"],
    "Rows": [len(coa), len(customers), len(items), len(sales_headers_df), len(purch_headers_df), len(journal_df), len(trial_balance), len(ar_aging), len(vat_monthly)]
})

from caas_jupyter_tools import display_dataframe_to_user
display_dataframe_to_user("Tổng quan sheet & số dòng", preview)

path
