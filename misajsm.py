# Re-create the synthetic MISA dataset and confirm the file exists, then show a quick preview.
import numpy as np
import pandas as pd
import random, os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from caas_jupyter_tools import display_dataframe_to_user

random.seed(7); np.random.seed(7)

def month_ends(start, end):
    cur = datetime(start.year, start.month, 1)
    while cur <= end:
        nm = cur + relativedelta(months=1)
        me = nm - timedelta(days=1)
        yield me if me <= end else end
        cur = nm

def choose(seq): 
    return seq[random.randrange(len(seq))]

def vn_date(d):
    return d.strftime("%Y-%m-%d")

# Parameters
start = datetime(2024,1,1)
end   = datetime(2025,8,24)

# Dimensions
coa = pd.DataFrame([
    ("111","Tiền mặt",3,"Asset","11"),("112","Tiền gửi ngân hàng",3,"Asset","11"),
    ("131","Phải thu KH",3,"Asset","13"),("1331","Thuế GTGT được khấu trừ",4,"Asset","13"),
    ("152","Nguyên liệu",3,"Asset","15"),("156","Hàng hóa",3,"Asset","15"),
    ("211","TSCĐ hữu hình",3,"Asset","21"),("214","HM lũy kế TSCĐ",3,"ContraAsset","21"),
    ("331","Phải trả NCC",3,"Liability","33"),("33311","Thuế GTGT đầu ra",5,"Liability","33"),
    ("334","Phải trả NLĐ",3,"Liability","33"),
    ("411","Vốn CSH",3,"Equity","41"),("421","LNSTCPP",3,"Equity","42"),
    ("511","Doanh thu BH&CCDV",3,"Revenue","51"),("521","Giảm trừ DT",3,"ContraRevenue","52"),
    ("632","Giá vốn hàng bán",3,"Expense","63"),("642","CP QLDN",3,"Expense","64"),
], columns=["account_code","account_name","level","account_type","parent_code"])

vat_codes = pd.DataFrame([("VAT10",0.10,"GTGT 10%"),("VAT8",0.08,"GTGT 8%"),("VAT5",0.05,"GTGT 5%"),("VAT0",0.0,"0%"),("NON",0.0,"Không chịu thuế")],
                         columns=["vat_code","vat_rate","description"])

warehouses = pd.DataFrame([("WH01","Kho Hà Nội"),("WH02","Kho TP.HCM")], columns=["warehouse_code","warehouse_name"])

customers = pd.DataFrame([
    ("CUS001","CT TNHH Minh An","0123456789","Hà Nội","CK","30 ngày"),
    ("CUS002","CT CP Đông Á","0123456790","TP.HCM","CK","45 ngày"),
    ("CUS003","CT TNHH Hoa Sen","0123456791","Đà Nẵng","TM/CK","30 ngày"),
    ("CUS004","CT CP ABC Group","0123456792","Hải Phòng","CK","60 ngày"),
], columns=["cus_code","name","tax_code","address","payment_method","payment_terms"])

vendors = pd.DataFrame([
    ("VEN001","NCC Thiên Long","0234567890","Hà Nội","CK","30 ngày"),
    ("VEN002","NCC Phương Nam","0234567891","TP.HCM","CK","45 ngày"),
    ("VEN003","NCC KingStar","0234567892","Đà Nẵng","TM/CK","30 ngày"),
], columns=["ven_code","name","tax_code","address","payment_method","payment_terms"])

items=[]
for i in range(1,13):
    std_cost=round(random.uniform(300000, 5000000),-3)
    price=round(std_cost*random.uniform(1.15,1.8),-3)
    items.append((f"SKU{i:03d}", f"Sản phẩm {i:03d}", choose(["cái","gói","bộ","license"]), std_cost, price, choose(vat_codes.vat_code.tolist())))
items = pd.DataFrame(items, columns=["item_code","item_name","uom","standard_cost","list_price","default_vat_code"])

# Transactions (smaller sample to ensure quick generation)
months = list(month_ends(start, end))
sales_headers=[]; sales_lines=[]; purch_headers=[]; purch_lines=[]; journal=[]

doc_seq=2000
def new_doc(prefix, dt):
    global doc_seq; doc_seq+=1; return f"{prefix}{dt.strftime('%y%m')}-{doc_seq}"

for m_end in months:
    m_start = datetime(m_end.year, m_end.month, 1)

    # Purchases: 5 per month
    for _ in range(5):
        dt = m_start + timedelta(days=random.randint(0,(m_end-m_start).days))
        ven = vendors.sample(1).iloc[0]; wh = warehouses.sample(1).iloc[0]; doc = new_doc("PN", dt)
        total_net=0; total_vat=0; lines=[]
        for li in range(1,3):
            it = items.sample(1).iloc[0]; qty=random.randint(5,20); price=it["standard_cost"]*random.uniform(0.95,1.05)
            vatc=it["default_vat_code"]; vat_rate=float(vat_codes.loc[vat_codes.vat_code==vatc,"vat_rate"].iloc[0])
            net=round(qty*price,0); vat=round(net*vat_rate,0); total_net+=net; total_vat+=vat
            lines.append((doc,li,it["item_code"],qty,it["uom"],round(price,0),vatc,net,vat,wh["warehouse_code"]))
        purch_headers.append((doc, vn_date(dt), ven["ven_code"], ven["name"], total_net, total_vat, total_net+total_vat, wh["warehouse_code"], ven["payment_terms"]))
        purch_lines.extend(lines)
        journal.extend([(vn_date(dt),doc,"156","Nợ",total_net,"Nhập mua hàng","","","CC01","",""),
                        (vn_date(dt),doc,"1331","Nợ",total_vat,"Thuế GTGT được khấu trừ","","","CC01","",""),
                        (vn_date(dt),doc,"331","Có",total_net+total_vat,f"Phải trả {ven['name']}","","","CC01","","")])

    # Sales: 7 per month
    for _ in range(7):
        dt = m_start + timedelta(days=random.randint(0,(m_end-m_start).days))
        cus = customers.sample(1).iloc[0]; wh = warehouses.sample(1).iloc[0]; doc = new_doc("SI", dt)
        total_net=0; total_vat=0; lines=[]
        for li in range(1,3):
            it = items.sample(1).iloc[0]; qty=random.randint(2,12); price=it["list_price"]*random.uniform(0.95,1.05)
            vatc=it["default_vat_code"]; vat_rate=float(vat_codes.loc[vat_codes.vat_code==vatc,"vat_rate"].iloc[0])
            net=round(qty*price,0); vat=round(net*vat_rate,0); total_net+=net; total_vat+=vat
            lines.append((doc,li,it["item_code"],qty,it["uom"],round(price,0),vatc,net,vat,wh["warehouse_code"]))
        gross=total_net+total_vat
        sales_headers.append((doc, vn_date(dt), cus["cus_code"], cus["name"], total_net, total_vat, gross, wh["warehouse_code"], cus["payment_terms"]))
        sales_lines.extend(lines)
        journal.extend([(vn_date(dt),doc,"131","Nợ",gross,f"Phải thu {cus['name']}","","","CC01","",""),
                        (vn_date(dt),doc,"511","Có",total_net,"Doanh thu bán hàng","","","CC01","",""),
                        (vn_date(dt),doc,"33311","Có",total_vat,"Thuế GTGT đầu ra","","","CC01","","")])

# Build core tables
sales_headers_df = pd.DataFrame(sales_headers, columns=["invoice_no","invoice_date","customer_code","customer_name","net_amount","vat_amount","gross_amount","warehouse_code","payment_terms"])
sales_lines_df   = pd.DataFrame(sales_lines,   columns=["invoice_no","line_no","item_code","quantity","uom","unit_price","vat_code","line_net","line_vat","warehouse_code"])
purch_headers_df = pd.DataFrame(purch_headers, columns=["bill_no","bill_date","vendor_code","vendor_name","net_amount","vat_amount","gross_amount","warehouse_code","payment_terms"])
purch_lines_df   = pd.DataFrame(purch_lines,   columns=["bill_no","line_no","item_code","quantity","uom","unit_cost","vat_code","line_net","line_vat","warehouse_code"])

trial = []
j = pd.DataFrame(journal, columns=["entry_date","document_no","account_code","drcr","amount","description","partner_code","item_code","cost_center","project_code","notes"])
j["month"]=pd.to_datetime(j["entry_date"]).dt.to_period("M").astype(str)
for (m,acc), g in j.groupby(["month","account_code"]):
    debit=g.loc[g.drcr=="Nợ","amount"].sum(); credit=g.loc[g.drcr=="Có","amount"].sum()
    trial.append((m,acc,debit,credit))
trial_balance = pd.DataFrame(trial, columns=["month","account_code","debit","credit"]).merge(coa[["account_code","account_name","account_type"]], on="account_code", how="left")

# Write Excel
path="/mnt/data/misa_simulated_dataset_quick.xlsx"
with pd.ExcelWriter(path, engine="xlsxwriter") as w:
    coa.to_excel(w, sheet_name="COA", index=False)
    vat_codes.to_excel(w, sheet_name="VAT_CODES", index=False)
    warehouses.to_excel(w, sheet_name="WAREHOUSES", index=False)
    customers.to_excel(w, sheet_name="CUSTOMERS", index=False)
    vendors.to_excel(w, sheet_name="VENDORS", index=False)
    items.to_excel(w, sheet_name="ITEMS", index=False)
    sales_headers_df.to_excel(w, sheet_name="SALES_HEADERS", index=False)
    sales_lines_df.to_excel(w, sheet_name="SALES_LINES", index=False)
    purch_headers_df.to_excel(w, sheet_name="PURCHASE_HEADERS", index=False)
    purch_lines_df.to_excel(w, sheet_name="PURCHASE_LINES", index=False)
    j.to_excel(w, sheet_name="JOURNAL", index=False)
    trial_balance.to_excel(w, sheet_name="TRIAL_BALANCE", index=False)

exists = os.path.exists(path)
display_dataframe_to_user("Preview số sheet & dòng", pd.DataFrame({
    "Sheet":["COA","CUSTOMERS","ITEMS","SALES_HEADERS","PURCHASE_HEADERS","JOURNAL","TRIAL_BALANCE"],
    "Rows":[len(coa),len(customers),len(items),len(sales_headers_df),len(purch_headers_df),len(j),len(trial_balance)]
}))
print("FILE_PATH:", path, "EXISTS:", exists)
