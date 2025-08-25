# Run an in-session visualization suite for the MISA dataset.
# It will load the existing dataset, render charts inline (matplotlib, one per figure),
# and also save all images + an HTML gallery + a ZIP bundle for download.

import os, json, zipfile
from datetime import datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from caas_jupyter_tools import display_dataframe_to_user

# 1) Detect dataset
data_path = None
for p in ["/mnt/data/misa_simulated_dataset_v1.xlsx","/mnt/data/misa_simulated_dataset_quick.xlsx"]:
    if os.path.exists(p):
        data_path = p
        break

if data_path is None:
    raise FileNotFoundError("Không tìm thấy file dữ liệu. Vui lòng tạo dataset trước.")

xls = pd.ExcelFile(data_path)
sheets = {s: pd.read_excel(data_path, sheet_name=s) for s in xls.sheet_names}

# 2) Helper functions
def to_month(df, col):
    d = df.copy()
    d[col] = pd.to_datetime(d[col], errors="coerce")
    d = d.dropna(subset=[col])
    d["month"] = d[col].dt.to_period("M").astype(str)
    return d

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)
    return path

out_dir = ensure_dir("/mnt/data/misa_visuals_runtime")

def save_fig_png(title_slug):
    fname = f"{title_slug}.png"
    fpath = os.path.join(out_dir, fname)
    plt.tight_layout()
    plt.savefig(fpath, dpi=140, bbox_inches="tight")
    return fpath

# 3) KPI table (if possible)
kpi_month = None
if "SALES_HEADERS" in sheets:
    sh = to_month(sheets["SALES_HEADERS"], "invoice_date")
    sales_month = sh.groupby("month")[["net_amount","vat_amount","gross_amount"]].sum().reset_index()
else:
    sales_month = pd.DataFrame(columns=["month","net_amount","vat_amount","gross_amount"])

if "PURCHASE_HEADERS" in sheets:
    ph = to_month(sheets["PURCHASE_HEADERS"], "bill_date")
    purch_month = ph.groupby("month")[["net_amount","vat_amount","gross_amount"]].sum().reset_index()
else:
    purch_month = pd.DataFrame(columns=["month","net_amount","vat_amount","gross_amount"])

if all(k in sheets for k in ["SALES_LINES","PURCHASE_LINES","SALES_HEADERS","PURCHASE_HEADERS"]):
    sl = sheets["SALES_LINES"].merge(sheets["SALES_HEADERS"][["invoice_no","invoice_date"]], on="invoice_no", how="left")
    pl = sheets["PURCHASE_LINES"].merge(sheets["PURCHASE_HEADERS"][["bill_no","bill_date"]], on="bill_no", how="left")
    sl["month"] = pd.to_datetime(sl["invoice_date"], errors="coerce").dt.to_period("M").astype(str)
    pl["month"] = pd.to_datetime(pl["bill_date"], errors="coerce").dt.to_period("M").astype(str)
    vat_out = sl.groupby("month")["line_vat"].sum().reset_index(name="vat_output")
    vat_in  = pl.groupby("month")["line_vat"].sum().reset_index(name="vat_input")
    vat_tbl = vat_out.merge(vat_in, on="month", how="outer").fillna(0.0)
    vat_tbl["vat_payable"] = vat_tbl["vat_output"] - vat_tbl["vat_input"]
else:
    vat_tbl = pd.DataFrame(columns=["month","vat_output","vat_input","vat_payable"])

kpi_month = sales_month[["month","gross_amount"]].rename(columns={"gross_amount":"sales_gross"}).merge(
    purch_month[["month","gross_amount"]].rename(columns={"gross_amount":"purch_gross"}),
    on="month", how="outer"
).merge(vat_tbl[["month","vat_payable"]], on="month", how="outer").fillna(0)

kpi_month = kpi_month.sort_values("month")

display_dataframe_to_user("Bảng KPI theo tháng (Sales/Purchases/VAT)", kpi_month)

# 4) Render charts inline (and save PNGs)
chart_files = []

# Sales gross by month
if not sales_month.empty:
    plt.figure()
    plt.plot(sales_month["month"], sales_month["gross_amount"], marker="o")
    plt.title("Doanh thu (Gross) theo tháng")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(rotation=45); plt.show()
    chart_files.append(save_fig_png("01_sales_gross_by_month"))

    # Invoice count
    inv_count = sh.groupby("month")["invoice_no"].count().reset_index(name="invoice_count")
    plt.figure()
    plt.bar(inv_count["month"], inv_count["invoice_count"])
    plt.title("Số lượng hóa đơn bán theo tháng")
    plt.xlabel("Tháng"); plt.ylabel("Số hóa đơn"); plt.xticks(rotation=45); plt.show()
    chart_files.append(save_fig_png("02_invoice_count_by_month"))

    # Avg invoice
    avg_val = sales_month[["month","gross_amount"]].merge(inv_count, on="month", how="left")
    avg_val["avg_invoice"] = avg_val["gross_amount"] / avg_val["invoice_count"].replace(0, np.nan)
    plt.figure()
    plt.plot(avg_val["month"], avg_val["avg_invoice"], marker="o")
    plt.title("Giá trị HĐ bán trung bình theo tháng")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(rotation=45); plt.show()
    chart_files.append(save_fig_png("03_avg_invoice_value"))

    # Top customers
    top_cus = sh.groupby(["customer_code","customer_name"])["gross_amount"].sum().reset_index().sort_values("gross_amount", ascending=False).head(10)
    plt.figure()
    plt.barh(top_cus["customer_name"], top_cus["gross_amount"])
    plt.title("Top 10 khách hàng theo doanh thu")
    plt.xlabel("VND"); plt.gca().invert_yaxis(); plt.show()
    chart_files.append(save_fig_png("04_top_customers_by_revenue"))

# Top items (if lines)
if "SALES_LINES" in sheets:
    sl = sheets["SALES_LINES"].copy()
    sl_top = sl.groupby("item_code")["line_net"].sum().reset_index().sort_values("line_net", ascending=False).head(10)
    labels = sl_top["item_code"]
    if "ITEMS" in sheets:
        sl_top = sl_top.merge(sheets["ITEMS"][["item_code","item_name"]], on="item_code", how="left")
        labels = sl_top["item_name"].fillna(sl_top["item_code"])
    plt.figure()
    plt.barh(labels, sl_top["line_net"])
    plt.title("Top 10 mặt hàng theo doanh thu (Net)")
    plt.xlabel("VND"); plt.gca().invert_yaxis(); plt.show()
    chart_files.append(save_fig_png("05_top_items_by_revenue_net"))

# Purchases gross by month
if not purch_month.empty:
    plt.figure()
    plt.bar(purch_month["month"], purch_month["gross_amount"])
    plt.title("Giá trị mua hàng (Gross) theo tháng")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(rotation=45); plt.show()
    chart_files.append(save_fig_png("06_purchases_gross_by_month"))

# Sales vs Purchases comparison
if not sales_month.empty and not purch_month.empty:
    comb = kpi_month.copy()
    plt.figure()
    x = np.arange(len(comb["month"])); width = 0.35
    plt.bar(x - width/2, comb["sales_gross"], width, label="Bán")
    plt.bar(x + width/2, comb["purch_gross"], width, label="Mua")
    plt.title("So sánh Bán vs Mua theo tháng (Gross)")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(x, comb["month"], rotation=45); plt.legend(); plt.show()
    chart_files.append(save_fig_png("07_sales_vs_purchases_gross"))

# VAT charts (if available)
if not vat_tbl.empty:
    v = vat_tbl.sort_values("month")
    plt.figure()
    plt.plot(v["month"], v["vat_output"], marker="o", label="Đầu ra")
    plt.plot(v["month"], v["vat_input"], marker="o", label="Đầu vào")
    plt.plot(v["month"], v["vat_payable"], marker="o", label="Phải nộp")
    plt.title("VAT đầu ra/đầu vào & phải nộp theo tháng")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(rotation=45); plt.legend(); plt.show()
    chart_files.append(save_fig_png("08_vat_output_input_payable"))

# Aging histograms (approx)
today = pd.Timestamp.today().normalize()
if "SALES_HEADERS" in sheets:
    sh_days = pd.to_datetime(sheets["SALES_HEADERS"]["invoice_date"], errors="coerce").dropna()
    days_since = (today - sh_days).dt.days
    plt.figure()
    plt.hist(days_since, bins=20)
    plt.title("Phân phối số ngày kể từ ngày xuất HĐ bán")
    plt.xlabel("Số ngày"); plt.ylabel("Số HĐ"); plt.show()
    chart_files.append(save_fig_png("09_hist_days_since_sales_invoice"))

if "PURCHASE_HEADERS" in sheets:
    ph_days = pd.to_datetime(sheets["PURCHASE_HEADERS"]["bill_date"], errors="coerce").dropna()
    days_since = (today - ph_days).dt.days
    plt.figure()
    plt.hist(days_since, bins=20)
    plt.title("Phân phối số ngày kể từ ngày HĐ mua")
    plt.xlabel("Số ngày"); plt.ylabel("Số HĐ"); plt.show()
    chart_files.append(save_fig_png("10_hist_days_since_purchase_bill"))

# GL totals by month
if "JOURNAL" in sheets:
    j = sheets["JOURNAL"].copy()
    j["entry_date"] = pd.to_datetime(j["entry_date"], errors="coerce")
    j = j.dropna(subset=["entry_date"])
    j["month"] = j["entry_date"].dt.to_period("M").astype(str)
    tot = j.groupby(["month","drcr"])["amount"].sum().reset_index()
    pivot = tot.pivot(index="month", columns="drcr", values="amount").fillna(0)
    plt.figure()
    if "Nợ" in pivot.columns:
        plt.plot(pivot.index, pivot["Nợ"], marker="o", label="Tổng Nợ")
    if "Có" in pivot.columns:
        plt.plot(pivot.index, pivot["Có"], marker="o", label="Tổng Có")
    plt.title("Tổng phát sinh Nợ/Có theo tháng (GL)")
    plt.xlabel("Tháng"); plt.ylabel("VND"); plt.xticks(rotation=45); plt.legend(); plt.show()
    chart_files.append(save_fig_png("11_gl_total_debit_credit"))

    acc = j.groupby("account_code")["amount"].sum().reset_index().sort_values("amount", ascending=False).head(20)
    plt.figure()
    plt.barh(acc["account_code"], acc["amount"])
    plt.title("Top 20 tài khoản theo tổng phát sinh")
    plt.xlabel("VND"); plt.gca().invert_yaxis(); plt.show()
    chart_files.append(save_fig_png("12_gl_top_accounts_total_turnover"))

# Trial balance net
if "TRIAL_BALANCE" in sheets:
    tb = sheets["TRIAL_BALANCE"].copy()
    tb_sum = tb.groupby(["account_code","account_name"])[["debit","credit"]].sum().reset_index()
    tb_sum["net"] = tb_sum["debit"] - tb_sum["credit"]
    top_abs = tb_sum.reindex(tb_sum["net"].abs().sort_values(ascending=False).index).head(20)
    plt.figure()
    plt.barh(top_abs["account_code"], top_abs["net"])
    plt.title("Trial Balance: Top 20 tài khoản theo số dư ròng (Nợ–Có)")
    plt.xlabel("VND"); plt.gca().invert_yaxis(); plt.show()
    chart_files.append(save_fig_png("13_trial_balance_top_net"))

# 5) Build HTML gallery and ZIP
html_path = "/mnt/data/misa_dashboard_runtime.html"
cards = []
for f in chart_files:
    cards.append(f"""
    <div style="border:1px solid #ddd;padding:12px;margin:12px;border-radius:8px">
      <img src="{os.path.relpath(f, start='/mnt/data')}" style="max-width:100%"/>
    </div>
    """)
html = f"""<!doctype html>
<html><head><meta charset="utf-8"><title>MISA Runtime Visualization</title></head>
<body style="font-family: Arial, sans-serif; max-width: 1100px; margin: 0 auto; padding: 20px;">
<h1>MISA Runtime Visualization</h1>
<p>Nguồn dữ liệu: <b>{os.path.basename(data_path)}</b></p>
{"".join(cards)}
</body></html>
"""
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html)

zip_path = "/mnt/data/misa_visuals_runtime.zip"
with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(html_path, arcname=os.path.basename(html_path))
    for f in chart_files:
        z.write(f, arcname=os.path.join("misa_visuals_runtime", os.path.basename(f)))

print("DATASET_USED:", data_path)
print("GALLERY_HTML:", html_path)
print("ZIP_BUNDLE:", zip_path)
print("NUM_CHARTS:", len(chart_files))
