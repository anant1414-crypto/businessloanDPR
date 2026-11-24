import io
import json
import sqlite3
import uuid
import pandas as pd
import razorpay
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------- 0. PAGE CONFIG & RAZORPAY / DB HELPERS ----------------
st.set_page_config(page_title="CM-YUVA & MUDRA DPR Generator", layout="wide")
DB_PATH = "orders.db"
APP_URL = st.secrets.get("APP_URL", "http://localhost:8501")

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                order_id TEXT PRIMARY KEY,
                scheme TEXT,
                payload_json TEXT,
                amount_inr INTEGER,
                plink_id TEXT,
                status TEXT DEFAULT 'PENDING'
            )
        """)

def get_rzp_client():
    return razorpay.Client(
        auth=(st.secrets["RAZORPAY_KEY_ID"], st.secrets["RAZORPAY_KEY_SECRET"])
    )

def create_razorpay_order(scheme: str, payload: dict, amount_inr: int, name: str, phone: str):
    init_db()
    order_id = str(uuid.uuid4())[:10].upper()
    client = get_rzp_client()
    clean_phone = "".join(filter(str.isdigit, phone))[-10:] or "9999999999"

    plink = client.payment_link.create({
        "amount": amount_inr * 100,
        "currency": "INR",
        "accept_partial": False,
        "reference_id": order_id,
        "description": f"{scheme} CA-Format DPR & Balance Sheet (.docx)",
        "customer": {"name": name or "Applicant", "contact": f"+91{clean_phone}"},
        "notify": {"sms": False, "email": False},
        "callback_url": f"{APP_URL}?order_id={order_id}",
        "callback_method": "get",
    })
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO orders (order_id, scheme, payload_json, amount_inr, plink_id, status) VALUES (?, ?, ?, ?, ?, ?)",
            (order_id, scheme, json.dumps(payload), amount_inr, plink["id"], "PENDING"),
        )
    return order_id, plink["short_url"]

def check_order_paid(order_id: str):
    init_db()
    with sqlite3.connect(DB_PATH) as conn:
        row = conn.execute(
            "SELECT scheme, payload_json, plink_id, status FROM orders WHERE order_id = ?",
            (order_id.strip().upper(),),
        ).fetchone()
    if not row:
        return None
    scheme, payload_json, plink_id, status = row
    if status != "PAID":
        client = get_rzp_client()
        plink_data = client.payment_link.fetch(plink_id)
        if plink_data.get("status") == "paid":
            status = "PAID"
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("UPDATE orders SET status = 'PAID' WHERE order_id = ?", (order_id.strip().upper(),))
    return {"order_id": order_id.strip().upper(), "scheme": scheme, "payload": json.loads(payload_json), "status": status}

def get_mudra_tier(loan_amt: float) -> str:
    if loan_amt <= 50000:
        return "Shishu (Up to Rs. 50,000)"
    elif loan_amt <= 500000:
        return "Kishore (Rs. 50,001 - Rs. 5,00,000)"
    elif loan_amt <= 1000000:
        return "Tarun (Rs. 5,00,001 - Rs. 10,00,000)"
    return "Tarun Plus (Rs. 10,00,001 - Rs. 20,00,000)"

# ---------------- 1. FINANCIAL & DOCX ENGINE ----------------
def compute_dpr_and_docx(p: dict):
    is_mudra = p["scheme"] == "PMMY MUDRA"
    facility_type = p.get("facility_type", "Term Loan + Cash Credit (TL + CC)") if is_mudra else "Composite Loan (TL + WC)"
    has_cc = is_mudra and ("CC" in facility_type)

    if not is_mudra:
        # STRICT CM-YUVA RULES: Total Project Cost = 5,00,000; Margin = 10% (50,000); Loan = 90% (4,50,000); Subsidy = 10% (50,000)
        total_project_cost = 500000
        tl_cost = min(500000, max(0, int(p["tl_cost"])))
        wc_cost = total_project_cost - tl_cost
        margin_pct = 10.0
        promoter_contrib = 50000
        total_bank_loan = 450000  # Strictly 90% of 5,00,000
        tl_loan = round(tl_cost * 0.90)  # Strictly 90% of TL component
        cc_loan = total_bank_loan - tl_loan  # Strictly 90% of WC component
        govt_subsidy = 50000
        tenure_months = 48
        interest_rate = 0.0
        proj_years = 4
        year_labels = ["YEAR-1", "YEAR-2", "YEAR-3", "YEAR-4"]
        base_fy = int(p.get("base_fy", 2026))
    else:
        # PMMY MUDRA RULES: Loan = 90% of Project Cost (10% Margin) for TL, CC, or TL+CC
        margin_pct = 10.0
        if facility_type == "Term Loan Only (TL)":
            tl_cost = int(p["tl_cost"])
            wc_cost = 0
        elif facility_type == "Cash Credit Only (CC)":
            tl_cost = 0
            wc_cost = int(p["wc_cost"])
        else:
            tl_cost = int(p["tl_cost"])
            wc_cost = int(p["wc_cost"])

        total_project_cost = tl_cost + wc_cost
        promoter_contrib = round(total_project_cost * 0.10)
        total_bank_loan = total_project_cost - promoter_contrib  # Strictly 90% of Project Cost
        tl_loan = round(tl_cost * 0.90)
        cc_loan = total_bank_loan - tl_loan
        govt_subsidy = 0
        tenure_months = int(p["tenure_months"]) if tl_loan > 0 else 60
        interest_rate = float(p["interest_rate"])
        proj_years = 5 if tenure_months >= 60 else 4
        base_fy = int(p.get("base_fy", 2026))

        if has_cc:
            # Provisional (Current Year), Estimated (Next Year), Projected (Subsequent Years)
            tags = ["PROVISIONAL", "ESTIMATED", "PROJECTED", "PROJECTED", "PROJECTED"]
            year_labels = [
                f"FY {base_fy + idx}-{str(base_fy + idx + 1)[-2:]} ({tags[idx]})"
                for idx in range(proj_years)
            ]
        else:
            year_labels = [f"YEAR-{i}" for i in range(1, proj_years + 1)]

    mudra_tier = get_mudra_tier(total_bank_loan) if is_mudra else "CM-YUVA Phase-1 (Rs. 5.00 Lakh)"

    # EMI strictly on Term Loan (tl_loan) only! CC limit is revolving and renewed every year.
    r_monthly = (interest_rate / 100.0) / 12.0
    if tl_loan > 0 and tenure_months > 0:
        if r_monthly > 0:
            monthly_emi = round(
                tl_loan * r_monthly * ((1 + r_monthly) ** tenure_months) / (((1 + r_monthly) ** tenure_months) - 1),
                2,
            )
        else:
            monthly_emi = round(tl_loan / tenure_months, 2)
    else:
        monthly_emi = 0.0

    # Month-by-month amortization on Term Loan (tl_loan) ONLY
    sched_bal = float(tl_loan)
    yearly_tl_prin = [0.0] * proj_years
    yearly_tl_int = [0.0] * proj_years
    monthly_grid = []

    for yr_idx in range(proj_years):
        yr_m_list = []
        for m in range(12):
            m_num = yr_idx * 12 + m + 1
            if tl_loan > 0 and m_num <= tenure_months and sched_bal > 0:
                m_int = round(sched_bal * r_monthly, 2)
                m_prin = round(min(sched_bal, monthly_emi - m_int), 2)
                if m_num == tenure_months:
                    m_prin = round(sched_bal, 2)
                sched_bal = max(0.0, round(sched_bal - m_prin, 2))
            else:
                m_int, m_prin = 0.0, 0.0
            yearly_tl_prin[yr_idx] += m_prin
            yearly_tl_int[yr_idx] += m_int
            yr_m_list.append((m_prin, m_int))
        monthly_grid.append(yr_m_list)

    # Annual CC Interest (revolving limit renewed annually, no principal amortization)
    annual_cc_int = round(cc_loan * (interest_rate / 100.0)) if (is_mudra and cc_loan > 0) else 0

    # Capex breakdown matching tl_cost
    if tl_cost > 0:
        ratio = tl_cost / 480000.0
        capex_items = [
            ("Primary Plant, Machinery & Core Equipment", round(245000 * ratio)),
            ("Storage, Raw Material Handling & Utility Setup", round(55000 * ratio)),
            ("Electrical Panel, Power Backup & Plumbing/Wiring Fittings", round(60000 * ratio)),
            ("Packaging, Containers, Tools & Operational Accessories", round(65000 * ratio)),
            ("Quality Testing, Sealing & Finishing Equipment", round(35000 * ratio)),
        ]
        allocated_sum = sum(amt for _, amt in capex_items)
        capex_items.append(
            ("Shop Furniture, Weighing Scale, Computer & Branding Board", round(tl_cost - allocated_sum))
        )
    else:
        capex_items = [("Nil - Pure Working Capital / Cash Credit Proposal", 0)]

    utilities_breakdown = [
        ("Travelling & Conveyance", p["u_travel"]),
        ("Telephone, Postage & Internet", p["u_telecom"]),
        ("Repairs & Maintenance", p["u_repairs"]),
        ("Insurance", p["u_insurance"]),
        ("Staff Welfare", p["u_welfare"]),
        ("Power & Fuel", p["u_power"]),
        ("Miscellaneous Expenses", p["u_misc"]),
    ]
    yr1_utilities = sum(val for _, val in utilities_breakdown)
    yr1_salaries = (p["staff1_salary"] + p["staff2_salary"]) * 12

    dep_list = []
    wdv = float(tl_cost)
    for _ in range(proj_years):
        d = round(wdv * 0.15)
        dep_list.append(d)
        wdv -= d

    sales_growth = [1.0, 1.15, 1.30, 1.48, 1.65][:proj_years]
    exp_growth = [1.0, 1.10, 1.20, 1.32, 1.45][:proj_years]

    pnl_rows, bs_rows, repayment_summary = [], [], []
    cum_capital = float(promoter_contrib)
    rem_tl = float(tl_loan)
    net_fixed = float(tl_cost)

    for i, (yr_label, s_mult, e_mult, dep) in enumerate(zip(year_labels, sales_growth, exp_growth, dep_list), start=1):
        sales = round(p["yr1_sales"] * s_mult)
        purchases = round(p["yr1_purchases"] * s_mult)
        salaries = round(yr1_salaries * e_mult)
        utilities = round(yr1_utilities * e_mult)
        tl_int_yr = round(yearly_tl_int[i - 1])
        cc_int_yr = annual_cc_int
        tl_prin_yr = round(yearly_tl_prin[i - 1])
        bank_charges = round(p["yr1_bank_charges"] * (1 + 0.25 * (i - 1)))

        tot_exp_before_dep = purchases + salaries + utilities + tl_int_yr + cc_int_yr + bank_charges
        profit_before_dep = sales - tot_exp_before_dep
        pat = profit_before_dep - dep

        # Debt service = TL Principal + TL Interest + CC Interest
        debt_service = tl_prin_yr + tl_int_yr + cc_int_yr
        cash_accrual_pre_int = pat + dep + tl_int_yr + cc_int_yr
        dscr = round(cash_accrual_pre_int / debt_service, 2) if debt_service > 0 else 0.0

        pnl_rows.append({
            "Year": yr_label,
            "Sales Realisation": sales,
            "Cost of Sales / Purchases": purchases,
            "Salaries": salaries,
            "Utilities & Others": utilities,
            "Intt. On Term Loan": tl_int_yr,
            "Intt. On Cash Credit (CC)": cc_int_yr,
            "Bank Charges": bank_charges,
            "Total Expenditure (Excl. Dep)": tot_exp_before_dep,
            "Profit Before Depreciation (Cash Accrual)": profit_before_dep,
            "Depreciation": dep,
            "Profit After Tax (PAT)": pat,
            "Total Debt Service": debt_service,
            "DSCR": dscr,
        })

        beg_tl = round(rem_tl)
        end_tl = max(0, beg_tl - tl_prin_yr)
        repayment_summary.append({
            "Year": yr_label,
            "Opening Term Loan": beg_tl,
            "TL Principal Repaid": tl_prin_yr,
            "TL Interest Paid": tl_int_yr,
            "Closing Term Loan": end_tl,
            "CC Limit (Renewed Annually)": cc_loan,
            "Annual CC Interest": cc_int_yr,
        })
        rem_tl = float(end_tl)

        drawings = round(p["yr1_drawings"] * (1 + 0.20 * (i - 1)))
        closing_cap = round(cum_capital + pat - drawings)
        net_fixed = round(net_fixed - dep)

        # In CM-YUVA, 10% subsidy (Rs. 50,000) is adjusted in Year-3; in Mudra, CC limit stays active on Balance Sheet
        subsidy_asset = govt_subsidy if (not is_mudra and i >= 3) else 0
        closing_cc_or_wc = cc_loan if is_mudra else (cc_loan if i < 4 else 0)
        total_liab = closing_cap + end_tl + closing_cc_or_wc

        # Working capital assets scale smoothly with sales growth in Provisional/Estimated/Projected years
        wc_base = max(wc_cost, round(sales * 0.12)) if (is_mudra and wc_cost == 0) else wc_cost
        inventory = round(wc_base * 0.65 * (1 + 0.08 * (i - 1)))
        debtors = round(wc_base * 0.25 * (1 + 0.08 * (i - 1)))
        non_cash_assets = net_fixed + subsidy_asset + inventory + debtors
        if total_liab >= non_cash_assets:
            cash = total_liab - non_cash_assets
        else:
            # Adjust inventory/debtors proportionally if cash would dip below zero
            cash = round(wc_base * 0.10)
            rem_wc_assets = max(0, total_liab - net_fixed - subsidy_asset - cash)
            inventory = round(rem_wc_assets * 0.72)
            debtors = rem_wc_assets - inventory
        total_assets = net_fixed + subsidy_asset + inventory + debtors + cash

        bs_rows.append({
            "Year": yr_label,
            "Opening Capital": round(cum_capital),
            "Add: Net Profit": pat,
            "Less: Drawings": drawings,
            "Closing Capital": closing_cap,
            "Term Loan Balance": end_tl,
            "Cash Credit (CC) Limit": closing_cc_or_wc,
            "Total Liabilities": total_liab,
            "Net Fixed Assets": net_fixed,
            "Govt. Subsidy Receivable/Adj": subsidy_asset,
            "Inventory / Stock": inventory,
            "Sundry Debtors & Advances": debtors,
            "Cash in Hand & Bank": cash,
            "Total Assets": total_assets,
        })
        cum_capital = float(closing_cap)

    yr1 = pnl_rows[0]
    var_cost = (
        yr1["Cost of Sales / Purchases"]
        + yr1["Salaries"]
        + yr1["Bank Charges"]
        + round(yr1["Utilities & Others"] * 0.80)
    )
    fixed_cost = yr1["Depreciation"] + yr1["Intt. On Term Loan"] + yr1["Intt. On Cash Credit (CC)"] + round(yr1["Utilities & Others"] * 0.20)
    pv_ratio = round(((yr1["Sales Realisation"] - var_cost) / yr1["Sales Realisation"]) * 100, 2) if yr1["Sales Realisation"] > 0 else 0.0
    bep_sales = round(fixed_cost / (pv_ratio / 100.0)) if pv_ratio > 0 else 0
    margin_of_safety = yr1["Sales Realisation"] - bep_sales
    yr1_net_worth = promoter_contrib + yr1["Profit After Tax (PAT)"]
    debt_equity = round(total_bank_loan / yr1_net_worth, 2) if yr1_net_worth > 0 else 0.0
    avg_dscr = round(sum(r["DSCR"] for r in pnl_rows) / float(proj_years), 2)

    ratios = {
        "Variable Cost (Yr 1)": var_cost,
        "Fixed Cost (Yr 1)": fixed_cost,
        "P/V Ratio (%)": pv_ratio,
        "Break-Even Sales (₹)": bep_sales,
        "Margin of Safety (₹)": margin_of_safety,
        "Debt-Equity Ratio": f"{debt_equity} : 1",
        "Average DSCR": avg_dscr,
    }

    # ---------------- BUILD WORD DOCUMENT (.DOCX) ----------------
    def add_draft_footer(document_obj):
        for sec in document_obj.sections:
            ftr = sec.footer
            fp = ftr.paragraphs[0] if ftr.paragraphs else ftr.add_paragraph()
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fr = fp.add_run(
                "DRAFT / UNAUDITED PROJECTION — Prepared by Applicant for Bank Appraisal Only. "
                "Subject to Independent CA Verification & Bank Due Diligence."
            )
            fr.font.size = Pt(8)
            fr.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)

    doc = Document()
    add_draft_footer(doc)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.2

    def make_table(headers, rows_data):
        num_cols = len(headers) if headers else len(rows_data[0])
        tbl = doc.add_table(rows=len(rows_data) + (1 if headers else 0), cols=num_cols)
        tbl.style = "Table Grid"
        if headers:
            for j, h in enumerate(headers):
                tbl.cell(0, j).text = str(h)
        offset = 1 if headers else 0
        for r_i, r_val in enumerate(rows_data):
            for c_i, c_val in enumerate(r_val):
                tbl.cell(r_i + offset, c_i).text = str(c_val)
        return tbl

    scheme_banner = (
        f"UNDER PRADHAN MANTRI MUDRA YOJANA (PMMY - {mudra_tier.upper()})\n"
        f"FACILITY: {facility_type.upper()} | GOVERNMENT OF INDIA"
        if is_mudra
        else "UNDER MUKHYAMANTRI YUVA UDYAMI VIKAS ABHIYAN (CM-YUVA)\nGOVERNMENT OF UTTAR PRADESH"
    )

    # PAGE 1: COVER PAGE
    doc.add_paragraph("\n\n")
    p_cov1 = doc.add_paragraph()
    p_cov1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov1 = p_cov1.add_run("DETAILED PROJECT REPORT CUM FEASIBILITY STUDY")
    r_cov1.bold = True
    r_cov1.font.size = Pt(20)
    r_cov1.font.color.rgb = RGBColor(0x1B, 0x4F, 0x72)

    p_cov2 = doc.add_paragraph()
    p_cov2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov2 = p_cov2.add_run(f"\n{scheme_banner}\n\n\nOF\n")
    r_cov2.bold = True
    r_cov2.font.size = Pt(12)

    p_cov3 = doc.add_paragraph()
    p_cov3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov3 = p_cov3.add_run(f"{p['unit_name'].upper()}\n")
    r_cov3.bold = True
    r_cov3.font.size = Pt(18)
    r_cov3.font.color.rgb = RGBColor(0x0B, 0x53, 0x45)

    p_cov4 = doc.add_paragraph()
    p_cov4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_cov4 = p_cov4.add_run(
        f"ACTIVITY: {p['activity'].upper()}\n"
        f"PROPRIETOR: {p['promoter_name'].upper()}\n"
        f"UNIT ADDRESS: {p['address'].upper()}\n\n"
        f"TOTAL PROJECT COST: Rs. {total_project_cost:,}/-\n"
        f"BANK FINANCE REQUIRED (90%): Rs. {total_bank_loan:,}/- (TL: Rs. {tl_loan:,}/- | CC/WC: Rs. {cc_loan:,}/-)\n\n\n"
        f"SUBMITTED TO:\n{p['bank_name'].upper()}\n{p['bank_branch'].upper()}\n\n"
        f"PREPARED BY (APPLICANT):\n{p['promoter_name'].upper()}\nDATE: {p['report_date']}"
    )
    r_cov4.bold = True
    r_cov4.font.size = Pt(11)
    doc.add_page_break()

    # PAGE 2: PROJECT AT A GLANCE
    doc.add_heading("1. PROJECT AT A GLANCE", level=1)
    subsidy_or_tier_row = (
        ["9. PMMY Mudra Category & Facility", f"{mudra_tier} | {facility_type}"]
        if is_mudra
        else ["9. Govt. Margin Money Subsidy", f"Rs. {govt_subsidy:,}/- (10% Grant under CM-YUVA Scheme)"]
    )
    if is_mudra:
        if tl_loan > 0 and cc_loan > 0:
            repay_str = f"TL: {tenure_months} EMIs @ Rs. {monthly_emi:,.2f}/- ({interest_rate}% p.a.) | CC: Renewed Annually"
        elif tl_loan > 0:
            repay_str = f"Term Loan: {tenure_months} Monthly EMIs @ Rs. {monthly_emi:,.2f}/- ({interest_rate}% p.a. Reducing)"
        else:
            repay_str = f"Cash Credit Limit: Rs. {cc_loan:,}/- @ {interest_rate}% p.a. (Revolving, Renewed Annually)"
    else:
        repay_str = f"48 Monthly Instalments @ Rs. {monthly_emi:,.2f}/- (0% Interest under CM-YUVA)"

    glance = [
        ["1. Name of the Unit", p["unit_name"]],
        ["2. Name of the Proprietor", f"{p['promoter_name']} ({p['category']})"],
        ["3. Address of the Unit", p["address"]],
        ["4. Status & Constitution", f"{p['unit_status']} Enterprise | Proprietorship"],
        ["5. Nature of Activity", p["activity"]],
        ["6. Total Cost of Project", f"Rs. {total_project_cost:,}/- (Fixed: Rs. {tl_cost:,}/- | WC: Rs. {wc_cost:,}/-)"],
        ["7. Promoter's Contribution (10%)", f"Rs. {promoter_contrib:,}/- ({margin_pct}%)"],
        ["8. Bank Loan Required (90%)", f"Rs. {total_bank_loan:,}/- (Term Loan: Rs. {tl_loan:,}/- & CC: Rs. {cc_loan:,}/-)"],
        subsidy_or_tier_row,
        ["10. Annual Sales Realisation (Yr 1)", f"Rs. {yr1['Sales Realisation']:,}/-"],
        ["11. Net Profit After Tax (Yr 1)", f"Rs. {yr1['Profit After Tax (PAT)']:,}/-"],
        [f"12. Average DSCR ({proj_years} Years)", f"{avg_dscr} : 1"],
        ["13. Break-Even Sales (Yr 1)", f"Rs. {bep_sales:,}/- (P/V Ratio: {pv_ratio}%)"],
        ["14. Employment Potential", f"0{p['employment_count']} Persons (Direct Skilled & Semi-Skilled)"],
        ["15. Repayment / Renewal Terms", repay_str],
        ["16. Financing Bank & Branch", f"{p['bank_name']}, {p['bank_branch']}"],
    ]
    make_table(["S. No. & Parameter", "Details"], glance)

    doc.add_heading("2. EXECUTIVE SUMMARY & PROMOTER BACKGROUND", level=1)
    scheme_intro = (
        f"Pradhan Mantri Mudra Yojana (PMMY - {mudra_tier}) of the Government of India"
        if is_mudra
        else "Mukhyamantri Yuva Udyami Vikas Abhiyan (CM-YUVA) of the Government of Uttar Pradesh"
    )
    doc.add_paragraph(
        f"{p['unit_name']} is a proposed micro-enterprise promoted by {p['promoter_name']} at {p['address']}. "
        f"The unit is set up under the {scheme_intro} to engage in {p['activity']}.\n\n"
        f"The promoter possesses hands-on operational understanding of the local market and customer demand. "
        f"With growing local consumption and organized service delivery, the project demonstrates strong cash generation "
        f"and comfortable debt-servicing capacity from the first year of operations."
    )
    doc.add_page_break()

    # PAGE 3: TECHNICAL FEASIBILITY & MARKET STRATEGY
    doc.add_heading("3. TECHNICAL FEASIBILITY & OPERATIONAL WORKFLOW", level=1)
    doc.add_paragraph(
        f"The unit is planned with modern, energy-efficient equipment suited for {p['activity']}, ensuring consistent "
        f"output quality, minimal wastage, and compliance with local municipal and environmental standards."
    )
    process_steps = [
        ["Stage 1: Raw Material / Input Procurement", "Sourcing certified raw materials and consumables from authorized regional suppliers."],
        ["Stage 2: Primary Processing / Preparation", "Standardized operational setup ensuring hygiene, precision, and quality control."],
        ["Stage 3: Core Production / Service Execution", "Deployment of dedicated plant & machinery / working capital cycle for daily operations."],
        ["Stage 4: Quality Inspection & Packaging", "Batch-level verification, labeling, and safe packaging for retail/institutional buyers."],
        ["Stage 5: Dispatch, Billing & Customer Retention", "Timely local delivery, digital UPI/invoice billing, and repeat order management."],
    ]
    make_table(["Operational Stage", "Technical Function & Quality Control"], process_steps)

    doc.add_heading("4. MARKET POTENTIAL & MARKETING STRATEGY", level=1)
    doc.add_paragraph(
        f"1. Institutional & Commercial Tie-Ups: Monthly supply contracts with offices, clinics, shops, and institutions near {p['address']}.\n"
        "2. Direct Retail & Walk-In Customers: Competitive pricing, consistent quality, and local visibility.\n"
        "3. Bulk & Event Orders: Serving seasonal, festive, and bulk local demand for incremental margin."
    )
    doc.add_page_break()

    # PAGE 4: COST OF PROJECT & MEANS OF FINANCING
    doc.add_heading("5. COST OF PROJECT & MEANS OF FINANCING", level=1)
    doc.add_heading("A. Fixed Capital / Plant & Machinery Breakdown", level=2)
    cost_rows = [[str(idx + 1), item, f"{amt:,}"] for idx, (item, amt) in enumerate(capex_items)]
    cost_rows.extend([
        ["", "Sub-Total: Fixed Capital / Plant & Machinery", f"{tl_cost:,}"],
        ["", "Working Capital Requirement (Stock, Consumables & Receivables)", f"{wc_cost:,}"],
        ["", "TOTAL COST OF PROJECT", f"{total_project_cost:,}"],
    ])
    make_table(["S. No.", f"Particulars (Premises: {p['premises_status']})", "Amount (Rs.)"], cost_rows)

    doc.add_heading("B. Working Capital Assessment", level=2)
    wc_p1 = round(wc_cost * 0.45)
    wc_p2 = round(wc_cost * 0.30)
    wc_p3 = wc_cost - wc_p1 - wc_p2
    wc_rows = [
        ["Raw Materials, Inventory & Operational Consumables", "15-30 Days", f"{wc_p1:,}"],
        ["Electricity, Fuel & Operational Overheads", "15 Days", f"{wc_p2:,}"],
        ["Sundry Debtors / Credit & Cash Buffer", "15 Days", f"{wc_p3:,}"],
        ["Total Working Capital Requirement", "", f"{wc_cost:,}"],
    ]
    make_table(["Working Capital Component", "Holding Period", "Amount (Rs.)"], wc_rows)

    doc.add_heading("C. Means of Financing (90% Bank Finance : 10% Margin)", level=2)
    means = [
        [f"1. Promoter's Own Contribution ({margin_pct}%)", f"Rs. {promoter_contrib:,}"],
        ["2. Term Loan from Bank (90% of Fixed Capital)", f"Rs. {tl_loan:,}"],
        ["3. Cash Credit / Working Capital Limit (90% of WC)", f"Rs. {cc_loan:,}"],
        ["TOTAL MEANS OF FINANCING", f"Rs. {total_project_cost:,}"],
    ]
    if not is_mudra:
        means.append(["Memo: UP Govt. Margin Money Subsidy (10% Adjusted in Loan)", f"Rs. {govt_subsidy:,}"])
    else:
        means.append([f"Memo: Covered under PMMY ({mudra_tier}) - Collateral Free", "CGTMSE / CGFMU"])
    make_table(["Source of Finance", "Amount (Rs.)"], means)
    doc.add_page_break()

    # PAGE 5: MANPOWER, UTILITIES & DEPRECIATION
    doc.add_heading("6. MANPOWER, UTILITIES & DEPRECIATION SCHEDULE", level=1)
    doc.add_heading("A. Employment & Salary Structure (Year 1)", level=2)
    staff_rows = [
        ["1", p["staff1_role"], "01", f"Rs. {p['staff1_salary']:,}", f"Rs. {p['staff1_salary'] * 12:,}"],
        ["2", p["staff2_role"], "01", f"Rs. {p['staff2_salary']:,}", f"Rs. {p['staff2_salary'] * 12:,}"],
        ["", "Total Manpower Cost", f"0{p['employment_count']}", f"Rs. {p['staff1_salary'] + p['staff2_salary']:,}", f"Rs. {yr1_salaries:,}"],
    ]
    make_table(["S. No.", "Designation", "Nos.", "Monthly (Rs.)", "Annual (Rs.)"], staff_rows)

    doc.add_heading("B. Utilities & Administrative Overheads (Year 1)", level=2)
    util_rows = [[str(i + 1), u_name, f"{u_val:,}"] for i, (u_name, u_val) in enumerate(utilities_breakdown)]
    util_rows.append(["", "Total Utilities & Others", f"{yr1_utilities:,}"])
    make_table(["S. No.", "Expense Head", "Annual Amount (Rs.)"], util_rows)

    doc.add_heading(f"C. Depreciation Calculation Chart ({proj_years} Years @ 15% WDV)", level=2)
    wdv_vals = []
    cur_w = tl_cost
    for d_val in dep_list:
        cur_w -= d_val
        wdv_vals.append(cur_w)
    wdv_calc = [
        ["Plant, Machinery & Equipment @ 15% WDV"] + [f"{d_v:,}" for d_v in dep_list],
        ["Net Written Down Value (Year-End)"] + [f"{w_v:,}" for w_v in wdv_vals],
    ]
    make_table(["Particulars (Rs.)"] + year_labels, wdv_calc)
    doc.add_page_break()

    # PAGE 6: PROJECTED PROFITABILITY STATEMENT
    doc.add_heading(f"7. PROJECTED PROFITABILITY STATEMENT ({proj_years} YEARS)", level=1)
    assump_text = (
        f"Assumptions: 300 working days/year; bank interest @ {interest_rate}% p.a. on reducing Term Loan balance and/or revolving Cash Credit limit under PMMY Mudra."
        if is_mudra
        else "Assumptions: 300 working days/year; 100% interest subvention for 4 years under CM-YUVA scheme (effective interest rate = NIL)."
    )
    doc.add_paragraph(assump_text)
    pnl_keys = [
        "Sales Realisation", "Cost of Sales / Purchases", "Salaries", "Utilities & Others",
        "Intt. On Term Loan", "Intt. On Cash Credit (CC)", "Bank Charges",
        "Total Expenditure (Excl. Dep)", "Profit Before Depreciation (Cash Accrual)",
        "Depreciation", "Profit After Tax (PAT)",
    ]
    pnl_table_rows = [[key] + [f"{row[key]:,}" for row in pnl_rows] for key in pnl_keys]
    make_table(["Particulars (Rs.)"] + year_labels, pnl_table_rows)

    doc.add_heading("8. CASH ACCRUAL & REPAYMENT CAPACITY", level=1)
    cash_keys = ["Profit After Tax (PAT)", "Depreciation", "Profit Before Depreciation (Cash Accrual)", "Total Debt Service", "DSCR"]
    cash_table_rows = [
        [key] + [f"{row[key]:,.2f}" if key == "DSCR" else f"{row[key]:,}" for row in pnl_rows]
        for key in cash_keys
    ]
    make_table(["Particulars (Rs.)"] + year_labels, cash_table_rows)
    doc.add_page_break()

    # PAGE 7: BREAK-EVEN & RATIO ANALYSIS
    doc.add_heading("9. BREAK-EVEN ANALYSIS & FINANCIAL RATIOS", level=1)
    doc.add_heading("A. Cost Bifurcation (Year 1)", level=2)
    bep_bifurcation = [
        ["1. Cost of Sales / Purchases", f"{yr1['Cost of Sales / Purchases']:,}", "0"],
        ["2. Salaries & Wages", f"{yr1['Salaries']:,}", "0"],
        ["3. Interest on Term Loan & CC", "0", f"{yr1['Intt. On Term Loan'] + yr1['Intt. On Cash Credit (CC)']:,}"],
        ["4. Bank Charges", f"{yr1['Bank Charges']:,}", "0"],
        ["5. Depreciation @ 15% WDV", "0", f"{yr1['Depreciation']:,}"],
        ["6. Utilities & Others (80% Variable : 20% Fixed)", f"{round(yr1['Utilities & Others'] * 0.80):,}", f"{round(yr1['Utilities & Others'] * 0.20):,}"],
        ["TOTAL COST BIFURCATION", f"{var_cost:,}", f"{fixed_cost:,}"],
    ]
    make_table(["Particulars (Rs.)", "Variable Cost (Rs.)", "Fixed Cost (Rs.)"], bep_bifurcation)

    doc.add_heading("B. Key Viability Ratios", level=2)
    ratio_rows = [
        ["1. Profit / Volume (P/V) Ratio = ((Sales - Variable Cost) / Sales) x 100", f"{pv_ratio}%"],
        ["2. Break-Even Sales = Fixed Cost / P/V Ratio", f"Rs. {bep_sales:,}/-"],
        ["3. Margin of Safety = Projected Sales - Break-Even Sales", f"Rs. {margin_of_safety:,}/-"],
        ["4. Debt-Equity Ratio (Loan : Promoter Capital + Yr 1 Profit)", f"{debt_equity} : 1"],
        [f"5. Average Debt Service Coverage Ratio (DSCR - {proj_years} Years)", f"{avg_dscr} : 1 (Benchmark >= 1.50)"],
    ]
    make_table(["Financial Ratio & Formula", "Computed Value"], ratio_rows)
    doc.add_page_break()

    # PAGE 8: PROVISIONAL, ESTIMATED & PROJECTED BALANCE SHEET
    bs_heading = (
        f"10. PROVISIONAL, ESTIMATED & PROJECTED BALANCE SHEET ({proj_years} YEARS)"
        if has_cc
        else f"10. PROJECTED BALANCE SHEET FOR {proj_years} YEARS"
    )
    doc.add_heading(bs_heading, level=1)
    bs_keys = [
        "Opening Capital", "Add: Net Profit", "Less: Drawings", "Closing Capital",
        "Term Loan Balance", "Cash Credit (CC) Limit", "Total Liabilities",
        "Net Fixed Assets", "Govt. Subsidy Receivable/Adj",
        "Inventory / Stock", "Sundry Debtors & Advances", "Cash in Hand & Bank", "Total Assets",
    ]
    bs_table_rows = [[key] + [f"{row[key]:,}" for row in bs_rows] for key in bs_keys]
    make_table(["Particulars (Rs.)"] + year_labels, bs_table_rows)
    if not is_mudra:
        doc.add_paragraph(
            f"\nNote: Under CM-YUVA, total project cost is Rs. 5,00,000/- with 10% promoter margin (Rs. 50,000/-), "
            f"90% bank loan (Rs. 4,50,000/-), and 10% UP Govt. Margin Money Subsidy (Rs. 50,000/-) adjusted as per scheme rules."
        )
    elif has_cc:
        doc.add_paragraph(
            f"\nNote: In accordance with bank appraisal norms for Cash Credit proposals, {year_labels[0]} represents Provisional financials "
            f"(current year), {year_labels[1]} represents Estimated financials, and subsequent columns represent Projected financials. "
            f"Cash Credit limit of Rs. {cc_loan:,}/- is revolving and renewed every year."
        )
    else:
        doc.add_paragraph(
            f"\nNote: Projected Balance Sheet prepared for {mudra_tier} ({facility_type}) with 90% bank loan (Rs. {total_bank_loan:,}/-) "
            f"and reducing-balance interest @ {interest_rate}% p.a."
        )
    doc.add_page_break()

    # PAGE 9: LOAN REPAYMENT SCHEDULE (TL) & ANNUAL CC RENEWAL
    if is_mudra and tl_loan == 0:
        doc.add_heading(f"11. CASH CREDIT (CC) LIMIT UTILIZATION & ANNUAL RENEWAL SCHEDULE ({proj_years} YEARS)", level=1)
        doc.add_paragraph(
            f"Since this is a pure Cash Credit (CC) working capital limit of Rs. {cc_loan:,}/- @ {interest_rate}% p.a., "
            "there is no monthly Term Loan EMI. The CC limit is revolving in nature and renewed every 12 months based on stock & sales turnover."
        )
        cc_sched_rows = [
            [
                r["Year"],
                f"Rs. {cc_loan:,}",
                "Annual Renewal (12 Months)",
                f"Rs. {r['Annual CC Interest']:,}",
                f"Rs. {cc_loan:,}",
            ]
            for r in repayment_summary
        ]
        make_table(
            ["Financial Year", "Sanctioned CC Limit", "Renewal Frequency", "Annual Interest @ " + f"{interest_rate}%", "Closing CC Limit"],
            cc_sched_rows,
        )
    else:
        repay_heading = (
            f"11. TERM LOAN REPAYMENT SCHEDULE ({tenure_months} MONTHS @ {interest_rate}% P.A.) & CC RENEWAL"
            if is_mudra
            else "11. LOAN REPAYMENT SCHEDULE (48 MONTHS @ NIL INTEREST)"
        )
        doc.add_heading(repay_heading, level=1)
        if is_mudra and cc_loan > 0:
            doc.add_paragraph(
                f"Note: Monthly EMI of Rs. {monthly_emi:,.2f}/- applies strictly to the Term Loan of Rs. {tl_loan:,}/- over {tenure_months} months. "
                f"The Cash Credit limit of Rs. {cc_loan:,}/- is renewed every year with annual interest of Rs. {annual_cc_int:,}/-."
            )
        m_rows = []
        for m_idx in range(12):
            row_cells = [f"Month {m_idx + 1}"]
            for yr_idx in range(proj_years):
                p_val, i_val = monthly_grid[yr_idx][m_idx]
                row_cells.extend([f"{p_val:,.2f}", f"{i_val:,.2f}"])
            m_rows.append(row_cells)

        tot_row = ["TL Annual Total"]
        bal_row = ["Closing TL Balance"]
        for yr_idx in range(proj_years):
            tot_row.extend([f"{round(yearly_tl_prin[yr_idx]):,}", f"{round(yearly_tl_int[yr_idx]):,}"])
            bal_row.extend([f"{repayment_summary[yr_idx]['Closing Term Loan']:,}", "-"])
        m_rows.append(tot_row)
        m_rows.append(bal_row)

        if is_mudra and cc_loan > 0:
            cc_row = ["CC Limit (Renewed)"]
            for yr_idx in range(proj_years):
                cc_row.extend([f"{cc_loan:,}", f"{annual_cc_int:,}"])
            m_rows.append(cc_row)

        sched_headers = ["Month"]
        for yr_idx in range(proj_years):
            sched_headers.extend([f"Yr-{yr_idx + 1} Prin.", f"Yr-{yr_idx + 1} Int."])
        make_table(sched_headers, m_rows)
    doc.add_page_break()

    # PAGE 10: SWOT, STATUTORY COMPLIANCE & DECLARATION
    doc.add_heading("12. SWOT ANALYSIS & RISK MITIGATION", level=1)
    fin_strength = (
        f"Collateral-free bank credit under PMMY ({mudra_tier} - {facility_type}) with strong DSCR ({avg_dscr}:1) and healthy cash accrual."
        if is_mudra
        else f"Zero interest burden for 4 years under CM-YUVA (Rs. 5L project, 90% loan, 10% subsidy) with strong DSCR ({avg_dscr}:1)."
    )
    swot_rows = [
        ["Strengths", fin_strength],
        ["Weaknesses", "Initial customer acquisition in competitive local pockets; working capital discipline required."],
        ["Opportunities", f"Growing local demand in and around {p['address']}; repeat retail and institutional orders."],
        ["Threats & Mitigation", "Power/input cost fluctuations mitigated via backup power and vendor rate contracts."],
    ]
    make_table(["SWOT Dimension", "Project Assessment"], swot_rows)

    doc.add_heading("13. STATUTORY LICENSES & MANDATORY CHECKLIST", level=1)
    doc.add_paragraph(
        "1. Udyam MSME Registration (Ministry of MSME, Govt. of India)\n"
        "2. Applicable Trade / FSSAI / Shop & Establishment Registration\n"
        "3. Registered Rent Agreement / Property Ownership Proof & Landlord NOC\n"
        "4. GST-Compliant Vendor Proforma Invoices / Stock Statement\n"
        "5. Applicant KYC (Aadhaar, PAN, Address Proof, Caste/Category Certificate if applicable)"
    )

    doc.add_heading("14. PROMOTER DECLARATION", level=1)
    doc.add_paragraph(
        f"I, {p['promoter_name']}, Proprietor of {p['unit_name']}, hereby declare that the financial projections and technical "
        f"particulars furnished in this Detailed Project Report are true and feasible to the best of my knowledge."
    )
    p_sig = doc.add_paragraph()
    r_sig = p_sig.add_run(
        f"\n\nPlace: {p['bank_branch']}\nDate: {p['report_date']}\t\t\t\t\tSignature of Proprietor\n\t\t\t\t\t\t({p['promoter_name']})\n\t\t\t\t\t\t{p['unit_name']}"
    )
    r_sig.bold = True

    # Build Standalone Financial Statements (.docx): Year-Wise Sales, P&L & Balance Sheet
    fin_doc = Document()
    add_draft_footer(fin_doc)
    f_style = fin_doc.styles["Normal"]
    f_style.font.name = "Calibri"
    f_style.font.size = Pt(10)

    def add_fin_table(headers, rows_data):
        tbl = fin_doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        tbl.style = "Table Grid"
        for j, h in enumerate(headers):
            tbl.cell(0, j).text = str(h)
        for r_i, r_val in enumerate(rows_data):
            for c_i, c_val in enumerate(r_val):
                tbl.cell(r_i + 1, c_i).text = str(c_val)

    fin_doc.add_heading(f"{p['unit_name'].upper()} — FINANCIAL STATEMENTS ({proj_years} YEARS)", level=1)
    fin_doc.add_paragraph(
        f"Proprietor: {p['promoter_name']} | Scheme: {p['scheme']} ({facility_type}) | "
        f"Project Cost: Rs. {total_project_cost:,}/- | Bank Finance: Rs. {total_bank_loan:,}/-"
    )

    fin_doc.add_heading("1. YEAR-WISE SALES & OPERATING COST SUMMARY (Rs.)", level=2)
    sales_keys = ["Sales Realisation", "Cost of Sales / Purchases", "Salaries", "Utilities & Others", "Profit Before Depreciation (Cash Accrual)"]
    add_fin_table(["Particulars (Rs.)"] + year_labels, [[k] + [f"{r[k]:,}" for r in pnl_rows] for k in sales_keys])

    fin_doc.add_heading("2. PROJECTED PROFIT & LOSS ACCOUNT (Rs.)", level=2)
    add_fin_table(["Particulars (Rs.)"] + year_labels, pnl_table_rows)

    fin_doc.add_heading("3. PROVISIONAL / ESTIMATED / PROJECTED BALANCE SHEET (Rs.)", level=2)
    add_fin_table(["Particulars (Rs.)"] + year_labels, bs_table_rows)

    fin_doc.add_heading("4. KEY VIABILITY RATIOS & DSCR SUMMARY", level=2)
    add_fin_table(["Ratio / Metric", "Computed Value"], ratio_rows)

    fin_buf = io.BytesIO()
    fin_doc.save(fin_buf)
    fin_buf.seek(0)

    # Automated Benchmark Guardrails
    yr1_margin_pct = round((yr1["Profit After Tax (PAT)"] / yr1["Sales Realisation"]) * 100, 2) if yr1["Sales Realisation"] > 0 else 0.0
    warnings = []
    if yr1_margin_pct > 35.0:
        warnings.append(f"High Net Profit Margin ({yr1_margin_pct}%): Banks typically scrutinize micro-unit net margins above 35%.")
    if avg_dscr < 1.25:
        warnings.append(f"Low Average DSCR ({avg_dscr}:1): Below standard bank viability benchmark of 1.25:1.")
    elif avg_dscr > 5.0:
        warnings.append(f"Unusually High DSCR ({avg_dscr}:1): Verify that Year-1 sales are not overstated or expenses understated.")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return {
        "fin_docx_bytes": fin_buf.getvalue(),
        "warnings": warnings,
        "yr1_margin_pct": yr1_margin_pct,
        "total_project_cost": total_project_cost,
        "total_bank_loan": total_bank_loan,
        "tl_loan": tl_loan,
        "cc_loan": cc_loan,
        "monthly_emi": monthly_emi,
        "bep_sales": bep_sales,
        "avg_dscr": avg_dscr,
        "mudra_tier": mudra_tier,
        "facility_type": facility_type,
        "proj_years": proj_years,
        "pnl_rows": pnl_rows,
        "ratios": ratios,
        "bs_rows": bs_rows,
        "repayment_summary": repayment_summary,
        "docx_bytes": buf.getvalue(),
    }


# ---------------- 2. CHECK IF RETURNING FROM RAZORPAY PAYMENT ----------------
query_order_id = st.query_params.get("order_id")
if query_order_id:
    order = check_order_paid(query_order_id)
    if order and order["status"] == "PAID":
        st.success(f"Payment Verified! Order ID: **{order['order_id']}** ({order['scheme']})")
        res = compute_dpr_and_docx(order["payload"])
        u_clean = order["payload"]["unit_name"].replace(" ", "_")
        s_clean = order["scheme"].replace(" ", "_")
        c_dl1, c_dl2 = st.columns(2)
        c_dl1.download_button(
            label="1. Download Complete Word DPR (.docx)",
            data=res["docx_bytes"],
            file_name=f"{s_clean}_DPR_{u_clean}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
        c_dl2.download_button(
            label="2. Download Year-Wise Sales, P&L & Balance Sheet (.docx)",
            data=res["fin_docx_bytes"],
            file_name=f"{s_clean}_Financials_{u_clean}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.info("Bookmark or save your Order ID above. You can re-download your DPR anytime from the sidebar.")
        if st.button("Create Another DPR"):
            st.query_params.clear()
            st.rerun()
        st.stop()
    elif order:
        st.warning(f"Payment for Order **{order['order_id']}** is still pending. Complete payment or refresh after paying.")

# ---------------- 3. MAIN UI & SIDEBAR INPUTS ----------------
st.title("CM-YUVA & PMMY MUDRA DPR Generator")
st.caption("Generates Bank-Ready CA-Format Word DPR (.docx) with Project at a Glance, P&L, Ratios, Balance Sheet & Repayment/Renewal Schedule")

st.sidebar.header("Already Paid? Retrieve DPR")
lookup_id = st.sidebar.text_input("Enter Order ID (e.g. A1B2C3D4E5)")
if st.sidebar.button("Fetch Paid DPR") and lookup_id:
    found = check_order_paid(lookup_id)
    if found and found["status"] == "PAID":
        st.query_params["order_id"] = found["order_id"]
        st.rerun()
    else:
        st.sidebar.error("Order not found or payment not completed yet.")

st.sidebar.divider()
st.sidebar.header("0. Select Loan Scheme")
scheme = st.sidebar.radio("Scheme", ["CM-YUVA (UP)", "PMMY MUDRA"])

facility_type = "Composite Loan (TL + WC)"
if scheme == "PMMY MUDRA":
    facility_type = st.sidebar.selectbox(
        "MUDRA Facility Type",
        [
            "Term Loan + Cash Credit (TL + CC)",
            "Term Loan Only (TL)",
            "Cash Credit Only (CC)",
        ],
    )

st.sidebar.header("1. Unit, Bank & Applicant Details")
unit_name = st.sidebar.text_input("Name of the Unit", "M/s Kashi Pure Aqua Industries")
promoter_name = st.sidebar.text_input("Proprietor Name", "Shri Rajesh Kumar")
mobile_no = st.sidebar.text_input("Applicant Mobile (10 Digits)", "9876543210")
activity = st.sidebar.text_input("Business Activity", "RO Drinking Water Purification & 20L Jar Supply")
address = st.sidebar.text_input("Unit Address", "Sigra, District: Varanasi, Uttar Pradesh")
unit_status = st.sidebar.selectbox("Status of Unit", ["NEW", "EXPANSION"])
premises_status = st.sidebar.selectbox("Shop / Premises Status", ["Rented", "Owned"])
category = st.sidebar.selectbox("Applicant Category", ["General", "OBC", "SC / ST / PwD"])
bank_name = st.sidebar.text_input("Financing Bank", "Punjab National Bank / State Bank of India")
bank_branch = st.sidebar.text_input("Branch & District", "Varanasi, District: Varanasi")
ca_firm = "Applicant Self-Prepared Draft"
base_fy = st.sidebar.number_input("Starting Financial Year (Provisional/Yr-1)", min_value=2024, max_value=2030, value=2026, step=1)
report_date = st.sidebar.text_input("Report Date", "21.11.2026")

st.sidebar.header("2. Project Cost & Loan Structure (₹)")
if scheme == "CM-YUVA (UP)":
    st.sidebar.info("CM-YUVA Rule: Project Cost fixed at ₹5,00,000 | 10% Margin (₹50,000) | 90% Loan (₹4,50,000) | 10% Govt Subsidy (₹50,000).")
    tl_cost = st.sidebar.number_input("Fixed Capital / Term Loan Cost (₹)", min_value=0, max_value=500000, value=450000, step=10000)
    wc_cost = 500000 - tl_cost
    st.sidebar.number_input("Working Capital Requirement (Auto-Balanced ₹)", value=wc_cost, disabled=True)
    margin_pct = 10.0
    tenure_months = 48
    interest_rate = 0.0
else:
    if facility_type == "Term Loan Only (TL)":
        tl_cost = st.sidebar.number_input("Fixed Capital / Term Loan Project Cost (₹)", min_value=10000, max_value=2222222, value=500000, step=10000)
        wc_cost = 0
    elif facility_type == "Cash Credit Only (CC)":
        tl_cost = 0
        wc_cost = st.sidebar.number_input("Working Capital / CC Project Cost (₹)", min_value=10000, max_value=2222222, value=500000, step=10000)
    else:
        tl_cost = st.sidebar.number_input("Fixed Capital / Term Loan Cost (₹)", min_value=0, max_value=2222222, value=300000, step=10000)
        wc_cost = st.sidebar.number_input("Working Capital / CC Cost (₹)", min_value=0, max_value=2222222, value=200000, step=10000)

    margin_pct = 10.0
    st.sidebar.caption("Bank Loan is strictly 90% of Total Project Cost (10% Promoter Margin).")
    if tl_cost > 0:
        tenure_months = st.sidebar.selectbox("Term Loan Repayment Tenure (Months)", [36, 48, 60], index=2)
    else:
        tenure_months = st.sidebar.selectbox("Projection Horizon (Months)", [48, 60], index=1)
    interest_rate = st.sidebar.number_input("Bank Interest Rate (% p.a.)", value=11.5, min_value=0.0, max_value=24.0, step=0.5)

st.sidebar.header("3. Year-1 Sales, Purchases & Drawings (₹)")
yr1_sales = st.sidebar.number_input("Annual Sales Realisation (₹)", value=900000, step=25000)
yr1_purchases = st.sidebar.number_input("Cost of Sales / Consumables (₹)", value=320000, step=10000)
yr1_bank_charges = st.sidebar.number_input("Annual Bank Charges (₹)", value=3000, step=500)
yr1_drawings = st.sidebar.number_input("Annual Proprietor Drawings (₹)", value=120000, step=10000)

st.sidebar.header("4. Staff & Employment (Monthly)")
employment_count = st.sidebar.number_input("Total Employment Potential", value=2, step=1)
staff1_role = st.sidebar.text_input("Staff 1 Role", "Plant Operator cum Technician")
staff1_salary = st.sidebar.number_input("Staff 1 Salary (₹/mo)", value=10000, step=500)
staff2_role = st.sidebar.text_input("Staff 2 Role", "Delivery Helper")
staff2_salary = st.sidebar.number_input("Staff 2 Salary (₹/mo)", value=6000, step=500)

st.sidebar.header("5. Utilities & Others (Yearly ₹)")
u_travel = st.sidebar.number_input("Travelling & Conveyance", value=7000, step=500)
u_telecom = st.sidebar.number_input("Telephone, Postage & Internet", value=4000, step=500)
u_repairs = st.sidebar.number_input("Repairs & Maintenance", value=4000, step=500)
u_insurance = st.sidebar.number_input("Insurance", value=3000, step=500)
u_welfare = st.sidebar.number_input("Staff Welfare", value=5000, step=500)
u_power = st.sidebar.number_input("Power & Fuel", value=12000, step=1000)
u_misc = st.sidebar.number_input("Miscellaneous Expenses", value=6000, step=500)

payload = {
    "scheme": scheme,
    "facility_type": facility_type,
    "unit_name": unit_name,
    "promoter_name": promoter_name,
    "mobile_no": mobile_no,
    "activity": activity,
    "address": address,
    "unit_status": unit_status,
    "premises_status": premises_status,
    "category": category,
    "bank_name": bank_name,
    "bank_branch": bank_branch,
    "ca_firm": ca_firm,
    "base_fy": base_fy,
    "report_date": report_date,
    "tl_cost": tl_cost,
    "wc_cost": wc_cost,
    "margin_pct": margin_pct,
    "tenure_months": tenure_months,
    "interest_rate": interest_rate,
    "yr1_sales": yr1_sales,
    "yr1_purchases": yr1_purchases,
    "yr1_bank_charges": yr1_bank_charges,
    "yr1_drawings": yr1_drawings,
    "employment_count": employment_count,
    "staff1_role": staff1_role,
    "staff1_salary": staff1_salary,
    "staff2_role": staff2_role,
    "staff2_salary": staff2_salary,
    "u_travel": u_travel,
    "u_telecom": u_telecom,
    "u_repairs": u_repairs,
    "u_insurance": u_insurance,
    "u_welfare": u_welfare,
    "u_power": u_power,
    "u_misc": u_misc,
}

res = compute_dpr_and_docx(payload)
n_yrs = res["proj_years"]

# ---------------- 4. LIVE FINANCIAL PREVIEW & RAZORPAY PAYWALL ----------------
if scheme == "PMMY MUDRA":
    cc_note = " | Includes Provisional, Estimated & Projected Balance Sheets" if "CC" in facility_type else ""
    st.info(f"MUDRA Category: **{res['mudra_tier']}** | Facility: **{facility_type}** | Interest: **{interest_rate}% p.a.**{cc_note}")
else:
    st.info(f"CM-YUVA Fixed Structure: **₹5,00,000 Project** | **₹4,50,000 Loan (90%)** (TL: ₹{res['tl_loan']:,} + WC: ₹{res['cc_loan']:,}) | **₹50,000 Subsidy (10%)** | **0% Interest**")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Project Cost", f"₹ {res['total_project_cost']:,}")
k2.metric("Total Loan (90%)", f"₹ {res['total_bank_loan']:,}")
if res["tl_loan"] > 0:
    k3.metric(f"TL EMI ({tenure_months}m)", f"₹ {res['monthly_emi']:,.2f}")
else:
    k3.metric("CC Limit (Annual)", f"₹ {res['cc_loan']:,}")
k4.metric("Break-Even Sales", f"₹ {res['bep_sales']:,}")
k5.metric("Average DSCR", f"{res['avg_dscr']}")

bs_tab_title = (
    f"Provisional, Estimated & Projected BS ({n_yrs}-Yr)"
    if (scheme == "PMMY MUDRA" and "CC" in facility_type)
    else f"{n_yrs}-Yr Balance Sheet"
)

tab1, tab2, tab3, tab4 = st.tabs(
    [f"{n_yrs}-Yr Profitability & DSCR", "Break-Even & Ratios", bs_tab_title, "TL Repayment & CC Renewal"]
)
tab1.dataframe(pd.DataFrame(res["pnl_rows"]), use_container_width=True, hide_index=True)
tab2.table(pd.DataFrame(list(res["ratios"].items()), columns=["Financial Metric", "Value"]))
tab3.dataframe(pd.DataFrame(res["bs_rows"]), use_container_width=True, hide_index=True)
tab4.dataframe(pd.DataFrame(res["repayment_summary"]), use_container_width=True, hide_index=True)

st.divider()
price_inr = 149 if (scheme == "PMMY MUDRA" and res["total_bank_loan"] <= 50000) else 299
st.subheader(f"Download Complete CA-Format Word DPR (.docx) — ₹{price_inr}")

for w_msg in res.get("warnings", []):
    st.warning(f"Benchmark Check: {w_msg}")

if "plink_count" not in st.session_state:
    st.session_state["plink_count"] = 0

attested = st.checkbox(
    "I confirm these projections are genuine estimates prepared by the applicant (unaudited draft) and not certified by a CA."
)

if st.button(f"Proceed to Pay ₹{price_inr} via UPI / Card", type="primary", disabled=not attested):
    if "RAZORPAY_KEY_ID" not in st.secrets:
        st.error("Add RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET, and APP_URL to .streamlit/secrets.toml first.")
    elif st.session_state["plink_count"] >= 3:
        st.error("Session limit reached (max 3 payment links per session). Refresh page or retrieve an existing Order ID.")
    else:
        st.session_state["plink_count"] += 1
        oid, pay_url = create_razorpay_order(scheme, payload, price_inr, promoter_name, mobile_no)
        st.success(f"Order Created! Your Order ID is **{oid}** (save this to re-download anytime).")
        st.link_button(f"Pay ₹{price_inr} Now on Razorpay", pay_url, type="primary")

# ---------------- 5. RAZORPAY KYC COMPLIANCE FOOTER ----------------
st.divider()
with st.expander("About Us, Terms & Conditions, Privacy Policy & Refund Policy"):
    st.markdown("""
    ### 1. About Us & Contact Details
    * **Legal Entity / Trade Name:** [Your Proprietorship Trade Name]
    * **Service Description:** Automated self-serve software utility for generating bank-ready Detailed Project Reports (DPR) and financial projections (`.docx`) for CM-YUVA and PMMY Mudra loan proposals.
    * **Registered Address:** [Full Office Address, Varanasi, Uttar Pradesh, India]
    * **Customer Support Email:** [support@yourdomain.in] | **Phone:** [+91-XXXXXXXXXX]

    ### 2. Terms & Conditions
    * This platform is an **automated document-formatting and financial calculation software tool**.
    * We are an independent private software provider and are **not affiliated with the Government of Uttar Pradesh, Government of India, DIC, SIDBI, or any bank/NBFC**.
    * Generation of a DPR does not guarantee loan sanction or subsidy approval; final credit decisions rest solely with the financing bank and competent authorities.

    ### 3. Privacy Policy
    * Applicant and business inputs entered on this portal are used strictly to compute financial ratios and generate your requested `.docx` project report.
    * All online payments are processed securely via **Razorpay**; we do not store your UPI PIN, card numbers, or net-banking credentials on our servers.

    ### 4. Refund & Cancellation Policy
    * Because this service delivers an **instant, customized digital document (`.docx`)** immediately upon payment verification, **orders cannot be cancelled or refunded once the report has been generated/downloaded**.
    * In case of a **duplicate payment deduction** or a verified technical failure preventing report download, a **100% refund** will be processed to your original payment method within **5–7 business days** upon contacting customer support with your Order ID.
    """)
