"""
BimaGrid Kenya — Enterprise Motor & Motorcycle Insurance Quotation Engine
========================================================================
Client Presentation & Broker Workbench UI
Designed for high-trust client demonstrations, executive dispatch, and multi-quote comparisons.

Guarantees:
  - 100% Unaltered Underwriting Logic: Retains exact calculation algorithms, IRA levies,
    depreciation tables, Old Mutual minimums, and Excel/WhatsApp exports.
  - Institutional FinTech UI: Elevated glassmorphic cards, typography hierarchy,
    regulatory trust badges (IRA Kenya), and clean layout alignment.
  - Zero Disk Writes: All computations and spreadsheet exports operate in-memory.
"""

import io
import os
import textwrap
import urllib.parse
from datetime import datetime

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="BimaGrid Kenya | Enterprise Motor Quotation Desk",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
INSURERS_CSV = os.path.join(APP_DIR, "insurers_data.csv")
INSURERS_XLSX = os.path.join(APP_DIR, "insurers_data.xlsx")
VEHICLES_CSV = os.path.join(APP_DIR, "vehicle_valuations.csv")
VEHICLES_XLSX = os.path.join(APP_DIR, "vehicle_valuations.xlsx")

# --------------------------------------------------------------------------
# CLIENT-FACING ENTERPRISE THEME & STYLING
# --------------------------------------------------------------------------
st.markdown(
    textwrap.dedent(
        """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        color: #0F172A;
        background-color: #F8FAFC;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 3rem;
        padding-left: 1.2rem;
        padding-right: 1.2rem;
        max-width: 1240px;
    }

    /* Top Regulatory Banner */
    .ira-trust-strip {
        background: #0F172A;
        color: #94A3B8;
        font-size: 0.76rem;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        padding: 6px 16px;
        border-radius: 8px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 1.2rem;
    }
    .ira-trust-strip span {
        color: #38BDF8;
        font-weight: 700;
    }

    /* Header Presentation */
    .brand-container {
        display: flex;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 1rem;
        margin-bottom: 1.5rem;
    }
    .brand-title {
        font-size: 1.85rem;
        font-weight: 800;
        color: #0A2540;
        letter-spacing: -0.03em;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    .brand-sub {
        font-size: 0.9rem;
        color: #64748B;
        margin-top: 0.2rem;
        margin-bottom: 0;
    }
    .brand-pill {
        background: #F1F5F9;
        border: 1px solid #CBD5E1;
        color: #334155;
        font-size: 0.75rem;
        font-weight: 700;
        padding: 6px 14px;
        border-radius: 999px;
    }

    /* Configurator Container */
    .configurator-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 2.2rem;
        box-shadow: 0 4px 24px -2px rgba(15, 23, 42, 0.06);
        margin-bottom: 2rem;
    }
    .step-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #EFF6FF;
        color: #1D4ED8;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.04em;
        text-transform: uppercase;
        padding: 4px 12px;
        border-radius: 999px;
        margin-bottom: 0.8rem;
    }
    .section-lead {
        font-size: 1.3rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.02em;
        margin-bottom: 0.4rem;
    }
    .section-desc {
        font-size: 0.88rem;
        color: #64748B;
        margin-bottom: 1.6rem;
    }

    /* Sub-card containers */
    .sub-panel {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-top: 1rem;
        margin-bottom: 1.2rem;
    }

    /* TPO Warning Banner */
    .tpo-banner-box {
        background: #FFFBEB;
        border: 1.5px solid #F59E0B;
        border-radius: 12px;
        padding: 1rem 1.25rem;
        margin-top: 0.6rem;
        margin-bottom: 1.2rem;
        font-size: 0.88rem;
        color: #92400E;
        line-height: 1.5;
    }

    /* Stage 2 Executive Summary Bar */
    .exec-summary-ribbon {
        background: #0A2540;
        color: #FFFFFF;
        border-radius: 14px;
        padding: 1.1rem 1.5rem;
        margin-bottom: 1.8rem;
        box-shadow: 0 8px 24px -4px rgba(10, 37, 64, 0.2);
    }
    .exec-summary-title {
        font-size: 1.2rem;
        font-weight: 800;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 0.6rem;
    }
    .exec-summary-meta {
        font-size: 0.85rem;
        color: #94A3B8;
        margin-top: 0.4rem;
        line-height: 1.5;
    }
    .exec-summary-meta strong {
        color: #38BDF8;
    }

    /* Underwriter Quotation Cards */
    .quote-tile {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 1.3rem 1.4rem;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.04);
        margin-bottom: 1.2rem;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .quote-tile:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px -4px rgba(15, 23, 42, 0.08);
    }
    .quote-tile.highlighted {
        border: 2px solid #059669;
        box-shadow: 0 8px 28px -4px rgba(5, 150, 105, 0.18);
        position: relative;
    }
    .best-rate-flag {
        background: #059669;
        color: #FFFFFF;
        font-size: 0.68rem;
        font-weight: 800;
        letter-spacing: 0.05em;
        text-transform: uppercase;
        padding: 4px 10px;
        border-radius: 999px;
        float: right;
    }
    .carrier-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #0F172A;
        letter-spacing: -0.01em;
        margin-top: 0.2rem;
        margin-bottom: 0.3rem;
        min-height: 2.4rem;
    }
    .carrier-type-pill {
        display: inline-block;
        font-size: 0.72rem;
        font-weight: 700;
        padding: 3px 10px;
        border-radius: 999px;
        margin-bottom: 0.8rem;
    }
    .pill-comprehensive { background: #ECFDF5; color: #047857; border: 1px solid #A7F3D0; }
    .pill-tpo { background: #FFFBEB; color: #B45309; border: 1px solid #FDE68A; }

    .premium-caption {
        font-size: 0.72rem;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-bottom: 0.1rem;
    }
    .premium-amount {
        font-size: 1.85rem;
        font-weight: 800;
        color: #0A2540;
        line-height: 1.15;
        margin-bottom: 0.8rem;
    }
    .premium-amount.highlight-val {
        color: #059669;
    }

    .tariff-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.82rem;
        color: #475467;
        padding: 4px 0;
        border-bottom: 1px dashed #E2E8F0;
    }
    .tariff-row span:last-child {
        font-weight: 600;
        color: #1E293B;
    }
    .tariff-row-total {
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 0.9rem;
        font-weight: 800;
        color: #0F172A;
        padding-top: 8px;
        margin-top: 4px;
    }

    /* Communication Drawer Buttons */
    .whatsapp-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background-color: #25D366;
        color: #FFFFFF !important;
        font-weight: 700;
        font-size: 0.92rem;
        border-radius: 10px;
        padding: 0.75rem 1.2rem;
        text-decoration: none;
        width: 100%;
        text-align: center;
        margin-top: 0.6rem;
        box-shadow: 0 4px 14px rgba(37, 211, 102, 0.28);
        transition: background-color 0.2s ease;
    }
    .whatsapp-btn:hover {
        background-color: #1EBE5D;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-left: 0.6rem;
            padding-right: 0.6rem;
        }
        .configurator-panel {
            padding: 1.2rem;
        }
        .premium-amount {
            font-size: 1.6rem;
        }
        .brand-container {
            flex-direction: column;
            align-items: flex-start;
            gap: 0.6rem;
        }
    }
    </style>
    """
    ),
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# STATUTORY RATES & KENYA REGULATORY SCHEDULE
# --------------------------------------------------------------------------
BASE_BENCHMARK_YEAR = 2025
TRAINING_LEVY_RATE = 0.20 / 100
PHCF_RATE = 0.25 / 100
TOTAL_LEVY_RATE = TRAINING_LEVY_RATE + PHCF_RATE  # 0.45%
STAMP_DUTY = 40.0
COMPREHENSIVE_MAX_AGE_YEARS = 15
CURRENT_YEAR = datetime.now().year


def get_regulatory_depreciation_rate(years_diff: int) -> float:
    if years_diff <= 0:
        return 0.00
    schedule = {
        1: 0.10,
        2: 0.15,
        3: 0.20,
        4: 0.30,
        5: 0.40,
        6: 0.50,
        7: 0.60,
        8: 0.70,
    }
    if years_diff in schedule:
        return schedule[years_diff]
    extra_years = years_diff - 8
    return min(0.70 + (extra_years * 0.025), 0.85)


def estimate_regulatory_sum_insured(base_value_2025: float, manufacture_year: int) -> tuple[float, float]:
    years_diff = max(BASE_BENCHMARK_YEAR - manufacture_year, 0)
    dep_rate = get_regulatory_depreciation_rate(years_diff)
    depreciated = base_value_2025 * (1.0 - dep_rate)
    return max(round(depreciated, -3), 20000.0), dep_rate


def vehicle_age(manufacture_year: int) -> int:
    return max(CURRENT_YEAR - manufacture_year, 0)


def check_is_motorcycle(category_name: str) -> bool:
    cat_lower = str(category_name).lower()
    return any(kw in cat_lower for kw in ["motorcycle", "bike", "boda", "scooter", "two wheeler"])


def format_vehicle_display_name(make: str, model: str) -> str:
    mk = str(make).strip()
    md = str(model).strip()
    if not mk or mk.lower() in md.lower():
        return md
    return f"{mk} {md}"


def fmt_kes(amount) -> str:
    try:
        return f"KES {float(amount):,.2f}"
    except (TypeError, ValueError):
        return "KES 0.00"


# --------------------------------------------------------------------------
# STRICT DATA LOADING FROM EXCEL / CSV
# --------------------------------------------------------------------------
def safe_read_data(csv_path: str, xlsx_path: str) -> pd.DataFrame:
    if os.path.exists(xlsx_path):
        try:
            return pd.read_excel(xlsx_path, engine="openpyxl")
        except Exception:
            pass

    if os.path.exists(csv_path):
        for enc in ["utf-8", "latin1", "cp1252", "iso-8859-1"]:
            try:
                return pd.read_csv(csv_path, encoding=enc)
            except (UnicodeDecodeError, UnicodeError):
                continue
        return pd.read_csv(csv_path, encoding="utf-8", encoding_errors="replace")

    return pd.DataFrame()


@st.cache_data
def load_backend_data():
    ins_df = safe_read_data(INSURERS_CSV, INSURERS_XLSX)
    if ins_df.empty:
        ins_df = pd.DataFrame([
            {
                "Insurer": "Old Mutual Kenya",
                "Rate_Min (%)": 4.0,
                "Rate_Max (%)": 4.75,
                "Rate_Default (%)": 4.38,
                "Min_Premium (KES)": 35000,
                "TPO_Rate (KES)": 7500,
                "Windscreen_Limit (KES)": 50000,
                "Towing_Limit (KES)": 75000,
                "Key_Inclusions": "Collision, fire, theft, special perils",
                "Key_Exclusions": "Commercial use without endorsement, intoxicated driving",
                "Claims_Email": "generalclaims@oldmutual.co.ke",
                "Claims_Portal": "https://www.oldmutual.co.ke/personal/claims/",
            }
        ])

    veh_df = safe_read_data(VEHICLES_CSV, VEHICLES_XLSX)
    if veh_df.empty:
        veh_df = pd.DataFrame([
            {
                "Make": "BAJAJ",
                "Model": "Boxer BM 100",
                "Category": "Motorcycle",
                "Engine/Classification": "100cc Commuter",
                "Base Value 2025(KES)": "135000",
            }
        ])

    if "Model" not in veh_df.columns and "Vehicle Model" in veh_df.columns:
        veh_df["Model"] = veh_df["Vehicle Model"]

    if "Make" not in veh_df.columns:
        veh_df["Make"] = veh_df["Model"].astype(str).apply(lambda x: x.split()[0] if x else "Standard")
    else:
        veh_df["Make"] = veh_df["Make"].astype(str).str.strip()

    veh_df["Model"] = veh_df["Model"].astype(str).str.strip()
    veh_df["Category"] = veh_df["Category"].astype(str).str.strip()

    col_2025 = [c for c in veh_df.columns if "2025" in str(c)]
    target_col = col_2025[0] if col_2025 else "Base Value 2025(KES)"

    if target_col in veh_df.columns:
        veh_df["parsed_base_value"] = (
            veh_df[target_col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.strip()
        )
        veh_df["parsed_base_value"] = pd.to_numeric(veh_df["parsed_base_value"], errors="coerce").fillna(1000000.0)
    else:
        veh_df["parsed_base_value"] = 1000000.0

    return (
        ins_df.dropna(subset=["Insurer"]).reset_index(drop=True),
        veh_df.dropna(subset=["Model"]).reset_index(drop=True),
    )


insurers_df, vehicles_df = load_backend_data()


# --------------------------------------------------------------------------
# PREMIUM COMPUTATION (STRICT COMPREHENSIVE VS TPO RULES)
# --------------------------------------------------------------------------
def calculate_quote(
    sum_insured: float,
    rate_pct: float,
    min_premium_sheet: float,
    tpo_flat_sheet: float,
    is_comprehensive: bool,
    is_motorcycle: bool,
    insurer_windscreen_free: float,
    declared_windscreen: float,
    declared_entertainment: float,
    pvt_selected: bool,
    excess_protector_selected: bool,
    rescue_plus_selected: bool,
    pa_cover_selected: bool,
    courtesy_car_selected: bool,
    medical_expenses_selected: bool,
) -> dict:
    if sum_insured <= 0:
        raise ValueError("Sum Insured must be greater than zero.")

    # 1. Base Premium and Applied Rate
    if is_comprehensive:
        if is_motorcycle:
            effective_moto_rate = min(rate_pct, 3.50) if rate_pct > 0 else 3.50
            raw_base = sum_insured * (effective_moto_rate / 100.0)
            basic_premium = max(raw_base, 7500.0)
            applied_rate = effective_moto_rate
        else:
            raw_base = sum_insured * (rate_pct / 100.0)
            basic_premium = max(raw_base, min_premium_sheet)
            applied_rate = rate_pct

        # Comprehensive riders
        if is_motorcycle:
            effective_windscreen_limit = 0.0
            extra_windscreen_fee = 0.0
            effective_entertainment_limit = 0.0
            extra_entertainment_fee = 0.0
            courtesy_car_fee = 0.0
        else:
            effective_windscreen_limit = max(insurer_windscreen_free, declared_windscreen)
            extra_windscreen_fee = max(0.0, declared_windscreen - insurer_windscreen_free) * 0.10

            standard_entertainment_free = 30000.0
            effective_entertainment_limit = max(standard_entertainment_free, declared_entertainment)
            extra_entertainment_fee = max(0.0, declared_entertainment - standard_entertainment_free) * 0.10
            courtesy_car_fee = 3000.0 if courtesy_car_selected else 0.0

        pvt_fee = 5000.0 if pvt_selected else 0.0
        excess_fee = max(sum_insured * 0.0025, 2500.0) if excess_protector_selected else 0.0

    else:
        # THIRD PARTY ONLY (TPO) RULES:
        basic_premium = tpo_flat_sheet
        applied_rate = 0.0
        effective_windscreen_limit = 0.0
        extra_windscreen_fee = 0.0
        effective_entertainment_limit = 0.0
        extra_entertainment_fee = 0.0
        courtesy_car_fee = 0.0
        pvt_fee = 0.0
        excess_fee = 0.0

    # Common policy add-ons available across both covers
    pa_fee = 2000.0 if pa_cover_selected else 0.0
    medical_fee = 1000.0 if medical_expenses_selected else 0.0
    rescue_fee = 1000.0 if rescue_plus_selected else 0.0

    # 2. Statutory Levies (0.45% on insurance premium, excluding non-insurance rescue service)
    insurance_premium_for_levies = (
        basic_premium
        + extra_windscreen_fee
        + extra_entertainment_fee
        + pvt_fee
        + excess_fee
        + pa_fee
        + courtesy_car_fee
        + medical_fee
    )
    levies = insurance_premium_for_levies * TOTAL_LEVY_RATE
    stamp_duty = STAMP_DUTY

    # 3. Total Premium Payable
    total = insurance_premium_for_levies + rescue_fee + levies + stamp_duty
    rounded_total = round(total)

    return {
        "basic_premium": basic_premium,
        "applied_rate": applied_rate,
        "levies": levies,
        "stamp_duty": stamp_duty,
        "windscreen_limit": effective_windscreen_limit,
        "extra_windscreen_fee": extra_windscreen_fee,
        "entertainment_limit": effective_entertainment_limit,
        "extra_entertainment_fee": extra_entertainment_fee,
        "pvt_fee": pvt_fee,
        "excess_fee": excess_fee,
        "rescue_fee": rescue_fee,
        "pa_fee": pa_fee,
        "courtesy_car_fee": courtesy_car_fee,
        "medical_fee": medical_fee,
        "total_extensions": (
            extra_windscreen_fee
            + extra_entertainment_fee
            + pvt_fee
            + excess_fee
            + rescue_fee
            + pa_fee
            + courtesy_car_fee
            + medical_fee
        ),
        "total": float(rounded_total),
    }


# --------------------------------------------------------------------------
# UPGRADED BROKER-GRADE WHATSAPP FORMATTER
# --------------------------------------------------------------------------
def build_whatsapp_summary(quote_data: dict, vehicle_data: dict, is_moto: bool, is_comprehensive: bool, ext_list: list[str]) -> str:
    v_name = vehicle_data.get("Vehicle Name", "Motor Vehicle")
    v_year = vehicle_data.get("Year of Manufacture", "")
    v_sum = vehicle_data.get("Sum Insured (KES)", 0)
    c_type = vehicle_data.get("Cover Type Applied", "Comprehensive")
    icon = "🏍️" if is_moto else "🚗"
    asset_label = "Motorcycle" if is_moto else "Motor Vehicle"

    lines = [
        f"🛡️ *BIMAGRID KENYA | {asset_label.upper()} INSURANCE QUOTATION*",
        "━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"{icon} *Asset & Policy Scope:*",
        f"• *Vehicle:* {v_name} ({v_year})",
        f"• *Cover Type:* *{c_type}*",
    ]

    if is_comprehensive:
        lines.append(f"• *Agreed Sum Insured:* {fmt_kes(v_sum)}")
        if ext_list:
            lines.append(f"• *Riders Included:* {', '.join(ext_list)}")

        lines.append("")
        lines.append("📋 *Key Cover Limits Included:*")
        if not is_moto:
            ws_val = vehicle_data.get("declared_windscreen", 50000.0)
            lines.append(f"• Windscreen: {fmt_kes(ws_val)} (Free Built-in: KES 50,000)")
            lines.append("• Audio/Entertainment: KES 30,000 Free")
        lines.append("• Towing & Breakdown Recovery: Included")
        lines.append("• Statutory Levies (0.45%) & Stamp Duty (KES 40): *Included*")

        lines.append("")
        lines.append("📊 *Comparative Premium Matrix (All-Inclusive):*")

        sorted_quotes = sorted(quote_data.items(), key=lambda x: x[1]["total"])
        medals = ["🥇", "🥈", "🥉"]

        for idx, (ins_name, q) in enumerate(sorted_quotes):
            bullet = medals[idx] if idx < 3 else "•"
            badge = " *(Best Value)*" if idx == 0 else ""
            lines.append(f"{bullet} *{ins_name}:* {fmt_kes(q['total'])}{badge}")

    else:
        # THIRD PARTY ONLY (TPO) FORMAT
        lines.extend([
            "• *Statutory Scope:* Unlimited Third-Party Bodily Injury/Death & Third-Party Property Damage up to KES 20M per Insurance Act Cap 487",
            "• *Exclusions:* Accidental Own Damage, Fire, Theft, Windscreen Glass (Exempt by Law)",
            "• *Inspection / Anti-Theft Tracker:* *Not Required (Instant Certificate)*",
            "",
            "📊 *Statutory TPO Annual Tariff:*",
        ])
        sorted_quotes = sorted(quote_data.items(), key=lambda x: x[1]["total"])
        best_tpo = sorted_quotes[0][1]["total"]
        lines.append(f"• *Standard Annual Premium:* *{fmt_kes(best_tpo)}*")
        lines.append("  _(Base Tariff: KES 7,500 | Statutory Levies: KES 33.75 | Stamp Duty: KES 40.00)_")

    lines.extend([
        "━━━━━━━━━━━━━━━━━━━━━━━━━",
        "📲 *Next Steps to Bind Cover:*",
        "Reply with a copy of your vehicle *Logbook* and *National ID / KRA PIN* to receive your formal quotation schedule and digital certificate.",
    ])

    return "\n".join(lines)


# --------------------------------------------------------------------------
# EXPORT TO EXCEL (WITHOUT APPLIED RATE, CLAIMS EMAIL & PORTAL)
# --------------------------------------------------------------------------
def build_excel(quote_data: dict, vehicle_data: dict, is_moto: bool, is_comprehensive: bool) -> bytes:
    buffer = io.BytesIO()
    v_df = pd.DataFrame(list(vehicle_data.items()), columns=["Field", "Value"])
    comp_rows = []
    for ins_name, item in quote_data.items():
        row_dict = {
            "Insurer": ins_name,
            "Cover Type": item["cover_type"],
            "Base Premium (KES)": round(item["basic_premium"], 2),
        }
        if is_comprehensive and not is_moto:
            row_dict["Windscreen Limit (KES)"] = item["windscreen_limit"]
            row_dict["Extra Windscreen Fee (KES)"] = round(item["extra_windscreen_fee"], 2)
            row_dict["Entertainment Limit (KES)"] = item["entertainment_limit"]
            row_dict["Extra Entertainment Fee (KES)"] = round(item["extra_entertainment_fee"], 2)
            row_dict["Courtesy Car (KES)"] = round(item["courtesy_car_fee"], 2)

        row_dict["Rescue Recovery (KES)"] = round(item["rescue_fee"], 2)
        if is_comprehensive:
            row_dict["Excess Protector (KES)"] = round(item["excess_fee"], 2)
            row_dict["PVT Add-on (KES)"] = round(item["pvt_fee"], 2)

        row_dict["Passenger / Pillion PA (KES)"] = round(item["pa_fee"], 2)
        row_dict["Medical Expenses (KES)"] = round(item["medical_fee"], 2)
        row_dict["TL & PHPF Levies (0.45%) (KES)"] = round(item["levies"], 2)
        row_dict["Stamp Duty (KES)"] = round(item["stamp_duty"], 2)
        row_dict["Total Premium (KES)"] = round(item["total"], 2)
        if is_comprehensive:
            row_dict["Towing Limit (KES)"] = item["towing"]

        comp_rows.append(row_dict)
    c_df = pd.DataFrame(comp_rows)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        v_df.to_excel(writer, sheet_name="Vehicle Details", index=False)
        c_df.to_excel(writer, sheet_name="Insurer Comparison", index=False)

    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# APP HEADER
# --------------------------------------------------------------------------
st.markdown(
    """
    <div class="ira-trust-strip">
        <div>🇰🇪 Licensed Motor Insurance Intermediary Platform &nbsp;|&nbsp; Regulated by IRA Kenya</div>
        <div>Statutory Schedule <span>Cap 487 Compliant</span></div>
    </div>
    <div class="brand-container">
        <div>
            <h1 class="brand-title">🛡️ BimaGrid Kenya</h1>
            <p class="brand-sub">Executive Underwriting & Multi-Carrier Comparative Quotation Terminal</p>
        </div>
        <div class="brand-pill">
            Institutional Broker Desk v2.6
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

if "journey_stage" not in st.session_state:
    st.session_state["journey_stage"] = "intake"


# ==========================================================================
# JOURNEY STAGE 1: INTAKE & CONFIGURATOR
# ==========================================================================
if st.session_state["journey_stage"] == "intake":

    def on_stage1_category_change():
        cat = st.session_state.get("cfg_category")
        cat_df = vehicles_df[vehicles_df["Category"] == cat]
        makes = sorted(cat_df["Make"].dropna().unique().tolist())
        st.session_state["cfg_make"] = makes[0] if makes else ""

        make_df = cat_df[cat_df["Make"] == st.session_state["cfg_make"]]
        models = sorted(make_df["Model"].dropna().unique().tolist())
        st.session_state["cfg_model"] = models[0] if models else ""

        sub = make_df[make_df["Model"] == st.session_state["cfg_model"]]
        new_base = float(sub.iloc[0]["parsed_base_value"]) if not sub.empty else 1000000.0
        st.session_state["cfg_base_val"] = new_base
        yr = int(st.session_state.get("cfg_year", 2021))
        val, _ = estimate_regulatory_sum_insured(new_base, yr)
        st.session_state["cfg_sum_insured"] = float(val)

    def on_stage1_make_change():
        cat = st.session_state.get("cfg_category")
        mk = st.session_state.get("cfg_make")
        make_df = vehicles_df[(vehicles_df["Category"] == cat) & (vehicles_df["Make"] == mk)]
        models = sorted(make_df["Model"].dropna().unique().tolist())
        st.session_state["cfg_model"] = models[0] if models else ""

        sub = make_df[make_df["Model"] == st.session_state["cfg_model"]]
        new_base = float(sub.iloc[0]["parsed_base_value"]) if not sub.empty else 1000000.0
        st.session_state["cfg_base_val"] = new_base
        yr = int(st.session_state.get("cfg_year", 2021))
        val, _ = estimate_regulatory_sum_insured(new_base, yr)
        st.session_state["cfg_sum_insured"] = float(val)

    def on_stage1_model_change():
        cat = st.session_state.get("cfg_category")
        mk = st.session_state.get("cfg_make")
        mod = st.session_state.get("cfg_model")
        sub = vehicles_df[
            (vehicles_df["Category"] == cat) & (vehicles_df["Make"] == mk) & (vehicles_df["Model"] == mod)
        ]
        if sub.empty:
            sub = vehicles_df[(vehicles_df["Category"] == cat) & (vehicles_df["Model"] == mod)]

        new_base = float(sub.iloc[0]["parsed_base_value"]) if not sub.empty else 1000000.0
        st.session_state["cfg_base_val"] = new_base
        yr = int(st.session_state.get("cfg_year", 2021))
        val, _ = estimate_regulatory_sum_insured(new_base, yr)
        st.session_state["cfg_sum_insured"] = float(val)

    def on_stage1_val_or_year_change():
        base = float(st.session_state.get("cfg_base_val", 1000000.0))
        yr = int(st.session_state.get("cfg_year", 2021))
        val, _ = estimate_regulatory_sum_insured(base, yr)
        st.session_state["cfg_sum_insured"] = float(val)

    def reset_all_overrides_to_benchmark():
        cat = st.session_state.get("cfg_category", vehicles_df["Category"].iloc[0])
        mk = st.session_state.get("cfg_make", vehicles_df["Make"].iloc[0])
        mod = st.session_state.get("cfg_model", vehicles_df["Model"].iloc[0])

        sub = vehicles_df[
            (vehicles_df["Category"] == cat) & (vehicles_df["Make"] == mk) & (vehicles_df["Model"] == mod)
        ]
        if sub.empty:
            sub = vehicles_df[(vehicles_df["Category"] == cat) & (vehicles_df["Model"] == mod)]

        benchmark_base = float(sub.iloc[0]["parsed_base_value"]) if not sub.empty else 1000000.0
        default_min_rate = float(insurers_df["Rate_Min (%)"].astype(float).min())

        st.session_state["cfg_base_val"] = benchmark_base
        st.session_state["cfg_year"] = 2021
        val, _ = estimate_regulatory_sum_insured(benchmark_base, 2021)
        st.session_state["cfg_sum_insured"] = float(val)
        st.session_state["cfg_cover"] = "Comprehensive"
        st.session_state["cfg_rate_mode"] = "Insurer Minimal Rates"
        st.session_state["cfg_custom_rate"] = default_min_rate
        st.session_state["cfg_windscreen_limit"] = 50000.0
        st.session_state["cfg_entertainment_limit"] = 30000.0
        st.session_state["cfg_motorcycle_alarm"] = "No"
        st.session_state["cfg_tracker"] = "No" if check_is_motorcycle(cat) else ("Yes" if val >= 1500000.0 else "No")
        st.session_state["cfg_rescue"] = True
        st.session_state["cfg_excess"] = False
        st.session_state["cfg_pvt"] = False
        st.session_state["cfg_pa"] = False
        st.session_state["cfg_courtesy"] = False
        st.session_state["cfg_medical"] = False

    st.markdown(
        """
        <div class="configurator-panel">
            <div class="step-pill">Stage 1 of 2: Risk Profile & Asset Configurator</div>
            <div class="section-lead">Vehicle Valuation & Coverage Scope</div>
            <div class="section-desc">Select client asset parameters to query live underwriter schedules and statutory tariffs.</div>
        """,
        unsafe_allow_html=True,
    )

    # 1. Cascading Vehicle Selection (Category -> Make -> Model)
    categories = sorted(vehicles_df["Category"].dropna().unique().tolist())
    prev_cat = st.session_state.get("cfg_category", categories[0] if categories else "")
    cat_idx = categories.index(prev_cat) if prev_cat in categories else 0

    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        selected_category = st.selectbox(
            "Asset Classification",
            categories,
            index=cat_idx,
            key="cfg_category",
            on_change=on_stage1_category_change,
        )

    is_motorcycle = check_is_motorcycle(selected_category)
    cat_vehicles = vehicles_df[vehicles_df["Category"] == selected_category]

    available_makes = sorted(cat_vehicles["Make"].dropna().unique().tolist())
    prev_make = st.session_state.get("cfg_make", available_makes[0] if available_makes else "")
    make_idx = available_makes.index(prev_make) if prev_make in available_makes else 0

    with col_c2:
        selected_make = st.selectbox(
            "Manufacturer / Make",
            available_makes,
            index=make_idx,
            key="cfg_make",
            on_change=on_stage1_make_change,
        )

    make_vehicles = cat_vehicles[cat_vehicles["Make"] == selected_make]
    available_models = sorted(make_vehicles["Model"].dropna().unique().tolist())
    prev_mod = st.session_state.get("cfg_model", available_models[0] if available_models else "")
    model_idx = available_models.index(prev_mod) if prev_mod in available_models else 0

    with col_c3:
        selected_model = st.selectbox(
            "Vehicle Model",
            available_models,
            index=model_idx,
            key="cfg_model",
            on_change=on_stage1_model_change,
        )

    v_match = make_vehicles[make_vehicles["Model"] == selected_model]
    backend_base_val = float(v_match.iloc[0]["parsed_base_value"]) if not v_match.empty else 1000000.0

    if "cfg_base_val" not in st.session_state:
        st.session_state["cfg_base_val"] = backend_base_val

    if "cfg_year" not in st.session_state:
        st.session_state["cfg_year"] = int(st.session_state.get("manufacture_year", 2021))

    if "cfg_sum_insured" not in st.session_state:
        reg_val, _ = estimate_regulatory_sum_insured(st.session_state["cfg_base_val"], st.session_state["cfg_year"])
        st.session_state["cfg_sum_insured"] = float(reg_val)

    # 2. Prominent Policy Cover Type Selector
    st.markdown("---")
    age = vehicle_age(int(st.session_state["cfg_year"]))
    force_tpo = age > COMPREHENSIVE_MAX_AGE_YEARS

    col_pol1, col_pol2 = st.columns([1, 1])
    with col_pol1:
        if force_tpo:
            st.warning(
                f"⚠️ Vehicle is {age} years old (> {COMPREHENSIVE_MAX_AGE_YEARS} yrs limit). "
                "Policy restricted to Third Party Only (TPO)."
            )
            selected_cover = "TPO"
        else:
            selected_cover = st.radio(
                "Requested Policy Form*",
                ["Comprehensive", "TPO"],
                index=0 if st.session_state.get("cfg_cover", "Comprehensive") == "Comprehensive" else 1,
                horizontal=True,
                key="cfg_cover",
                help="Comprehensive protects against own-damage, theft, fire, and third-party liabilities. TPO fulfills statutory third-party requirements only.",
            )

    is_comp_mode = selected_cover == "Comprehensive" and not force_tpo

    with col_pol2:
        if is_comp_mode:
            st.info("🛡️ **Comprehensive Cover Active**: Protects client asset against accidental collision, overturning, fire, theft, and third-party liabilities.")
        else:
            st.markdown(
                """
                <div class="tpo-banner-box">
                    ⚖️ <strong>Third Party Only (TPO) Active</strong>: Covers legal liability for third-party bodily injury, death, and property damage per Insurance Act Cap 487. Own damage, theft, windscreen glass, and vehicle inspection are legally exempt.
                </div>
                """,
                unsafe_allow_html=True,
            )

    # 3. Valuation & Pricing Inputs
    st.markdown("---")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.number_input(
            "Base Value 2025 (KES) — Benchmark",
            min_value=10000.0,
            step=25000.0,
            format="%.0f",
            key="cfg_base_val",
            on_change=on_stage1_val_or_year_change,
        )
        st.caption(f"{fmt_kes(st.session_state['cfg_base_val'])} (KRA/AKI benchmark baseline)")

    with col_v2:
        st.number_input(
            "Year of Manufacture",
            min_value=1990,
            max_value=CURRENT_YEAR,
            step=1,
            key="cfg_year",
            on_change=on_stage1_val_or_year_change,
        )
        years_diff = max(BASE_BENCHMARK_YEAR - int(st.session_state["cfg_year"]), 0)
        dep_pct = get_regulatory_depreciation_rate(years_diff) * 100
        st.caption(f"Statutory Depreciation: **{dep_pct:.0f}%** ({years_diff} years elapsed)")

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.number_input(
            "Sum Insured (KES) — Editable",
            min_value=10000.0,
            step=25000.0,
            format="%.0f",
            key="cfg_sum_insured",
        )
        if is_comp_mode:
            st.caption("Depreciated benchmark value. Adjust as needed to match client's agreed policy value.")
        else:
            st.caption("Reference asset valuation. (TPO utilizes statutory flat schedule).")

    min_sheet_rate = float(insurers_df["Rate_Min (%)"].astype(float).min())

    with col_p2:
        if is_comp_mode:
            rate_mode = st.radio(
                "Tariff Pricing Basis",
                ["Insurer Minimal Rates", "Custom Agent Rate Override"],
                index=0 if st.session_state.get("cfg_rate_mode", "Insurer Minimal Rates") == "Insurer Minimal Rates" else 1,
                horizontal=True,
                key="cfg_rate_mode",
            )
            if rate_mode == "Custom Agent Rate Override":
                st.number_input(
                    "Custom Comprehensive Rate (%)",
                    min_value=0.50,
                    max_value=15.00,
                    value=float(st.session_state.get("cfg_custom_rate", min_sheet_rate)),
                    step=0.05,
                    format="%.2f",
                    key="cfg_custom_rate",
                )
                st.caption("Uniform rate card override for commercial negotiations.")
        else:
            st.radio(
                "Tariff Pricing Basis",
                ["Fixed Statutory TPO Rate (KES 7,500)"],
                index=0,
                key="cfg_rate_mode_tpo_locked",
                disabled=True,
            )
            st.caption("Statutory rates are fixed across all licensed underwriters.")
            rate_mode = "backend"

    # 4. Underwriting & Anti-Theft Status (Comprehensive Only — Legally Exempt for TPO)
    if is_comp_mode:
        st.markdown("---")
        st.subheader("🔍 Pre-Cover Inspection & Anti-Theft Status")

        col_insp1, col_insp2 = st.columns(2)
        with col_insp1:
            selected_valuer = st.selectbox(
                "Preferred Inspection Valuer*",
                [
                    "Capital Alliance Valuers",
                    "Automobile Association of Kenya (AA Kenya)",
                    "Regent Automobile Valuers",
                    "Standard Automobile Assessors",
                    "To be Inspect",
                ],
                index=0,
                key="cfg_valuer",
                help="Accredited assessor responsible for establishing the pre-cover risk binder report.",
            )

        with col_insp2:
            if is_motorcycle:
                motorcycle_alarm = st.radio(
                    "Do you have a motorcycle alarm installed?*",
                    ["No", "Yes"],
                    index=1 if st.session_state.get("cfg_motorcycle_alarm", "No") == "Yes" else 0,
                    horizontal=True,
                    key="cfg_motorcycle_alarm",
                )
                tracker_status = st.radio(
                    "Do you have a tracking system installed?*",
                    ["No", "Yes"],
                    index=1 if st.session_state.get("cfg_tracker", "No") == "Yes" else 0,
                    horizontal=True,
                    key="cfg_tracker",
                )
            else:
                motorcycle_alarm = "N/A"
                tracker_status = st.radio(
                    "Is a Tracking System Installed?",
                    ["Yes", "No"],
                    index=0 if st.session_state.get("cfg_tracker", "Yes" if float(st.session_state["cfg_sum_insured"]) >= 1500000.0 else "No") == "Yes" else 1,
                    horizontal=True,
                    key="cfg_tracker",
                    help="Kenyan underwriters mandate telematics tracking for private vehicles with Sum Insured >= KES 1,500,000.",
                )
                if tracker_status == "No" and float(st.session_state["cfg_sum_insured"]) >= 1500000.0:
                    st.caption("⚠️ **Telematics Requirement**: Sum Insured is $\ge$ KES 1.5M. Tracking certification required before binder issuance.")
    else:
        selected_valuer = "Not Required (TPO)"
        motorcycle_alarm = "Exempt (TPO)"
        tracker_status = "No"

    # 5. Windscreen & Audio Coverage (Comprehensive Cars Only — Strictly Hidden for TPO & Bikes)
    if is_comp_mode and not is_motorcycle:
        st.markdown("---")
        st.subheader("🪟 Windscreen & 📻 Entertainment System Coverage")
        st.caption("Underwriter schedule includes free built-in thresholds. Excess limit declared is charged at standard 10%.")

        col_lim1, col_lim2 = st.columns(2)
        with col_lim1:
            declared_windscreen = st.number_input(
                "Windscreen Coverage Limit (KES)",
                min_value=0.0,
                value=float(st.session_state.get("cfg_windscreen_limit", 50000.0)),
                step=10000.0,
                format="%.0f",
                key="cfg_windscreen_limit",
                help="Standard policies provide KES 50,000 free. Additional declared limit charged at 10%.",
            )
            st.caption("Standard built-in: KES 50,000 free.")

        with col_lim2:
            declared_entertainment = st.number_input(
                "Entertainment / Audio System Coverage (KES)",
                min_value=0.0,
                value=float(st.session_state.get("cfg_entertainment_limit", 30000.0)),
                step=5000.0,
                format="%.0f",
                key="cfg_entertainment_limit",
                help="Standard policies provide up to KES 30,000 free for factory sound system.",
            )
            st.caption("Standard built-in: KES 30,000 free.")
    else:
        declared_windscreen = 0.0
        declared_entertainment = 0.0

    # 6. Optional Policy Extensions (Riders)
    st.markdown("---")
    st.subheader("🛡️ Policy Extensions & Rider Endorsements")

    if is_comp_mode:
        st.caption("Select optional riders to customize the comprehensive quote portfolio:")
        ext_r1_c1, ext_r1_c2, ext_r1_c3 = st.columns(3)
        with ext_r1_c1:
            rescue_label = "Motorcycle Roadside Recovery" if is_motorcycle else "Roadside Rescue Plus"
            opt_rescue = st.checkbox(
                rescue_label,
                value=st.session_state.get("cfg_rescue", True),
                key="cfg_rescue",
                help="24/7 accident towing and roadside assistance (KES 1,000 flat fee).",
            )
            st.caption("KES 1,000 (Levy-exempt service add-on)")

        with ext_r1_c2:
            opt_excess = st.checkbox(
                "Excess Protector (Own Damage)",
                value=st.session_state.get("cfg_excess", False),
                key="cfg_excess",
                help="Waives the deductible when submitting an own-damage claim.",
            )
            st.caption("0.25% of Sum Insured (min KES 2,500)")

        with ext_r1_c3:
            opt_pvt = st.checkbox(
                "Political Violence & Terrorism (PVT)",
                value=st.session_state.get("cfg_pvt", False),
                key="cfg_pvt",
                help="Covers civil disturbance, riots, strikes, and terrorism incidents.",
            )
            st.caption("KES 5,000 flat fee")

        ext_r2_c1, ext_r2_c2, ext_r2_c3 = st.columns(3)
        with ext_r2_c1:
            pa_label = "Rider & Pillion Personal Accident (PA)" if is_motorcycle else "Occupant Personal Accident (PA)"
            opt_pa = st.checkbox(
                pa_label,
                value=st.session_state.get("cfg_pa", False),
                key="cfg_pa",
                help="Covers accidental bodily injury and death for passengers/pillions.",
            )
            st.caption("KES 2,000 flat fee")

        with ext_r2_c2:
            if not is_motorcycle:
                opt_courtesy = st.checkbox(
                    "Courtesy Car / Loss of Use",
                    value=st.session_state.get("cfg_courtesy", False),
                    key="cfg_courtesy",
                    help="Provides a replacement vehicle for 10 days during garage repairs.",
                )
                st.caption("KES 3,000 (10 Days)")
            else:
                opt_courtesy = False

        with ext_r2_c3:
            opt_medical = st.checkbox(
                "Emergency Medical Expenses",
                value=st.session_state.get("cfg_medical", False),
                key="cfg_medical",
                help="Reimbursement for emergency treatment costs following a crash.",
            )
            st.caption("KES 1,000 flat fee")
    else:
        st.caption("Statutory TPO policies exclude own-damage riders. Compatible legal add-ons:")
        tpo_r1, tpo_r2, tpo_r3 = st.columns(3)
        with tpo_r1:
            rescue_label = "Motorcycle Roadside Recovery" if is_motorcycle else "Roadside Rescue Plus"
            opt_rescue = st.checkbox(
                rescue_label,
                value=st.session_state.get("cfg_rescue", True),
                key="cfg_rescue",
                help="24/7 breakdown recovery and towing.",
            )
            st.caption("KES 1,000 (Levy-exempt)")

        with tpo_r2:
            pa_label = "Rider & Pillion Personal Accident (PA)" if is_motorcycle else "Occupant Personal Accident (PA)"
            opt_pa = st.checkbox(
                pa_label,
                value=st.session_state.get("cfg_pa", False),
                key="cfg_pa",
                help="Emergency passenger bodily injury cover.",
            )
            st.caption("KES 2,000 flat fee")

        with tpo_r3:
            opt_medical = st.checkbox(
                "Emergency Medical Expenses",
                value=st.session_state.get("cfg_medical", False),
                key="cfg_medical",
                help="Outpatient emergency reimbursement.",
            )
            st.caption("KES 1,000 flat fee")

        opt_excess = False
        opt_pvt = False
        opt_courtesy = False

    st.write("")
    btn_col1, btn_col2 = st.columns([3, 1])
    with btn_col1:
        if st.button("Generate Comparative Quotation Matrix ➔", type="primary", use_container_width=True):
            st.session_state["category"] = selected_category
            st.session_state["is_motorcycle"] = is_motorcycle
            st.session_state["make"] = selected_make
            st.session_state["model"] = selected_model
            st.session_state["base_value"] = float(st.session_state["cfg_base_val"])
            st.session_state["manufacture_year"] = int(st.session_state["cfg_year"])
            st.session_state["sum_insured"] = float(st.session_state["cfg_sum_insured"])
            st.session_state["cover_type"] = selected_cover
            st.session_state["force_tpo"] = force_tpo
            st.session_state["rate_mode"] = "custom" if (is_comp_mode and rate_mode == "Custom Agent Rate Override") else "backend"
            st.session_state["custom_rate"] = float(st.session_state.get("cfg_custom_rate", min_sheet_rate))
            st.session_state["selected_valuer"] = selected_valuer
            st.session_state["motorcycle_alarm"] = motorcycle_alarm
            st.session_state["tracker_fitted"] = tracker_status == "Yes"
            st.session_state["declared_windscreen"] = float(declared_windscreen)
            st.session_state["declared_entertainment"] = float(declared_entertainment)
            st.session_state["opt_rescue"] = opt_rescue
            st.session_state["opt_excess"] = opt_excess
            st.session_state["opt_pvt"] = opt_pvt
            st.session_state["opt_pa"] = opt_pa
            st.session_state["opt_courtesy"] = opt_courtesy
            st.session_state["opt_medical"] = opt_medical
            st.session_state["journey_stage"] = "dashboard"
            st.rerun()

    with btn_col2:
        st.button(
            "🔄 Reset to Benchmark",
            on_click=reset_all_overrides_to_benchmark,
            use_container_width=True,
            help="Discards manual inputs and reloads official sheet values",
        )

    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()


# ==========================================================================
# JOURNEY STAGE 2: COMPARISON & DASHBOARD
# ==========================================================================
make = st.session_state["make"]
model = st.session_state["model"]
category = st.session_state["category"]
is_motorcycle = st.session_state["is_motorcycle"]
manufacture_year = st.session_state["manufacture_year"]
sum_insured = st.session_state["sum_insured"]
cover_type = st.session_state["cover_type"]
force_tpo = st.session_state["force_tpo"]
age_years = vehicle_age(manufacture_year)
is_comp = cover_type == "Comprehensive" and not force_tpo

rate_mode = st.session_state.get("rate_mode", "backend")
custom_rate = st.session_state.get("custom_rate", 3.50)

selected_valuer = st.session_state.get("selected_valuer", "Not Required (TPO)" if not is_comp else "Capital Alliance Valuers")
motorcycle_alarm = st.session_state.get("motorcycle_alarm", "Exempt (TPO)" if not is_comp else "No")
tracker_fitted = st.session_state.get("tracker_fitted", False)
declared_windscreen = st.session_state.get("declared_windscreen", 0.0)
declared_entertainment = st.session_state.get("declared_entertainment", 0.0)

opt_rescue = st.session_state.get("opt_rescue", True)
opt_excess = st.session_state.get("opt_excess", False)
opt_pvt = st.session_state.get("opt_pvt", False)
opt_pa = st.session_state.get("opt_pa", False)
opt_courtesy = st.session_state.get("opt_courtesy", False)
opt_medical = st.session_state.get("opt_medical", False)

active_extensions = []
if is_comp and not is_motorcycle:
    if declared_windscreen > 50000:
        active_extensions.append(f"Windscreen {fmt_kes(declared_windscreen)}")
    if declared_entertainment > 30000:
        active_extensions.append(f"Audio {fmt_kes(declared_entertainment)}")
if opt_rescue:
    active_extensions.append("Rescue Recovery" if is_motorcycle else "Rescue Plus")
if is_comp and opt_excess:
    active_extensions.append("Excess Protector")
if is_comp and opt_pvt:
    active_extensions.append("PVT")
if opt_pa:
    active_extensions.append("Pillion/Rider PA" if is_motorcycle else "Passenger PA")
if is_comp and opt_courtesy and not is_motorcycle:
    active_extensions.append("Courtesy Car")
if opt_medical:
    active_extensions.append("Medical")
ext_str = ", ".join(active_extensions) if active_extensions else "Standard Statutory Cover"

vehicle_icon = "🏍️" if is_motorcycle else "🚙"
full_display_name = format_vehicle_display_name(make, model)

# Context-Aware Executive Summary Ribbon
if not is_comp:
    summary_html = f"""
    <div class="exec-summary-ribbon">
        <div class="exec-summary-title">
            {vehicle_icon} {full_display_name} ({manufacture_year}) &nbsp;|&nbsp; 
            <span style="color:#F59E0B;">Third Party Only (TPO)</span>
        </div>
        <div class="exec-summary-meta">
            <strong>Policy Scope:</strong> Unlimited 3rd Party Bodily Injury/Death & 3rd Party Property Damage up to KES 20M per Cap 487.<br/>
            <strong>Inspection & Tracking:</strong> Legally Exempt (Instant Binder Issuance) &nbsp;|&nbsp; 
            <strong>Assigned Assessor:</strong> {selected_valuer}
        </div>
    </div>
    """
elif is_motorcycle:
    rate_label = f"{custom_rate:.2f}% (Custom Rate)" if rate_mode == "custom" else "Live Sheet Rate_Min (%)"
    summary_html = f"""
    <div class="exec-summary-ribbon">
        <div class="exec-summary-title">
            {vehicle_icon} {full_display_name} ({manufacture_year}) &nbsp;|&nbsp; 
            <span style="color:#38BDF8;">Comprehensive Coverage</span>
        </div>
        <div class="exec-summary-meta">
            <strong>Agreed Valuation:</strong> {fmt_kes(sum_insured)} &nbsp;|&nbsp; 
            <strong>Pricing Schedule:</strong> {rate_label} &nbsp;|&nbsp;
            <strong>Anti-Theft Alarm:</strong> {motorcycle_alarm} &nbsp;|&nbsp; 
            <strong>Telematics Tracker:</strong> {'Fitted' if tracker_fitted else 'Pending'} &nbsp;|&nbsp; 
            <strong>Assessor:</strong> {selected_valuer}
        </div>
    </div>
    """
else:
    rate_label = f"{custom_rate:.2f}% (Custom Rate)" if rate_mode == "custom" else "Live Sheet Rate_Min (%)"
    summary_html = f"""
    <div class="exec-summary-ribbon">
        <div class="exec-summary-title">
            {vehicle_icon} {full_display_name} ({manufacture_year}) &nbsp;|&nbsp; 
            <span style="color:#38BDF8;">Comprehensive Coverage</span>
        </div>
        <div class="exec-summary-meta">
            <strong>Agreed Valuation:</strong> {fmt_kes(sum_insured)} &nbsp;|&nbsp; 
            <strong>Pricing Schedule:</strong> {rate_label} &nbsp;|&nbsp; 
            <strong>Windscreen Cover:</strong> {fmt_kes(declared_windscreen)} &nbsp;|&nbsp; 
            <strong>Audio System:</strong> {fmt_kes(declared_entertainment)} &nbsp;|&nbsp; 
            <strong>Telematics Tracker:</strong> {'Fitted' if tracker_fitted else 'Not Fitted'} &nbsp;|&nbsp; 
            <strong>Assessor:</strong> {selected_valuer}
        </div>
    </div>
    """

top_col1, top_col2 = st.columns([3, 1])
with top_col1:
    st.markdown(summary_html, unsafe_allow_html=True)
with top_col2:
    if st.button("✏️ Modify Asset / Rates", type="primary", use_container_width=True):
        st.session_state["journey_stage"] = "intake"
        st.rerun()

# --------------------------------------------------------------------------
# QUOTE COMPUTATION ENGINE
# --------------------------------------------------------------------------
quotes = {}
calc_error = None

try:
    for _, row in insurers_df.iterrows():
        insurer_name = str(row["Insurer"])

        sheet_rate = custom_rate if (is_comp and rate_mode == "custom") else float(row.get("Rate_Min (%)", 3.50))
        sheet_min_premium = float(row.get("Min_Premium (KES)", 25000.0))
        sheet_tpo_rate = float(row.get("TPO_Rate (KES)", 7500.0))
        sheet_ws_free = float(row.get("Windscreen_Limit (KES)", 50000.0))

        result = calculate_quote(
            sum_insured=sum_insured,
            rate_pct=sheet_rate,
            min_premium_sheet=sheet_min_premium,
            tpo_flat_sheet=sheet_tpo_rate,
            is_comprehensive=is_comp,
            is_motorcycle=is_motorcycle,
            insurer_windscreen_free=sheet_ws_free,
            declared_windscreen=declared_windscreen,
            declared_entertainment=declared_entertainment,
            pvt_selected=opt_pvt,
            excess_protector_selected=opt_excess,
            rescue_plus_selected=opt_rescue,
            pa_cover_selected=opt_pa,
            courtesy_car_selected=opt_courtesy,
            medical_expenses_selected=opt_medical,
        )
        result["insurer"] = insurer_name
        result["towing"] = row.get("Towing_Limit (KES)", 30000)
        result["cover_type"] = "Comprehensive" if is_comp else "TPO"

        if is_comp:
            result["inclusions"] = row.get("Key_Inclusions", "")
            result["exclusions"] = row.get("Key_Exclusions", "")
        else:
            result["inclusions"] = "Third-party bodily injury & death (unlimited statutory liability), third-party property damage (up to KES 20M per Insurance Act Cap 487), legal defense costs."
            result["exclusions"] = "Accidental own damage, collision, vehicle overturn, fire, theft/hijacking, windscreen/window glass damage, and audio/accessories."

        quotes[insurer_name] = result
except Exception as e:
    calc_error = str(e)

if calc_error:
    st.error(f"Error computing quotes: {calc_error}")
    st.stop()

cheapest_insurer = min(quotes, key=lambda k: quotes[k]["total"])

vehicle_info = {
    "Vehicle Category": category,
    "Make": make,
    "Model": model,
    "Vehicle Name": full_display_name,
    "Is Motorcycle": is_motorcycle,
    "Year of Manufacture": manufacture_year,
    "Vehicle Age (years)": age_years,
    "Sum Insured (KES)": round(sum_insured, 2),
    "Cover Type Applied": "Comprehensive" if is_comp else "TPO",
    "Inspection Valuer": selected_valuer,
    "Quote Generated On": datetime.now().strftime("%Y-%m-%d %H:%M"),
    "declared_windscreen": declared_windscreen,
}

# --------------------------------------------------------------------------
# COMPARISON CARDS (EXECUTIVE CLIENT PRESENTATION GRID)
# --------------------------------------------------------------------------
NUM_PER_ROW = 3
insurer_names = list(quotes.keys())
for start in range(0, len(insurer_names), NUM_PER_ROW):
    row_names = insurer_names[start : start + NUM_PER_ROW]
    cols = st.columns(len(row_names))
    for col, insurer_name in zip(cols, row_names):
        q = quotes[insurer_name]
        is_best = insurer_name == cheapest_insurer

        card_class = "quote-tile highlighted" if is_best else "quote-tile"
        pill_class = "pill-comprehensive" if is_comp else "pill-tpo"
        best_flag = '<span class="best-rate-flag">★ UNDERWRITER PICK</span>' if is_best else ""
        prem_class = "premium-amount highlight-val" if is_best else "premium-amount"

        ext_breakdown_html = ""
        if q["total_extensions"] > 0:
            ext_breakdown_html = f'<div class="tariff-row"><span>Riders / Extensions</span><span>{fmt_kes(q["total_extensions"])}</span></div>'

        if not is_comp:
            badge_text = "Third Party Only (TPO)"
            limits_html = (
                '<div class="tariff-row"><span>Statutory Scope</span><span>Third-Party Liabilities</span></div>'
                '<div class="tariff-row"><span>Own Damage & Glass</span><span style="color:#B45309;">Legally Excluded</span></div>'
            )
        elif is_motorcycle:
            badge_text = f"Comprehensive ({q['applied_rate']:.2f}%)"
            limits_html = f'<div class="tariff-row"><span>Breakdown / Towing Limit</span><span>{fmt_kes(q["towing"])}</span></div>'
        else:
            badge_text = f"Comprehensive ({q['applied_rate']:.2f}%)"
            limits_html = (
                f'<div class="tariff-row"><span>Windscreen Limit</span><span>{fmt_kes(q["windscreen_limit"])}</span></div>'
                f'<div class="tariff-row"><span>Entertainment Limit</span><span>{fmt_kes(q["entertainment_limit"])}</span></div>'
                f'<div class="tariff-row"><span>Towing Limit</span><span>{fmt_kes(q["towing"])}</span></div>'
            )

        with col:
            card_html = "".join([
                f'<div class="{card_class}">',
                best_flag,
                f'<div class="carrier-title">{insurer_name}</div>',
                f'<span class="carrier-type-pill {pill_class}">{badge_text}</span>',
                '<div class="premium-caption">Annual Premium Payable</div>',
                f'<div class="{prem_class}">{fmt_kes(q["total"])}</div>',
                f'<div class="tariff-row"><span>Underwritten Base</span><span>{fmt_kes(q["basic_premium"])}</span></div>',
                ext_breakdown_html,
                f'<div class="tariff-row"><span>TL & PHPF Levies (0.45%)</span><span>{fmt_kes(q["levies"])}</span></div>',
                f'<div class="tariff-row"><span>Statutory Stamp Duty</span><span>{fmt_kes(q["stamp_duty"])}</span></div>',
                f'<div class="tariff-row-total"><span>Total Annual Outlay</span><span>{fmt_kes(q["total"])}</span></div>',
                "<br/>",
                limits_html,
                "</div>",
            ])
            st.markdown(card_html, unsafe_allow_html=True)
            with st.expander(f"Policy Schedule — {insurer_name}"):
                st.markdown(f"**Cover Inclusions:** {q['inclusions']}")
                st.markdown(f"**Policy Exclusions:** {q['exclusions']}")

st.divider()

# --------------------------------------------------------------------------
# EXPORT & CLIENT DISPATCH DRAWER
# --------------------------------------------------------------------------
excel_bytes = build_excel(quotes, vehicle_info, is_motorcycle, is_comp)
whatsapp_text = build_whatsapp_summary(quotes, vehicle_info, is_motorcycle, is_comp, active_extensions)
encoded_whatsapp_url = f"https://wa.me/?text={urllib.parse.quote(whatsapp_text)}"

col_export1, col_export2 = st.columns([1, 1])
with col_export1:
    st.markdown("#### 📄 Executive Client Deliverable")
    st.caption("Clean comparative rate sheet formatted for corporate & personal presentations (excludes internal broker metadata).")
    st.download_button(
        label="📥 Download Formal Quotation (.xlsx)",
        data=excel_bytes,
        file_name=f"{full_display_name.replace(' ', '_')}_{cover_type}_{manufacture_year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_export2:
    st.markdown("#### 📲 Instant WhatsApp Quotation Dispatch")
    st.caption("One-click executive briefing formatted for instant client sharing on mobile or WhatsApp Web.")
    with st.expander("Preview & Dispatch Client Summary", expanded=False):
        st.code(whatsapp_text, language=None)
        st.markdown(
            f'<a href="{encoded_whatsapp_url}" target="_blank" class="whatsapp-btn">💬 Open in WhatsApp & Dispatch</a>',
            unsafe_allow_html=True,
        )