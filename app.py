"""
Kenyan Motor Insurance Agent Tool (v2 — CSV-driven, fully editable)
====================================================================
A mobile-responsive, no-login Streamlit app for insurance agents to generate
instant, side-by-side motor insurance quotes across 10 Kenyan insurers.

All insurer tariffs and vehicle valuations are loaded from two CSV files
shipped alongside this script (insurers_data.csv, vehicle_valuations.csv).
Agents can edit every field live from the browser via data editors — no code
changes needed when an insurer updates their rates.

Run with:
    streamlit run app.py

Files expected in the same folder:
    insurers_data.csv
    vehicle_valuations.csv
"""

import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Kenya Motor Insurance Agent Tool",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_DIR = os.path.dirname(os.path.abspath(__file__))
INSURERS_CSV = os.path.join(APP_DIR, "insurers_data.csv")
VEHICLES_CSV = os.path.join(APP_DIR, "vehicle_valuations.csv")

# --------------------------------------------------------------------------
# CUSTOM CSS
# --------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 1300px;
    }
    .quote-card {
        background: #ffffff;
        border: 1px solid #E3E8EF;
        border-radius: 14px;
        padding: 1.1rem 1.2rem 1.3rem 1.2rem;
        box-shadow: 0 2px 10px rgba(16, 24, 40, 0.06);
        margin-bottom: 1rem;
        height: 100%;
    }
    .quote-card.best {
        border: 2px solid #1447E6;
        box-shadow: 0 4px 18px rgba(20, 71, 230, 0.18);
    }
    .insurer-name {
        font-size: 1.05rem;
        font-weight: 700;
        color: #101828;
        margin-bottom: 0.15rem;
        min-height: 2.6rem;
    }
    .cover-badge {
        display: inline-block;
        font-size: 0.72rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        padding: 2px 10px;
        border-radius: 999px;
        margin-bottom: 0.6rem;
    }
    .badge-comp { background: #E7F4E8; color: #1A7F37; }
    .badge-tpo  { background: #FFF3DC; color: #B45309; }
    .best-badge {
        background: #1447E6; color: #fff; font-size: 0.68rem; font-weight: 700;
        padding: 2px 9px; border-radius: 999px; float: right;
    }
    .total-premium-label {
        font-size: 0.78rem;
        color: #667085;
        margin-top: 0.6rem;
        margin-bottom: 0.05rem;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }
    .total-premium-value {
        font-size: 1.85rem;
        font-weight: 800;
        color: #1447E6;
        line-height: 1.15;
        margin-bottom: 0.4rem;
    }
    .breakdown-line {
        display: flex;
        justify-content: space-between;
        font-size: 0.82rem;
        color: #475467;
        padding: 2px 0;
        border-bottom: 1px dashed #EAECF0;
    }
    .breakdown-total {
        display: flex;
        justify-content: space-between;
        font-size: 0.85rem;
        font-weight: 700;
        color: #101828;
        padding-top: 4px;
    }
    .policy-detail {
        font-size: 0.78rem;
        color: #475467;
        margin-top: 0.3rem;
    }
    section[data-testid="stSidebar"] { background-color: #F9FAFB; }
    .app-title { font-size: 1.6rem; font-weight: 800; color: #101828; margin-bottom: 0; }
    .app-subtitle { color: #667085; font-size: 0.92rem; margin-top: 0.1rem; margin-bottom: 1.2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# CONSTANTS
# --------------------------------------------------------------------------
BASE_YEAR = 2017
VALUE_ADJUSTMENT_PER_YEAR = 0.10   # +/-10% per year vs. base-year value
TRAINING_LEVY_RATE = 0.20 / 100
PHCF_RATE = 0.25 / 100
STAMP_DUTY = 40.0
COMPREHENSIVE_MAX_AGE_YEARS = 15
CURRENT_YEAR = datetime.now().year

INSURER_COLS = [
    "Insurer", "Rate_Min (%)", "Rate_Max (%)", "Rate_Default (%)",
    "Min_Premium (KES)", "TPO_Rate (KES)", "Windscreen_Limit (KES)",
    "Towing_Limit (KES)", "Key_Inclusions", "Key_Exclusions",
    "Claims_Email", "Claims_Portal",
]
VEHICLE_COLS = ["Vehicle Model", "Category", "Engine/Classification", "Base Value 2017 (KES)"]


# --------------------------------------------------------------------------
# DATA LOADING (CSV -> session_state, editable thereafter)
# --------------------------------------------------------------------------
@st.cache_data
def load_default_insurers() -> pd.DataFrame:
    if os.path.exists(INSURERS_CSV):
        return pd.read_csv(INSURERS_CSV)
    # Fallback minimal defaults if the CSV is missing, so the app never crashes
    return pd.DataFrame([
        {"Insurer": "Jubilee", "Rate_Min (%)": 3.5, "Rate_Max (%)": 4.0, "Rate_Default (%)": 3.75,
         "Min_Premium (KES)": 15000, "TPO_Rate (KES)": 7500, "Windscreen_Limit (KES)": 50000,
         "Towing_Limit (KES)": 50000, "Key_Inclusions": "", "Key_Exclusions": "",
         "Claims_Email": "", "Claims_Portal": ""},
    ])


@st.cache_data
def load_default_vehicles() -> pd.DataFrame:
    if os.path.exists(VEHICLES_CSV):
        return pd.read_csv(VEHICLES_CSV)
    return pd.DataFrame([
        {"Vehicle Model": "Toyota Vitz", "Category": "Car",
         "Engine/Classification": "1300cc Hatchback", "Base Value 2017 (KES)": 1100000},
    ])


if "insurers_df" not in st.session_state:
    st.session_state.insurers_df = load_default_insurers().copy()

if "vehicles_df" not in st.session_state:
    st.session_state.vehicles_df = load_default_vehicles().copy()


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def estimate_sum_insured(base_value: float, manufacture_year: int) -> float:
    years_diff = manufacture_year - BASE_YEAR
    adjusted = base_value * (1 + VALUE_ADJUSTMENT_PER_YEAR * years_diff)
    return max(round(adjusted, -3), 0)


def vehicle_age(manufacture_year: int) -> int:
    return max(CURRENT_YEAR - manufacture_year, 0)


def calculate_quote(sum_insured: float, rate_pct: float, min_premium: float,
                     tpo_flat: float, is_comprehensive: bool) -> dict:
    """Formula: (Sum Insured * Rate) + Levies + Stamp Duty, per IRA rules."""
    if sum_insured <= 0:
        raise ValueError("Sum Insured must be greater than zero.")

    if is_comprehensive:
        basic_premium = max(sum_insured * (rate_pct / 100), min_premium)
    else:
        basic_premium = tpo_flat

    training_levy = basic_premium * TRAINING_LEVY_RATE
    phcf = basic_premium * PHCF_RATE
    total = basic_premium + training_levy + phcf + STAMP_DUTY

    return {
        "basic_premium": basic_premium,
        "training_levy": training_levy,
        "phcf": phcf,
        "stamp_duty": STAMP_DUTY,
        "total": total,
    }


def fmt_kes(amount) -> str:
    try:
        return f"KES {float(amount):,.0f}"
    except (TypeError, ValueError):
        return "KES 0"


# --------------------------------------------------------------------------
# HEADER
# --------------------------------------------------------------------------
st.markdown('<p class="app-title">🚗 Kenya Motor Insurance Agent Tool</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="app-subtitle">Instant, side-by-side motor insurance quotes across 10 insurers — '
    'edit rates and vehicle values live if your data changes, export to Excel, share on WhatsApp.</p>',
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# EDIT DATA (FRONT-END, CSV-BACKED)
# --------------------------------------------------------------------------
with st.expander("✏️ Edit Insurer Rates & Benefits (updates instantly)", expanded=False):
    st.caption(
        "Edit any cell directly in the table below — rate, minimum premium, TPO rate, limits, "
        "inclusions/exclusions, or claims contacts. Changes apply immediately to every quote card, "
        "the Excel export, and the WhatsApp summary. Use 'Reset to Original Data' to undo."
    )
    edited_insurers = st.data_editor(
        st.session_state.insurers_df,
        column_config={
            "Rate_Min (%)": st.column_config.NumberColumn(format="%.2f", min_value=0.0, max_value=15.0, step=0.05),
            "Rate_Max (%)": st.column_config.NumberColumn(format="%.2f", min_value=0.0, max_value=15.0, step=0.05),
            "Rate_Default (%)": st.column_config.NumberColumn(
                "Rate Used for Quote (%)", format="%.2f", min_value=0.0, max_value=15.0, step=0.05
            ),
            "Min_Premium (KES)": st.column_config.NumberColumn(format="%d", min_value=0, step=500),
            "TPO_Rate (KES)": st.column_config.NumberColumn(format="%d", min_value=0, step=250),
            "Windscreen_Limit (KES)": st.column_config.NumberColumn(format="%d", min_value=0, step=5000),
            "Towing_Limit (KES)": st.column_config.NumberColumn(format="%d", min_value=0, step=5000),
        },
        num_rows="dynamic",   # agent can add/remove insurers
        use_container_width=True,
        key="insurer_editor",
    )
    st.session_state.insurers_df = edited_insurers

    if st.button("↺ Reset Insurer Data to Original CSV"):
        load_default_insurers.clear()
        st.session_state.insurers_df = load_default_insurers().copy()
        st.rerun()

with st.expander("✏️ Edit Vehicle Valuation Base Values (updates instantly)", expanded=False):
    st.caption(
        "Base values are for the year 2017 and adjusted "
        f"±{VALUE_ADJUSTMENT_PER_YEAR*100:.0f}% per year of difference to estimate Sum Insured."
    )
    edited_vehicles = st.data_editor(
        st.session_state.vehicles_df,
        column_config={
            "Base Value 2017 (KES)": st.column_config.NumberColumn(format="%d", min_value=0, step=10000),
        },
        num_rows="dynamic",
        use_container_width=True,
        key="vehicle_editor",
    )
    st.session_state.vehicles_df = edited_vehicles

    if st.button("↺ Reset Vehicle Data to Original CSV"):
        load_default_vehicles.clear()
        st.session_state.vehicles_df = load_default_vehicles().copy()
        st.rerun()

insurers_df = st.session_state.insurers_df.dropna(subset=["Insurer"]).reset_index(drop=True)
vehicles_df = st.session_state.vehicles_df.dropna(subset=["Vehicle Model"]).reset_index(drop=True)

if insurers_df.empty:
    st.error("No insurer data available. Please add at least one insurer in the edit table above.")
    st.stop()
if vehicles_df.empty:
    st.error("No vehicle data available. Please add at least one vehicle in the edit table above.")
    st.stop()

# --------------------------------------------------------------------------
# SIDEBAR — PRIMARY INPUTS
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("🧾 Vehicle Details")

    categories = sorted(vehicles_df["Category"].dropna().unique().tolist())
    vehicle_category = st.radio("Vehicle Category", categories, horizontal=True)

    category_vehicles = vehicles_df[vehicles_df["Category"] == vehicle_category]
    model = st.selectbox("Model", category_vehicles["Vehicle Model"].tolist())

    manufacture_year = st.number_input(
        "Year of Manufacture",
        min_value=1990,
        max_value=CURRENT_YEAR,
        value=2019,
        step=1,
    )

    vehicle_row = category_vehicles[category_vehicles["Vehicle Model"] == model].iloc[0]
    base_value = float(vehicle_row["Base Value 2017 (KES)"])
    auto_value = estimate_sum_insured(base_value, int(manufacture_year))

    st.caption(
        f"Auto-valuation: {model} ({BASE_YEAR} base value {fmt_kes(base_value)}) "
        f"adjusted {VALUE_ADJUSTMENT_PER_YEAR*100:.0f}%/year → **{fmt_kes(auto_value)}**"
    )

    sum_insured = st.number_input(
        "Sum Insured (KES) — auto-filled, editable",
        min_value=0.0,
        value=float(auto_value),
        step=10000.0,
        format="%.0f",
    )

    age_years = vehicle_age(int(manufacture_year))
    force_tpo = age_years > COMPREHENSIVE_MAX_AGE_YEARS

    if force_tpo:
        st.warning(
            f"⚠️ Vehicle is {age_years} years old (>{COMPREHENSIVE_MAX_AGE_YEARS} yrs). "
            "Comprehensive cover disabled — Third Party Only (TPO) quotes will be shown."
        )
        cover_type = "TPO"
    else:
        cover_type = st.radio("Cover Type", ["Comprehensive", "TPO"], horizontal=True)

    st.divider()
    st.caption(
        "Statutory levies applied to every quote: "
        f"Training Levy {TRAINING_LEVY_RATE*100:.2f}%, "
        f"PHCF {PHCF_RATE*100:.2f}%, "
        f"Stamp Duty {fmt_kes(STAMP_DUTY)}."
    )
    st.caption(f"Loaded {len(insurers_df)} insurers from your data.")

# --------------------------------------------------------------------------
# QUOTE ENGINE
# --------------------------------------------------------------------------
st.subheader("📊 Quote Comparison")

quotes = {}
calc_error = None

try:
    for _, row in insurers_df.iterrows():
        insurer_name = str(row["Insurer"])
        is_comprehensive = cover_type == "Comprehensive" and not force_tpo
        result = calculate_quote(
            sum_insured=sum_insured,
            rate_pct=float(row["Rate_Default (%)"]),
            min_premium=float(row["Min_Premium (KES)"]),
            tpo_flat=float(row["TPO_Rate (KES)"]),
            is_comprehensive=is_comprehensive,
        )
        result["insurer"] = insurer_name
        result["windscreen"] = row.get("Windscreen_Limit (KES)", 0)
        result["towing"] = row.get("Towing_Limit (KES)", 0)
        result["inclusions"] = row.get("Key_Inclusions", "")
        result["exclusions"] = row.get("Key_Exclusions", "")
        result["claims_email"] = row.get("Claims_Email", "")
        result["cover_type"] = "Comprehensive" if is_comprehensive else "TPO"
        quotes[insurer_name] = result
except ValueError as e:
    calc_error = str(e)

if calc_error:
    st.error(f"⚠️ {calc_error} Please enter a Sum Insured greater than zero in the sidebar.")
else:
    cheapest_insurer = min(quotes, key=lambda k: quotes[k]["total"])

    NUM_PER_ROW = 3
    insurer_names = list(quotes.keys())
    for start in range(0, len(insurer_names), NUM_PER_ROW):
        row_names = insurer_names[start:start + NUM_PER_ROW]
        cols = st.columns(len(row_names))
        for col, insurer_name in zip(cols, row_names):
            q = quotes[insurer_name]
            is_best = insurer_name == cheapest_insurer
            badge_class = "badge-comp" if q["cover_type"] == "Comprehensive" else "badge-tpo"
            best_tag = '<span class="best-badge">BEST PRICE</span>' if is_best else ""
            card_class = "quote-card best" if is_best else "quote-card"

            with col:
                st.markdown(
                    f"""
                    <div class="{card_class}">
                        {best_tag}
                        <div class="insurer-name">{insurer_name}</div>
                        <span class="cover-badge {badge_class}">{q['cover_type']}</span>
                        <div class="total-premium-label">Total Premium Payable</div>
                        <div class="total-premium-value">{fmt_kes(q['total'])}</div>
                        <div class="breakdown-line"><span>Basic Premium</span><span>{fmt_kes(q['basic_premium'])}</span></div>
                        <div class="breakdown-line"><span>Training Levy (0.20%)</span><span>{fmt_kes(q['training_levy'])}</span></div>
                        <div class="breakdown-line"><span>PHCF (0.25%)</span><span>{fmt_kes(q['phcf'])}</span></div>
                        <div class="breakdown-line"><span>Stamp Duty</span><span>{fmt_kes(q['stamp_duty'])}</span></div>
                        <div class="breakdown-total"><span>Total</span><span>{fmt_kes(q['total'])}</span></div>
                        <br/>
                        <div class="breakdown-line"><span>Windscreen Limit</span><span>{fmt_kes(q['windscreen'])}</span></div>
                        <div class="breakdown-line"><span>Towing Limit</span><span>{fmt_kes(q['towing'])}</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                with st.expander(f"Policy details — {insurer_name}"):
                    st.markdown(f"**Inclusions:** {q['inclusions']}")
                    st.markdown(f"**Exclusions:** {q['exclusions']}")
                    st.markdown(f"**Claims email:** {q['claims_email']}")

    st.divider()

    # ----------------------------------------------------------------
    # EXCEL EXPORT
    # ----------------------------------------------------------------
    def build_excel(quotes: dict, vehicle_info: dict) -> bytes:
        buffer = io.BytesIO()

        vehicle_df = pd.DataFrame(list(vehicle_info.items()), columns=["Field", "Value"])

        comparison_rows = []
        for insurer_name, q in quotes.items():
            comparison_rows.append({
                "Insurer": insurer_name,
                "Cover Type": q["cover_type"],
                "Rate Used (%) / TPO Flat (KES)": (
                    f"{insurers_df.loc[insurers_df['Insurer'] == insurer_name, 'Rate_Default (%)'].values[0]}%"
                    if q["cover_type"] == "Comprehensive"
                    else fmt_kes(insurers_df.loc[insurers_df['Insurer'] == insurer_name, 'TPO_Rate (KES)'].values[0])
                ),
                "Basic Premium (KES)": round(q["basic_premium"], 2),
                "Training Levy (KES)": round(q["training_levy"], 2),
                "PHCF (KES)": round(q["phcf"], 2),
                "Stamp Duty (KES)": round(q["stamp_duty"], 2),
                "Total Premium (KES)": round(q["total"], 2),
                "Windscreen Limit (KES)": q["windscreen"],
                "Towing Limit (KES)": q["towing"],
                "Key Inclusions": q["inclusions"],
                "Key Exclusions": q["exclusions"],
                "Claims Email": q["claims_email"],
            })
        comparison_df = pd.DataFrame(comparison_rows)

        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            vehicle_df.to_excel(writer, sheet_name="Vehicle Details", index=False)
            comparison_df.to_excel(writer, sheet_name="Insurer Comparison", index=False)

            for sheet_name, df in [("Vehicle Details", vehicle_df), ("Insurer Comparison", comparison_df)]:
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns, start=1):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(max_len, 45)

        buffer.seek(0)
        return buffer.getvalue()

    vehicle_info = {
        "Vehicle Category": vehicle_category,
        "Model": model,
        "Year of Manufacture": int(manufacture_year),
        "Vehicle Age (years)": age_years,
        "Sum Insured (KES)": round(sum_insured, 2),
        "Cover Type Applied": "TPO" if force_tpo else cover_type,
        "Quote Generated On": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    excel_bytes = build_excel(quotes, vehicle_info)

    col_a, col_b = st.columns(2)

    with col_a:
        st.download_button(
            label="📥 Download Comparison as Excel",
            data=excel_bytes,
            file_name=f"{model.replace(' ', '_')}_{manufacture_year}_insurance_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # ----------------------------------------------------------------
    # WHATSAPP SUMMARY
    # ----------------------------------------------------------------
    with col_b:
        whatsapp_lines = [
            "🚗 *Motor Insurance Quote*",
            f"Vehicle: {model} {int(manufacture_year)} | Value: {fmt_kes(sum_insured)}",
            "",
        ]
        for insurer_name, q in quotes.items():
            whatsapp_lines.append(f"*{insurer_name}* ({q['cover_type']}): {fmt_kes(q['total'])}")
        whatsapp_lines.append("")
        whatsapp_lines.append(f"✅ Best Price: {cheapest_insurer} — {fmt_kes(quotes[cheapest_insurer]['total'])}")
        whatsapp_text = "\n".join(whatsapp_lines)

        st.text_area(
            "WhatsApp-ready summary (tap inside, then Ctrl/Cmd+A, Ctrl/Cmd+C to copy)",
            value=whatsapp_text,
            height=180,
        )
        st.caption(
            "Streamlit cannot access the phone/system clipboard directly for security reasons — "
            "select all text above and copy manually, then paste into WhatsApp."
        )

st.divider()
st.caption(
    "⚠️ Rates and limits shown are editable working defaults for internal agent use only. "
    "Always confirm final premiums against each insurer's current official rate card before binding cover."
)
