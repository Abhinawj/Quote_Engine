# Kenya Motor Insurance Agent Tool

A mobile-responsive, no-login Streamlit app for insurance agents to generate
instant, side-by-side motor insurance quotes across 10 Kenyan insurers —
with live rate editing, Excel export, and a WhatsApp-ready summary.

---

## 1. What's in this folder

| File | Purpose |
|---|---|
| `app.py` | The Streamlit application (main entry point). |
| `insurers_data.csv` | Default data for all 10 insurers — rates, min premiums, TPO rates, windscreen/towing limits, inclusions, exclusions, claims contacts. Loaded at startup. |
| `vehicle_valuations.csv` | Default 2017 base values for 9 vehicle models, used for auto-valuation. Loaded at startup. |
| `requirements.txt` | Python package dependencies. |
| `README.md` | This file. |

**All four files must stay in the same folder** — `app.py` loads the two
CSVs by relative path at startup.

---

## 2. Requirements

- Python 3.9 or later
- pip

---

## 3. Installation

```bash
# 1. (Recommended) create a virtual environment
python3 -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

# 2. Install dependencies
pip install -r requirements.txt
```

`requirements.txt` installs:
- `streamlit` — the web app framework
- `pandas` — data handling
- `openpyxl` — Excel file generation

---

## 4. Running the app

```bash
streamlit run app.py
```

Streamlit will print a local URL (typically `http://localhost:8501`) —
open it in a browser. On the same Wi-Fi network, other devices (e.g. an
agent's phone) can reach it via the **Network URL** Streamlit also prints.

No login or account is required.

---

## 5. How to use it

1. **Sidebar — Vehicle Details**: pick Car or Motorcycle, choose a model,
   enter the year of manufacture. Sum Insured auto-fills (2017 base value
   ±10%/year) but can be edited manually.
2. Vehicles older than **15 years** automatically force Third Party Only
   (TPO) cover — Comprehensive is disabled for those.
3. **Edit Insurer Rates & Benefits / Edit Vehicle Valuation Base Values**
   (expanders above the quotes): edit any cell live — rate, minimum
   premium, TPO rate, limits, inclusions, exclusions, claims contacts, or
   base vehicle values. Rows can be added or deleted. Changes apply
   instantly to every quote card. Use **"Reset to Original CSV"** to
   discard edits and reload the shipped defaults.
4. **Quote cards**: one per insurer, showing the full premium breakdown
   (basic premium, Training Levy, PHCF, stamp duty, total) and the
   cheapest option is marked **BEST PRICE**. Click "Policy details" on a
   card to see inclusions/exclusions/claims email.
5. **Download Comparison as Excel**: generates a two-sheet workbook
   (Vehicle Details + Insurer Comparison) in memory and serves it directly
   — nothing is saved to the server's disk.
6. **WhatsApp summary**: a ready-to-copy text block. Select all
   (Ctrl/Cmd+A) and copy (Ctrl/Cmd+C) inside the box, then paste into
   WhatsApp — browsers don't allow apps to write to the system clipboard
   automatically for security reasons, so this one manual step is
   unavoidable.

---

## 6. Important: edits are session-only, not permanent

Right now, changes made in the on-screen editors apply **only to your
current browser session**. Closing the tab or restarting the app reloads
the original `insurers_data.csv` / `vehicle_valuations.csv` values.

If you want an agent's edit (e.g. a corrected rate) to persist for
everyone, going forward, after the app restarts, you have two options:

- **Simplest**: after editing in the browser, manually update the
  corresponding row in `insurers_data.csv` or `vehicle_valuations.csv`
  and save — the new values become the default on next launch.
- **Automatic**: ask for the app to be extended to write changes back to
  the CSV (or a small database) automatically when edited. Not included
  in this version, to keep the base app simple and avoid accidental
  overwrites in a shared/deployed setting.

---

## 7. Updating the data

To change insurer rates or add/remove insurers or vehicle models
permanently, edit `insurers_data.csv` or `vehicle_valuations.csv` directly
in a spreadsheet program or text editor, keeping the existing column
headers unchanged, then restart the app.

---

## 8. Deploying (optional)

To make this available to your team without everyone installing Python
locally, you can deploy it to:
- **Streamlit Community Cloud** (free, connects to a GitHub repo)
- Any host that runs Python (Render, Railway, a small VPS, etc.) with
  `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`

No environment variables or secrets are required — the app has no login
and no external API calls.

---

## 9. Disclaimer

Rates, minimum premiums, limits, and claims contacts shipped in the
default CSVs are working defaults for internal agent use only. Confirm
figures against each insurer's current official rate card before
quoting or binding cover with a client.
