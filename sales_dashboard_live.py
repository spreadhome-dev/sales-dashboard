"""
LIVE Sales Order Dashboard - Shoppers Stop Ltd
Run:  python sales_dashboard_live.py
Open: http://localhost:8050   (or http://<this-pc-ip>:8050 from other PCs)

Requires:  pip install flask requests pandas openpyxl
"""

import os
import io
import time
import threading
import html as html_escape
from datetime import datetime

import requests
import urllib3
import pandas as pd
from requests.auth import HTTPBasicAuth
from urllib.parse import quote
from flask import Flask, request, jsonify, send_file, Response

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Config ─────────────────────────────────────────────────────
USERNAME = os.environ.get("BC_USER", "Manoj")
PASSWORD = os.environ.get("BC_PASS", "$Manoj@2024")
BASE_URL = "https://192.168.1.222:7048/BC230/ODataV4"
COMPANY = "SPREAD"
SERVICE_NAME = "Sales_Order_A"
CUSTOMER_NAME = "Shoppers Stop Ltd"
DATE_FROM = "2026-01-01"

BASE_STOCK_CSV = "base_stock.csv"
PORT = 8050
CACHE_SECONDS = 120        # re-use BC data for 2 min between page loads
AUTO_RELOAD_SECONDS = 300  # browser refreshes itself every 5 min

if not PASSWORD:
    PASSWORD = input("Enter BC password: ")

AUTH = HTTPBasicAuth(USERNAME, PASSWORD)

FINAL_COLUMNS = [
    "No", "Order_Date", "Sell_to_Customer_No", "Sell_to_Customer_Name",
    "Base Stock", "External_Document_No", "Sell_to_Post_Code",
    "Sell_to_Contact", "Bill_to_Customer_No", "Bill_to_Name",
    "Bill_to_Post_Code", "Bill_to_Contact", "Ship_to_Name",
    "Ship_to_Post_Code", "Ship_to_Contact", "Posting_Date",
    "Shortcut_Dimension_2_Code", "Location_Code", "Document_Date",
    "Status", "Amount", "Amount_Including_VAT",
]
SELECT_COLUMNS = [c for c in FINAL_COLUMNS if c != "Base Stock"]

app = Flask(__name__)
_lock = threading.Lock()
_cache = {"df": None, "time": 0, "complete": True}

# ── BC data fetch ──────────────────────────────────────────────
def fetch_bc_data():
    odata_filter = (
        f"Sell_to_Customer_Name eq '{CUSTOMER_NAME}' "
        f"and Posting_Date ge {DATE_FROM}"
    )
    select_clause = ",".join(SELECT_COLUMNS)
    url = (
        f"{BASE_URL}/Company('{quote(COMPANY)}')/{quote(SERVICE_NAME)}"
        f"?$filter={quote(odata_filter)}&$select={quote(select_clause, safe=',')}"
    )
    rows, complete = [], True
    session = requests.Session()
    session.auth = AUTH
    session.verify = False
    session.headers.update({"Accept": "application/json"})

    while url:
        try:
            r = session.get(url, timeout=300)
            if r.status_code != 200:
                print(f"HTTP {r.status_code}: {r.text[:300]}")
                complete = False
                break
            data = r.json()
            rows.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        except Exception as e:
            print(f"Fetch error: {e}")
            complete = False
            break
    return rows, complete

# ── Base Stock storage (server-side csv) ───────────────────────
def load_base_stock():
    if os.path.exists(BASE_STOCK_CSV):
        try:
            bs = pd.read_csv(BASE_STOCK_CSV, dtype=str).fillna("")
            bs.columns = [c.strip() for c in bs.columns]
            return dict(zip(bs["No"], bs["Base Stock"]))
        except Exception as e:
            print(f"base_stock.csv read error: {e}")
    return {}

def save_base_stock(store: dict):
    df = pd.DataFrame(
        [(k, v) for k, v in sorted(store.items()) if str(v).strip() != ""],
        columns=["No", "Base Stock"],
    )
    df.to_csv(BASE_STOCK_CSV, index=False)

# ── Build dataframe (with cache) ───────────────────────────────
def get_dataframe(force=False):
    with _lock:
        fresh = (time.time() - _cache["time"]) < CACHE_SECONDS
        if _cache["df"] is not None and fresh and not force:
            df = _cache["df"].copy()
            complete = _cache["complete"]
        else:
            rows, complete = fetch_bc_data()
            if rows:
                df = pd.DataFrame(rows)
                if "@odata.etag" in df.columns:
                    df = df.drop(columns=["@odata.etag"])
                df["No"] = df["No"].astype(str)
                for dcol in ["Order_Date", "Posting_Date", "Document_Date"]:
                    if dcol in df.columns:
                        df[dcol] = pd.to_datetime(df[dcol], errors="coerce").dt.date
                _cache["df"] = df.copy()
                _cache["time"] = time.time()
                _cache["complete"] = complete
            elif _cache["df"] is not None:
                df = _cache["df"].copy()   # fall back to last good data
                complete = False
            else:
                return None, complete

    # merge base stock fresh every time (instant reflection of edits)
    store = load_base_stock()
    df["Base Stock"] = df["No"].map(store).fillna("")
    cols = [c for c in FINAL_COLUMNS if c in df.columns]
    return df[cols].copy(), complete

# ── Routes ─────────────────────────────────────────────────────
@app.route("/save", methods=["POST"])
def save():
    data = request.get_json(force=True)
    no = str(data.get("no", "")).strip()
    value = str(data.get("value", "")).strip()
    if not no:
        return jsonify(ok=False), 400
    with _lock:
        store = load_base_stock()
        if value == "":
            store.pop(no, None)
        else:
            store[no] = value
        save_base_stock(store)
    return jsonify(ok=True)

@app.route("/download")
def download():
    df, _ = get_dataframe()
    if df is None:
        return "No data", 503
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Sales_Orders_ShoppersStop.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.route("/")
def index():
    df, complete = get_dataframe()
    if df is None:
        return Response(
            "<h2>Could not reach Business Central. Check BC server / credentials "
            "and refresh this page.</h2>",
            mimetype="text/html",
        )

    updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
    warn = "" if complete else "<p style='color:#c00'><b>WARNING: showing partial or cached data (BC fetch problem)</b></p>"
    total_amt = pd.to_numeric(df["Amount"], errors="coerce").sum()
    total_vat = pd.to_numeric(df["Amount_Including_VAT"], errors="coerce").sum()

    cols = list(df.columns)
    thead = "".join(f"<th>{html_escape.escape(str(c))}</th>" for c in cols)
    body_rows = []
    for _, row in df.iterrows():
        no_val = html_escape.escape(str(row["No"]))
        tds = []
        for c in cols:
            val = "" if pd.isna(row[c]) else str(row[c])
            if c == "Base Stock":
                tds.append(
                    f'<td class="bs-cell"><input type="text" class="bs-input" '
                    f'data-no="{no_val}" value="{html_escape.escape(val)}" '
                    f'placeholder="--"></td>'
                )
            else:
                tds.append(f"<td>{html_escape.escape(val)}</td>")
        body_rows.append("<tr>" + "".join(tds) + "</tr>")
    tbody = "\n".join(body_rows)

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Sales Orders - {CUSTOMER_NAME}</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 20px; background: #f5f6fa; }}
  h1 {{ font-size: 20px; color: #2c3e50; margin-bottom: 4px; }}
  .meta {{ color: #666; font-size: 13px; margin-bottom: 12px; }}
  .toolbar {{ display: flex; gap: 10px; margin-bottom: 14px; align-items: center; }}
  .btn {{ background: #2c3e50; color: #fff; border: none; border-radius: 6px;
          padding: 9px 16px; font-size: 13px; cursor: pointer; text-decoration: none; }}
  .btn:hover {{ background: #3d5771; }}
  #savemsg {{ font-size: 12px; color: #2e7d32; visibility: hidden; }}
  .cards {{ display: flex; gap: 16px; margin-bottom: 16px; }}
  .card {{ background: #fff; border-radius: 8px; padding: 14px 20px;
           box-shadow: 0 1px 3px rgba(0,0,0,.12); }}
  .card .label {{ font-size: 12px; color: #888; }}
  .card .value {{ font-size: 20px; font-weight: 600; color: #2c3e50; }}
  .tablewrap {{ max-height: 75vh; overflow: auto; box-shadow: 0 1px 3px rgba(0,0,0,.12); }}
  table.data {{ border-collapse: collapse; width: 100%; background: #fff; font-size: 12px; }}
  table.data th {{ background: #2c3e50; color: #fff; padding: 8px 10px;
                   position: sticky; top: 0; text-align: left; white-space: nowrap; z-index: 2; }}
  table.data td {{ padding: 6px 10px; border-bottom: 1px solid #eee; }}
  table.data tr:hover td {{ background: #eef4fb; }}
  .bs-cell {{ background: #fffbe8; }}
  .bs-input {{ width: 90px; padding: 4px 6px; border: 1px solid #d8c874;
               border-radius: 4px; background: #fffdf2; font-size: 12px; }}
  .bs-input.saved {{ border-color: #4caf50; background: #f2fff2; }}
</style>
</head>
<body>
<h1>Sales Orders &mdash; {CUSTOMER_NAME}</h1>
<div class="meta">Live data &nbsp;|&nbsp; Last fetched: {updated}</div>
{warn}
<div class="toolbar">
  <a class="btn" href="/download">&#11015; Download Excel</a>
  <span id="savemsg">&#10003; saved</span>
</div>
<div class="cards">
  <div class="card"><div class="label">Orders</div><div class="value">{len(df)}</div></div>
  <div class="card"><div class="label">Total Amount</div><div class="value">{total_amt:,.2f}</div></div>
  <div class="card"><div class="label">Total Incl. VAT</div><div class="value">{total_vat:,.2f}</div></div>
</div>
<div class="tablewrap">
<table class="data">
<thead><tr>{thead}</tr></thead>
<tbody>
{tbody}
</tbody>
</table>
</div>

<script>
let refreshTimer = null;
function scheduleRefresh() {{
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(() => location.reload(), {AUTO_RELOAD_SECONDS * 1000});
}}
scheduleRefresh();

let saveTimers = {{}};
document.querySelectorAll('.bs-input').forEach(inp => {{
  if (inp.value.trim() !== '') inp.classList.add('saved');
  inp.addEventListener('input', () => {{
    scheduleRefresh();                 // don't reload while typing
    clearTimeout(saveTimers[inp.dataset.no]);
    saveTimers[inp.dataset.no] = setTimeout(() => {{
      fetch('/save', {{
        method: 'POST',
        headers: {{'Content-Type': 'application/json'}},
        body: JSON.stringify({{no: inp.dataset.no, value: inp.value}})
      }}).then(r => {{
        if (r.ok) {{
          inp.classList.add('saved');
          const m = document.getElementById('savemsg');
          m.style.visibility = 'visible';
          clearTimeout(m._t);
          m._t = setTimeout(() => m.style.visibility = 'hidden', 1200);
        }}
      }}).catch(() => {{}});
    }}, 500);                          // save half a second after typing stops
  }});
}});
</script>
</body>
</html>"""

# ── Run the app ──────────────────────────────────────────────────
if __name__ == "__main__":
    # For local development only
    app.run(host="0.0.0.0", port=PORT, debug=True)
else:
    # For production (Gunicorn on Render)
    # Gunicorn will handle the server, we just need the app object
    pass
