"""
LIVE Sales Order Dashboard - Shoppers Stop Ltd
Deployed on Render.com
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

# Disable SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── Config ─────────────────────────────────────────────────────
USERNAME = os.environ.get("BC_USER", "Manoj")
PASSWORD = os.environ.get("BC_PASS", "")
BASE_URL = os.environ.get("BC_URL", "https://192.168.1.222:7048/BC230/ODataV4")
COMPANY = os.environ.get("BC_COMPANY", "SPREAD")
SERVICE_NAME = "Sales_Order_A"
CUSTOMER_NAME = os.environ.get("CUSTOMER_NAME", "Shoppers Stop Ltd")
DATE_FROM = os.environ.get("DATE_FROM", "2026-01-01")

BASE_STOCK_CSV = "base_stock.csv"
PORT = int(os.environ.get("PORT", 8050))
CACHE_SECONDS = 120
AUTO_RELOAD_SECONDS = 300

if not PASSWORD:
    print("⚠️ WARNING: BC_PASS environment variable not set!")

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

# ── Logging ──────────────────────────────────────────────────
import logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Health Check ──────────────────────────────────────────────
@app.route("/health")
def health():
    return jsonify({
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "cache_age": f"{time.time() - _cache['time']:.0f}s",
        "data_available": _cache["df"] is not None
    })

# ── BC data fetch ──────────────────────────────────────────────
def fetch_bc_data():
    logger.info(f"Fetching data for {CUSTOMER_NAME} from {DATE_FROM}")
    
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
                logger.error(f"HTTP {r.status_code}: {r.text[:300]}")
                complete = False
                break
            data = r.json()
            rows.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        except Exception as e:
            logger.error(f"Fetch error: {e}")
            complete = False
            break
    
    logger.info(f"Fetched {len(rows)} rows, complete={complete}")
    return rows, complete

# ── Base Stock storage ───────────────────────────────────────
def load_base_stock():
    if os.path.exists(BASE_STOCK_CSV):
        try:
            bs = pd.read_csv(BASE_STOCK_CSV, dtype=str).fillna("")
            bs.columns = [c.strip() for c in bs.columns]
            return dict(zip(bs["No"], bs["Base Stock"]))
        except Exception as e:
            logger.error(f"base_stock.csv read error: {e}")
    return {}

def save_base_stock(store: dict):
    df = pd.DataFrame(
        [(k, v) for k, v in sorted(store.items()) if str(v).strip() != ""],
        columns=["No", "Base Stock"],
    )
    df.to_csv(BASE_STOCK_CSV, index=False)
    logger.info(f"Saved {len(df)} base stock entries")

# ── Build dataframe ───────────────────────────────────────────
def get_dataframe(force=False):
    with _lock:
        fresh = (time.time() - _cache["time"]) < CACHE_SECONDS
        if _cache["df"] is not None and fresh and not force:
            logger.info("Using cached data")
            df = _cache["df"].copy()
            complete = _cache["complete"]
        else:
            logger.info("Fetching fresh data")
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
                logger.info(f"Cached {len(df)} rows")
            elif _cache["df"] is not None:
                logger.warning("Using fallback cached data")
                df = _cache["df"].copy()
                complete = False
            else:
                logger.error("No data available")
                return None, complete

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

@app.route("/test-bc")
def test_bc():
    """Test Business Central connection"""
    try:
        response = requests.get(
            BASE_URL,
            auth=AUTH,
            verify=False,
            timeout=10
        )
        return jsonify({
            "status": "success",
            "bc_status_code": response.status_code,
            "bc_url": BASE_URL,
            "message": "BC server is reachable"
        })
    except Exception as e:
        return jsonify({
            "status": "error",
            "error": str(e),
            "bc_url": BASE_URL,
            "message": "Cannot reach BC server"
        }), 500

@app.route("/")
def index():
    df, complete = get_dataframe()
    if df is None:
        return Response("""
        <h2>⚠️ Could not reach Business Central</h2>
        <p>Please check:</p>
        <ul>
            <li><a href="/test-bc">Test BC Connection</a></li>
            <li>BC server is accessible from Render</li>
            <li>Credentials are correct in Environment Variables</li>
        </ul>
        <p><small>Contact IT team to make BC accessible from internet</small></p>
        """, mimetype="text/html", status=503)

    updated = datetime.now().strftime("%d-%b-%Y %H:%M:%S")
    warn = "" if complete else "<p style='color:#c00'><b>⚠️ Showing cached data</b></p>"
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
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 20px; background: #f5f6fa; }}
  h1 {{ font-size: 20px; color: #2c3e50; margin-bottom: 4px; }}
  .meta {{ color: #666; font-size: 13px; margin-bottom: 12px; }}
  .toolbar {{ display: flex; gap: 10px; margin-bottom: 14px; align-items: center; flex-wrap: wrap; }}
  .btn {{ background: #2c3e50; color: #fff; border: none; border-radius: 6px;
          padding: 9px 16px; font-size: 13px; cursor: pointer; text-decoration: none; display: inline-block; }}
  .btn:hover {{ background: #3d5771; }}
  #savemsg {{ font-size: 12px; color: #2e7d32; visibility: hidden; }}
  .cards {{ display: flex; gap: 16px; margin-bottom: 16px; flex-wrap: wrap; }}
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
  .footer {{ margin-top: 16px; font-size: 11px; color: #999; text-align: center; }}
</style>
</head>
<body>
<h1>📊 Sales Orders — {CUSTOMER_NAME}</h1>
<div class="meta">Live data | Last fetched: {updated}</div>
{warn}
<div class="toolbar">
  <a class="btn" href="/download">⬇️ Download Excel</a>
  <a class="btn" href="/test-bc" target="_blank">🔌 Test BC</a>
  <span id="savemsg">✅ saved</span>
</div>
<div class="cards">
  <div class="card"><div class="label">📋 Orders</div><div class="value">{len(df)}</div></div>
  <div class="card"><div class="label">💰 Total Amount</div><div class="value">{total_amt:,.2f}</div></div>
  <div class="card"><div class="label">🧾 Total Incl. VAT</div><div class="value">{total_vat:,.2f}</div></div>
</div>
<div class="tablewrap">
<table class="data">
<thead><tr>{thead}</tr></thead>
<tbody>
{tbody}
</tbody>
</table>
</div>
<div class="footer">Deployed on Render | Auto-refreshes every {AUTO_RELOAD_SECONDS//60} minutes</div>

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
    scheduleRefresh();
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
    }}, 500);
  }});
}});
</script>
</body>
</html>"""

# ── Run the app ──────────────────────────────────────────────────
if __name__ == "__main__":
    logger.info(f"🚀 Starting Sales Dashboard on port {PORT}")
    logger.info(f"🔗 BC Server: {BASE_URL}")
    logger.info(f"🏢 Customer: {CUSTOMER_NAME}")
    
    from waitress import serve
    serve(app, host="0.0.0.0", port=PORT)