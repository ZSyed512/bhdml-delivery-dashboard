
# BHDML Delivery Dashboard — Mon–Fri (Multi‑Day)

**What you get**
- **5 top-level tabs** (Mon–Fri)
- Each day has **route tabs** (max 14), excluding any route whose name contains **"COPO"**
- Per route: **Client Name (First Last)**, **Address** (Line1 + Line2 + Building), **Phone** (Mobile else Home),
  **Quantity**, **Service Type**, **Diet Type**, **Delivered** checkbox (default checked)
- Per‑day **export to Excel**: one sheet per route, Delivered totals row, and a Summary sheet

## Replit quick start
1. Create a new Python Repl.
2. Upload `app.py`, `requirements.txt`, `README.md`.
3. Shell: `pip install -r requirements.txt`
4. Run: `streamlit run app.py --server.port=3000 --server.address=0.0.0.0`
5. Upload each weekday file, review per-route tabs, toggle **Delivered**, and export.
