
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="BHDML Delivery Dashboard", layout="wide")
st.title("BHDML Delivery Dashboard — Mon–Fri")

with st.expander("How this works", expanded=True):
    st.markdown(
        "**Upload** one Excel per weekday (PeerPlace *Report* export).  \n"
        "The app creates 5 big tabs (Mon–Fri). Each day contains **route tabs** (omitting any routes containing **'COPO'**).  \n"
        "Every route tab shows: **Client Name (First Last)**, **Address** (Line1 + Line2 + Building), **Phone** (Mobile else Home), "
        "**Quantity**, **Service Type**, **Diet Type**, and a **Delivered** checkbox you can toggle.  \n\n"
        "**Export** per-day to an Excel with one sheet per route."
    )

# Sidebar controls
exclude_copo = st.sidebar.checkbox("Exclude 'COPO' routes", value=True, help="Omits any route whose name contains 'COPO' (case-insensitive).")
default_delivered_checked = st.sidebar.checkbox("Default Delivered = checked", value=True, help="Pre-check Delivered for all clients. Uncheck where not delivered.")

# --- Helpers ---
EXPECTED_COLS = [
    "Delivery Route","Client ID","Last Name","First Name",
    "Address Line 1","Address Line 2","Building",
    "Home Phone","Mobile Phone","Quantity",
    "Service Type","Diet Type"
]

def read_report(file):
    xls = pd.ExcelFile(file)
    # Find header row by scanning for 'Delivery Route' in col 0
    df0 = pd.read_excel(xls, sheet_name="Report", header=None)
    header_row_idx = None
    for i in range(min(25, len(df0))):
        if str(df0.iloc[i, 0]).strip() == "Delivery Route":
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Could not find header row containing 'Delivery Route' on the 'Report' sheet.")
    df = pd.read_excel(xls, sheet_name="Report", header=header_row_idx)
    # Keep only expected columns that are present
    present = [c for c in EXPECTED_COLS if c in df.columns]
    df = df.dropna(how="all", subset=[c for c in ["Delivery Route","Client ID","Last Name","First Name"] if c in df.columns])
    return df[present].copy()

def normalize_for_display(df):
    # Build Client Name
    if "First Name" in df.columns and "Last Name" in df.columns:
        df.insert(0, "Client Name", (df["First Name"].astype(str) + " " + df["Last Name"].astype(str)).str.strip())
    # Build Address (Line1 + Line2 + Building)
    adr_parts = []
    if "Address Line 1" in df.columns: adr_parts.append(df["Address Line 1"].astype(str))
    if "Address Line 2" in df.columns: adr_parts.append(df["Address Line 2"].fillna("").astype(str))
    if "Building" in df.columns: adr_parts.append(df["Building"].fillna("").astype(str))
    if adr_parts:
        addr = adr_parts[0]
        for part in adr_parts[1:]:
            addr = addr + ((" " + part).where(part.str.strip()!="", ""))
        df.insert(1, "Address", addr.str.strip())
    # Phone (prefer Mobile Phone else Home Phone)
    phone = None
    if "Mobile Phone" in df.columns and "Home Phone" in df.columns:
        phone = df["Mobile Phone"].fillna("")
        phone = phone.mask(phone.str.strip()=="", df["Home Phone"].fillna(""))
    elif "Mobile Phone" in df.columns:
        phone = df["Mobile Phone"]
    elif "Home Phone" in df.columns:
        phone = df["Home Phone"]
    if phone is not None:
        df.insert(2, "Phone", phone.astype(str).str.strip())
    # Keep only desired display columns
    keep = [c for c in ["Delivery Route","Client ID","Client Name","Address","Phone","Quantity","Service Type","Diet Type"] if c in df.columns]
    df = df[keep].copy()
    # Delivered checkbox
    df["Delivered"] = bool(st.session_state.get("default_delivered_checked", True))
    return df

def route_list(df):
    routes = df["Delivery Route"].dropna().astype(str).unique().tolist() if "Delivery Route" in df.columns else []
    if exclude_copo:
        routes = [r for r in routes if "COPO" not in r.upper()]
    # Preserve order of appearance
    order = []
    seen = set()
    for r in df["Delivery Route"]:
        if pd.isna(r): continue
        rs = str(r)
        if exclude_copo and "COPO" in rs.upper():
            continue
        if rs not in seen:
            seen.add(rs)
            order.append(rs)
    return order[:14]

def to_excel_by_route(df):
    # Export one workbook with each route as a sheet
    import xlsxwriter  # ensure engine available
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        center = wb.add_format({"align":"center","valign":"vcenter","border":1})
        left = wb.add_format({"align":"left","valign":"vcenter","border":1})
        total_fmt = wb.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1})
        widths = {
            "Delivery Route": 18, "Client ID": 14, "Client Name": 22, "Address": 34, "Phone": 16,
            "Quantity": 10, "Service Type": 14, "Diet Type": 16, "Delivered": 12
        }
        summary = []
        for route in route_list(df):
            rdf = df[df["Delivery Route"].astype(str) == route].copy()
            if "Delivered" in rdf.columns:
                rdf["Delivered"] = rdf["Delivered"].map(lambda x: "X" if bool(x) else "")
            sheet = "".join(ch for ch in str(route) if ch.isalnum())[:31] or "Route"
            rdf.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.set_row(0, None, hdr_fmt)
            for ci, col in enumerate(rdf.columns):
                ws.set_column(ci, ci, widths.get(col, 14), center if col in ["Quantity","Delivered","Service Type","Diet Type"] else left)
            nrows = len(rdf)
            total_row = nrows + 1
            ws.write(total_row, 0, "TOTALS", total_fmt)
            if "Delivered" in rdf.columns:
                dcol = rdf.columns.get_loc("Delivered")
                def col_letter(idx0):
                    letters = ""
                    idx = idx0 + 1
                    while idx:
                        idx, rem = divmod(idx - 1, 26)
                        letters = chr(65 + rem) + letters
                    return letters
                colL = col_letter(dcol)
                ws.write_formula(total_row, dcol, f'=COUNTIF({colL}2:{colL}{nrows+1},"X")', total_fmt)
            summary.append({"Route": route, "Clients": nrows})
        if summary:
            s = pd.DataFrame(summary)
            s.to_excel(writer, sheet_name="Summary", index=False)
            sws = writer.sheets["Summary"]
            sws.set_row(0, None, hdr_fmt)
            sws.set_column(0, 0, 24, left)
            sws.set_column(1, 1, 12, center)
    output.seek(0)
    return output

# --- Uploaders for each day ---
st.subheader("Upload Your Week")
cols = st.columns(5)
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
uploads = {}
for i, d in enumerate(days):
    with cols[i]:
        uploads[d] = st.file_uploader(f"{d} file", type=["xlsx"], key=f"u_{d}")

# store sidebar default checkbox state
st.session_state["default_delivered_checked"] = default_delivered_checked

available = [(d, f) for d, f in uploads.items() if f is not None]
if not available:
    st.info("Upload at least one day to begin.")
    st.stop()

top_tabs = st.tabs([d for d, _ in available])

for (day, file), tab in zip(available, top_tabs):
    with tab:
        try:
            raw = read_report(file)
        except Exception as e:
            st.error(f"{day}: {e}")
            continue
        
        disp = normalize_for_display(raw)
        routes = route_list(disp)

        st.markdown(f"### {day} — Routes ({len(routes)})")
        sub_tabs = st.tabs(routes if routes else ["No Routes"])

        if not routes:
            with sub_tabs[0]:
                st.write("No routes found.")
        else:
            state_key = f"data_{day}"
            if state_key not in st.session_state:
                disp["Delivered"] = bool(default_delivered_checked)
                st.session_state[state_key] = disp.copy()
            for rt, rtab in zip(routes, sub_tabs):
                with rtab:
                    df_day = st.session_state[state_key]
                    rdf = df_day[df_day["Delivery Route"].astype(str) == rt].copy()
                    show_cols = [c for c in ["Delivery Route","Client ID","Client Name","Address","Phone","Quantity","Service Type","Diet Type","Delivered"] if c in rdf.columns]
                    col_cfg = {}
                    if "Delivered" in show_cols:
                        col_cfg["Delivered"] = st.column_config.CheckboxColumn("Delivered", default=default_delivered_checked, help="Uncheck if NOT delivered")
                    edited = st.data_editor(
                        rdf[show_cols],
                        key=f"editor_{day}_{rt}",
                        num_rows="fixed",
                        use_container_width=True,
                        column_config=col_cfg
                    )
                    # Sync by Client ID
                    if "Client ID" in edited.columns and "Client ID" in df_day.columns:
                        upd = edited.set_index("Client ID")
                        full = df_day.set_index("Client ID")
                        if "Delivered" in upd.columns and "Delivered" in full.columns:
                            full.loc[upd.index, "Delivered"] = upd["Delivered"].astype(bool)
                        st.session_state[state_key] = full.reset_index()

            st.divider()
            st.subheader(f"Export — {day}")
            st.caption("Exports one Excel with a sheet per route (Delivered column included, totals at bottom, Summary tab).")
            if st.button(f"Download {day} workbook", key=f"dl_{day}"):
                xls_bytes = to_excel_by_route(st.session_state[state_key])
                st.download_button(
                    f"Save {day}_Delivery_Log.xlsx",
                    data=xls_bytes,
                    file_name=f"{day}_Delivery_Log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"save_{day}"
                )
