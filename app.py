import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, timedelta

st.set_page_config(page_title="BHDML Delivery Dashboard — Routes → Days", layout="wide")
st.title("BHDML Delivery Dashboard — Routes → Mon–Fri")

with st.expander("How this works", expanded=True):
    st.markdown(
        "- Upload **one Excel per weekday** (PeerPlace *Report* export).\n"
        "- The app builds **Route tabs** (up to 14, omitting **'COPO'** if enabled). Inside each route you get **Mon–Fri** tabs.\n"
        "- Each row shows: **Client Name (First Last)**, **Address** (Line1 + Line2 + Building), **Phone** (Mobile→Home), "
        "**Meals** (editable), **Service Type**, **Diet Type**, and a **Delivered** checkbox.\n"
        "- Use **Add New Client** (inside a Route) to append a client to the selected **Day**.\n"
        "- **Save Week State** to JSON and **Load** it later to continue where you left off.\n"
        "- **Export**: per-route per-day workbook, or whole route (Mon–Fri) workbook with one sheet per day + Summary.\n"
    )

# --------------------- Sidebar Controls ---------------------
st.sidebar.header("Settings")
def monday_of_today():
    t = date.today()
    return t - timedelta(days=t.weekday())

week_monday = st.sidebar.date_input(
    "Week start (Monday)",
    value=monday_of_today(),
    help="Labels the Mon–Fri tabs with dates."
)
default_delivered_checked = st.sidebar.checkbox(
    "Default Delivered = checked",
    value=True,
    help="Pre-check Delivered for all clients; uncheck exceptions."
)
exclude_copo = st.sidebar.checkbox(
    "Exclude 'COPO' routes",
    value=True,
    help="Omits any route whose name contains 'COPO' (case-insensitive)."
)

# --------------------- Constants & Helpers ---------------------
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
DAY_SHORT = dict(zip(DAYS, ["Mon", "Tue", "Wed", "Thu", "Fri"]))

EXPECTED_COLS = [
    "Delivery Route","Client ID","Last Name","First Name",
    "Address Line 1","Address Line 2","Building",
    "Home Phone","Mobile Phone","Quantity",
    "Service Type","Diet Type"
]

def day_date(monday: date, offset: int) -> date:
    return monday + timedelta(days=offset)

def read_report(xfile) -> pd.DataFrame:
    xls = pd.ExcelFile(xfile)
    df0 = pd.read_excel(xls, sheet_name="Report", header=None)
    header_row_idx = None
    for i in range(min(25, len(df0))):
        if str(df0.iloc[i, 0]).strip() == "Delivery Route":
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Could not find header row containing 'Delivery Route' on the 'Report' sheet.")
    df = pd.read_excel(xls, sheet_name="Report", header=header_row_idx)

    present = [c for c in EXPECTED_COLS if c in df.columns]
    df = df.dropna(how="all", subset=[c for c in ["Delivery Route","Client ID","Last Name","First Name"] if c in df.columns])
    df = df[present].copy()

    # Client Name (First + Last)
    if "First Name" in df.columns and "Last Name" in df.columns:
        df.insert(0, "Client Name", (df["First Name"].astype(str) + " " + df["Last Name"].astype(str)).str.strip())

    # Address (Line1 + Line2 + Building)
    parts = []
    if "Address Line 1" in df.columns: parts.append(df["Address Line 1"].astype(str))
    if "Address Line 2" in df.columns: parts.append(df["Address Line 2"].fillna("").astype(str))
    if "Building" in df.columns:       parts.append(df["Building"].fillna("").astype(str))
    if parts:
        addr = parts[0]
        for p in parts[1:]:
            addr = addr + ((" " + p).where(p.str.strip()!="", ""))
        df.insert(1, "Address", addr.str.strip())

    # Phone: prefer Mobile else Home
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

    # Meals from Quantity
    if "Quantity" in df.columns:
        df["Meals"] = pd.to_numeric(df["Quantity"], errors="coerce").fillna(0).astype(int)
    else:
        df["Meals"] = 0

    # Delivered default
    df["Delivered"] = bool(default_delivered_checked)

    keep = [c for c in ["Delivery Route","Client ID","Client Name","Address","Phone","Meals","Service Type","Diet Type","Delivered"] if c in df.columns]
    return df[keep].copy()

def ensure_rowid(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df
    if "RowID" not in df.columns:
        df = df.reset_index(drop=True)
        df.insert(0, "RowID", df.index)
    return df

def filter_routes_from_week(week_state: dict, exclude_copo_flag: bool, cap: int = 14) -> list:
    # Union of routes across all available days, preserving first-seen order.
    seen = set(); order = []
    for d in DAYS:
        df = week_state.get(d)
        if isinstance(df, pd.DataFrame) and not df.empty and "Delivery Route" in df.columns:
            for r in df["Delivery Route"]:
                if pd.isna(r): continue
                rs = str(r)
                if exclude_copo_flag and "COPO" in rs.upper(): 
                    continue
                if rs not in seen:
                    seen.add(rs)
                    order.append(rs)
    return order[:cap]

def safe_sheet_name(name: str) -> str:
    s = "".join(ch for ch in str(name) if ch.isalnum())
    return s[:31] or "Sheet"

def to_excel_route_day(df_day: pd.DataFrame, route_name: str) -> BytesIO:
    """Export a single route for a single day as one-sheet workbook with totals."""
    import xlsxwriter
    out = BytesIO()
    df_day = ensure_rowid(df_day)
    cols = [c for c in ["Delivery Route","Client ID","Client Name","Address","Phone","Meals","Service Type","Diet Type","Delivered"] if c in df_day.columns]
    rdf = df_day[df_day["Delivery Route"].astype(str) == route_name][cols].copy()
    if "Delivered" in rdf.columns:
        rdf["Delivered"] = rdf["Delivered"].map(lambda x: "X" if bool(x) else "")
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        total = wb.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1})
        center = wb.add_format({"align":"center","valign":"vcenter","border":1})
        left = wb.add_format({"align":"left","valign":"vcenter","border":1})
        sheet = safe_sheet_name(route_name)
        rdf.to_excel(writer, sheet_name=sheet, index=False)
        ws = writer.sheets[sheet]
        ws.set_row(0, None, hdr)
        widths = {"Delivery Route":18,"Client ID":14,"Client Name":22,"Address":34,"Phone":16,"Meals":10,"Service Type":14,"Diet Type":16,"Delivered":12}
        for ci, col in enumerate(rdf.columns):
            ws.set_column(ci, ci, widths.get(col, 14), center if col in ["Meals","Delivered","Service Type","Diet Type"] else left)
        n = len(rdf)
        ws.write(n+1, 0, "TOTALS", total)
        # Delivered count
        if "Delivered" in rdf.columns:
            dcol = rdf.columns.get_loc("Delivered")
            def col_letter(i0):
                s=""; i=i0+1
                while i: i, rem = divmod(i-1,26); s=chr(65+rem)+s
                return s
            dL = col_letter(dcol)
            ws.write_formula(n+1, dcol, f'=COUNTIF({dL}2:{dL}{n+1},"X")', total)
        # Meals sum
        if "Meals" in rdf.columns:
            mcol = rdf.columns.get_loc("Meals")
            def col_letter(i0):
                s=""; i=i0+1
                while i: i, rem = divmod(i-1,26); s=chr(65+rem)+s
                return s
            mL = col_letter(mcol)
            ws.write_formula(n+1, mcol, f"=SUM({mL}2:{mL}{n+1})", total)
    out.seek(0)
    return out

def to_excel_route_week(week_state: dict, route_name: str) -> BytesIO:
    """Export a single route across Mon–Fri, one sheet per day + Summary."""
    import xlsxwriter
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        wb = writer.book
        hdr = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1})
        total = wb.add_format({"bold": True, "bg_color": "#FCE4D6", "border": 1})
        center = wb.add_format({"align":"center","valign":"vcenter","border":1})
        left = wb.add_format({"align":"left","valign":"vcenter","border":1})
        widths = {"Delivery Route":18,"Client ID":14,"Client Name":22,"Address":34,"Phone":16,"Meals":10,"Service Type":14,"Diet Type":16,"Delivered":12}

        cols = ["Delivery Route","Client ID","Client Name","Address","Phone","Meals","Service Type","Diet Type","Delivered"]
        summary_rows = []
        for d in DAYS:
            df_day = week_state.get(d)
            if not isinstance(df_day, pd.DataFrame) or df_day.empty:
                summary_rows.append({"Day": d, "Rows": 0, "Meals": 0, "Delivered": 0})
                continue
            rdf = df_day[df_day["Delivery Route"].astype(str) == route_name][cols].copy()
            delivered_count = int(rdf.get("Delivered", pd.Series(dtype=bool)).sum())
            meals_sum = int(pd.to_numeric(rdf.get("Meals", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
            summary_rows.append({"Day": d, "Rows": len(rdf), "Meals": meals_sum, "Delivered": delivered_count})
            # convert Delivered to X for export
            if "Delivered" in rdf.columns:
                rdf["Delivered"] = rdf["Delivered"].map(lambda x: "X" if bool(x) else "")
            sheet = safe_sheet_name(f"{DAY_SHORT[d]}_{route_name}")
            rdf.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.set_row(0, None, hdr)
            for ci, col in enumerate(rdf.columns):
                ws.set_column(ci, ci, widths.get(col, 14), center if col in ["Meals","Delivered","Service Type","Diet Type"] else left)
            n = len(rdf)
            ws.write(n+1, 0, "TOTALS", total)
            # Totals formulas
            if "Delivered" in rdf.columns:
                dcol = rdf.columns.get_loc("Delivered")
                def colL(i0):
                    s=""; i=i0+1
                    while i: i, rem = divmod(i-1,26); s=chr(65+rem)+s
                    return s
                ws.write_formula(n+1, dcol, f'=COUNTIF({colL(dcol)}2:{colL(dcol)}{n+1},"X")', total)
            if "Meals" in rdf.columns:
                mcol = rdf.columns.get_loc("Meals")
                def colL(i0):
                    s=""; i=i0+1
                    while i: i, rem = divmod(i-1,26); s=chr(65+rem)+s
                    return s
                ws.write_formula(n+1, mcol, f"=SUM({colL(mcol)}2:{colL(mcol)}{n+1})", total)

        # Summary sheet
        s = pd.DataFrame(summary_rows)
        s.to_excel(writer, sheet_name="Summary", index=False)
        sws = writer.sheets["Summary"]
        sws.set_row(0, None, hdr)
        sws.set_column(0, 0, 10, left)
        sws.set_column(1, 3, 12, center)
    out.seek(0)
    return out

# --------------------- State (Week-level Save/Load) ---------------------
def empty_week_state():
    return {d: None for d in DAYS}

if "week_state" not in st.session_state:
    st.session_state.week_state = empty_week_state()

st.subheader("Upload Your Week")
cols = st.columns(5)
uploads = {}
for i, d in enumerate(DAYS):
    with cols[i]:
        uploads[d] = st.file_uploader(f"{d} file (Report export)", type=["xlsx"], key=f"u_{d}")

# Load files into state
for d in DAYS:
    if uploads[d] is not None:
        try:
            df = read_report(uploads[d])
            df = ensure_rowid(df)
            st.session_state.week_state[d] = df
        except Exception as e:
            st.error(f"{d}: {e}")

# Save / Load JSON
st.markdown("### Save / Load Week State")
c1, c2 = st.columns(2)
with c1:
    if st.button("Download Week State (.json)"):
        import json
        payload = {}
        for d in DAYS:
            df = st.session_state.week_state.get(d)
            payload[d] = df.to_dict(orient="records") if isinstance(df, pd.DataFrame) else None
        b = BytesIO()
        b.write(json.dumps(payload).encode("utf-8"))
        b.seek(0)
        st.download_button("Save week_state.json", data=b, file_name="week_state.json", mime="application/json")
with c2:
    up = st.file_uploader("Load Week State (.json)", type=["json"], key="load_week_json")
    if up is not None:
        import json
        try:
            raw = json.load(up)
            new_state = {}
            for d in DAYS:
                recs = raw.get(d)
                if recs is None:
                    new_state[d] = None
                else:
                    df = pd.DataFrame.from_records(recs)
                    df = ensure_rowid(df)
                    if "Delivered" not in df.columns:
                        df["Delivered"] = bool(default_delivered_checked)
                    if "Meals" in df.columns:
                        df["Meals"] = pd.to_numeric(df["Meals"], errors="coerce").fillna(0).astype(int)
                    new_state[d] = df
            st.session_state.week_state = new_state
            st.success("Week state loaded.")
        except Exception as e:
            st.error(f"Load failed: {e}")

st.divider()

# --------------------- Routes (top level) ---------------------
# derive routes union across week
routes_all = filter_routes_from_week(st.session_state.week_state, exclude_copo, cap=14)

if not routes_all:
    st.info("Upload at least one day's file to see routes.")
    st.stop()

route_tabs = st.tabs(routes_all)

# Precompute date labels
date_labels = {DAYS[i]: day_date(week_monday, i).strftime("%a %m/%d/%Y") for i in range(5)}

for route_name, rtab in zip(routes_all, route_tabs):
    with rtab:
        st.markdown(f"### Route: {route_name}")

        # Add New Client (for this route; choose a Day)
        st.markdown("#### Add New Client to this Route")
        with st.form(f"add_client_{route_name}", clear_on_submit=True):
            cA, cB, cC = st.columns(3)
            with cA: first = st.text_input("First Name")
            with cB: last = st.text_input("Last Name")
            with cC: client_id = st.text_input("Client ID (optional)")
            address1 = st.text_input("Address Line 1")
            cD, cE = st.columns(2)
            with cD: address2 = st.text_input("Address Line 2", value="")
            with cE: building = st.text_input("Building", value="")
            phone = st.text_input("Phone (Mobile or Home)")
            cF, cG, cH = st.columns(3)
            with cF: meals = st.number_input("Meals", min_value=0, value=1, step=1)
            with cG: service_type = st.selectbox("Service Type", ["Weekday","City Meal"])
            with cH: diet_type = st.text_input("Diet Type", value="")
            day_choice = st.selectbox("Day", DAYS, index=0, help="Which day to add the client to")
            delivered_default = st.checkbox("Delivered", value=default_delivered_checked)

            submitted = st.form_submit_button("Add Client")
            if submitted:
                df_day = st.session_state.week_state.get(day_choice)
                if not isinstance(df_day, pd.DataFrame):
                    df_day = pd.DataFrame(columns=["RowID","Delivery Route","Client ID","Client Name","Address","Phone","Meals","Service Type","Diet Type","Delivered"])
                client_name = f"{first.strip()} {last.strip()}".strip()
                new_row = {
                    "Delivery Route": route_name,
                    "Client ID": client_id.strip() if client_id.strip() else "",
                    "Client Name": client_name,
                    "Address": " ".join([x for x in [address1.strip(), address2.strip(), building.strip()] if x]),
                    "Phone": phone.strip(),
                    "Meals": int(meals),
                    "Service Type": service_type,
                    "Diet Type": diet_type.strip(),
                    "Delivered": bool(delivered_default),
                }
                df_day = pd.concat([df_day, pd.DataFrame([new_row])], ignore_index=True)
                df_day = ensure_rowid(df_day)
                st.session_state.week_state[day_choice] = df_day
                st.success(f"Added {client_name} to {route_name} on {day_choice}.")

        # Now build Mon–Fri tabs inside this Route
        day_tabs = st.tabs([f"{d} — {date_labels[d]}" for d in DAYS])

        for d, dtab in zip(DAYS, day_tabs):
            with dtab:
                df_day = st.session_state.week_state.get(d)
                if not isinstance(df_day, pd.DataFrame) or df_day.empty:
                    st.info(f"No data for {d}.")
                    continue

                df_day = ensure_rowid(df_day)
                st.session_state.week_state[d] = df_day

                rdf = df_day[df_day["Delivery Route"].astype(str) == route_name].copy()
                if rdf.empty:
                    st.write("No rows on this day for this route.")
                    continue

                # Columns to show; include RowID for stable writeback
                show_cols = [c for c in [
                    "RowID","Delivery Route","Client ID","Client Name","Address","Phone",
                    "Meals","Service Type","Diet Type","Delivered"
                ] if c in rdf.columns]

                col_cfg = {}
                if "RowID" in show_cols:
                    col_cfg["RowID"] = st.column_config.NumberColumn("RowID", help="Internal key", disabled=True)
                if "Meals" in show_cols:
                    col_cfg["Meals"] = st.column_config.NumberColumn("Meals", help="Edit to reassign meals", min_value=0, step=1)
                if "Delivered" in show_cols:
                    col_cfg["Delivered"] = st.column_config.CheckboxColumn("Delivered", default=default_delivered_checked, help="Uncheck if NOT delivered")

                edited = st.data_editor(
                    rdf[show_cols],
                    key=f"editor_{route_name}_{d}",
                    num_rows="fixed",
                    use_container_width=True,
                    column_config=col_cfg
                )

                # Write back by RowID
                full = st.session_state.week_state[d]
                if "RowID" in edited.columns and "RowID" in full.columns:
                    upd = edited[["RowID"] + [c for c in ["Delivered","Meals"] if c in edited.columns]].copy()
                    upd = upd.set_index("RowID")
                    base = full.set_index("RowID")
                    if "Delivered" in upd.columns and "Delivered" in base.columns:
                        base.loc[upd.index, "Delivered"] = upd["Delivered"].astype(bool)
                    if "Meals" in upd.columns and "Meals" in base.columns:
                        base.loc[upd.index, "Meals"] = pd.to_numeric(upd["Meals"], errors="coerce").fillna(0).astype(int)
                    st.session_state.week_state[d] = base.reset_index()

                # Totals for this route/day
                df_route = st.session_state.week_state[d]
                df_route = df_route[df_route["Delivery Route"].astype(str) == route_name]
                delivered_count = int(df_route.get("Delivered", pd.Series(dtype=bool)).sum())
                total_meals = int(pd.to_numeric(df_route.get("Meals", pd.Series(dtype=int)), errors="coerce").fillna(0).sum())
                st.caption(f"**Totals ({d})** — Delivered: {delivered_count} • Meals: {total_meals}")

                # Export buttons
                c1, c2 = st.columns(2)
                with c1:
                    if st.button(f"Export {route_name} — {d}", key=f"exp_route_day_{route_name}_{d}"):
                        x = to_excel_route_day(st.session_state.week_state[d], route_name)
                        st.download_button(
                            label=f"Save {route_name}_{DAY_SHORT[d]}.xlsx",
                            data=x,
                            file_name=f"{safe_sheet_name(route_name)}_{DAY_SHORT[d]}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_route_day_{route_name}_{d}"
                        )
                with c2:
                    if st.button(f"Export {route_name} — Mon–Fri", key=f"exp_route_week_{route_name}"):
                        x = to_excel_route_week(st.session_state.week_state, route_name)
                        st.download_button(
                            label=f"Save {route_name}_Mon-Fri.xlsx",
                            data=x,
                            file_name=f"{safe_sheet_name(route_name)}_Mon-Fri.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_route_week_{route_name}"
                        )
