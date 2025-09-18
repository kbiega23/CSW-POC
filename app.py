# app.py — Winsert Savings Calculator (Office)
# Flicker-free wizard: Step 1 non-form, Step 2 non-form with deferred Graph write,
# Steps 3–4 in containers with .empty() before rerun. Inline climate/WWR metrics,
# blue theme, token cache, Graph Excel session reuse, tidy Results with input summary.

import os
import re
import requests
import streamlit as st
from msal import PublicClientApplication, SerializableTokenCache

# =========================
# Config & sanity checks
# =========================
SCOPES = ["User.Read", "Files.ReadWrite"]  # keep reserved scopes (offline_access) out

guid = re.compile(r"^[0-9a-fA-F-]{36}$")
bad = []
if not guid.match(st.secrets.get("TENANT_ID", "")): bad.append("TENANT_ID")
if not guid.match(st.secrets.get("CLIENT_ID", "")): bad.append("CLIENT_ID")
wp = st.secrets.get("WORKBOOK_PATH", "")
if not (wp.startswith("/me/drive/root:") or wp.startswith("/sites/")):
    bad.append("WORKBOOK_PATH")
if bad:
    st.error("Secrets not set correctly: " + ", ".join(bad))
    st.stop()

TENANT_ID     = st.secrets["TENANT_ID"]
CLIENT_ID     = st.secrets["CLIENT_ID"]
WORKBOOK_PATH = st.secrets["WORKBOOK_PATH"]
GRAPH = "https://graph.microsoft.com/v1.0"

st.set_page_config(page_title="Winsert Savings Calculator – Office", layout="centered")

# =========================
# Blue look & feel + progress bar + optional logo
# =========================
PRIMARY = "#1a66ff"
ACCENT  = "#0d47a1"
st.markdown(
    f"""
    <style>
      .block-container {{ max-width: 950px; }}
      h1, h2, h3, h4 {{ color: {ACCENT}; }}
      div.stButton > button {{
        background: {PRIMARY}; color: #fff; border: 0; border-radius: 10px; padding: 0.6rem 1.0rem;
        white-space: nowrap;
      }}
      div.stButton > button:hover {{ filter: brightness(0.95); }}
      .stProgress > div > div > div > div {{ background-color: {PRIMARY}; }}
      [data-testid="stMetric"] {{
        background: #f7faff; border: 1px solid #e6efff; border-radius: 12px; padding: .65rem;
      }}
      hr {{ border-top: 1px solid #e6efff; }}
      thead th {{ white-space: nowrap; }}
    </style>
    """,
    unsafe_allow_html=True,
)

def show_logo():
    logo_url = st.secrets.get("LOGO_URL", "").strip()
    with st.container():
        cols = st.columns([1, 9])
        with cols[0]:
            try:
                if logo_url:
                    st.image(logo_url, use_container_width=True)
                elif os.path.exists("logo.png"):
                    st.image("logo.png", use_container_width=True)
            except Exception:
                pass
        with cols[1]:
            st.title("Winsert Savings Calculator")

def progress_bar(step: int, total: int = 5):
    st.progress(step / total, text=f"Step {step} of {total}")

# =========================
# Token cache helpers
# =========================
def _get_cache_obj() -> SerializableTokenCache:
    cache = SerializableTokenCache()
    serialized = st.session_state.get("msal_cache_serialized")
    if serialized:
        try:
            cache.deserialize(serialized)
        except Exception:
            cache = SerializableTokenCache()
    return cache

def _save_cache_obj(cache: SerializableTokenCache):
    try:
        st.session_state["msal_cache_serialized"] = cache.serialize()
    except Exception:
        pass

def _build_pca(cache: SerializableTokenCache) -> PublicClientApplication:
    return PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

def acquire_token() -> str:
    cache = _get_cache_obj()
    app = _build_pca(cache)
    accounts = app.get_accounts()
    if accounts:
        res = app.acquire_token_silent(SCOPES, account=accounts[0])
        if "access_token" in res:
            _save_cache_obj(cache)
            return res["access_token"]
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("Failed to start device login.")
        st.stop()
    st.info(f"To authenticate, open **{flow['verification_uri']}** and enter code **{flow['user_code']}**")
    res = app.acquire_token_by_device_flow(flow)
    if "access_token" not in res:
        st.error(f"Auth error: {res.get('error_description', res)}")
        st.stop()
    _save_cache_obj(cache)
    return res["access_token"]

# =========================
# Graph helpers (with session reuse)
# =========================
def _item_base(token: str) -> str:
    if "drive_item" not in st.session_state:
        r = requests.get(f"{GRAPH}{WORKBOOK_PATH}", headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        st.session_state["drive_item"] = r.json()
    di = st.session_state["drive_item"]
    return f"{GRAPH}/drives/{di['parentReference']['driveId']}/items/{di['id']}/workbook"

def create_session(token: str) -> str:
    sid = st.session_state.get("excel_session_id")
    if sid:
        return sid
    r = requests.post(_item_base(token) + "/createSession",
                      headers={"Authorization": f"Bearer {token}"},
                      json={"persistChanges": True})
    r.raise_for_status()
    sid = r.json()["id"]
    st.session_state["excel_session_id"] = sid
    return sid

def close_session(token: str):
    sid = st.session_state.pop("excel_session_id", None)
    if not sid:
        return
    try:
        requests.post(_item_base(token) + "/closeSession",
                      headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid})
    except Exception:
        pass

def set_cell(token: str, sid: str, addr: str, value):
    payload = {"values": [[value]]}
    r = requests.patch(
        _item_base(token) + f"/worksheets('Office')/range(address='{addr}')",
        headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
        json=payload
    )
    r.raise_for_status()

def get_cell(token: str, sid: str, addr: str):
    r = requests.get(
        _item_base(token) + f"/worksheets('Office')/range(address='{addr}')",
        headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
    )
    r.raise_for_status()
    js = r.json()
    return js["values"][0][0] if js.get("values") else None

def calculate(token: str, sid: str):
    r = requests.post(
        _item_base(token) + "/application/calculate",
        headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
        json={"calculationType": "Full"}
    )
    r.raise_for_status()

# =========================
# Lists sheet: States & Cities (cached)
# =========================
def get_states_and_cities(token: str, sid: str):
    if "states_list" in st.session_state and "state_to_cities" in st.session_state:
        return st.session_state["states_list"], st.session_state["state_to_cities"]

    rng = "Lists!C1:BA100"
    r = requests.get(
        _item_base(token) + f"/worksheets('Lists')/range(address='{rng}')",
        headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
    )
    r.raise_for_status()
    values = r.json().get("values", [])
    num_cols = len(values[0]) if values else 0
    state_to_cities, states = {}, []
    for c in range(num_cols):
        col = [row[c] if c < len(row) else None for row in values]
        if not col or not col[0]:
            continue
        state = str(col[0]).strip()
        if not state:
            continue
        states.append(state)
        cities = [str(v).strip() for v in col[1:] if v not in (None, "")]
        state_to_cities[state] = cities

    st.session_state["states_list"] = states
    st.session_state["state_to_cities"] = state_to_cities
    return states, state_to_cities

# =========================
# Formatting helpers
# =========================
def _to_float(x):
    try:
        if x in (None, ""): return None
        if isinstance(x, (int, float)): return float(x)
        s = str(x).replace(",", "").strip()
        if s.endswith("%"):
            return float(s[:-1]) / 100.0
        return float(s)
    except Exception:
        return None

def fmt_pct_one_decimal(x):
    v = _to_float(x)
    if v is None: return "—"
    if v <= 1.0:
        v = v * 100.0
    return f"{v:.1f}%"

def fmt_int_units(x, units=""):
    v = _to_float(x)
    if v is None: return "—"
    return f"{round(v):,}{(' ' + units) if units else ''}"

def fmt_money_two_decimals(x):
    v = _to_float(x)
    if v is None: return "—"
    return f"${v:,.2f}"

def fmt_money_no_decimals(x):
    v = _to_float(x)
    if v is None: return "—"
    return f"${round(v):,}"

def fmt_int(x):
    v = _to_float(x)
    if v is None: return "—"
    return f"{round(v):,}"

def fmt_rate4(x):
    v = _to_float(x)
    if v is None: return "—"
    return f"${v:,.4f}"

# =========================
# Wizard helpers
# =========================
def set_step(n: int):
    st.session_state["step"] = n

def get_step() -> int:
    return st.session_state.get("step", 1)

# Step 1: Apply state
def apply_state(token, sid, state):
    if state:
        set_cell(token, sid, "C18", state)
        st.session_state["selected_state"] = state
        return True
    return False

# Step 2 helpers
def write_city_to_workbook(token, sid, city):
    set_cell(token, sid, "C19", city)

def check_climate(token, sid, city):
    # Write city (for climate calc), then compute HDD/CDD
    write_city_to_workbook(token, sid, city)
    calculate(token, sid)
    hdd = get_cell(token, sid, "C23")
    cdd = get_cell(token, sid, "C24")
    return hdd, cdd

# Step 3: Building inputs; optionally calc WWR
def save_building_inputs(token, sid, calc_for_wwr=False):
    bldg_area = st.session_state.get("bldg_area")
    floors    = st.session_state.get("floors")
    hvac      = st.session_state.get("hvac")
    existw    = st.session_state.get("existw")
    fuel      = st.session_state.get("fuel")
    cool      = st.session_state.get("cool")
    hours     = st.session_state.get("hours")
    csw_sf    = st.session_state.get("csw_sf")

    ready = all([
        isinstance(bldg_area, (int, float)) and bldg_area > 0,
        isinstance(floors, (int, float)) and floors >= 0,
        hvac, existw, fuel, cool,
        isinstance(hours, (int, float)) and hours > 0,
        isinstance(csw_sf, (int, float)) and csw_sf > 0,
    ])
    if not ready:
        return False, None

    set_cell(token, sid, "F18", bldg_area)
    set_cell(token, sid, "F19", floors)
    set_cell(token, sid, "F20", hvac)
    set_cell(token, sid, "F24", existw)
    set_cell(token, sid, "F21", fuel)
    set_cell(token, sid, "F22", cool)
    set_cell(token, sid, "F23", hours)
    set_cell(token, sid, "F27", csw_sf)

    if calc_for_wwr:
        calculate(token, sid)
        wwr = get_cell(token, sid, "F28")
        st.session_state["wwr_latest"] = wwr
        st.session_state["wwr_checked"] = True
        return True, wwr

    st.session_state["wwr_checked"] = False
    return True, None

# Ensure we commit a deferred city (if user skipped “Check Climate”)
def ensure_city_committed(token, sid):
    pending = st.session_state.get("pending_city")
    if pending:
        write_city_to_workbook(token, sid, pending)
        del st.session_state["pending_city"]

# Step 4: Rates; store inputs and commit any pending city
def save_rates(token, sid):
    cswtyp    = st.session_state.get("cswtyp")
    elec_rate = st.session_state.get("elec_rate")
    gas_rate  = st.session_state.get("gas_rate")
    ready = all([
        bool(cswtyp),
        (elec_rate is not None) and (elec_rate >= 0.0),
        (gas_rate  is not None) and (gas_rate  >= 0.0),
    ])
    if not ready:
        return False
    ensure_city_committed(token, sid)  # <-- commit deferred city here for results
    set_cell(token, sid, "F26", cswtyp)
    set_cell(token, sid, "C27", elec_rate)
    set_cell(token, sid, "C28", gas_rate)
    st.session_state["step4_applied"] = True
    return True

def read_results(token, sid):
    return {
        "eui_savings":  get_cell(token, sid, "E14"),
        "elec_savings": get_cell(token, sid, "F31"),
        "gas_savings":  get_cell(token, sid, "F33"),
        "elec_cost":    get_cell(token, sid, "C35"),
        "gas_cost":     get_cell(token, sid, "C36"),
        "total_savings":get_cell(token, sid, "F36"),
        "wwr":          get_cell(token, sid, "F28"),
    }

# Robust input summary direct from workbook (includes WWR under Hours)
def read_input_summary_from_workbook(token, sid):
    cells = [
        ("State",                               "C18",  "text"),
        ("City",                                "C19",  "text"),
        ("Building Area (ft²)",                 "F18",  "int"),
        ("Number of Floors",                    "F19",  "int"),
        ("HVAC System Type",                    "F20",  "text"),
        ("Existing Window Type",                "F24",  "text"),
        ("Heating Fuel",                        "F21",  "text"),
        ("Cooling Installed?",                  "F22",  "text"),
        ("Annual Operating Hours (hrs/yr)",     "F23",  "int"),
        ("Window-to-Wall Ratio",                "F28",  "pct1"),  # here per request
        ("CSW Installed (ft²)",                 "F27",  "int"),
        ("Type of CSW Analyzed",                "F26",  "text"),
        ("Electric Rate ($/kWh)",               "C27",  "rate"),
        ("Natural Gas Rate ($/therm)",          "C28",  "rate"),
    ]
    rows = []
    for label, addr, kind in cells:
        v = get_cell(token, sid, addr)
        if kind == "int":
            vf = _to_float(v); disp = "—" if vf is None else f"{int(round(vf)):,}"
        elif kind == "rate":
            disp = fmt_rate4(v)
        elif kind == "pct1":
            disp = fmt_pct_one_decimal(v)
        else:
            disp = "—" if v in (None, "") else str(v)
        rows.append({"Input": label, "Value": disp})
    return rows

# =========================
# APP
# =========================
token = acquire_token()
sid = create_session(token)

def show_header_and_progress():
    show_logo()
    step = get_step()
    progress_bar(step, total=5)
    st.markdown(f"### Step {step} of 5")

show_header_and_progress()
step = get_step()

try:
    # Load state/city lists up front (cached)
    states_list, state_to_cities = get_states_and_cities(token, sid)

    # -----------------
    # STEP 1: State (non-form for snappy nav)
    # -----------------
    if step == 1:
        st.header("1) Select State")
        default_state = st.session_state.get("selected_state")
        index = states_list.index(default_state) if default_state in states_list else 0
        st.selectbox("State", states_list, index=index, key="state_sel")
        cols = st.columns([1, 9])
        if cols[-1].button("Next →"):
            if apply_state(token, sid, st.session_state.get("state_sel")):
                set_step(2); st.rerun()
            else:
                st.warning("Please select a state to continue.")

    # -----------------
    # STEP 2: City (non-form inside container) — no Graph call on Next
    # -----------------
    elif step == 2:
        st.header("2) Select City")

        box2 = st.container()
        metrics2 = st.empty()  # placeholder for inline HDD/CDD

        with box2:
            cols_top = st.columns([1, 9])
            with cols_top[0]:
                if st.button("← Back", use_container_width=True):
                    box2.empty(); set_step(1); st.rerun()

            state = st.session_state.get("selected_state")
            if not state:
                st.info("Please select a State first.")
            else:
                # Reset city when state changes
                if st.session_state.get("last_state_for_city") != state:
                    st.session_state.pop("city_sel", None)
                    st.session_state["last_state_for_city"] = state

                city_options = state_to_cities.get(state, [])
                st.selectbox("City", city_options, key="city_sel")

                row = st.columns([2.2, 2.2, 2.2, 3.4, 2.0])
                check2 = row[0].button("Check Climate")
                next2  = row[-1].button("Next →")

        # Actions (outside box for instant clearing/navigation)
        if step == 2:
            if check2:
                city = st.session_state.get("city_sel")
                if not city:
                    st.warning("Please select a city to continue.")
                else:
                    hdd, cdd = check_climate(token, sid, city)
                    with metrics2.container():
                        row2 = st.columns([2.2, 2.2, 2.2, 3.4, 2.0])
                        row2[1].metric("Heating Degree Days", fmt_int(hdd))
                        row2[2].metric("Cooling Degree Days", fmt_int(cdd))
                    # stay on step 2
            elif next2:
                city = st.session_state.get("city_sel")
                if not city:
                    st.warning("Please select a city to continue.")
                else:
                    # Defer writing city to Excel for speed; commit at Step 4
                    st.session_state["pending_city"] = city
                    box2.empty(); set_step(3); st.rerun()

    # -------------------------
    # STEP 3: Building details (FORM in a container) — inline WWR
    # -------------------------
    elif step == 3:
        st.header("3) Building Information")

        box3 = st.container()
        metrics3 = st.empty()

        with box3.form("step3_building_form", clear_on_submit=False):
            cols_top = st.columns([1, 9])
            with cols_top[0]:
                if st.form_submit_button("← Back"):
                    box3.empty(); set_step(2); st.rerun()

            col1, col2 = st.columns(2)
            col1.number_input("Building Area (ft²)", min_value=0.0, step=100.0, key="bldg_area")
            col2.number_input("Number of Floors",   min_value=0,   step=1,      key="floors")
            col1.selectbox("HVAC System Type", [
                "Packaged VAV with electric reheat",
                "Packaged VAV with hydronic reheat",
                "Built-up VAV with hydronic reheat",
                "Other",
            ], key="hvac")
            col2.selectbox("Existing Window Type", ["Single pane", "Double pane"], key="existw")
            col1.selectbox("Heating Fuel", ["Electric", "Natural Gas", "None"],    key="fuel")
            col2.selectbox("Cooling Installed?", ["Yes", "No"],                    key="cool")
            col1.number_input("Annual Operating Hours (hrs/yr)", min_value=0, step=100, key="hours")
            col2.number_input("CSW Installed (ft²)",             min_value=0.0, step=50.0, key="csw_sf")

            row = st.columns([2.2, 2.2, 2.2, 3.4, 2.0])
            check3 = row[0].form_submit_button("Check WWR")
            next3  = row[-1].form_submit_button("Next →")

        if step == 3:
            if check3:
                ok, wwr = save_building_inputs(token, sid, calc_for_wwr=True)
                if not ok:
                    st.warning("Please complete all building fields.")
                else:
                    with metrics3.container():
                        row2 = st.columns([2.2, 2.2, 2.2, 3.4, 2.0])
                        row2[1].metric("Window-to-Wall Ratio", fmt_pct_one_decimal(wwr))
            elif next3:
                ok, _ = save_building_inputs(token, sid, calc_for_wwr=False)
                if ok:
                    box3.empty(); set_step(4); st.rerun()
                else:
                    st.warning("Please complete all building fields.")

    # ---------------------
    # STEP 4: Rates & CSW (FORM in a container) — flicker-free jump to results
    # ---------------------
    elif step == 4:
        st.header("4) Secondary Window & Rates")

        box4 = st.container()
        with box4.form("step4_rates_form", clear_on_submit=False):
            cols_top = st.columns([1, 9])
            with cols_top[0]:
                if st.form_submit_button("← Back"):
                    box4.empty(); set_step(3); st.rerun()

            col1, col2 = st.columns(2)
            col1.selectbox("Type of CSW Analyzed", ["Single", "Double"], key="cswtyp")
            col2.number_input("Electric Rate ($/kWh)", min_value=0.0, step=0.01, key="elec_rate")
            col1.number_input("Natural Gas Rate ($/therm)", min_value=0.0, step=0.01, key="gas_rate")
            row = st.columns([2.2, 2.2, 2.2, 3.4, 2.0])
            next4 = row[-1].form_submit_button("See Results →")

        if next4:
            if save_rates(token, sid):
                box4.empty(); set_step(5); st.rerun()
            else:
                st.warning("Please complete the CSW type and both rates.")

    # --------------
    # STEP 5: Results
    # --------------
    elif step == 5:
        st.header("Results")

        try:
            calculate(token, sid)  # one calc to include everything
            res = read_results(token, sid)

            # Group 1: Energy savings (EUI %, Electric kWh, Gas therms)
            st.subheader("Energy Savings")
            g1 = st.columns(3)
            g1[0].metric("EUI Savings",       fmt_pct_one_decimal(res["eui_savings"]))
            g1[1].metric("Electric Savings",  fmt_int_units(res["elec_savings"], "kWh/yr"))
            g1[2].metric("Gas Savings",       fmt_int_units(res["gas_savings"], "therms/yr"))

            st.divider()

            # Group 2: Annual Utility Savings (Electric $, Gas $, Total $)
            st.subheader("Annual Utility Savings")
            g2 = st.columns(3)
            g2[0].metric("Electric Cost Savings", fmt_money_two_decimals(res["elec_cost"]))
            g2[1].metric("Gas Cost Savings",      fmt_money_two_decimals(res["gas_cost"]))
            g2[2].metric("Total Savings",         fmt_money_no_decimals(res["total_savings"]))

            st.divider()

            # Summary table — read back from workbook cells (includes WWR under Hours)
            st.subheader("Your Inputs (Summary)")
            summary_rows = read_input_summary_from_workbook(token, sid)
            st.table(summary_rows)

        except requests.HTTPError as e:
            st.error(f"HTTP {e.response.status_code}: {e.response.text[:400]}")

        cols = st.columns([1, 1, 6, 2])
        with cols[0]:
            if st.button("← Back", use_container_width=True):
                set_step(4); st.rerun()
        with cols[-1]:
            if st.button("Start Over", use_container_width=True):
                close_session(token)
                keep = {"msal_cache_serialized", "drive_item"}
                for k in list(st.session_state.keys()):
                    if k not in keep:
                        del st.session_state[k]
                set_step(1); st.rerun()

finally:
    pass
