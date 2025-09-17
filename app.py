# app.py — Winsert Savings Calculator (Office)
# Wizard UI with blue branding, token cache, Lists-driven State/City,
# Step 2 & Step 3 as FORMS (single calc per step), and friendly result formatting.

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
        background: {PRIMARY}; color: #fff; border: 0; border-radius: 10px; padding: 0.5rem 1rem;
      }}
      div.stButton > button:hover {{ filter: brightness(0.95); }}
      .stProgress > div > div > div > div {{ background-color: {PRIMARY}; }}
      [data-testid="stMetric"] {{
        background: #f7faff; border: 1px solid #e6efff; border-radius: 12px; padding: .75rem;
      }}
      hr {{ border-top: 1px solid #e6efff; }}
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

def progress_bar(step: int, total: int = 4):
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
# Graph helpers
# =========================
def _item_base(token: str) -> str:
    if "drive_item" not in st.session_state:
        r = requests.get(f"{GRAPH}{WORKBOOK_PATH}", headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        st.session_state["drive_item"] = r.json()
    di = st.session_state["drive_item"]
    return f"{GRAPH}/drives/{di['parentReference']['driveId']}/items/{di['id']}/workbook"

def create_session(token: str) -> str:
    r = requests.post(_item_base(token) + "/createSession",
                      headers={"Authorization": f"Bearer {token}"},
                      json={"persistChanges": True})  # current behavior: write back
    r.raise_for_status()
    return r.json()["id"]

def close_session(token: str, sid: str):
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
# Lists sheet: States & Cities
# =========================
def get_states_and_cities(token: str, sid: str):
    """Reads Lists!C1:BA100 → row1 = states, rows below = cities."""
    rng = "Lists!C1:BA100"
    r = requests.get(
        _item_base(token) + f"/worksheets('Lists')/range(address='{rng}')",
        headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
    )
    r.raise_for_status()
    values = r.json().get("values", [])
    if not values:
        return [], {}
    num_cols = len(values[0])
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
    """Show as percentage with 1 decimal; treat <=1 as fraction, >1 as already percent."""
    v = _to_float(x)
    if v is None: return "—"
    if v <= 1.0:
        v = v * 100.0
    return f"{v:.1f}%"

def fmt_int_units(x, units=""):
    """Nearest whole number with thousands separator + units."""
    v = _to_float(x)
    if v is None: return "—"
    return f"{round(v):,}{(' ' + units) if units else ''}"

def fmt_money_two_decimals(x):
    """$ with cents, thousands separator."""
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

# =========================
# Wizard helpers
# =========================
def set_step(n: int):
    st.session_state["step"] = n

def get_step() -> int:
    return st.session_state.get("step", 1)

# Step 1 immediate write/recalc helper
def apply_step1_and_get_hdd_cdd(token, sid, state, city):
    if state and city:
        set_cell(token, sid, "C18", state)
        set_cell(token, sid, "C19", city)
        calculate(token, sid)
        hdd = get_cell(token, sid, "C23")
        cdd = get_cell(token, sid, "C24")
        return hdd, cdd
    return None, None

# Step 2 single-shot apply via FORM
def apply_step2(token, sid):
    """Validate Step 2 inputs, write them once, calc, return (ok, WWR)."""
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

    calculate(token, sid)
    wwr = get_cell(token, sid, "F28")

    st.session_state["wwr_latest"] = wwr
    st.session_state["step2_applied"] = True
    return True, wwr

# Step 3 single-shot apply via FORM
def apply_step3(token, sid):
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

    set_cell(token, sid, "F26", cswtyp)
    set_cell(token, sid, "C27", elec_rate)
    set_cell(token, sid, "C28", gas_rate)
    calculate(token, sid)

    st.session_state["step3_applied"] = True
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

# =========================
# APP
# =========================
token = acquire_token()
sid = create_session(token)

show_logo()
step = get_step()
progress_bar(step)
st.markdown(f"### Step {step} of 4")

try:
    # Load state/city lists up front
    states_list, state_to_cities = get_states_and_cities(token, sid)

    # -----------------
    # STEP 1: Location (immediate HDD/CDD)
    # -----------------
    if step == 1:
        st.header("1) Select Location")
        col1, col2 = st.columns(2)

        def _on_state_change():
            st.session_state.pop("city_sel", None)

        state = col1.selectbox("State", states_list, key="state_sel", on_change=_on_state_change)
        city  = col2.selectbox("City", state_to_cities.get(state, []), key="city_sel")

        hdd, cdd = apply_step1_and_get_hdd_cdd(token, sid, state, city)
        st.divider()
        st.subheader("Climate check")
        if hdd is not None and cdd is not None:
            cols = st.columns(2)
            cols[0].metric("Heating Degree Days", fmt_int(hdd))   # unitless per your request
            cols[1].metric("Cooling Degree Days", fmt_int(cdd))   # unitless
            st.caption("If these look off for the chosen city, pick a different city/state.")
        else:
            st.info("Choose both State and City to see HDD/CDD.")

        cols = st.columns([1, 1, 6, 2])
        with cols[-1]:
            if st.button("Next →", use_container_width=True):
                set_step(2); st.rerun()

    # -------------------------
    # STEP 2: Building details (FORM)
    # -------------------------
    elif step == 2:
        st.header("2) Building Information")

        # Back button (outside the form so it doesn't submit)
        cols_top = st.columns([1, 9])
        with cols_top[0]:
            if st.button("← Back", use_container_width=True):
                set_step(1); st.rerun()

        with st.form("step2_form", clear_on_submit=False):
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

            st.caption("Fill all fields, then click **Apply & Check WWR** once.")
            fcols = st.columns([1, 1, 6, 2])
            apply2 = fcols[0].form_submit_button("Apply & Check WWR")
            next2  = fcols[-1].form_submit_button("Next →")

        if apply2 or next2:
            ok, wwr = apply_step2(token, sid)
            if not ok:
                st.warning("Please complete all building fields before proceeding.")
            else:
                st.success("Building info applied.")
                st.metric("Window-to-Wall Ratio", fmt_pct_one_decimal(wwr))
                if next2:
                    set_step(3); st.rerun()
        elif st.session_state.get("step2_applied") and st.session_state.get("wwr_latest") is not None:
            st.divider()
            st.subheader("Envelope check")
            st.metric("Window-to-Wall Ratio", fmt_pct_one_decimal(st.session_state["wwr_latest"]))
            st.caption("If WWR seems off, adjust inputs and click Apply again.")

    # ---------------------
    # STEP 3: Rates & CSW (FORM)
    # ---------------------
    elif step == 3:
        st.header("3) Secondary Window & Rates")

        # Back button (outside the form)
        cols_top = st.columns([1, 9])
        with cols_top[0]:
            if st.button("← Back", use_container_width=True):
                set_step(2); st.rerun()

        with st.form("step3_form", clear_on_submit=False):
            col1, col2 = st.columns(2)
            col1.selectbox("Type of CSW Analyzed", ["Single", "Double"], key="cswtyp")
            col2.number_input("Electric Rate ($/kWh)", min_value=0.0, step=0.01, key="elec_rate")
            col1.number_input("Natural Gas Rate ($/therm)", min_value=0.0, step=0.01, key="gas_rate")

            st.caption("Enter all values, then click **Apply Rates** once.")
            fcols = st.columns([1, 1, 6, 2])
            apply3 = fcols[0].form_submit_button("Apply Rates")
            next3  = fcols[-1].form_submit_button("See Results →")

        if apply3 or next3:
            ok = apply_step3(token, sid)
            if not ok:
                st.warning("Please complete the CSW type and both rates before proceeding.")
            else:
                st.success("Rates applied.")
                if next3:
                    set_step(4); st.rerun()

    # --------------
    # STEP 4: Results
    # --------------
    elif step == 4:
        st.header("4) Results")
        try:
            calculate(token, sid)
            res = read_results(token, sid)

            # Group 1: Energy savings (EUI %, Electric kWh, Gas therms)
            st.subheader("Energy Savings")
            g1 = st.columns(3)
            g1[0].metric("EUI Savings",       fmt_pct_one_decimal(res["eui_savings"]))
            g1[1].metric("Electric Savings",  fmt_int_units(res["elec_savings"], "kWh/yr"))
            g1[2].metric("Gas Savings",       fmt_int_units(res["gas_savings"], "therms/yr"))

            st.divider()

            # Group 2: Cost impact (Electric $, Gas $, Total $)
            st.subheader("Cost Impact")
            g2 = st.columns(3)
            g2[0].metric("Electric Cost Savings", fmt_money_two_decimals(res["elec_cost"]))
            g2[1].metric("Gas Cost Savings",      fmt_money_two_decimals(res["gas_cost"]))
            g2[2].metric("Total Savings",         fmt_money_no_decimals(res["total_savings"]))

            # Optional: show WWR again
            st.divider()
            st.subheader("Envelope")
            st.metric("Window-to-Wall Ratio", fmt_pct_one_decimal(res["wwr"]))

        except requests.HTTPError as e:
            st.error(f"HTTP {e.response.status_code}: {e.response.text[:400]}")

        cols = st.columns([1, 1, 6, 2])
        with cols[0]:
            if st.button("← Back", use_container_width=True):
                set_step(3); st.rerun()
        with cols[-1]:
            if st.button("Start Over", use_container_width=True):
                # Keep login & drive item, clear the rest
                for k in list(st.session_state.keys()):
                    if k not in ("msal_cache_serialized", "drive_item"):
                        del st.session_state[k]
                set_step(1); st.rerun()

finally:
    # Current behavior: close workbook session each run (changes persist)
    close_session(token, sid)
