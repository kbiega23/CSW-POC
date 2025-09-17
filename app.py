# app.py — Winsert Savings Calculator (Office) — Wizard UI with friendly formatting
import re
import requests
import streamlit as st
from msal import PublicClientApplication, SerializableTokenCache

# =========================
# Config & sanity checks
# =========================
SCOPES = ["User.Read", "Files.ReadWrite"]  # Do NOT include 'offline_access'

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

st.set_page_config(page_title="Winsert Savings Calculator – Office (Wizard)", layout="centered")

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
                      json={"persistChanges": True})
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

def fmt_number_no_decimals(x, units=""):
    v = _to_float(x)
    if v is None: return "—"
    return f"{round(v):,}{(' ' + units) if units else ''}"

# =========================
# Wizard helpers
# =========================
def set_step(n: int):
    st.session_state["step"] = n

def get_step() -> int:
    return st.session_state.get("step", 1)

def nav_buttons(back_to=None, next_to=None, next_label="Next"):
    cols = st.columns([1, 1, 6, 2])
    with cols[0]:
        if back_to and st.button("← Back", use_container_width=True):
            set_step(back_to); st.rerun()
    with cols[-1]:
        if next_to and st.button(next_label, use_container_width=True):
            set_step(next_to); st.rerun()

def write_step1_and_get_hdd_cdd(token, sid, state, city):
    if state and city:
        set_cell(token, sid, "C18", state)
        set_cell(token, sid, "C19", city)
        calculate(token, sid)
        hdd = get_cell(token, sid, "C23")
        cdd = get_cell(token, sid, "C24")
        return hdd, cdd
    return None, None

def write_step2_and_get_wwr(token, sid, building_area, floors, hvac, existw, fuel, cool, hours, csw_sf):
    ready = all([
        building_area is not None and building_area > 0,
        floors is not None and floors >= 0,
        hvac, existw, fuel, cool,
        hours is not None and hours > 0,
        csw_sf is not None and csw_sf > 0
    ])
    if ready:
        set_cell(token, sid, "F18", building_area)
        set_cell(token, sid, "F19", floors)
        set_cell(token, sid, "F20", hvac)
        set_cell(token, sid, "F24", existw)
        set_cell(token, sid, "F21", fuel)
        set_cell(token, sid, "F22", cool)
        set_cell(token, sid, "F23", hours)
        set_cell(token, sid, "F27", csw_sf)
        calculate(token, sid)
        wwr = get_cell(token, sid, "F28")
        return wwr
    return None

def write_step3_inputs(token, sid, cswtyp, elec_rate, gas_rate):
    if cswtyp:
        set_cell(token, sid, "F26", cswtyp)
    if elec_rate is not None:
        set_cell(token, sid, "C27", elec_rate)
    if gas_rate is not None:
        set_cell(token, sid, "C28", gas_rate)
    calculate(token, sid)

def read_results(token, sid):
    return {
        "eui_savings": get_cell(token, sid, "E14"),
        "elec_savings": get_cell(token, sid, "F31"),
        "gas_savings": get_cell(token, sid, "F33"),
        "elec_cost": get_cell(token, sid, "C35"),
        "gas_cost": get_cell(token, sid, "C36"),
        "total_savings": get_cell(token, sid, "F36"),
        "wwr": get_cell(token, sid, "F28"),
    }

# =========================
# APP
# =========================
token = acquire_token()
sid = create_session(token)

# Global title for every step
st.title("Winsert Savings Calculator")

# Step indicator
step = get_step()
st.markdown(f"### Step {step} of 4")

try:
    # Load state/city lists up front
    states_list, state_to_cities = get_states_and_cities(token, sid)

    # -----------------
    # STEP 1: Location
    # -----------------
    if step == 1:
        st.header("1) Select Location")
        col1, col2 = st.columns(2)

        def _on_state_change():
            st.session_state.pop("city_sel", None)

        state = col1.selectbox("State", states_list, key="state_sel", on_change=_on_state_change)
        city = col2.selectbox("City", state_to_cities.get(state, []), key="city_sel")

        # Immediate gut check: HDD/CDD once both chosen
        hdd, cdd = write_step1_and_get_hdd_cdd(token, sid, state, city)
        st.divider()
        st.subheader("Climate check")
        if hdd is not None and cdd is not None:
            cols = st.columns(2)
            cols[0].metric("Heating Degree Days", fmt_number_no_decimals(hdd, "°F·days"))
            cols[1].metric("Cooling Degree Days", fmt_number_no_decimals(cdd, "°F·days"))
            st.caption("If these look off for the chosen city, pick a different city/state.")
        else:
            st.info("Choose both State and City to see HDD/CDD.")

        nav_buttons(back_to=None, next_to=2, next_label="Next →")

    # -------------------------
    # STEP 2: Building details
    # -------------------------
    elif step == 2:
        st.header("2) Building Information")
        col1, col2 = st.columns(2)

        building_area = col1.number_input("Building Area (ft²)", min_value=0.0, step=100.0, key="bldg_area")
        floors        = col2.number_input("Number of Floors", min_value=0, step=1, key="floors")

        # HVAC System Type included here
        hvac = col1.selectbox("HVAC System Type", [
            "Packaged VAV with electric reheat",
            "Packaged VAV with hydronic reheat",
            "Built-up VAV with hydronic reheat",
            "Other",
        ], key="hvac")

        existw = col2.selectbox("Existing Window Type", ["Single pane", "Double pane"], key="existw")
        fuel   = col1.selectbox("Heating Fuel", ["Electric", "Natural Gas", "None"], key="fuel")
        cool   = col2.selectbox("Cooling Installed?", ["Yes", "No"], key="cool")
        hours  = col1.number_input("Annual Operating Hours (hrs/yr)", min_value=0, step=100, key="hours")
        csw_sf = col2.number_input("CSW Installed (ft²)", min_value=0.0, step=50.0, key="csw_sf")

        # Show WWR once all building info present (as percent, 1 decimal)
        wwr = write_step2_and_get_wwr(token, sid, building_area, floors, hvac, existw, fuel, cool, hours, csw_sf)
        st.divider()
        st.subheader("Envelope check")
        if wwr is not None:
            st.metric("Window-to-Wall Ratio", fmt_pct_one_decimal(wwr))
            st.caption("If WWR seems off, adjust inputs before proceeding.")
        else:
            st.info("Enter all building details to estimate Window-to-Wall Ratio.")

        nav_buttons(back_to=1, next_to=3, next_label="Next →")

    # ---------------------
    # STEP 3: Rates & CSW
    # ---------------------
    elif step == 3:
        st.header("3) Secondary Window & Rates")
        col1, col2 = st.columns(2)

        cswtyp    = col1.selectbox("Type of CSW Analyzed", ["Single", "Double"], key="cswtyp")
        elec_rate = col2.number_input("Electric Rate ($/kWh)", min_value=0.0, step=0.01, key="elec_rate")
        gas_rate  = col1.number_input("Natural Gas Rate ($/therm)", min_value=0.0, step=0.01, key="gas_rate")

        # Write step 3 inputs on next to precompute results
        cols = st.columns([1, 1, 6, 2])
        with cols[0]:
            if st.button("← Back", use_container_width=True):
                set_step(2); st.rerun()
        with cols[-1]:
            if st.button("See Results →", use_container_width=True):
                try:
                    write_step3_inputs(token, sid, cswtyp, elec_rate, gas_rate)
                except requests.HTTPError as e:
                    st.error(f"HTTP {e.response.status_code}: {e.response.text[:400]}")
                set_step(4); st.rerun()

    # --------------
    # STEP 4: Results
    # --------------
    elif step == 4:
        st.header("4) Results")
        try:
            calculate(token, sid)
            res = read_results(token, sid)

            # EUI Savings & WWR as % (1 decimal)
            stats = st.columns(2)
            stats[0].metric("EUI Savings", fmt_pct_one_decimal(res["eui_savings"]))
            stats[1].metric("Window-to-Wall Ratio", fmt_pct_one_decimal(res["wwr"]))

            st.divider()
            st.subheader("Energy & Cost Savings (Annual)")
            c1, c2 = st.columns(2)
            with c1:
                st.metric("Electric Savings", fmt_int_units(res["elec_savings"], "kWh/yr"))
                st.metric("Gas Savings", fmt_int_units(res["gas_savings"], "therms/yr"))
                st.metric("Total Savings", fmt_int_units(res["total_savings"], "$/yr"))
            with c2:
                st.metric("Electric Cost Savings", fmt_money_two_decimals(res["elec_cost"]))
                st.metric("Gas Cost Savings", fmt_money_two_decimals(res["gas_cost"]))

        except requests.HTTPError as e:
            st.error(f"HTTP {e.response.status_code}: {e.response.text[:400]}")

        cols = st.columns([1, 1, 6, 2])
        with cols[0]:
            if st.button("← Back", use_container_width=True):
                set_step(3); st.rerun()
        with cols[-1]:
            if st.button("Start Over", use_container_width=True):
                # Clear step + keep token cache so you don't have to log in again
                for k in list(st.session_state.keys()):
                    if k not in ("msal_cache_serialized", "drive_item"):
                        del st.session_state[k]
                set_step(1); st.rerun()

finally:
    # Close workbook session to be tidy (values are persisted due to persistChanges=True)
    close_session(token, sid)
