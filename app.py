# app.py
import re
import requests
import streamlit as st
from msal import PublicClientApplication, SerializableTokenCache

# =========================
# Config & sanity checks
# =========================
SCOPES = ["User.Read", "Files.ReadWrite"]  # Do NOT include 'offline_access' here

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

st.set_page_config(page_title="CSW Savings Estimator – Office (POC)", layout="centered")
st.title("CSW Savings Estimator – Office (POC)")

# =========================
# MSAL token cache helpers
# =========================
def _get_cache_obj() -> SerializableTokenCache:
    """Load (or create) an MSAL token cache object from session_state."""
    cache = SerializableTokenCache()
    serialized = st.session_state.get("msal_cache_serialized")
    if serialized:
        try:
            cache.deserialize(serialized)
        except Exception:
            # Corrupt/empty cache: start fresh
            cache = SerializableTokenCache()
    return cache

def _save_cache_obj(cache: SerializableTokenCache):
    """Persist the MSAL token cache back into session_state."""
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
    """
    Acquire an access token using MSAL Device Code flow, but:
    - Try silent acquisition first (uses cached refresh token).
    - Only prompt with a device code if needed.
    """
    cache = _get_cache_obj()
    app = _build_pca(cache)

    # 1) Silent first
    accounts = app.get_accounts()
    if accounts:
        res = app.acquire_token_silent(SCOPES, account=accounts[0])
        if "access_token" in res:
            _save_cache_obj(cache)
            return res["access_token"]

    # 2) Device code only if silent failed
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("Failed to start device login.")
        st.stop()
    st.info(f"To authenticate, open **{flow['verification_uri']}** and enter code **{flow['user_code']}**")
    res = app.acquire_token_by_device_flow(flow)  # blocks until you complete login

    if "access_token" not in res:
        st.error(f"Auth error: {res.get('error_description', res)}")
        st.stop()

    _save_cache_obj(cache)
    return res["access_token"]

# =========================
# Graph helpers
# =========================
def _item_base(token: str) -> str:
    """Return the workbook base URL (drives/{driveId}/items/{itemId}/workbook)."""
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
    """
    Reads Lists!C1:BA100:
      - Row 1: state names (headers)
      - Rows below: cities
    Returns: (states_list, {state: [cities...]})
    """
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
    state_to_cities = {}
    states = []
    for c in range(num_cols):
        col = [row[c] if c < len(row) else None for row in values]
        if not col:
            continue
        header = col[0]
        if not header:
            continue
        state = str(header).strip()
        if not state:
            continue
        states.append(state)
        cities = [str(v).strip() for v in col[1:] if v not in (None, "")]
        state_to_cities[state] = cities
    return states, state_to_cities

# =========================
# App flow
# =========================
# Acquire (or silently refresh) token once per session/rerun
token = acquire_token()

# Create a workbook session for this run (closed on success/failure)
sid = create_session(token)

try:
    # Pull State/City options dynamically
    states_list, state_to_cities = get_states_and_cities(token, sid)

    col1, col2 = st.columns(2)

    # Reset city when state changes
    def _on_state_change():
        st.session_state.pop("city_sel", None)

    state = col1.selectbox("State", states_list, key="state_sel", on_change=_on_state_change)
    city = col2.selectbox("City", state_to_cities.get(state, []), key="city_sel")

    # Other dropdowns (hardcoded here, can be wired to Lists later)
    hvac = col1.selectbox("HVAC System Type", [
        "Packaged VAV with electric reheat",
        "Packaged VAV with hydronic reheat",
        "Built-up VAV with hydronic reheat",
        "Other",
    ])
    fuel = col2.selectbox("Heating Fuel", ["Electric", "Natural Gas", "None"])
    cool = col1.selectbox("Cooling Installed?", ["Yes", "No"])
    existw = col2.selectbox("Type of Existing Window", ["Single pane", "Double pane"])
    cswtyp = col1.selectbox("Type of CSW Analyzed", ["Single", "Double"])

    elec_rate = col2.number_input("Electric Rate ($/kWh)", min_value=0.0, step=0.01)
    gas_rate  = col1.number_input("Natural Gas Rate ($/therm)", min_value=0.0, step=0.01)
    bldg_area = col2.number_input("Building Area (ft²)", min_value=0.0, step=100.0)
    floors    = col1.number_input("No. of Floors", min_value=0, step=1)
    hours     = col2.number_input("Annual Operating Hours", min_value=0, step=100)
    csw_sf    = col1.number_input("Sq.ft. of CSW Installed", min_value=0.0, step=50.0)

    if st.button("Calculate"):
        try:
            # Write inputs
            set_cell(token, sid, "C18", state)
            set_cell(token, sid, "C19", city)
            set_cell(token, sid, "F20", hvac)
            set_cell(token, sid, "F21", fuel)
            set_cell(token, sid, "F22", cool)
            set_cell(token, sid, "F24", existw)
            set_cell(token, sid, "F26", cswtyp)

            set_cell(token, sid, "F18", bldg_area)
            set_cell(token, sid, "F19", floors)
            set_cell(token, sid, "F23", hours)
            set_cell(token, sid, "F27", csw_sf)
            set_cell(token, sid, "C27", elec_rate)
            set_cell(token, sid, "C28", gas_rate)

            # Recalculate and read outputs
            calculate(token, sid)
            out = {
                "Location HDD (C23)": get_cell(token, sid, "C23"),
                "Location CDD (C24)": get_cell(token, sid, "C24"),
                "Est. Window Wall Ratio (F28)": get_cell(token, sid, "F28"),
                "Electric Savings (F31)": get_cell(token, sid, "F31"),
                "Gas Savings (F33)": get_cell(token, sid, "F33"),
                "EUI Savings (E14)": get_cell(token, sid, "E14"),
                "Electric Cost Savings (C35)": get_cell(token, sid, "C35"),
                "Gas Cost Savings (C36)": get_cell(token, sid, "C36"),
                "Total Savings (F36)": get_cell(token, sid, "F36"),
            }
            st.subheader("Results")
            for k, v in out.items():
                st.write(f"**{k}**: {v}")
        except requests.HTTPError as e:
            st.error(f"HTTP {e.response.status_code}: {e.response.text[:400]}")
finally:
    close_session(token, sid)
