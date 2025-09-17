import os, time, requests, streamlit as st
from msal import PublicClientApplication

# ------------------ CONFIG (put in Streamlit secrets) ------------------
TENANT_ID     = st.secrets["TENANT_ID"]
CLIENT_ID     = st.secrets["CLIENT_ID"]
WORKBOOK_PATH = st.secrets["WORKBOOK_PATH"]  # e.g. "/me/drive/root:/Documents/CSW Savings Calculator 2_0_0_Unlocked.xlsx"

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Files.ReadWrite",
    "offline_access",
]


# ------------------ AUTH: Device Code flow ------------------
def acquire_token():
    if "token_cache" in st.session_state:
        tok = st.session_state["token_cache"]
        if tok and time.time() < tok["expires_at"] - 60:
            return tok["access_token"]

    app = PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}")
    accounts = app.get_accounts()
    if accounts:
        res = app.acquire_token_silent(SCOPES, account=accounts[0])
        if "access_token" in res:
            st.session_state["token_cache"] = {"access_token": res["access_token"], "expires_at": time.time()+res.get("expires_in", 3600)}
            return res["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error("Failed to start device login.")
        st.stop()

    st.info(f"Go to **{flow['verification_uri']}** and enter code **{flow['user_code']}**")
    res = app.acquire_token_by_device_flow(flow)
    if "access_token" not in res:
        st.error(res.get("error_description", res))
        st.stop()
    st.session_state["token_cache"] = {"access_token": res["access_token"], "expires_at": time.time()+res.get("expires_in", 3600)}
    return res["access_token"]

def item_base(token):
    if "drive_item" not in st.session_state:
        url = f"{GRAPH}{WORKBOOK_PATH}"
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        st.session_state["drive_item"] = r.json()
    di = st.session_state["drive_item"]
    drive_id = di["parentReference"]["driveId"]
    item_id  = di["id"]
    return f"{GRAPH}/drives/{drive_id}/items/{item_id}/workbook"

def create_session(token):
    r = requests.post(item_base(token)+"/createSession", headers={"Authorization": f"Bearer {token}"}, json={"persistChanges": True})
    r.raise_for_status()
    return r.json()["id"]

def close_session(token, sid):
    try:
        requests.post(item_base(token)+"/closeSession", headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid})
    except:
        pass

def set_cell(token, sid, addr, value):
    payload = {"values": [[value]]}
    r = requests.patch(item_base(token)+f"/worksheets('Office')/range(address='{addr}')",
                       headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
                       json=payload)
    r.raise_for_status()

def get_cell(token, sid, addr):
    r = requests.get(item_base(token)+f"/worksheets('Office')/range(address='{addr}')",
                     headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid})
    r.raise_for_status()
    return r.json()["values"][0][0]

def calculate(token, sid):
    r = requests.post(item_base(token)+"/application/calculate",
                      headers={"Authorization": f"Bearer {token}", "workbook-session-id": sid},
                      json={"calculationType": "Full"})
    r.raise_for_status()

# ------------------ UI ------------------
st.set_page_config(page_title="CSW Savings Estimator – Office (POC)", layout="centered")
st.title("CSW Savings Estimator – Office (POC)")

token = acquire_token()
session_id = create_session(token)

col1, col2 = st.columns(2)
state  = col1.text_input("State")
city   = col2.text_input("City")
hvac   = col1.selectbox("HVAC System Type", [
    "Packaged VAV with electric reheat",
    "Packaged VAV with hydronic reheat",
    "Built-up VAV with hydronic reheat",
    "Other",
])
fuel   = col2.selectbox("Heating Fuel", ["Electric", "Natural Gas", "None"])
cool   = col1.selectbox("Cooling Installed?", ["Yes", "No"])
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
        set_cell(token, session_id, "C18", state)
        set_cell(token, session_id, "C19", city)
        set_cell(token, session_id, "F20", hvac)
        set_cell(token, session_id, "F21", fuel)
        set_cell(token, session_id, "F22", cool)
        set_cell(token, session_id, "F24", existw)
        set_cell(token, session_id, "F26", cswtyp)
        set_cell(token, session_id, "F18", bldg_area)
        set_cell(token, session_id, "F19", floors)
        set_cell(token, session_id, "F23", hours)
        set_cell(token, session_id, "F27", csw_sf)
        set_cell(token, session_id, "C27", elec_rate)
        set_cell(token, session_id, "C28", gas_rate)
        calculate(token, session_id)
        out = {
            "Location HDD (C23)": get_cell(token, session_id, "C23"),
            "Location CDD (C24)": get_cell(token, session_id, "C24"),
            "Est. Window Wall Ratio (F28)": get_cell(token, session_id, "F28"),
            "Electric Savings (F31)": get_cell(token, session_id, "F31"),
            "Gas Savings (F33)": get_cell(token, session_id, "F33"),
            "EUI Savings (E14)": get_cell(token, session_id, "E14"),
            "Electric Cost Savings (C35)": get_cell(token, session_id, "C35"),
            "Gas Cost Savings (C36)": get_cell(token, session_id, "C36"),
            "Total Savings (F36)": get_cell(token, session_id, "F36"),
        }
        st.subheader("Results")
        for k, v in out.items():
            st.write(f"**{k}**: {v}")
    except Exception as e:
        st.error(f"Error: {e}")
    finally:
        close_session(token, session_id)
else:
    st.caption("Enter inputs, click Calculate. First run will prompt you to authenticate.")
