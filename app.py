# app.py
import os
import io
import tempfile
from pathlib import Path
import requests
import streamlit as st
import msal
from dotenv import load_dotenv

load_dotenv()

st.set_page_config(page_title="CNET Reports", layout="wide")
st.title("CNET Reports")

# =========================
# ENV
# =========================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

SP_HOSTNAME = os.getenv("SP_HOSTNAME")      # groupcastillo.sharepoint.com
SP_SITE_PATH = os.getenv("SP_SITE_PATH")    # /sites/GroupCastilloTeamSite
SP_DRIVE_NAME = os.getenv("SP_DRIVE_NAME", "Documents")

TARGET_EXCEL_PATH = os.getenv(
    "SP_FILE_PATH",
    "General/12433087 CANADA INC-MASTER/21-Work Orders-Complaints-Request/WorkOrders-Complaints-Master-2025-v1.xlsm"
)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

# =========================
# HELPERS
# =========================
def die(msg):
    st.error(msg)
    st.stop()

def graph_get(url, token):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json()

def graph_download(url, token):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.content

# =========================
# AUTH
# =========================
def get_token():
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        die(flow)

    st.info(f"Open {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        die(result)

    return result["access_token"]

# =========================
# SHAREPOINT
# =========================
def resolve_drive_id(token):
    site = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{SP_HOSTNAME}:{SP_SITE_PATH}",
        token
    )
    drives = graph_get(
        f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives",
        token
    )["value"]

    drive = next((d for d in drives if d["name"] == SP_DRIVE_NAME), drives[0])
    return drive["id"]

def download_excel(token):
    drive_id = resolve_drive_id(token)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{TARGET_EXCEL_PATH}:/content"
    content = graph_download(url, token)

    tmp = Path(tempfile.gettempdir()) / "cnet_reports"
    tmp.mkdir(exist_ok=True)
    local = tmp / Path(TARGET_EXCEL_PATH).name
    local.write_bytes(content)

    return str(local)

# =========================
# UI
# =========================
if "excel_path" not in st.session_state:

    if st.button("Connect to Microsoft"):
        token = get_token()
        path = download_excel(token)
        st.session_state["excel_path"] = path
        st.success("Excel downloaded")
        st.rerun()

    st.stop()

st.success("Excel ready")
st.write("Local cache:", st.session_state["excel_path"])
st.write("Open Tickets page from sidebar ⬅")
