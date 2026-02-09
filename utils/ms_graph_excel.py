# utils/ms_graph_excel.py
import os
import tempfile
from pathlib import Path

import msal
import requests
import streamlit as st

# -------------------------
# ENV
# -------------------------
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

SP_HOSTNAME = os.getenv("SP_HOSTNAME")
SP_SITE_PATH = os.getenv("SP_SITE_PATH")
SP_DRIVE_NAME = os.getenv("SP_DRIVE_NAME", "Documents")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

# -------------------------
# TOKEN CACHE (disk)
# -------------------------
def _cache_path() -> Path:
    # persistent across reruns, avoids re-login until refresh token expires
    d = Path(tempfile.gettempdir()) / "cnet_reports"
    d.mkdir(exist_ok=True)
    return d / "msal_token_cache.bin"

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = _cache_path()
    if p.exists():
        cache.deserialize(p.read_text(encoding="utf-8"))
    return cache

def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        _cache_path().write_text(cache.serialize(), encoding="utf-8")

def _msal_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    if not TENANT_ID or not CLIENT_ID:
        raise RuntimeError("Missing TENANT_ID / CLIENT_ID in environment.")
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

def _get_token_silent_or_interactive() -> str:
    cache = _load_cache()
    app = _msal_app(cache)

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    # Not silent-capable yet → needs device login (cannot be fully invisible)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(str(flow))

    # Minimal UI for first-time auth
    with st.expander("Microsoft sign-in required", expanded=True):
        st.info(f"Open {flow['verification_uri']} and enter code: {flow['user_code']}")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(str(result))

    _save_cache(cache)
    return result["access_token"]

# -------------------------
# GRAPH HELPERS
# -------------------------
def _graph_get(url: str, token: str):
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json()

def _graph_download(url: str, token: str) -> bytes:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.content

def _resolve_drive_id(token: str) -> str:
    if not SP_HOSTNAME or not SP_SITE_PATH:
        raise RuntimeError("Missing SP_HOSTNAME / SP_SITE_PATH in environment.")

    site = _graph_get(f"https://graph.microsoft.com/v1.0/sites/{SP_HOSTNAME}:{SP_SITE_PATH}", token)
    drives = _graph_get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", token)["value"]
    drive = next((d for d in drives if d.get("name") == SP_DRIVE_NAME), drives[0])
    return drive["id"]

# -------------------------
# PUBLIC API (cached download)
# -------------------------
@st.cache_data(show_spinner=False)
def download_excel_cached(sp_relative_path: str, ttl_seconds: int) -> str:
    """
    Downloads file from SharePoint and stores it in a local temp cache.
    Streamlit cache TTL controls refresh frequency.
    """
    token = _get_token_silent_or_interactive()
    drive_id = _resolve_drive_id(token)

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_relative_path}:/content"
    content = _graph_download(url, token)

    out_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    out_dir.mkdir(exist_ok=True)

    local = out_dir / Path(sp_relative_path).name
    local.write_bytes(content)
    return str(local)
