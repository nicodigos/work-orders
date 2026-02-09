# app.py
import os
import tempfile
from pathlib import Path

import msal
import requests
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# ==========================================
# APP CONFIG
# ==========================================
st.set_page_config(page_title="CNET Reports", layout="wide")
st.title("CNET Reports")

# ==========================================
# ENV
# ==========================================
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.Read.All"]

# ==========================================
# TOKEN CACHE (disk) so pages can be silent
# ==========================================
def _token_cache_path() -> Path:
    d = Path(tempfile.gettempdir()) / "cnet_reports"
    d.mkdir(exist_ok=True)
    return d / "msal_token_cache.bin"

def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = _token_cache_path()
    if p.exists():
        cache.deserialize(p.read_text(encoding="utf-8"))
    return cache

def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        _token_cache_path().write_text(cache.serialize(), encoding="utf-8")

def _msal_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    if not TENANT_ID or not CLIENT_ID:
        st.error("Missing TENANT_ID / CLIENT_ID in .env")
        st.stop()
    return msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

def _acquire_token_interactive() -> str:
    cache = _load_cache()
    app = _msal_app(cache)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        st.error(str(flow))
        st.stop()

    st.info(f"Open {flow['verification_uri']} and enter code: {flow['user_code']}")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        st.error(str(result))
        st.stop()

    _save_cache(cache)
    return result["access_token"]

def _acquire_token_silent() -> str | None:
    cache = _load_cache()
    app = _msal_app(cache)

    accounts = app.get_accounts()
    if not accounts:
        return None

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        _save_cache(cache)
        return result["access_token"]

    return None

# ==========================================
# MAIN UI
# ==========================================
# If token exists silently, user is "connected"
token = _acquire_token_silent()
connected = token is not None

c1, c2 = st.columns([1, 2])
with c1:
    if connected:
        st.success("Connected to Microsoft")
    else:
        st.warning("Not connected")

with c2:
    st.caption("Login happens here. Pages will reuse the cached login silently.")

if not connected:
    if st.button("Connect to Microsoft"):
        token = _acquire_token_interactive()
        st.success("Connected. Open a page from the sidebar ⬅")
        st.rerun()
else:
    st.info("Open a page from the sidebar ⬅")

# Store token in session for convenience (pages still use disk cache)
if token:
    st.session_state["graph_token"] = token

st.divider()
st.markdown(
    """
    **Refresh policy (handled in pages):**
    - Tickets: every 30 minutes
    - Banks Periodics: every 3 hours
    """
)
