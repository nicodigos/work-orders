import streamlit as st
from uuid import uuid4

from utils.ms_graph_excel import finish_redirect_flow, get_redirect_login_url, get_token_silent

# ==========================================
# APP CONFIG
# ==========================================
st.set_page_config(page_title="CNET Reports", layout="wide")
st.title("CNET Reports")

# ==========================================
# MAIN UI
# ==========================================
qp = st.query_params
auth_code = qp.get("code")
auth_state = qp.get("state")

if auth_code:
    expected_state = st.session_state.get("oauth_state")
    try:
        if expected_state and auth_state != expected_state:
            raise RuntimeError("Invalid OAuth state. Please try connecting again.")
        token = finish_redirect_flow(str(auth_code))
        st.session_state["graph_token"] = token
        st.session_state.pop("oauth_state", None)
        st.query_params.clear()
        st.success("Connected to Microsoft")
        st.rerun()
    except Exception as e:
        st.query_params.clear()
        st.error(f"Login failed: {e}")

token = get_token_silent()
connected = token is not None

c1, c2 = st.columns([1, 2])
with c1:
    if connected:
        st.success("Connected to Microsoft")
    else:
        st.warning("Not connected")

with c2:
    st.caption("Login happens here. Pages reuse this session silently.")

if not connected:
    if "oauth_state" not in st.session_state:
        st.session_state["oauth_state"] = str(uuid4())
    try:
        login_url = get_redirect_login_url(st.session_state["oauth_state"])
        st.link_button("Connect to Microsoft", login_url, type="primary")
    except Exception as e:
        st.error(f"Could not build login URL: {e}")
else:
    st.info("Open a page from the sidebar")

if token:
    st.session_state["graph_token"] = token

st.divider()
st.markdown(
    """
    **Refresh policy (handled in pages):**
    - Tickets: every 30 minutes
    """
)
