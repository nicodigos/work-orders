import streamlit as st

from utils.ms_graph_excel import finish_device_flow, get_token_silent, start_device_flow

# ==========================================
# APP CONFIG
# ==========================================
st.set_page_config(page_title="CNET Reports", layout="wide")
st.title("CNET Reports")

# ==========================================
# MAIN UI
# ==========================================
token = get_token_silent()
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
        try:
            app, cache, flow = start_device_flow()
            st.info(f"Open {flow['verification_uri']} and enter code: {flow['user_code']}")
            token = finish_device_flow(app, cache, flow)
        except Exception as e:
            st.error(str(e))
            st.stop()

        st.success("Connected. Open a page from the sidebar")
        st.rerun()
else:
    st.info("Open a page from the sidebar")

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
