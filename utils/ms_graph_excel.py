import os
import tempfile
from pathlib import Path
from urllib.parse import quote

import msal
import requests
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

SCOPES = ["User.Read", "Files.Read.All"]


# -------------------------
# TOKEN CACHE (disk)
# -------------------------
def _cache_path() -> Path:
    d = Path(tempfile.gettempdir()) / "cnet_reports"
    d.mkdir(exist_ok=True)
    return d / "msal_token_cache.bin"


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = _cache_path()
    if p.exists():
        cache.deserialize(p.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        _cache_path().write_text(cache.serialize(), encoding="utf-8")


def _msal_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    if not tenant_id or not client_id:
        raise RuntimeError("Missing TENANT_ID / CLIENT_ID in environment.")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    return msal.PublicClientApplication(
        client_id,
        authority=authority,
        token_cache=cache,
    )


# -------------------------
# AUTH API
# -------------------------
def get_token_silent() -> str | None:
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


def get_token_silent_or_raise(not_authenticated_message: str, expired_message: str) -> str:
    cache = _load_cache()
    app = _msal_app(cache)

    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError(not_authenticated_message)

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if result and "access_token" in result:
        _save_cache(cache)
        return result["access_token"]

    raise RuntimeError(expired_message)


def start_device_flow() -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache, dict]:
    cache = _load_cache()
    app = _msal_app(cache)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(str(flow))
    return app, cache, flow


def finish_device_flow(
    app: msal.PublicClientApplication,
    cache: msal.SerializableTokenCache,
    flow: dict,
) -> str:
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(str(result))

    _save_cache(cache)
    return result["access_token"]


def get_token_silent_or_interactive() -> str:
    token = get_token_silent()
    if token:
        return token

    app, cache, flow = start_device_flow()
    with st.expander("Microsoft sign-in required", expanded=True):
        st.info(f"Open {flow['verification_uri']} and enter code: {flow['user_code']}")
    return finish_device_flow(app, cache, flow)


# -------------------------
# GRAPH HELPERS
# -------------------------
def graph_get(url: str, token: str) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json()


def graph_download(url: str, token: str) -> bytes:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.content


def graph_put_bytes(url: str, token: str, content: bytes) -> dict:
    r = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}"},
        data=content,
        timeout=120,
    )
    if r.status_code >= 400:
        raise RuntimeError(r.text)
    return r.json() if r.content else {}


def resolve_drive_id(token: str) -> str:
    sp_hostname = os.getenv("SP_HOSTNAME")
    sp_site_path = os.getenv("SP_SITE_PATH")
    sp_drive_name = os.getenv("SP_DRIVE_NAME", "Documents")

    if not sp_hostname or not sp_site_path:
        raise RuntimeError("Missing SP_HOSTNAME / SP_SITE_PATH in environment.")

    site = graph_get(f"https://graph.microsoft.com/v1.0/sites/{sp_hostname}:{sp_site_path}", token)
    drives = graph_get(f"https://graph.microsoft.com/v1.0/sites/{site['id']}/drives", token)["value"]
    drive = next((d for d in drives if d.get("name") == sp_drive_name), drives[0])
    return drive["id"]


def list_children_by_path(drive_id: str, sp_path: str, token: str) -> list[dict]:
    sp_path_enc = quote(sp_path.strip("/"), safe="/")
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_path_enc}:/children"
        f"?$top=200&$select=id,name,folder,file,webUrl"
    )

    out: list[dict] = []
    while url:
        data = graph_get(url, token)
        out.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return out


def download_sharepoint_file_bytes(
    sp_relative_path: str,
    token: str,
    drive_id: str | None = None,
) -> bytes:
    did = drive_id or resolve_drive_id(token)
    url = f"https://graph.microsoft.com/v1.0/drives/{did}/root:/{sp_relative_path}:/content"
    return graph_download(url, token)


def download_drive_item_content(drive_id: str, item_id: str, token: str) -> bytes:
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    return graph_download(url, token)


def upload_sharepoint_file_bytes(
    sp_relative_path: str,
    content: bytes,
    token: str,
    drive_id: str | None = None,
) -> dict:
    did = drive_id or resolve_drive_id(token)
    url = f"https://graph.microsoft.com/v1.0/drives/{did}/root:/{sp_relative_path}:/content"
    return graph_put_bytes(url, token, content)


def write_temp_file(filename: str, content: bytes) -> str:
    out_dir = Path(tempfile.gettempdir()) / "cnet_reports"
    out_dir.mkdir(exist_ok=True)
    local = out_dir / filename
    local.write_bytes(content)
    return str(local)


# -------------------------
# PUBLIC API (cached download)
# -------------------------
@st.cache_data(show_spinner=False)
def download_sharepoint_excel_cached(
    sp_relative_path: str,
    ttl_seconds: int,
) -> str:
    # Include ttl_seconds as a cache key component for caller-controlled refresh.
    _ = ttl_seconds
    token = get_token_silent_or_interactive()
    content = download_sharepoint_file_bytes(sp_relative_path, token)
    return write_temp_file(Path(sp_relative_path).name, content)


@st.cache_data(show_spinner=False)
def download_sharepoint_excel_cached_silent(
    sp_relative_path: str,
    ttl_seconds: int,
    not_authenticated_message: str,
    expired_message: str,
) -> str:
    # Include ttl_seconds as a cache key component for caller-controlled refresh.
    _ = ttl_seconds
    token = get_token_silent_or_raise(not_authenticated_message, expired_message)
    content = download_sharepoint_file_bytes(sp_relative_path, token)
    return write_temp_file(Path(sp_relative_path).name, content)


# Backward-compatible alias

def download_excel_cached(sp_relative_path: str, ttl_seconds: int) -> str:
    return download_sharepoint_excel_cached(sp_relative_path, ttl_seconds)
