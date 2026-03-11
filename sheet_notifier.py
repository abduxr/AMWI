from __future__ import annotations

import argparse
import base64
import hashlib
import json
import sys
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from typing import Any

import pandas as pd
from dateutil import tz

from google.oauth2 import service_account
from googleapiclient.discovery import build
import msal
import requests
from playwright.sync_api import sync_playwright

try:
    import tomllib  # py3.11+
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore[no-redef]

import smtplib


IST = tz.gettz("Asia/Kolkata")
SHEETS_READONLY_SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


@dataclass(frozen=True)
class Config:
    source_type: str

    excel_path: Path | None
    excel_worksheet: str | None

    gs_spreadsheet_id: str | None
    gs_worksheet: str | None
    gs_range_a1: str | None
    gs_service_account_json: Path | None

    sp_share_url: str | None
    sp_worksheet: str | None
    sp_tenant_id: str | None
    sp_client_id: str | None
    sp_client_secret: str | None

    spb_share_url: str | None
    spb_storage_state_path: Path | None
    spb_worksheet: str | None

    assignee_name: str
    col_assignee: str
    col_sa: str
    col_to: str
    col_start: str
    col_end: str
    col_customer: str
    col_status: str
    col_release: str

    from_addr: str
    to_addr: str
    subject_prefix: str

    smtp_host: str
    smtp_port: int
    smtp_username: str
    smtp_password: str
    smtp_use_tls: bool


def _require(d: dict[str, Any], key: str) -> Any:
    if key not in d:
        raise KeyError(key)
    return d[key]


def load_config(path: Path) -> Config:
    raw = tomllib.loads(path.read_text(encoding="utf-8"))

    source = raw.get("source") or {"type": "excel"}
    source_type = str(source.get("type") or "excel").strip().lower()

    excel = raw.get("excel") or raw.get("sheet") or {}
    gs = raw.get("google_sheets") or {}
    sp = raw.get("sharepoint") or {}
    spb = raw.get("sharepoint_browser") or {}

    filt = _require(raw, "filter")
    email = _require(raw, "email")
    smtp = _require(raw, "smtp")

    excel_worksheet = str(excel.get("worksheet") or "").strip() or None
    gs_worksheet = str(gs.get("worksheet") or "").strip() or None
    gs_range_a1 = str(gs.get("range_a1") or "").strip() or None
    sp_worksheet = str(sp.get("worksheet") or "").strip() or None
    spb_worksheet = str(spb.get("worksheet") or "").strip() or None

    return Config(
        source_type=source_type,
        excel_path=Path(str(excel.get("path") or "")).expanduser() if str(excel.get("path") or "").strip() else None,
        excel_worksheet=excel_worksheet,
        gs_spreadsheet_id=str(gs.get("spreadsheet_id") or "").strip() or None,
        gs_worksheet=gs_worksheet,
        gs_range_a1=gs_range_a1,
        gs_service_account_json=Path(str(gs.get("service_account_json") or "")).expanduser()
        if str(gs.get("service_account_json") or "").strip()
        else None,
        sp_share_url=str(sp.get("share_url") or "").strip() or None,
        sp_worksheet=sp_worksheet,
        sp_tenant_id=str(sp.get("tenant_id") or "").strip() or None,
        sp_client_id=str(sp.get("client_id") or "").strip() or None,
        sp_client_secret=str(sp.get("client_secret") or "").strip() or None,
        spb_share_url=str(spb.get("share_url") or "").strip() or None,
        spb_storage_state_path=Path(str(spb.get("storage_state_path") or "")).expanduser()
        if str(spb.get("storage_state_path") or "").strip()
        else None,
        spb_worksheet=spb_worksheet,
        assignee_name=str(_require(filt, "assignee_name")),
        col_assignee=str(_require(filt, "col_assignee")),
        col_sa=str(_require(filt, "col_sa")),
        col_to=str(_require(filt, "col_to")),
        col_start=str(_require(filt, "col_start")),
        col_end=str(_require(filt, "col_end")),
        col_customer=str(_require(filt, "col_customer")),
        col_status=str(_require(filt, "col_status")),
        col_release=str(_require(filt, "col_release")),
        from_addr=str(_require(email, "from_addr")),
        to_addr=str(_require(email, "to_addr")),
        subject_prefix=str(email.get("subject_prefix") or "[MW Digest]"),
        smtp_host=str(_require(smtp, "host")),
        smtp_port=int(smtp.get("port") or 587),
        smtp_username=str(_require(smtp, "username")),
        smtp_password=str(_require(smtp, "password")),
        smtp_use_tls=bool(smtp.get("use_tls", True)),
    )


def _read_excel(cfg: Config) -> pd.DataFrame:
    if not cfg.excel_path:
        raise KeyError("excel.path is required when source.type = 'excel'")
    if not cfg.excel_path.exists():
        raise FileNotFoundError(str(cfg.excel_path))

    df = pd.read_excel(
        cfg.excel_path,
        sheet_name=cfg.excel_worksheet if cfg.excel_worksheet else 0,
        engine="openpyxl",
    )
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _read_google_sheets(cfg: Config) -> pd.DataFrame:
    if not cfg.gs_spreadsheet_id:
        raise KeyError("google_sheets.spreadsheet_id is required when source.type = 'google_sheets'")
    if not cfg.gs_service_account_json:
        raise KeyError("google_sheets.service_account_json is required when source.type = 'google_sheets'")
    if not cfg.gs_service_account_json.exists():
        raise FileNotFoundError(str(cfg.gs_service_account_json))

    creds = service_account.Credentials.from_service_account_file(
        str(cfg.gs_service_account_json),
        scopes=[SHEETS_READONLY_SCOPE],
    )
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    if cfg.gs_range_a1:
        a1 = cfg.gs_range_a1
    elif cfg.gs_worksheet:
        a1 = f"{cfg.gs_worksheet}!A:ZZ"
    else:
        a1 = "A:ZZ"

    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=cfg.gs_spreadsheet_id, range=a1, valueRenderOption="UNFORMATTED_VALUE")
        .execute()
    )
    values = result.get("values") or []
    if not values:
        return pd.DataFrame()

    headers = [str(h).strip() for h in values[0]]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _share_url_to_share_id(share_url: str) -> str:
    raw = share_url.encode("utf-8")
    b64 = base64.b64encode(raw).decode("ascii")
    b64 = b64.rstrip("=").replace("/", "_").replace("+", "-")
    return f"u!{b64}"


def _graph_token(cfg: Config) -> str:
    if not cfg.sp_tenant_id:
        raise KeyError("sharepoint.tenant_id is required when source.type = 'sharepoint'")
    if not cfg.sp_client_id:
        raise KeyError("sharepoint.client_id is required when source.type = 'sharepoint'")
    if not cfg.sp_client_secret:
        raise KeyError("sharepoint.client_secret is required when source.type = 'sharepoint'")

    app = msal.ConfidentialClientApplication(
        client_id=cfg.sp_client_id,
        client_credential=cfg.sp_client_secret,
        authority=f"https://login.microsoftonline.com/{cfg.sp_tenant_id}",
    )
    result = app.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise RuntimeError(f"Failed to get Graph token: {result.get('error')} {result.get('error_description')}")
    return str(result["access_token"])


def _read_sharepoint_excel(cfg: Config) -> pd.DataFrame:
    if not cfg.sp_share_url:
        raise KeyError("sharepoint.share_url is required when source.type = 'sharepoint'")

    token = _graph_token(cfg)
    share_id = _share_url_to_share_id(cfg.sp_share_url)

    # Download the workbook content from the share link
    url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if r.status_code >= 400:
        raise RuntimeError(f"Graph download failed ({r.status_code}): {r.text[:500]}")

    bio = BytesIO(r.content)
    df = pd.read_excel(bio, sheet_name=cfg.sp_worksheet if cfg.sp_worksheet else 0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def init_sharepoint_browser_auth(config_path: Path) -> int:
    """
    One-time interactive login to SharePoint to save cookies for headless runs.
    """
    cfg = load_config(config_path)
    if not cfg.spb_share_url:
        raise KeyError("sharepoint_browser.share_url is required")
    if not cfg.spb_storage_state_path:
        raise KeyError("sharepoint_browser.storage_state_path is required")

    cfg.spb_storage_state_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.goto(cfg.spb_share_url, wait_until="domcontentloaded")

        # User completes SSO/MFA manually, then closes the browser window.
        page.wait_for_timeout(120_000)

        context.storage_state(path=str(cfg.spb_storage_state_path))
        context.close()
        browser.close()

    return 0

def _read_sharepoint_browser_excel(cfg: Config) -> pd.DataFrame:
    if not cfg.spb_share_url:
        raise KeyError("sharepoint_browser.share_url is required when source.type = 'sharepoint_browser'")
    if not cfg.spb_storage_state_path:
        raise KeyError("sharepoint_browser.storage_state_path is required when source.type = 'sharepoint_browser'")
    if not cfg.spb_storage_state_path.exists():
        raise FileNotFoundError(
            f"{cfg.spb_storage_state_path} not found. Run: python sheet_notifier.py --init-sharepoint-auth --config config.toml"
        )

    # Force SharePoint to download the file directly
    download_url = cfg.spb_share_url
    if "?" in download_url:
        download_url += "&download=1"
    else:
        download_url += "?download=1"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=str(cfg.spb_storage_state_path))

        # Retry download a few times (SharePoint sometimes fails)
        response = None
        for _ in range(3):
            response = context.request.get(download_url)
            if response.status == 200:
                break

        if response is None or response.status != 200:
            raise RuntimeError(f"SharePoint download failed. Status: {response.status if response else 'no response'}")

        content = response.body()

        context.close()
        browser.close()

    bio = BytesIO(content)
    df = pd.read_excel(
        bio,
        sheet_name=cfg.spb_worksheet if cfg.spb_worksheet else 0,
        engine="openpyxl",
    )

    df.columns = [str(c).strip() for c in df.columns]
    return df

def _read_sheet(cfg: Config) -> pd.DataFrame:
    if cfg.source_type == "excel":
        return _read_excel(cfg)
    if cfg.source_type in {"google_sheets", "google", "gsheets"}:
        return _read_google_sheets(cfg)
    if cfg.source_type in {"sharepoint", "sp"}:
        return _read_sharepoint_excel(cfg)
    if cfg.source_type in {"sharepoint_browser", "sp_browser", "spb"}:
        return _read_sharepoint_browser_excel(cfg)
    raise ValueError("source.type must be 'excel', 'google_sheets', 'sharepoint', or 'sharepoint_browser'")


def _parse_dt(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, datetime):
        dt = value
    else:
        s = str(value).strip().strip('"')
        if not s:
            return ""
        dt = pd.to_datetime(s, errors="coerce")
        if pd.isna(dt):
            return s
        dt = dt.to_pydatetime()

    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    return dt.astimezone(IST).strftime("%Y-%m-%d %H:%M IST")


def _norm(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip().strip('"')


def _row_key(cfg: Config, row: dict[str, Any]) -> str:
    sa = _norm(row.get(cfg.col_sa))
    to = _norm(row.get(cfg.col_to))
    if sa and to:
        return f"{sa}|{to}"
    if to:
        return to
    if sa:
        return sa
    stable = json.dumps(row, sort_keys=True, default=str, ensure_ascii=False)
    return hashlib.sha256(stable.encode("utf-8")).hexdigest()[:16]


def _row_fingerprint(cfg: Config, row: dict[str, Any]) -> str:
    parts = {
        "assignee": _norm(row.get(cfg.col_assignee)),
        "sa": _norm(row.get(cfg.col_sa)),
        "to": _norm(row.get(cfg.col_to)),
        "release": _norm(row.get(cfg.col_release)),
        "start": _parse_dt(row.get(cfg.col_start)),
        "end": _parse_dt(row.get(cfg.col_end)),
        "customer": _norm(row.get(cfg.col_customer)),
        "status": _norm(row.get(cfg.col_status)),
    }
    s = json.dumps(parts, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def _select_rows(cfg: Config, df: pd.DataFrame) -> list[dict[str, Any]]:
    required_cols = [
        cfg.col_assignee,
        cfg.col_sa,
        cfg.col_to,
        cfg.col_release,
        cfg.col_start,
        cfg.col_end,
        cfg.col_customer,
        cfg.col_status,
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing columns in sheet: {missing}. Found: {list(df.columns)}")

    sub = df[df[cfg.col_assignee].astype(str).str.strip() == cfg.assignee_name.strip()]
    rows = sub.to_dict(orient="records")
    return rows


def _load_state(state_path: Path) -> dict[str, str]:
    if not state_path.exists():
        return {}
    try:
        return json.loads(state_path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_state(state_path: Path, state: dict[str, str]) -> None:
    state_path.write_text(json.dumps(state, indent=2, sort_keys=True), encoding="utf-8")


def _format_table(cfg: Config, rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "<p>(none)</p>"

    html = """
    <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial; font-size: 13px;">
        <tr style="background-color:#f2f2f2;">
            <th>SA</th>
            <th>TO</th>
            <th>Release</th>
            <th>Start</th>
            <th>End</th>
            <th>Customer</th>
            <th>Status</th>
        </tr>
    """

    for r in rows:
        html += f"""
        <tr>
            <td>{_norm(r.get(cfg.col_sa))}</td>
            <td>{_norm(r.get(cfg.col_to))}</td>
            <td>{_norm(r.get(cfg.col_release))}</td>
            <td>{_parse_dt(r.get(cfg.col_start))}</td>
            <td>{_parse_dt(r.get(cfg.col_end))}</td>
            <td>{_norm(r.get(cfg.col_customer))}</td>
            <td>{_norm(r.get(cfg.col_status))}</td>
        </tr>
        """

    html += "</table>"
    return html


def _send_email(cfg: Config, subject: str, body: str) -> None:
    msg = EmailMessage()
    msg["From"] = cfg.from_addr
    msg["To"] = cfg.to_addr
    msg["Subject"] = subject

    msg.set_content("This email requires an HTML capable viewer.")

    msg.add_alternative(body, subtype="html")

    with smtplib.SMTP(cfg.smtp_host, cfg.smtp_port, timeout=30) as s:
        if cfg.smtp_use_tls:
            s.starttls()
        if cfg.smtp_username:
            s.login(cfg.smtp_username, cfg.smtp_password)
        s.send_message(msg)


def run_once(config_path: Path, state_path: Path, send_if_no_changes: bool) -> int:
    cfg = load_config(config_path)
    df = _read_sheet(cfg)
    rows = _select_rows(cfg, df)

    prev = _load_state(state_path)
    current: dict[str, str] = {}

    new_rows: list[dict[str, Any]] = []
    changed_rows: list[dict[str, Any]] = []

    for r in rows:
        key = _row_key(cfg, r)
        fp = _row_fingerprint(cfg, r)
        current[key] = fp
        if key not in prev:
            new_rows.append(r)
        elif prev[key] != fp:
            changed_rows.append(r)

    removed = sorted(set(prev.keys()) - set(current.keys()))

    now_ist = datetime.now(tz=IST).strftime("%Y-%m-%d %H:%M IST")
    subject = f"{cfg.subject_prefix} {cfg.assignee_name} - {len(new_rows)} new, {len(changed_rows)} updated"

    body_parts = [
        f"MW digest for: {cfg.assignee_name}",
        f"Generated at: {now_ist}",
        "",
        f"New MWs: {len(new_rows)}",
        _format_table(cfg, new_rows) if new_rows else "(none)",
        "",
        f"Updated MWs: {len(changed_rows)}",
        _format_table(cfg, changed_rows) if changed_rows else "(none)",
    ]
    if removed:
        body_parts += ["", f"Removed since last run: {len(removed)}", "\n".join(removed)]

    body = "\n".join(body_parts) + "\n"

    if new_rows or changed_rows or removed or send_if_no_changes:
        _send_email(cfg, subject, body)

    _save_state(state_path, current)
    return 0


def main() -> int:
    p = argparse.ArgumentParser(description="Read Excel MW sheet and email daily digest.")
    p.add_argument("--config", default="config.toml", help="Path to TOML config (default: config.toml)")
    p.add_argument("--state", default=".state.json", help="Path to state file (default: .state.json)")
    p.add_argument("--send-if-no-changes", action="store_true", help="Send even when nothing changed")
    p.add_argument(
        "--init-sharepoint-auth",
        action="store_true",
        help="One-time interactive SharePoint login (for source.type=sharepoint_browser)",
    )
    args = p.parse_args()

    try:
        if args.init_sharepoint_auth:
            return init_sharepoint_browser_auth(Path(args.config))
        return run_once(Path(args.config), Path(args.state), bool(args.send_if_no_changes))
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
