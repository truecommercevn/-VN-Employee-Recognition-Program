"""
Sync Recognition data from SharePoint Excel files → data.json on GitHub.

Reads:
  - Top Talent Of the Quarter.xlsx  (quarterly winners across all sheets)
  - Top Mover.xlsx                  (monthly winners across all sheets)

From SharePoint site:
  tcapc.sharepoint.com/sites/VNPeopleteam-Publicfolder
  └── Shared Documents/
      └── VN Employee Recognition Page/

Authenticates with Azure AD app (client credentials flow, Sites.Selected).
Outputs data.json next to this script for the page to fetch.

Env vars required:
  AZURE_TENANT_ID
  AZURE_CLIENT_ID
  AZURE_CLIENT_SECRET
"""

import os
import json
import sys
import re
import msal
import requests
from pathlib import Path

# ─── CONFIG ────────────────────────────────────────────────────────────────
SP_HOSTNAME = "tcapc.sharepoint.com"
SP_SITE_PATH = "/sites/VNPeopleteam-Publicfolder"
FOLDER_PATH = "VN Employee Recognition Page"  # under "Shared Documents" (default drive)

FILE_TOP_TALENT = "Top Talent Of the Quarter.xlsx"
FILE_TOP_MOVER  = "Top Mover.xlsx"

# Avatar colors rotate through these (av-1 … av-8)
AVATAR_COLORS = [f"av-{i}" for i in range(1, 9)]

# Map value name → CSS class for Top Mover tags
VALUE_TAG_MAP = {
    "Customer Drives Our Purpose": "tag-cus",
    "Ownership + Impact":          "tag-own",
    "Win Together":                "tag-col",
    "Adapt + Empower":             "tag-adp",
    "Relentless Innovation":       "tag-inn",
}

GRAPH = "https://graph.microsoft.com/v1.0"

# ─── AUTH ──────────────────────────────────────────────────────────────────
def get_token():
    tenant = os.environ["AZURE_TENANT_ID"]
    app = msal.ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{tenant}",
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error')} - {result.get('error_description')}")
    return result["access_token"]


def gget(url, token):
    """Graph API GET with auth header."""
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    if r.status_code >= 400:
        raise RuntimeError(f"GET {url} → {r.status_code}\n{r.text[:500]}")
    return r.json()


# ─── HELPERS ───────────────────────────────────────────────────────────────
def initials(name: str) -> str:
    """Derive 2-letter initials. 'Cuong Truong' → 'CT', 'Lam H Nguyen' → 'LN'."""
    parts = [p for p in name.strip().split() if p]
    if not parts: return "??"
    if len(parts) == 1: return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def clean_dept(dept) -> str:
    """Strip leading '###-' code from department name."""
    if not dept: return ""
    return re.sub(r"^\d+-", "", str(dept)).strip()


def color_for(idx: int) -> str:
    return AVATAR_COLORS[idx % len(AVATAR_COLORS)]


# ─── GRAPH API: SITE & FILES ───────────────────────────────────────────────
def get_site_id(token):
    """Resolve site ID from hostname:/sites/path."""
    url = f"{GRAPH}/sites/{SP_HOSTNAME}:{SP_SITE_PATH}"
    return gget(url, token)["id"]


def get_file_item_id(token, site_id, file_name):
    """Get drive item ID by path under the default drive root."""
    path = f"{FOLDER_PATH}/{file_name}"
    url = f"{GRAPH}/sites/{site_id}/drive/root:/{requests.utils.quote(path)}"
    return gget(url, token)["id"]


def list_worksheets(token, site_id, item_id):
    """Return list of {name, id} for all worksheets in a workbook."""
    url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/workbook/worksheets"
    data = gget(url, token)
    return [{"name": ws["name"], "id": ws["id"]} for ws in data.get("value", [])]


def read_sheet_values(token, site_id, item_id, sheet_name):
    """Return used-range values as 2D list (rows of cells)."""
    enc = requests.utils.quote(sheet_name)
    url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/workbook/worksheets('{enc}')/usedRange?$select=values"
    return gget(url, token).get("values", [])


# ─── DATA TRANSFORM ────────────────────────────────────────────────────────
def parse_top_talent(token, site_id, item_id) -> dict:
    """{ 'Q1 2026': [winner, ...], 'Q2 2025': [...], ... }"""
    result = {}
    for ws in list_worksheets(token, site_id, item_id):
        sheet = ws["name"]
        rows = read_sheet_values(token, site_id, item_id, sheet)
        if not rows: continue
        header = [str(h or "").strip().lower() for h in rows[0]]
        # Find column indices (UKG ID | Email | Name | Job Title | Department)
        def col(*names):
            for n in names:
                if n in header: return header.index(n)
            return -1
        i_email = col("email")
        i_name  = col("name")
        i_role  = col("job title", "title", "role")
        i_dept  = col("department", "dept")
        winners = []
        for idx, row in enumerate(rows[1:]):
            name = (row[i_name] if i_name >= 0 and i_name < len(row) else "") or ""
            if not str(name).strip(): continue
            email = (row[i_email] if i_email >= 0 and i_email < len(row) else "") or ""
            role  = (row[i_role]  if i_role  >= 0 and i_role  < len(row) else "") or ""
            dept  = (row[i_dept]  if i_dept  >= 0 and i_dept  < len(row) else "") or ""
            winners.append({
                "name":     str(name).strip(),
                "email":    str(email).strip(),
                "role":     str(role).strip(),
                "dept":     clean_dept(dept),
                "initials": initials(str(name)),
                "color":    color_for(idx),
            })
        result[sheet] = winners
    return result


def parse_top_mover(token, site_id, item_id) -> dict:
    """{ 'May 2026': [winner with value+desc, ...], ... }"""
    result = {}
    for ws in list_worksheets(token, site_id, item_id):
        sheet = ws["name"]
        rows = read_sheet_values(token, site_id, item_id, sheet)
        if not rows: continue
        header = [str(h or "").strip().lower() for h in rows[0]]
        def col(*names):
            for n in names:
                if n in header: return header.index(n)
            return -1
        i_email = col("email")
        i_name  = col("name")
        i_role  = col("job title", "title", "role")
        i_dept  = col("department", "dept")
        i_value = col("value", "core value")
        i_just  = col("justification", "description", "desc")
        winners = []
        for idx, row in enumerate(rows[1:]):
            name = (row[i_name] if i_name >= 0 and i_name < len(row) else "") or ""
            if not str(name).strip(): continue
            email = (row[i_email] if i_email >= 0 and i_email < len(row) else "") or ""
            role  = (row[i_role]  if i_role  >= 0 and i_role  < len(row) else "") or ""
            dept  = (row[i_dept]  if i_dept  >= 0 and i_dept  < len(row) else "") or ""
            value = str(row[i_value]).strip() if i_value >= 0 and i_value < len(row) and row[i_value] else ""
            just  = str(row[i_just]).strip()  if i_just  >= 0 and i_just  < len(row) and row[i_just]  else ""
            winners.append({
                "name":       str(name).strip(),
                "email":      str(email).strip(),
                "role":       str(role).strip(),
                "dept":       clean_dept(dept),
                "initials":   initials(str(name)),
                "color":      color_for(idx),
                "value":      value,
                "valueClass": VALUE_TAG_MAP.get(value, "tag-own"),  # fallback
                "desc":       just,
            })
        result[sheet] = winners
    return result


# ─── MAIN ──────────────────────────────────────────────────────────────────
def main():
    print("→ Authenticating with Azure AD…")
    token = get_token()
    print("  ✅ Got access token")

    print(f"→ Resolving site: {SP_HOSTNAME}{SP_SITE_PATH}")
    site_id = get_site_id(token)
    print(f"  ✅ Site ID: {site_id[:40]}…")

    print(f"→ Reading {FILE_TOP_TALENT}…")
    talent_id = get_file_item_id(token, site_id, FILE_TOP_TALENT)
    top_talent_periods = parse_top_talent(token, site_id, talent_id)
    total_t = sum(len(v) for v in top_talent_periods.values())
    print(f"  ✅ {len(top_talent_periods)} sheets, {total_t} winners total")

    print(f"→ Reading {FILE_TOP_MOVER}…")
    try:
        mover_id = get_file_item_id(token, site_id, FILE_TOP_MOVER)
        top_mover_periods = parse_top_mover(token, site_id, mover_id)
        total_m = sum(len(v) for v in top_mover_periods.values())
        print(f"  ✅ {len(top_mover_periods)} sheets, {total_m} winners total")
    except RuntimeError as e:
        print(f"  ⚠️ {FILE_TOP_MOVER} not found or error — skipping: {e}")
        top_mover_periods = {}

    out = {
        "topTalent": {"periods": top_talent_periods},
        "topMover":  {"periods": top_mover_periods},
    }

    out_path = Path(__file__).parent.parent / "data.json"
    out_path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"→ Wrote {out_path} ({out_path.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
