import os
import base64
import requests
import pandas as pd
from dotenv import load_dotenv
from utils import is_url

load_dotenv()

TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

def _ensure_env_vars():
    missing = [k for k, v in {
        "AZURE_TENANT_ID": TENANT_ID,
        "AZURE_CLIENT_ID": CLIENT_ID,
        "AZURE_CLIENT_SECRET": CLIENT_SECRET
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")

def get_access_token(tenant_id=TENANT_ID, client_id=CLIENT_ID, client_secret=CLIENT_SECRET):
    _ensure_env_vars()
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()['access_token']


def _encode_share_url(url):
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii")
    return "u!" + encoded.rstrip("=")


def _worksheet_name(name):
    return str(name).replace("'", "''")


def _request_json(method, url, token, **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    r = requests.request(method, url, headers=headers, **kwargs)
    r.raise_for_status()
    return r.json() if r.content else {}


def get_drive_item_from_share_url(share_url, token):
    """
    Resolve a SharePoint/OneDrive sharing URL into a Graph driveItem.
    """
    share_id = _encode_share_url(share_url)
    graph_url = f"{GRAPH_ROOT}/shares/{share_id}/driveItem"
    return _request_json("GET", graph_url, token)


def get_workbook_item_url(workbook_url, token):
    """
    Return the Graph URL for an Excel workbook item from a SharePoint/OneDrive sharing URL.
    """
    if not is_url(workbook_url):
        raise RuntimeError("SHAREPOINT_PATH must be a full SharePoint/OneDrive URL.")

    item = get_drive_item_from_share_url(workbook_url, token)
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    return f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}"


def get_used_range_for_workbook(workbook_url, sheet_name, token):
    """
    workbook_url: SharePoint/OneDrive sharing URL.
    sheet_name: sheet name, e.g. "Sheet1"
    """
    workbook_item_url = get_workbook_item_url(workbook_url, token)
    graph_url = (
        f"{workbook_item_url}/workbook/worksheets('{_worksheet_name(sheet_name)}')"
        "/usedRange()"
    )
    return _request_json("GET", graph_url, token)

def range_json_to_dataframe(range_json):
    """
    Convert Graph worksheet usedRange JSON to pandas.DataFrame.
    Uses the first row as header if it looks like header strings.
    """
    values = range_json.get('values', [])
    if not values:
        return pd.DataFrame()

    header = values[0]
    data_rows = values[1:] if len(values) > 1 else []

    # Heuristic: if any header cell is a non-empty string, treat row 0 as header
    use_header = any(isinstance(h, str) and h.strip() for h in header)
    if use_header:
        df = pd.DataFrame(data_rows, columns=[str(h).strip() for h in header])
    else:
        df = pd.DataFrame(values)

    # normalize columns: strip whitespace and make consistent
    df.columns = df.columns.astype(str).str.strip()
    return df

def load_sheet_as_dataframe(workbook_url, sheet_name='Sheet1'):
    """
    Convenience loader. Returns a pandas.DataFrame for the sheet's usedRange.
    workbook_url must be a SharePoint/OneDrive sharing URL.
    """
    token = get_access_token()
    range_json = get_used_range_for_workbook(workbook_url, sheet_name, token)
    return range_json_to_dataframe(range_json)


def write_dataframe_to_workbook(df, workbook_url, sheet_name='Sheet1'):
    """
    Write a pandas.DataFrame back to a SharePoint workbook's sheet.
    Converts df to a 2D list (values array) and updates the current used range.
    """
    token = get_access_token()

    # Prepare header + data rows
    values = [df.columns.tolist()] + df.values.tolist()

    workbook_item_url = get_workbook_item_url(workbook_url, token)
    worksheet = _worksheet_name(sheet_name)
    used_range = _request_json(
        "GET",
        f"{workbook_item_url}/workbook/worksheets('{worksheet}')/usedRange()",
        token
    )
    address = used_range.get("address")
    if not address:
        raise RuntimeError("Could not determine the used range address for the workbook.")

    # usedRange is a workbook function. Update the concrete range address it returned.
    graph_url = (
        f"{workbook_item_url}/workbook/worksheets('{worksheet}')"
        f"/range(address='{_worksheet_name(address)}')"
    )
    headers = {'Content-Type': 'application/json'}

    payload = {'values': values}
    return _request_json("PATCH", graph_url, token, json=payload, headers=headers)
