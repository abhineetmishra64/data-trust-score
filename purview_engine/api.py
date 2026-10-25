"""Authentication and paginated HTTP access to Microsoft Purview."""
import requests

from .settings import API_VERSION, BASE_URL, CLIENT_ID, CLIENT_SECRET, SOURCES, TENANT_ID


class PurviewError(Exception):
    """A problem already explained in plain words for the user."""


HTTP_HINTS = {
    400: "Bad request. Check tenant ID, client ID, secret and scope.",
    401: "Unauthorized. The secret may be wrong or expired.",
    403: "Forbidden. The app needs a Purview role (e.g. Data Curator).",
    404: "Not found. Check the URL and api-version.",
    429: "Too many requests. Wait a bit and retry.",
}


def explain_error(error, step):
    """Turn a requests exception into a PurviewError with a readable message."""
    if isinstance(error, requests.exceptions.Timeout):
        return PurviewError(f"{step}: timed out. Check your internet/VPN and retry.")
    if isinstance(error, requests.exceptions.SSLError):
        return PurviewError(f"{step}: SSL error. A proxy or firewall may be blocking traffic.")
    if isinstance(error, requests.exceptions.ConnectionError):
        return PurviewError(f"{step}: cannot connect. Check your internet/VPN.")
    if isinstance(error, requests.exceptions.HTTPError):
        code = error.response.status_code
        try:
            body = error.response.json()
            detail = body.get("error_description") or body.get("error") or body
        except ValueError:
            detail = error.response.text[:300]
        return PurviewError(f"{step}: HTTP {code}. {HTTP_HINTS.get(code, 'Unexpected response.')}\nDetails: {detail}")
    return PurviewError(f"{step}: {error}")


def request_access_token():
    """Log in as the app (client credentials) and return a bearer token."""
    try:
        response = requests.post(
            f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
            data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
                  "client_secret": CLIENT_SECRET, "scope": "https://purview.azure.net/.default"},
            timeout=30)
        response.raise_for_status()
        return response.json()["access_token"]
    except requests.exceptions.RequestException as error:
        raise explain_error(error, "Getting access token") from error
    except (ValueError, KeyError) as error:
        raise PurviewError("Getting access token: unexpected response format.") from error


def fetch_all_items(source, report=lambda text: None):
    """Fetch all pages of one supported source and report progress."""
    report("Signing in to Microsoft Entra ID...")
    headers = {"Authorization": f"******"}
    return fetch_pages(SOURCES[source]["path"], headers, SOURCES[source]["label"], report)


def fetch_pages(path, headers, label, report=lambda text: None, extra_params=None):
    """GET BASE_URL + path and follow 'nextLink', returning the combined 'value' list."""
    url = BASE_URL + path
    params, items, page_number = {"api-version": API_VERSION, **(extra_params or {})}, [], 0
    try:
        while url:
            page_number += 1
            report(f"Fetching {label}, page {page_number} ({len(items)} so far)...")
            response = requests.get(url, headers=headers, params=params, timeout=60)
            response.raise_for_status()
            page = response.json()
            items += page.get("value", [])
            url, params = page.get("nextLink"), None
    except requests.exceptions.RequestException as error:
        raise explain_error(error, f"Fetching {label}") from error
    except (ValueError, AttributeError) as error:
        raise PurviewError(f"Fetching {label}: response was not valid JSON.") from error
    return items
