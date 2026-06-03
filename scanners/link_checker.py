"""
link_checker.py
Extracts all URLs from parsed document lines and checks if they are broken.

Supports optional authorization for protected links via:
  - ******
  - Basic auth (username/password)
  - Custom headers (e.g., API keys)

Auth can be applied globally or scoped to specific hosts so that credentials
are NOT leaked to third-party domains.
"""

import re
import requests
from urllib.parse import urlparse
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from concurrent.futures import ThreadPoolExecutor, as_completed


# Regex pattern to find URLs
URL_PATTERN = re.compile(
    r'https?://[^\s<>"\'\)\],;|\\}]+'
)

DEFAULT_USER_AGENT = "Mozilla/5.0 (DocScanner/1.0)"
DEFAULT_TIMEOUT = 10


def extract_links(parsed_lines: list) -> list:
    """Extract all URLs with their locations."""
    links = []
    for entry in parsed_lines:
        found_urls = URL_PATTERN.findall(entry["text"])
        for url in found_urls:
            url = url.rstrip(".,;:!?)")
            links.append({
                "page": entry["page"],
                "line": entry["line"],
                "url": url,
            })
    return links


def _build_session() -> requests.Session:
    """Create a requests session with sensible retries."""
    session = requests.Session()
    retry = Retry(
        total=2,
        backoff_factor=0.5,
        status_forcelist=(500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "HEAD"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def _resolve_auth_for_url(url: str, auth_config: dict):
    """
    Decide which auth/headers to apply to the given URL.

    auth_config schema (all optional):
        {
            "bearer_token": "xxx",                  # applied as Authorization: Bearer
            "basic_auth": ("user", "pass"),         # HTTPBasicAuth tuple
            "headers": {"X-API-Key": "..."},        # extra headers
            "allowed_hosts": ["api.example.com"],   # if set, auth ONLY sent to these hosts
        }

    Returns: (auth_tuple_or_None, headers_dict)
    """
    headers = {"User-Agent": DEFAULT_USER_AGENT}
    auth = None

    if not auth_config:
        return auth, headers

    host = (urlparse(url).hostname or "").lower()
    allowed_hosts = [h.lower() for h in auth_config.get("allowed_hosts") or []]

    if allowed_hosts and host not in allowed_hosts:
        return auth, headers

    if auth_config.get("bearer_token"):
        headers["Authorization"] = f"Bearer {auth_config['bearer_token']}"

    if auth_config.get("basic_auth"):
        auth = tuple(auth_config["basic_auth"])

    if auth_config.get("headers"):
        headers.update(auth_config["headers"])

    return auth, headers


def check_single_link(link_info: dict, auth_config: dict = None, session: requests.Session = None) -> dict:
    """Check a single URL (with optional auth) and return status."""
    url = link_info["url"]
    close_session = False
    if session is None:
        session = _build_session()
        close_session = True

    auth, headers = _resolve_auth_for_url(url, auth_config or {})

    try:
        response = session.head(
            url,
            timeout=DEFAULT_TIMEOUT,
            allow_redirects=True,
            headers=headers,
            auth=auth,
        )
        if response.status_code >= 400:
            response = session.get(
                url,
                timeout=DEFAULT_TIMEOUT,
                allow_redirects=True,
                stream=True,
                headers=headers,
                auth=auth,
            )
            response.close()

        status_code = response.status_code
        is_broken = status_code >= 400
    except requests.exceptions.Timeout:
        status_code = "Timeout"
        is_broken = True
    except requests.exceptions.ConnectionError:
        status_code = "Connection Error"
        is_broken = True
    except requests.exceptions.RequestException as e:
        status_code = f"Error: {str(e)[:50]}"
        is_broken = True
    finally:
        if close_session:
            session.close()

    authorized = auth is not None or "Authorization" in headers or any(
        k.lower() not in ("user-agent",) and k.lower() != "authorization"
        for k in headers
        if k.lower() != "user-agent"
    )

    return {
        "page": link_info["page"],
        "line": link_info["line"],
        "url": url,
        "status_code": status_code,
        "is_broken": is_broken,
        "authorized": authorized,
    }


def check_all_links(parsed_lines: list, progress_callback=None, auth_config: dict = None) -> dict:
    """
    Extract and check all links in the document.

    Args:
        parsed_lines: list of {"page", "line", "text"} dicts.
        progress_callback: optional callable(completed, total).
        auth_config: optional dict, see _resolve_auth_for_url for schema.

    Returns:
        {
            "all_links": [...],
            "broken_links": [...],
            "total_links_checked": int,
            "broken_link_count": int
        }
    """
    links = extract_links(parsed_lines)

    if not links:
        return {
            "all_links": [],
            "broken_links": [],
            "total_links_checked": 0,
            "broken_link_count": 0,
        }

    results = []
    session = _build_session()

    try:
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {
                executor.submit(check_single_link, link, auth_config, session): link
                for link in links
            }
            completed = 0
            for future in as_completed(futures):
                result = future.result()
                results.append(result)
                completed += 1
                if progress_callback:
                    progress_callback(completed, len(links))
    finally:
        session.close()

    broken = [r for r in results if r["is_broken"]]

    return {
        "all_links": results,
        "broken_links": broken,
        "total_links_checked": len(results),
        "broken_link_count": len(broken),
    }