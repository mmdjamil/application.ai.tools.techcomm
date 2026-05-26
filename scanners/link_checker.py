"""
link_checker.py
Extracts all URLs from parsed document lines and checks if they are broken.
"""

import re
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed


# Regex pattern to find URLs
URL_PATTERN = re.compile(
    r'https?://[^\s<>"\'\)\],;|\\}]+'
)


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


def check_single_link(link_info: dict) -> dict:
    """Check a single URL and return status."""
    url = link_info["url"]
    try:
        response = requests.head(
            url,
            timeout=10,
            allow_redirects=True,
            headers={"User-Agent": "Mozilla/5.0 (DocScanner/1.0)"}
        )
        if response.status_code >= 400:
            response = requests.get(
                url,
                timeout=10,
                allow_redirects=True,
                stream=True,
                headers={"User-Agent": "Mozilla/5.0 (DocScanner/1.0)"}
            )
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

    return {
        "page": link_info["page"],
        "line": link_info["line"],
        "url": url,
        "status_code": status_code,
        "is_broken": is_broken,
    }


def check_all_links(parsed_lines: list, progress_callback=None) -> dict:
    """
    Extract and check all links in the document.

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

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(check_single_link, link): link for link in links}
        completed = 0
        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            completed += 1
            if progress_callback:
                progress_callback(completed, len(links))

    broken = [r for r in results if r["is_broken"]]

    return {
        "all_links": results,
        "broken_links": broken,
        "total_links_checked": len(results),
        "broken_link_count": len(broken),
    }