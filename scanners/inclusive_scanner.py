"""
inclusive_scanner.py
Scans parsed document lines for non-inclusive words and returns results.
"""

import re


# Define non-inclusive terms and their replacements
NON_INCLUSIVE_TERMS = {
    "blacklist":  "denylist",
    "whitelist":  "allowlist",
    "master":     "primary / initiator",
    "slave":      "secondary / target",
}


def scan_for_inclusive_language(parsed_lines: list) -> dict:
    """
    Scans each line for non-inclusive words.

    Returns:
        {
            "findings": [...],
            "total_non_inclusive_count": int,
            "total_word_count": int
        }
    """
    findings = []
    total_word_count = 0

    for entry in parsed_lines:
        page = entry["page"]
        line = entry["line"]
        text = entry["text"]

        # Count words
        words_in_line = text.split()
        total_word_count += len(words_in_line)

        # Search for each non-inclusive term (case-insensitive, whole word)
        for term, replacement in NON_INCLUSIVE_TERMS.items():
            pattern = re.compile(r'\b' + re.escape(term) + r'\b', re.IGNORECASE)
            matches = pattern.findall(text)
            for match in matches:
                context = text.strip()
                if len(context) > 120:
                    idx = text.lower().find(term.lower())
                    start = max(0, idx - 40)
                    end = min(len(text), idx + len(term) + 40)
                    context = "..." + text[start:end] + "..."

                findings.append({
                    "page": page,
                    "line": line,
                    "found_word": match,
                    "suggested_replacement": replacement,
                    "context": context,
                })

    return {
        "findings": findings,
        "total_non_inclusive_count": len(findings),
        "total_word_count": total_word_count,
    }