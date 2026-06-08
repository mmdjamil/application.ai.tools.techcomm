"""
inclusive_scanner.py
Scans parsed document lines for non-inclusive words and returns results.

Each finding now also includes a `suggested_sentence` — the original
sentence with the non-inclusive term replaced — so the UI can let the
user review/approve the rewrite and export a corrected copy of the
document.
"""

import re


# Define non-inclusive terms and their replacements.
# A replacement may contain multiple options separated by " / " — the
# first option is treated as the default word to insert when rewriting
# a sentence. The full string is still shown to the user for context.
NON_INCLUSIVE_TERMS = {
    "blacklist":  "denylist",
    "whitelist":  "allowlist",
    "master":     "primary / initiator",
    "slave":      "secondary / target",
}


def pick_default_replacement(replacement: str) -> str:
    """
    Pick a single-word default replacement from a NON_INCLUSIVE_TERMS
    value such as "primary / initiator" -> "primary".
    """
    for sep in ("/", ","):
        if sep in replacement:
            return replacement.split(sep)[0].strip()
    return replacement.strip()


def _match_case(original: str, replacement: str) -> str:
    """
    Make the replacement match the casing pattern of the original token.
    - "BLACKLIST" -> "DENYLIST"
    - "Blacklist" -> "Denylist"
    - "blacklist" -> "denylist"
    - mixed       -> replacement returned unchanged
    """
    if not original:
        return replacement
    if original.isupper():
        return replacement.upper()
    if original[:1].isupper() and original[1:].islower():
        return replacement[:1].upper() + replacement[1:]
    if original.islower():
        return replacement.lower()
    return replacement


def apply_replacement_to_text(text: str, term: str, replacement: str) -> str:
    """
    Replace every whole-word, case-insensitive occurrence of `term`
    in `text` with `replacement` (the default single-word form),
    preserving the casing pattern of each match.
    """
    default_replacement = pick_default_replacement(replacement)
    pattern = re.compile(r"\b" + re.escape(term) + r"\b", re.IGNORECASE)
    return pattern.sub(
        lambda m: _match_case(m.group(0), default_replacement), text
    )


def apply_all_accepted_replacements(text: str, accepted: list) -> str:
    """
    Apply a list of (term, replacement) tuples to `text` in order.
    Used when generating a corrected copy of a document so that lines
    containing several non-inclusive terms are fully rewritten.
    """
    out = text
    for term, replacement in accepted:
        out = apply_replacement_to_text(out, term, replacement)
    return out


def scan_for_inclusive_language(parsed_lines: list) -> dict:
    """
    Scans each line for non-inclusive words.

    Returns:
        {
            "findings": [
                {
                    "page": int,
                    "line": int,
                    "found_word": str,
                    "term": str,                  # canonical key in NON_INCLUSIVE_TERMS
                    "suggested_replacement": str, # full replacement string
                    "context": str,               # short context excerpt
                    "original_sentence": str,     # full original line
                    "suggested_sentence": str,    # original line with this term replaced
                },
                ...
            ],
            "total_non_inclusive_count": int,
            "total_word_count": int,
        }
    """
    findings = []
    total_word_count = 0

    for entry in parsed_lines:
        page = entry["page"]
        line = entry["line"]
        text = entry["text"]

        total_word_count += len(text.split())

        for term, replacement in NON_INCLUSIVE_TERMS.items():
            pattern = re.compile(r"\b" + re.escape(term) + r"\b", re.IGNORECASE)
            matches = pattern.findall(text)
            if not matches:
                continue

            suggested_sentence = apply_replacement_to_text(text, term, replacement)

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
                    "term": term,
                    "suggested_replacement": replacement,
                    "context": context,
                    "original_sentence": text,
                    "suggested_sentence": suggested_sentence,
                })

    return {
        "findings": findings,
        "total_non_inclusive_count": len(findings),
        "total_word_count": total_word_count,
    }