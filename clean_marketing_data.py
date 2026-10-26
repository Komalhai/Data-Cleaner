#!/usr/bin/env python3
"""
clean_marketing_data.py
========================
A reusable cleaner/validator for B2B marketing contact spreadsheets of ANY
length and ANY header names.

USAGE
-----
    python clean_marketing_data.py input.xlsx
    python clean_marketing_data.py input.xlsx -o cleaned_output.xlsx
    python clean_marketing_data.py input.xlsx --mapping my_mapping.json
    python clean_marketing_data.py input.xlsx --decision-titles titles.json

WHAT IT DOES
------------
1. HEADER MAPPING   - auto-detects which column is company/email/phone/etc,
                       no matter what the original file calls them, using an
                       exact + synonym + fuzzy-match dictionary.
2. VALIDATION       - checks every email and phone number, flags placeholder/
                       fake values (dummy@..., 9999999999, repeated-digit
                       numbers, values reused across many different
                       companies).
3. COMPLETENESS     - scores each row on how much useful info it has.
4. DEDUPLICATION    - finds duplicate companies even when spelled differently
                       ("A D Industrial" vs "A.D. Industrial Enterprises"),
                       and groups multiple contacts under one company so you
                       don't message the same company 10 times.
5. DECISION-MAKER   - if a designation/job-title column exists, ranks
                       contacts by relevance to a keyword list you supply.
6. COMPLIANCE FLAGS - warns if there's no consent / source / date-collected
                       column at all, since that's a business-process gap no
                       script can invent data for.
7. OUTPUT           - one Excel workbook with separate sheets: a ready-to-use
                       master list, a needs-review list, duplicate clusters,
                       the header mapping it used (for transparency), and a
                       summary.

This script does not require you to edit code for a new file's headers -
it re-detects them every run. If it mismaps something, override it with
--mapping (see mapping_template.json).
"""

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz

try:
    import phonenumbers
    HAVE_PHONENUMBERS = True
except ImportError:
    HAVE_PHONENUMBERS = False


# =============================================================================
# 1. STANDARD SCHEMA + HEADER SYNONYM DICTIONARY
# =============================================================================
# Every possible messy header you might see gets mapped to one of these
# canonical field names. Add more synonyms here as you meet new files -
# this list is the main thing you'll extend over time.

STANDARD_SCHEMA = {
    "company_name": [
        "company", "company name", "companyname", "organisation", "organization",
        "business name", "firm", "firm name", "account name", "customer name",
        "name 1", "name1", "client name",
    ],
    "contact_person": [
        "contact person", "contact name", "poc", "point of contact", "contact",
        "salutation/ contact person", "salutation contact person", "person name",
        "attention", "attn",
    ],
    "designation": [
        "designation", "job title", "title", "role", "position", "job role",
    ],
    "email": [
        "email", "e-mail", "email1", "email address", "work email",
        "contact email", "mail id", "mailid",
    ],
    "phone_mobile": [
        "mobile", "mob", "cell", "mobile number", "mob_num", "mob num",
        "phone", "phone number", "contact no", "contact number",
        "mob_num to print on invoice", "mobile no",
    ],
    "phone_landline": [
        "telephone", "landline", "office phone", "std code",
        "telephone no with std code", "tel no", "off phone",
    ],
    "address_line1": ["street", "address", "address line 1", "address1", "street1", "addr1"],
    "address_line2": ["street2", "address line 2", "address2", "addr2"],
    "address_line3": ["street3", "address line 3", "address3", "addr3"],
    "address_line4": ["street4", "address line 4", "address4", "addr4"],
    "postal_code": ["postal code", "pincode", "pin", "pin code", "zip", "zip code", "zipcode"],
    "city": ["city", "town"],
    "state": ["state", "province", "region"],
    "industry": ["industry", "sector", "business type", "category", "vertical"],
    "source": ["source", "lead source", "data source", "channel"],
    "date_collected": ["date", "date collected", "created date", "added on", "acquired date"],
    "consent_status": ["consent", "consent status", "opt-in", "opt in", "gdpr consent", "marketing consent"],
}

# Columns we specifically expect might be absent -> drives compliance warnings
COMPLIANCE_FIELDS = ["source", "date_collected", "consent_status"]

# Key fields used for the completeness score (weights sum doesn't need to be 100,
# it's just relative importance)
COMPLETENESS_WEIGHTS = {
    "company_name": 20,
    "contact_person": 15,
    "email": 20,
    "phone_mobile": 15,
    "phone_landline": 5,
    "address_line1": 10,
    "city": 10,
    "postal_code": 5,
}

# =============================================================================
# 2. VALIDATION RULES
# =============================================================================

EMAIL_REGEX = re.compile(r"^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$")

PLACEHOLDER_EMAIL_PATTERNS = [
    "dummy", "test@", "sample", "xxx", "noreply", "no-reply", "n/a",
    "example.com", "abc@", "asdf",
]

GENERIC_EMAIL_PREFIXES = [
    "info", "sales", "admin", "contact", "hello", "support", "office",
    "enquiry", "inquiry", "mail",
]

# Numbers that are almost always fake/placeholder in Indian B2B sheets
PLACEHOLDER_PHONE_PATTERNS = [
    re.compile(r"^(\d)\1{9,}$"),   # same digit repeated (9999999999, 8888888888...)
    re.compile(r"^1234567890$"),
    re.compile(r"^0000000000$"),
]


def clean_str(val):
    """Normalize any cell to a trimmed string, or '' for empty/NaN."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    if s.lower() in ("nan", "none", "null", "n/a", "na", "-"):
        return ""
    return s


def normalize_header(h):
    """Lowercase, strip punctuation/extra spaces for matching purposes."""
    h = str(h).lower()
    h = re.sub(r"[^a-z0-9 ]", " ", h)
    h = re.sub(r"\s+", " ", h).strip()
    return h


def validate_email(email):
    """
    Returns dict: valid_format(bool), is_placeholder(bool), is_generic(bool), note(str)
    """
    email = clean_str(email)
    if not email:
        return {"value": "", "valid_format": False, "is_placeholder": False,
                "is_generic": False, "note": "missing"}

    low = email.lower()
    valid_format = bool(EMAIL_REGEX.match(email))
    is_placeholder = any(p in low for p in PLACEHOLDER_EMAIL_PATTERNS)
    local_part = low.split("@")[0] if "@" in low else low
    is_generic = local_part in GENERIC_EMAIL_PREFIXES or any(
        local_part.startswith(g) for g in GENERIC_EMAIL_PREFIXES
    )

    if not valid_format:
        note = "invalid format"
    elif is_placeholder:
        note = "placeholder/fake email"
    elif is_generic:
        note = "generic mailbox (not a named person)"
    else:
        note = "ok"

    return {"value": email, "valid_format": valid_format, "is_placeholder": is_placeholder,
            "is_generic": is_generic, "note": note}


def validate_phone(phone, default_region="IN"):
    """
    Returns dict: value, digits_only, valid(bool), is_placeholder(bool), note(str)

    IMPORTANT: "valid" here means "a genuine Indian mobile number" - not
    just "a structurally plausible phone number of any country". A
    correctly-formatted foreign number or an Indian landline is explicitly
    marked invalid, because this tool is for India-focused marketing
    outreach. (Uses the same India-only classifier as the WhatsApp-list
    feature below, so both parts of the tool agree on what counts as valid.)
    """
    raw = clean_str(phone)
    if not raw:
        return {"value": "", "digits_only": "", "valid": False,
                "is_placeholder": False, "note": "missing"}

    raw2 = raw.replace(".0", "") if raw.endswith(".0") else raw
    digits = re.sub(r"\D", "", raw2)

    c = classify_phone_number(raw2)
    is_placeholder = c["reason"] == "Placeholder/fake pattern (repeated or sequential digits)"
    valid = c["status"] == "Valid"

    if is_placeholder:
        note = "placeholder/fake number"
    elif c["country"] == "Foreign (non-India)":
        note = "foreign number - excluded (India-only marketing list)"
    elif c["number_type"] == "Landline/Other" and c["country"] == "India":
        note = "Indian landline - not a mobile number"
    elif not valid:
        note = f"invalid - {c['reason']}" if c["reason"] else "invalid / unparseable"
    else:
        note = "ok"

    return {"value": raw, "digits_only": digits, "valid": valid,
            "is_placeholder": is_placeholder, "note": note}


LEGAL_SUFFIXES = [
    "PRIVATE LIMITED", "PVT LTD", "PVT. LTD.", "PVT LTD.", "LIMITED", "LTD",
    "LTD.", "LLP", "CORP", "CORPORATION", "CO.", "CO", "INDIA", "INC",
]


# =============================================================================
# 2b. MULTI-VALUE CELL SPLITTING
# =============================================================================
# Handles cells that actually contain more than one phone number or email,
# separated in any of the ways people type them: newlines, commas, "and",
# slashes, or just extra spaces with no punctuation at all
# (e.g. "9870340080  8384802066"). Each value becomes its own row later on,
# so validation runs against a single clean number/email instead of a
# concatenated blob that fails every length/format check.

def _split_phone_tokens(s):
    """Split a phone-ish string on whitespace, then re-merge short digit
    fragments (e.g. '98703' + '40080') back into one number, so a single
    number written with internal spacing isn't mistaken for two numbers."""
    tokens = [t for t in re.split(r"\s+", s.strip()) if t]
    results, buffer = [], ""
    for tok in tokens:
        digits = re.sub(r"\D", "", tok)
        if not digits:
            continue
        if len(digits) >= 8:
            if buffer:
                results.append(buffer)
                buffer = ""
            results.append(digits)
        else:
            buffer += digits
            if len(buffer) >= 10:
                results.append(buffer)
                buffer = ""
    if buffer:
        results.append(buffer)
    return results


def split_multi_value(raw, kind="phone"):
    """Return a list of individual values found in one cell. kind is
    'phone' or 'email'. A cell with just one value returns a 1-item list;
    an empty cell returns []."""
    s = clean_str(raw)
    if not s:
        return []
    s = re.sub(r"\r\n|\r|\n", "|", s)
    s = re.sub(r"\s+and\s+", "|", s, flags=re.IGNORECASE)
    s = re.sub(r"[,;/&]", "|", s)
    parts = [p.strip() for p in s.split("|") if p.strip()]

    if kind == "phone":
        expanded = []
        for p in parts:
            expanded.extend(_split_phone_tokens(p))
        return expanded
    else:  # email
        expanded = []
        for p in parts:
            expanded.extend(t for t in re.split(r"\s+", p) if t)
        return expanded


def expand_multi_value_rows(records):
    """Given a list of standardized row dicts, split any row whose
    phone_mobile and/or email cell contains multiple values into several
    rows - one value per row, every other field duplicated as-is.

    If a row has the SAME number of phones as emails (and more than one of
    each), they're paired by position (1st phone with 1st email, etc.) on
    the assumption they were typed as matching pairs. Otherwise phones and
    emails are expanded independently, which naturally produces one row per
    combination when both are multi-valued in the same row (rare) - flagged
    via multi_value_split_note either way so it's never silent."""
    result = []
    for rec in records:
        phones = split_multi_value_preserving_format(rec.get("phone_mobile", ""), "phone")
        emails = split_multi_value(rec.get("email", ""), "email")
        n_phone, n_email = len(phones), len(emails)

        if n_phone <= 1 and n_email <= 1:
            result.append(rec)
            continue

        if n_phone == n_email and n_phone > 1:
            for i in range(n_phone):
                newrec = dict(rec)
                newrec["phone_mobile"] = phones[i]
                newrec["email"] = emails[i]
                newrec["multi_value_split_note"] = (
                    f"split {i+1} of {n_phone} - one cell had multiple phone+email pairs"
                )
                result.append(newrec)
            continue

        phone_list = phones if n_phone > 1 else [rec.get("phone_mobile", "")]
        base_rows = []
        for i, ph in enumerate(phone_list):
            newrec = dict(rec)
            newrec["phone_mobile"] = ph
            if n_phone > 1:
                newrec["multi_value_split_note"] = f"phone {i+1} of {n_phone} from one cell"
            base_rows.append(newrec)

        for br in base_rows:
            email_list = split_multi_value(br.get("email", ""), "email")
            if len(email_list) > 1:
                for j, em in enumerate(email_list):
                    r2 = dict(br)
                    r2["email"] = em
                    note = f"email {j+1} of {len(email_list)} from one cell"
                    prev = r2.get("multi_value_split_note", "")
                    r2["multi_value_split_note"] = f"{prev}; {note}" if prev else note
                    result.append(r2)
            else:
                result.append(br)
    return result


def normalize_company_name(name):
    """Uppercase, strip punctuation and legal suffixes -> matching key."""
    name = clean_str(name).upper()
    name = re.sub(r"\(\s*P\s*\)", "PVT", name)  # "(P) LTD" -> "PVT LTD" before suffix stripping
    name = re.sub(r"[.,()&/\\\-]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    for suf in LEGAL_SUFFIXES:
        pattern = r"\b" + re.escape(suf).replace(r"\ ", r"\s+") + r"\b"
        name = re.sub(pattern, "", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


# =============================================================================
# 2c. WHATSAPP / INDIAN-MOBILE-ONLY CLASSIFICATION
# =============================================================================
# Structural validation only - this NEVER claims a number is actually
# registered on WhatsApp. It only decides whether a number has the shape of
# a valid Indian mobile number (10 digits, starts 6-9, no foreign/landline
# pattern). Real WhatsApp registration can only be confirmed through the
# official WhatsApp Business/API verification process.

INDIAN_MOBILE_PREFIXES = ("6", "7", "8", "9")
PLACEHOLDER_10DIGIT = re.compile(r"^(\d)\1{9}$")
SEQUENTIAL_10DIGIT = {"1234567890", "0123456789"}


def _group_phone_segments(s):
    """Like _split_phone_tokens, but returns the ORIGINAL substrings
    (formatting intact - '+', dashes, spaces within one number) instead of
    bare digits, so classify_phone_number can still see a leading '+91' or
    '00' to detect country codes correctly."""
    tokens = [t for t in s.strip().split() if t]
    groups, buffer = [], []
    buffer_digits = 0
    for tok in tokens:
        d = re.sub(r"\D", "", tok)
        if not d:
            continue
        if len(d) >= 8:
            if buffer:
                groups.append(buffer)
                buffer, buffer_digits = [], 0
            groups.append([tok])
        else:
            buffer.append(tok)
            buffer_digits += len(d)
            if buffer_digits >= 10:
                groups.append(buffer)
                buffer, buffer_digits = [], 0
    if buffer:
        groups.append(buffer)
    return [" ".join(g) for g in groups]


def split_multi_value_preserving_format(raw, kind="phone"):
    """Same multi-value splitting as split_multi_value, but keeps each
    number's original punctuation ('+', '00', dashes) intact instead of
    reducing to bare digits - needed here so classify_phone_number can
    still detect a country-code prefix."""
    s = clean_str(raw)
    if not s:
        return []
    s = re.sub(r"\r\n|\r|\n", "|", s)
    s = re.sub(r"\s+and\s+", "|", s, flags=re.IGNORECASE)
    s = re.sub(r"[,;/&]", "|", s)
    parts = [p.strip() for p in s.split("|") if p.strip()]

    if kind == "phone":
        expanded = []
        for p in parts:
            expanded.extend(_group_phone_segments(p))
        return expanded
    else:
        expanded = []
        for p in parts:
            expanded.extend(t for t in re.split(r"\s+", p) if t)
        return expanded


def classify_phone_number(raw_token):
    """
    Classify a single already-split phone value (may still carry punctuation
    - this function does its own digit cleaning, so it's safe to call on a
    raw cell value directly too).

    Returns: {cleaned, country, number_type, status, reason}
      status  in {"Valid", "Invalid"}   (a later dedup pass may relabel some
                                          "Valid" rows to "Duplicate")
      country in {"India", "Foreign (non-India)", "Unknown"}
      number_type in {"Mobile", "Landline/Other", "Unknown"}
    """
    original = clean_str(raw_token)
    if not original:
        return {"cleaned": "", "country": "", "number_type": "",
                 "status": "Invalid", "reason": "Blank/missing"}

    digits = re.sub(r"\D", "", original)
    if not digits:
        return {"cleaned": "", "country": "Unknown", "number_type": "Unknown",
                 "status": "Invalid", "reason": "No digits found"}

    core = None
    foreign_hint = False

    if digits.startswith("0091") and len(digits) == 14:
        core = digits[4:]
    elif digits.startswith("91") and len(digits) == 12:
        core = digits[2:]
    elif len(digits) == 10:
        core = digits
    elif digits.startswith("0") and len(digits) == 11:
        core = digits[1:]  # trunk-prefix style, common for both landline and (incorrectly) mobile
    elif (original.strip().startswith("+") or digits.startswith("00")) and not digits.lstrip("0").startswith("91"):
        foreign_hint = True

    if foreign_hint:
        return {"cleaned": "", "country": "Foreign (non-India)", "number_type": "Unknown",
                 "status": "Invalid", "reason": "Foreign country code detected"}

    if core is None:
        return {"cleaned": "", "country": "Unknown", "number_type": "Unknown",
                 "status": "Invalid", "reason": f"Unrecognized/incomplete format ({len(digits)} digits)"}

    if len(core) != 10:
        return {"cleaned": "", "country": "India", "number_type": "Unknown",
                 "status": "Invalid", "reason": f"Not exactly 10 digits after cleaning ({len(core)} digits)"}

    if PLACEHOLDER_10DIGIT.match(core) or core in SEQUENTIAL_10DIGIT:
        return {"cleaned": core, "country": "India", "number_type": "Unknown",
                 "status": "Invalid", "reason": "Placeholder/fake pattern (repeated or sequential digits)"}

    if core[0] in INDIAN_MOBILE_PREFIXES:
        return {"cleaned": core, "country": "India", "number_type": "Mobile",
                 "status": "Valid", "reason": ""}

    return {"cleaned": core, "country": "India", "number_type": "Landline/Other",
             "status": "Invalid", "reason": "Not a valid Indian mobile prefix (likely landline)"}


def find_matching_columns(columns, synonyms, threshold=85):
    """Like build_header_mapping's matcher, but returns EVERY column that
    matches (not just the single best one) - needed here because a file may
    have several genuinely-separate phone columns (Mobile, Alternate Mobile,
    WhatsApp Number...) that should ALL be scanned for numbers."""
    matches = []
    for col in columns:
        if BLANK_HEADER_PATTERN.match(str(col)):
            continue
        norm = normalize_header(col)
        best = 0
        for syn in synonyms:
            if norm == syn:
                best = 100
                break
            best = max(best, fuzz.ratio(norm, syn))
        if best >= threshold:
            matches.append(col)
    return matches


def build_whatsapp_extraction(raw_df):
    """
    Scans every phone-like column in raw_df, splits multi-value cells,
    classifies every individual number, and de-duplicates valid Indian
    mobile numbers. Returns (records_df, summary_dict).
    """
    columns = list(raw_df.columns)
    phone_cols = find_matching_columns(
        columns, STANDARD_SCHEMA["phone_mobile"] + STANDARD_SCHEMA["phone_landline"]
    )
    name_cols = find_matching_columns(columns, STANDARD_SCHEMA["company_name"])
    contact_cols = find_matching_columns(columns, STANDARD_SCHEMA["contact_person"])
    name_col = name_cols[0] if name_cols else None
    contact_col = contact_cols[0] if contact_cols else None

    total_records = len(raw_df)
    records = []

    for idx, row in raw_df.iterrows():
        company = clean_str(row[name_col]) if name_col else ""
        contact = clean_str(row[contact_col]) if contact_col else ""
        display_name = company or contact or f"Row {idx + 2}"  # +2 = header row + 1-indexed

        found_any_cell_content = False
        for pc in phone_cols:
            raw_cell = clean_str(row[pc])
            if not raw_cell:
                continue
            found_any_cell_content = True
            tokens = split_multi_value_preserving_format(raw_cell, "phone")
            if not tokens:
                tokens = [raw_cell]  # let classify_phone_number produce a clear rejection reason
            for tok in tokens:
                c = classify_phone_number(tok)
                records.append({
                    "customer_company_name": display_name,
                    "original_phone_number": raw_cell,
                    "cleaned_indian_mobile_number": c["cleaned"] if c["status"] == "Valid" else "",
                    "country": c["country"],
                    "number_type": c["number_type"],
                    "validation_status": c["status"],
                    "reason_for_rejection": c["reason"],
                })

        if phone_cols and not found_any_cell_content:
            records.append({
                "customer_company_name": display_name,
                "original_phone_number": "",
                "cleaned_indian_mobile_number": "",
                "country": "", "number_type": "",
                "validation_status": "Invalid", "reason_for_rejection": "Blank/missing",
            })

    # de-duplicate: keep first occurrence of each valid cleaned number
    seen = set()
    for r in records:
        if r["validation_status"] == "Valid":
            if r["cleaned_indian_mobile_number"] in seen:
                r["validation_status"] = "Duplicate"
                r["reason_for_rejection"] = "Duplicate of an earlier valid number"
            else:
                seen.add(r["cleaned_indian_mobile_number"])

    records_df = pd.DataFrame(records, columns=[
        "customer_company_name", "original_phone_number", "cleaned_indian_mobile_number",
        "country", "number_type", "validation_status", "reason_for_rejection",
    ])

    def count(mask_fn):
        return int(records_df.apply(mask_fn, axis=1).sum()) if len(records_df) else 0

    summary = {
        "total_records_received": total_records,
        "phone_columns_used": phone_cols,
        "total_phone_numbers_found": count(lambda r: r["original_phone_number"] != ""),
        "valid_indian_mobile_numbers": count(lambda r: r["validation_status"] == "Valid"),
        "indian_landline_numbers_removed": count(
            lambda r: r["reason_for_rejection"] == "Not a valid Indian mobile prefix (likely landline)"),
        "foreign_numbers_removed": count(lambda r: r["reason_for_rejection"] == "Foreign country code detected"),
        "invalid_numbers_removed": count(lambda r: r["validation_status"] == "Invalid" and r["reason_for_rejection"] not in (
            "Blank/missing", "Foreign country code detected", "Not a valid Indian mobile prefix (likely landline)")),
        "duplicate_numbers_removed": count(lambda r: r["validation_status"] == "Duplicate"),
        "blank_missing_numbers": count(lambda r: r["reason_for_rejection"] == "Blank/missing"),
    }
    if not phone_cols:
        summary["warning"] = "No phone-like column was detected in this file - nothing to extract."

    return records_df, summary


# =============================================================================
# 3. HEADER MAPPING
# =============================================================================

BLANK_HEADER_PATTERN = re.compile(r"^unnamed[:\s]", re.IGNORECASE)

# Targets that legitimately CAN be filled by more than one raw column
# (e.g. an address commonly spans several columns). Any target not listed
# here may only be claimed by the single best-matching column; a second,
# weaker match is left unmapped and flagged for manual review instead of
# silently overwriting/competing with the first.
MULTI_COLUMN_TARGETS = {"address_line1", "address_line2", "address_line3", "address_line4"}


def build_header_mapping(columns, manual_overrides=None, fuzzy_threshold=85):
    """
    Given raw column names, return:
      mapping: {raw_column: standard_field or None}
      report: list of dicts describing how/why each mapping was chosen

    Rules:
    - Manual overrides always win.
    - A blank/auto-generated header ("Unnamed: 1", pandas' name for a column
      that had no header text) is NEVER auto-mapped - its meaning can't be
      inferred from an empty header, so it's always left for manual review.
    - If two columns both best-match the same single-valued target (e.g. two
      columns both look like "email"), only the higher-confidence one is
      mapped; the other is left unmapped and flagged as a conflict, rather
      than silently overwriting or merging data.
    """
    manual_overrides = manual_overrides or {}
    mapping = {}
    report = []
    candidates = []  # (col, target, score, method) for non-override columns

    for col in columns:
        if col in manual_overrides:
            target = manual_overrides[col]
            mapping[col] = target
            report.append({"raw_column": col, "mapped_to": target, "method": "manual override"})
            continue

        if BLANK_HEADER_PATTERN.match(str(col)):
            mapping[col] = None
            report.append({"raw_column": col, "mapped_to": None,
                            "method": "unmapped - blank original header, cannot infer meaning",
                            "confidence": 0})
            continue

        norm_col = normalize_header(col)
        best_target, best_score, best_method = None, 0, None

        for target, synonyms in STANDARD_SCHEMA.items():
            for syn in synonyms:
                if norm_col == syn:
                    best_target, best_score, best_method = target, 100, "exact match"
                    break
                score = fuzz.ratio(norm_col, syn)
                if score > best_score:
                    best_target, best_score, best_method = target, score, "fuzzy match"
            if best_method == "exact match":
                break

        candidates.append((col, best_target, best_score, best_method))

    # Resolve conflicts: for single-valued targets, keep only the strongest match.
    best_by_target = {}
    for col, target, score, method in candidates:
        if target is None or score < fuzzy_threshold:
            continue
        if target in MULTI_COLUMN_TARGETS:
            continue  # these can all be kept, handled below
        if target not in best_by_target or score > best_by_target[target][1]:
            best_by_target[target] = (col, score)

    for col, target, score, method in candidates:
        if target is None or score < fuzzy_threshold:
            mapping[col] = None
            report.append({"raw_column": col, "mapped_to": None,
                            "method": "unmapped - no confident match, needs manual review",
                            "confidence": score})
        elif target in MULTI_COLUMN_TARGETS:
            mapping[col] = target
            report.append({"raw_column": col, "mapped_to": target, "method": method, "confidence": score})
        elif best_by_target[target][0] == col:
            mapping[col] = target
            report.append({"raw_column": col, "mapped_to": target, "method": method, "confidence": score})
        else:
            mapping[col] = None
            report.append({"raw_column": col, "mapped_to": None,
                            "method": f"unmapped - conflicted with a stronger match for '{target}'",
                            "confidence": score})

    return mapping, report


# =============================================================================
# 4. MAIN PIPELINE
# =============================================================================

def load_file(path):
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path)
    elif path.suffix.lower() in (".csv", ".tsv"):
        sep = "\t" if path.suffix.lower() == ".tsv" else ","
        return pd.read_csv(path, sep=sep)
    else:
        raise ValueError(f"Unsupported file type: {path.suffix}")


def run_pipeline(input_path, manual_mapping=None, decision_titles=None,
                  fuzzy_dup_threshold=90):
    raw_df = load_file(input_path)
    raw_df = raw_df.where(pd.notnull(raw_df), "")

    mapping, mapping_report = build_header_mapping(list(raw_df.columns), manual_mapping)

    # Build a standardized dataframe. If multiple raw columns map to the same
    # target, keep the first non-empty value found, and note the conflict.
    std_fields = list(STANDARD_SCHEMA.keys())
    std_df = pd.DataFrame(index=raw_df.index, columns=std_fields, data="")
    field_sources = defaultdict(list)

    for raw_col, target in mapping.items():
        if target is None:
            continue
        field_sources[target].append(raw_col)
        col_vals = raw_df[raw_col].apply(clean_str)
        std_df[target] = std_df[target].where(std_df[target] != "", col_vals)

    unmapped_columns = [c for c, t in mapping.items() if t is None]

    # ---- Split any cells that actually hold multiple phones/emails --------
    std_records = std_df.to_dict("records")
    std_records = expand_multi_value_rows(std_records)

    # ---- Row-level validation -------------------------------------------------
    rows = []
    for i, row in enumerate(std_records):
        email_check = validate_email(row["email"])
        phone_m = validate_phone(row["phone_mobile"])
        phone_l = validate_phone(row["phone_landline"])

        completeness = 0
        max_score = sum(COMPLETENESS_WEIGHTS.values())
        for field, weight in COMPLETENESS_WEIGHTS.items():
            if clean_str(row[field]):
                completeness += weight
        completeness_pct = round(100 * completeness / max_score, 1)

        best_phone_valid = phone_m["valid"] or phone_l["valid"]

        issues = []
        if not row["company_name"]:
            issues.append("missing company name")
        if not email_check["value"]:
            issues.append("missing email")
        elif not email_check["valid_format"]:
            issues.append("invalid email format")
        elif email_check["is_placeholder"]:
            issues.append("placeholder/fake email")
        elif email_check["is_generic"]:
            issues.append("generic (non-personal) email")
        if not row["phone_mobile"] and not row["phone_landline"]:
            issues.append("missing phone")
        elif not best_phone_valid:
            issues.append("invalid/placeholder phone")
        if not row["city"] and not row["address_line1"]:
            issues.append("missing location info")

        rows.append({
            "row_index": i,
            "company_name": row["company_name"],
            "company_key": normalize_company_name(row["company_name"]),
            "contact_person": row["contact_person"],
            "designation": row["designation"],
            "email": email_check["value"],
            "email_note": email_check["note"],
            "phone_mobile": row["phone_mobile"],
            "phone_mobile_note": phone_m["note"],
            "phone_landline": row["phone_landline"],
            "phone_landline_note": phone_l["note"],
            "address_line1": row["address_line1"],
            "address_line2": row["address_line2"],
            "city": row["city"],
            "state": row["state"],
            "postal_code": row["postal_code"],
            "industry": row["industry"],
            "source": row["source"],
            "date_collected": row["date_collected"],
            "consent_status": row["consent_status"],
            "completeness_pct": completeness_pct,
            "phone_ok": best_phone_valid,
            "email_ok": email_check["valid_format"] and not email_check["is_placeholder"],
            "issues": "; ".join(issues) if issues else "none",
            "multi_value_split_note": row.get("multi_value_split_note", ""),
        })

    work_df = pd.DataFrame(rows)

    # ---- Cross-row: flag emails/phones reused across many different companies -
    email_company_counts = (
        work_df[work_df["email"] != ""]
        .groupby("email")["company_key"].nunique()
    )
    shared_emails = set(email_company_counts[email_company_counts >= 3].index)
    work_df["email_shared_flag"] = work_df["email"].apply(
        lambda e: "shared across many companies (likely placeholder)" if e in shared_emails else ""
    )

    phone_company_counts = (
        work_df[work_df["phone_mobile"] != ""]
        .groupby("phone_mobile")["company_key"].nunique()
    )
    shared_phones = set(phone_company_counts[phone_company_counts >= 3].index)
    work_df["phone_shared_flag"] = work_df["phone_mobile"].apply(
        lambda p: "shared across many companies (likely placeholder)" if p in shared_phones else ""
    )

    # ---- Decision-maker relevance (only if designation data + keyword list) ---
    decision_titles = decision_titles or []
    has_designation_data = work_df["designation"].str.strip().ne("").any()

    def decision_score(designation):
        if not decision_titles or not designation:
            return None
        designation_low = designation.lower()
        best = max((fuzz.partial_ratio(designation_low, kw.lower()) for kw in decision_titles),
                   default=0)
        return best

    if has_designation_data and decision_titles:
        work_df["decision_maker_score"] = work_df["designation"].apply(decision_score)
    else:
        work_df["decision_maker_score"] = None

    # ---- Deduplication: exact key match = definite; fuzzy + shared contact = likely
    groups = defaultdict(list)
    for idx, key in work_df["company_key"].items():
        groups[key if key else f"__blank_{idx}"].append(idx)

    # merge fuzzy-similar keys into the same cluster when they also share phone/email/city
    keys = [k for k in groups.keys() if not k.startswith("__blank_")]
    parent = {k: k for k in keys}

    def find(k):
        while parent[k] != k:
            parent[k] = parent[parent[k]]
            k = parent[k]
        return k

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    # Two separate bars, both deliberately conservative to avoid false merges
    # in datasets where many unrelated companies share a city or generic
    # industry words (e.g. "Shree Sai Steels" vs "Shree Ji Steels" - both
    # steel companies in the same industrial belt, but NOT the same company):
    #   - AUTO_MERGE_BY_CONTACT: name is at least somewhat similar AND the two
    #     groups share an actual email or phone number - a strong signal the
    #     records are the same real-world contact/company.
    #   - AUTO_MERGE_BY_NAME_ONLY: name match is almost exact (punctuation/
    #     spacing differences only) - safe to merge without needing a shared
    #     contact.
    # Anything in between is left un-merged but recorded as a "possible
    # duplicate - review" note so nothing is silently lost either way.
    AUTO_MERGE_BY_CONTACT_NAME_FLOOR = 80
    AUTO_MERGE_BY_NAME_ONLY_FLOOR = 97
    POSSIBLE_DUP_FLOOR = fuzzy_dup_threshold

    possible_dup_notes = defaultdict(list)

    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            k1, k2 = keys[i], keys[j]
            score = fuzz.token_sort_ratio(k1, k2)
            if score < POSSIBLE_DUP_FLOOR:
                continue

            rows1 = work_df.loc[groups[k1]]
            rows2 = work_df.loc[groups[k2]]
            shares_contact = bool(
                (set(rows1["email"]) & set(rows2["email"])) - {""}
                or (set(rows1["phone_mobile"]) & set(rows2["phone_mobile"])) - {""}
            )

            if score >= AUTO_MERGE_BY_NAME_ONLY_FLOOR:
                union(k1, k2)
            elif shares_contact and score >= AUTO_MERGE_BY_CONTACT_NAME_FLOOR:
                union(k1, k2)
            else:
                # Not confident enough to auto-merge - flag both sides for a human to check.
                possible_dup_notes[k1].append((k2, score))
                possible_dup_notes[k2].append((k1, score))

    cluster_of = {}
    for k in keys:
        cluster_of[k] = find(k)
    work_df["cluster_key"] = work_df["company_key"].apply(
        lambda k: cluster_of.get(k, k if k else f"blank_{id(k)}")
    )

    def possible_dup_text(key):
        notes = possible_dup_notes.get(key)
        if not notes:
            return ""
        # only mention it if the other side wasn't itself absorbed into this cluster already
        parts = sorted(set(n for n, s in notes if cluster_of.get(n, n) != cluster_of.get(key, key)),
                        )
        return "; ".join(parts[:3])

    work_df["possible_duplicate_of"] = work_df["company_key"].apply(possible_dup_text)

    # Rank contacts within a cluster: prefer valid email + valid phone + higher
    # completeness + higher decision-maker score
    def rank_key(row):
        return (
            row["email_ok"],
            row["phone_ok"],
            row["decision_maker_score"] if row["decision_maker_score"] is not None else -1,
            row["completeness_pct"],
        )

    work_df["_rank"] = work_df.apply(rank_key, axis=1)
    work_df["contact_rank_in_company"] = (
        work_df.groupby("cluster_key")["_rank"]
        .rank(method="first", ascending=False)
        .astype(int)
    )
    work_df["is_primary_contact"] = work_df["contact_rank_in_company"] == 1
    work_df["contacts_at_this_company"] = work_df.groupby("cluster_key")["cluster_key"].transform("count")
    work_df.drop(columns=["_rank"], inplace=True)

    duplicate_flagged = work_df["contacts_at_this_company"] > 1

    # ---- Compliance-level warnings (file-wide, not row-wide) -------------------
    compliance_warnings = []
    for field in COMPLIANCE_FIELDS:
        if not any(mapping.get(c) == field for c in mapping):
            label = {
                "source": "data SOURCE (where each record came from)",
                "date_collected": "DATE COLLECTED (how fresh each record is)",
                "consent_status": "CONSENT STATUS (marketing opt-in/compliance)",
            }[field]
            compliance_warnings.append(f"No column found for {label}.")

    summary = {
        "total_rows": len(work_df),
        "unique_companies_after_dedup": work_df["cluster_key"].nunique(),
        "rows_flagged_as_duplicate_contact": int(duplicate_flagged.sum()),
        "missing_email": int((work_df["email"] == "").sum()),
        "invalid_or_placeholder_email": int(
            work_df["email_note"].isin(["invalid format", "placeholder/fake email"]).sum()
        ),
        "generic_email": int((work_df["email_note"] == "generic mailbox (not a named person)").sum()),
        "missing_phone": int(((work_df["phone_mobile"] == "") & (work_df["phone_landline"] == "")).sum()),
        "invalid_or_placeholder_phone": int((~work_df["phone_ok"]).sum()),
        "avg_completeness_pct": round(work_df["completeness_pct"].mean(), 1),
        "rows_with_issues": int((work_df["issues"] != "none").sum()),
        "rows_split_from_multi_value_cells": int((work_df["multi_value_split_note"] != "").sum()),
        "unmapped_columns": unmapped_columns,
        "compliance_warnings": compliance_warnings,
        "decision_maker_data_available": bool(has_designation_data),
    }

    return {
        "work_df": work_df,
        "mapping_report": mapping_report,
        "summary": summary,
    }


# =============================================================================
# 5. EXCEL OUTPUT
# =============================================================================

def write_output(result, output_path, whatsapp_df=None, whatsapp_summary=None):
    work_df = result["work_df"].copy()
    mapping_report_df = pd.DataFrame(result["mapping_report"])
    summary = result["summary"]

    display_cols = [
        "cluster_key", "is_primary_contact", "contacts_at_this_company",
        "company_name", "contact_person", "designation",
        "email", "email_note", "email_shared_flag",
        "phone_mobile", "phone_mobile_note", "phone_shared_flag",
        "phone_landline", "phone_landline_note",
        "address_line1", "address_line2", "city", "state", "postal_code",
        "industry", "source", "date_collected", "consent_status",
        "completeness_pct", "decision_maker_score", "issues", "possible_duplicate_of",
        "multi_value_split_note",
    ]
    display_cols = [c for c in display_cols if c in work_df.columns]

    master_df = work_df[work_df["is_primary_contact"] | (work_df["contacts_at_this_company"] == 1)]
    master_df = master_df[display_cols].sort_values("company_name")

    needs_review_mask = (work_df["issues"] != "none") | (work_df["possible_duplicate_of"] != "")
    review_df = work_df[needs_review_mask][display_cols].sort_values(
        ["completeness_pct"], ascending=True
    )

    dup_df = work_df[work_df["contacts_at_this_company"] > 1][display_cols].sort_values(
        ["cluster_key", "is_primary_contact"], ascending=[True, False]
    )

    summary_rows = [
        ["Total records in file", summary["total_rows"]],
        ["Unique companies after de-duplication", summary["unique_companies_after_dedup"]],
        ["Contact rows that are duplicates of another contact at same company",
         summary["rows_flagged_as_duplicate_contact"]],
        ["Rows missing email entirely", summary["missing_email"]],
        ["Rows with invalid or placeholder/fake email", summary["invalid_or_placeholder_email"]],
        ["Rows with generic (non-personal) email", summary["generic_email"]],
        ["Rows missing phone entirely", summary["missing_phone"]],
        ["Rows with invalid or placeholder/fake phone", summary["invalid_or_placeholder_phone"]],
        ["Average completeness score (%)", summary["avg_completeness_pct"]],
        ["Rows with at least one flagged issue", summary["rows_with_issues"]],
        ["Decision-maker/job-title data available?", "Yes" if summary["decision_maker_data_available"] else "No"],
        ["", ""],
        ["UNMAPPED COLUMNS (not recognized - review manually):",
         ", ".join(summary["unmapped_columns"]) or "none"],
        ["", ""],
        ["COMPLIANCE WARNINGS:", ""],
    ]
    for w in summary["compliance_warnings"]:
        summary_rows.append(["", w])
    if not summary["compliance_warnings"]:
        summary_rows.append(["", "none - source/date/consent columns all present"])

    summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        master_df.to_excel(writer, sheet_name="Clean_Master_List", index=False)
        review_df.to_excel(writer, sheet_name="Needs_Review", index=False)
        dup_df.to_excel(writer, sheet_name="Duplicate_Contacts", index=False)
        mapping_report_df.to_excel(writer, sheet_name="Header_Mapping_Used", index=False)

        if whatsapp_df is not None:
            RENAME = {
                "customer_company_name": "Customer/Company Name",
                "original_phone_number": "Original Phone Number",
                "cleaned_indian_mobile_number": "Cleaned Indian Mobile Number",
                "country": "Country", "number_type": "Number Type",
                "validation_status": "Validation Status",
                "reason_for_rejection": "Reason for Rejection",
            }
            all_numbers_df = whatsapp_df.rename(columns=RENAME)
            ready_df = all_numbers_df[all_numbers_df["Validation Status"] == "Valid"].drop(
                columns=["Original Phone Number", "Reason for Rejection"]
            )

            ws_summary_rows = [
                ["Total records received", whatsapp_summary["total_records_received"]],
                ["Phone-like columns scanned", ", ".join(whatsapp_summary["phone_columns_used"]) or "none found"],
                ["Total phone numbers found", whatsapp_summary["total_phone_numbers_found"]],
                ["Valid Indian mobile numbers (unique, WhatsApp-ready)", whatsapp_summary["valid_indian_mobile_numbers"]],
                ["Indian landline numbers removed", whatsapp_summary["indian_landline_numbers_removed"]],
                ["Foreign numbers removed", whatsapp_summary["foreign_numbers_removed"]],
                ["Invalid numbers removed", whatsapp_summary["invalid_numbers_removed"]],
                ["Duplicate numbers removed", whatsapp_summary["duplicate_numbers_removed"]],
                ["Blank/missing numbers", whatsapp_summary["blank_missing_numbers"]],
                ["", ""],
                ["NOTE:", "Structural validation only. This does NOT confirm WhatsApp"
                          " registration - that requires the official WhatsApp Business/API"
                          " verification process."],
            ]
            if "warning" in whatsapp_summary:
                ws_summary_rows.insert(0, ["WARNING", whatsapp_summary["warning"]])
            pd.DataFrame(ws_summary_rows, columns=["Metric", "Value"]).to_excel(
                writer, sheet_name="WhatsApp_Summary", index=False)
            ready_df.to_excel(writer, sheet_name="WhatsApp_Ready_List", index=False)
            all_numbers_df.to_excel(writer, sheet_name="All_Numbers_Classified", index=False)

    _format_workbook(output_path)


def _format_workbook(path):
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = load_workbook(path)
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10)

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        max_row = ws.max_row
        max_col = ws.max_column
        if max_col == 0:
            continue
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            width = max(12, min(40, len(str(cell.value or "")) + 6))
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        for row in ws.iter_rows(min_row=2, max_row=max_row):
            for cell in row:
                cell.font = body_font

    wb.save(path)


# =============================================================================
# 6. CLI
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description="Clean and validate a marketing contact spreadsheet.")
    parser.add_argument("input_file", help="Path to the input .xlsx/.csv file")
    parser.add_argument("-o", "--output", default=None, help="Output .xlsx path")
    parser.add_argument("--mapping", default=None, help="JSON file overriding header->field mapping")
    parser.add_argument("--decision-titles", default=None,
                         help="JSON file: list of job-title keywords that count as decision-makers")
    parser.add_argument("--dup-threshold", type=int, default=90,
                         help="Fuzzy match threshold (0-100) for company-name duplicate detection")
    parser.add_argument("--no-whatsapp-list", action="store_true",
                         help="Skip building the WhatsApp-ready Indian-mobile-only sheets")
    args = parser.parse_args()

    manual_mapping = None
    if args.mapping:
        with open(args.mapping) as f:
            manual_mapping = json.load(f)

    decision_titles = None
    if args.decision_titles:
        with open(args.decision_titles) as f:
            decision_titles = json.load(f)

    output_path = args.output or str(Path(args.input_file).with_name(
        Path(args.input_file).stem + "_CLEANED.xlsx"))

    print(f"Reading {args.input_file} ...")
    result = run_pipeline(args.input_file, manual_mapping, decision_titles, args.dup_threshold)

    whatsapp_df, whatsapp_summary = None, None
    if not args.no_whatsapp_list:
        print("Extracting WhatsApp-ready Indian mobile numbers ...")
        raw_df = load_file(args.input_file)
        raw_df = raw_df.where(pd.notnull(raw_df), "")
        whatsapp_df, whatsapp_summary = build_whatsapp_extraction(raw_df)

    print("Writing output ...")
    write_output(result, output_path, whatsapp_df, whatsapp_summary)

    s = result["summary"]
    print("\n=== SUMMARY ===")
    for k, v in s.items():
        print(f"{k}: {v}")
    if whatsapp_summary:
        print("\n=== WHATSAPP-READY NUMBER SUMMARY ===")
        for k, v in whatsapp_summary.items():
            print(f"{k}: {v}")
    print(f"\nDone. Output saved to: {output_path}")


if __name__ == "__main__":
    main()
