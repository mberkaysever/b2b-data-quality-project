#!/usr/bin/env python3
"""
Quality check and qualification of B2B client data for electronic invoicing.

Reads the raw synthetic dataset, applies business rules (SIREN, SIRET, VAT,
email, duplicates), and produces a qualified export plus a management summary.
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

# Legal suffixes stripped before company-name similarity comparison
LEGAL_SUFFIXES = ("sarl", "sas", "sasu", "eurl", "sa", "sci", "snc")
NAME_SIMILARITY_THRESHOLD = 90.0  # percent

# --- Paths ---
PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_FILE = PROJECT_ROOT / "data" / "b2b_clients_raw.xlsx"
OUTPUT_QUALIFIED = PROJECT_ROOT / "outputs" / "b2b_clients_qualified.xlsx"
OUTPUT_SUMMARY = PROJECT_ROOT / "outputs" / "data_quality_summary.xlsx"

# --- Qualification labels (fixed vocabulary for reporting) ---
STATUS_READY = "Ready for ERP update"
STATUS_NEEDS_CORRECTION = "Needs correction"
STATUS_PHONE_FOLLOWUP = "Needs phone follow-up"
STATUS_DUPLICATE = "Potential duplicate to review"

PRIORITY_HIGH = "High"
PRIORITY_MEDIUM = "Medium"
PRIORITY_LOW = "Low"

ACTION_UPDATE_ERP = "Update ERP"
ACTION_CORRECT_FORMAT = "Correct formatting"
ACTION_REQUEST_FISCAL = "Request missing fiscal information"
ACTION_CALL_CONTACT = "Call billing/admin contact"
ACTION_REVIEW_DUPLICATE = "Review duplicate account"
ACTION_EXCLUDE_INACTIVE = "Exclude inactive account"

# Columns read as text to preserve leading zeros on identifiers
TEXT_COLUMNS = [
    "client_id", "company_name", "siren", "siret", "vat_number",
    "billing_email", "billing_contact_name", "phone", "city",
    "postal_code", "country", "account_status", "last_invoice_date",
    "source_file", "notes",
]


def clean_text(value) -> str:
    """Normalize cell value to stripped string."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def digits_only(value: str) -> str:
    """Keep only digits from an identifier field."""
    return re.sub(r"\D", "", value)


def normalize_company_name_for_matching(name: str) -> str:
    """
    Normalize company name for duplicate similarity:
    lowercase, single spaces, no punctuation, legal suffixes removed.
    """
    name = clean_text(name).lower()
    name = re.sub(r"[^\w\s]", " ", name, flags=re.UNICODE)
    name = re.sub(r"\s+", " ", name).strip()
    for suffix in LEGAL_SUFFIXES:
        name = re.sub(rf"\b{suffix}\b", "", name, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", name).strip()


def name_similarity_percent(name_a: str, name_b: str) -> float:
    """Return similarity score between two normalized names (0–100)."""
    a = normalize_company_name_for_matching(name_a)
    b = normalize_company_name_for_matching(name_b)
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio() * 100


def count_duplicates_legacy(df: pd.DataFrame) -> int:
    """
    Previous duplicate logic (aggressive): same normalized name, SIREN, or SIRET.
    Used only for before/after reporting.
    """
    flags = [False] * len(df)
    simple_norm = df["company_name"].map(
        lambda n: re.sub(r"\s+", " ", clean_text(n).lower())
    )
    name_counts = simple_norm.value_counts()
    dup_names = set(name_counts[name_counts > 1].index) - {""}

    siren_digits = df["siren"].map(digits_only)
    valid_siren = siren_digits.where(df["siren"].map(is_valid_siren), "")
    dup_sirens = set(valid_siren[valid_siren != ""].value_counts().pipe(
        lambda s: s[s > 1].index
    ))

    siret_digits = df["siret"].map(digits_only)
    valid_siret = siret_digits.where(df["siret"].map(is_valid_siret), "")
    dup_sirets = set(valid_siret[valid_siret != ""].value_counts().pipe(
        lambda s: s[s > 1].index
    ))

    for i in range(len(df)):
        if simple_norm.iloc[i] in dup_names:
            flags[i] = True
        elif valid_siren.iloc[i] in dup_sirens:
            flags[i] = True
        elif valid_siret.iloc[i] in dup_sirets:
            flags[i] = True
    return sum(flags)


def mark_potential_duplicates(
    df: pd.DataFrame, issue_lists: list[list[str]]
) -> list[bool]:
    """
    Flag duplicates only when:
    1) SIRET is exactly identical to another record, OR
    2) SIREN is identical AND normalized company names are >90% similar.

    Company name similarity alone never triggers a duplicate flag.
    """
    n = len(df)
    dup_flags = [False] * n

    # Rule 1: exact SIRET match (non-empty, character-identical after trim)
    siret_clean = df["siret"].map(clean_text)
    siret_counts = siret_clean[siret_clean != ""].value_counts()
    dup_siret_values = set(siret_counts[siret_counts > 1].index)

    for i, siret in enumerate(siret_clean):
        if siret in dup_siret_values:
            dup_flags[i] = True
            if "duplicate_siret" not in issue_lists[i]:
                issue_lists[i].append("duplicate_siret")

    # Rule 2: same SIREN + company name similarity > 90%
    siren_clean = df["siren"].map(clean_text)
    siren_groups: dict[str, list[int]] = {}
    for i, siren in enumerate(siren_clean):
        if siren:
            siren_groups.setdefault(siren, []).append(i)

    for siren, indices in siren_groups.items():
        if len(indices) < 2:
            continue
        names = [df.iloc[i]["company_name"] for i in indices]
        for a in range(len(indices)):
            for b in range(a + 1, len(indices)):
                if name_similarity_percent(names[a], names[b]) > NAME_SIMILARITY_THRESHOLD:
                    i_a, i_b = indices[a], indices[b]
                    dup_flags[i_a] = True
                    dup_flags[i_b] = True
                    tag = "duplicate_siren_and_name"
                    if tag not in issue_lists[i_a]:
                        issue_lists[i_a].append(tag)
                    if tag not in issue_lists[i_b]:
                        issue_lists[i_b].append(tag)

    return dup_flags


def is_valid_siren(siren: str) -> bool:
    """SIREN must be exactly 9 digits."""
    d = digits_only(siren)
    return len(d) == 9 and d.isdigit()


def is_valid_siret(siret: str) -> bool:
    """SIRET must be exactly 14 digits."""
    d = digits_only(siret)
    return len(d) == 14 and d.isdigit()


def is_missing_siret(siret: str) -> bool:
    return clean_text(siret) == ""


def is_missing_vat(vat: str) -> bool:
    return clean_text(vat) == ""


def is_valid_vat(vat: str) -> bool:
    """VAT must start with FR and be exactly 13 characters."""
    v = clean_text(vat)
    return v.startswith("FR") and len(v) == 13


def is_valid_email(email: str) -> bool:
    """Email must contain @ and a domain with at least one dot."""
    e = clean_text(email)
    if not e or "@" not in e or " " in e:
        return False
    local, _, domain = e.partition("@")
    if not local or not domain or "." not in domain:
        return False
    if domain.startswith(".") or domain.endswith(".") or ".." in e:
        return False
    # Basic domain shape: label.tld
    domain_parts = domain.split(".")
    if any(not part for part in domain_parts):
        return False
    return True


def has_formatting_issue(company_name: str) -> bool:
    """
    Detect inconsistent capitalization or spacing in company_name.
    Canonical form: stripped, single spaces, title-style not all upper/lower.
    """
    raw = clean_text(company_name)
    if not raw:
        return True
    if raw != company_name.strip():
        return True
    if "  " in company_name:
        return True
    if raw == raw.lower() or raw == raw.upper():
        return True
    # Chaotic mixed case (not typical French company casing)
    letters = [c for c in raw if c.isalpha()]
    if letters:
        ups = sum(1 for c in letters if c.isupper())
        ratio = ups / len(letters)
        if 0.15 < ratio < 0.85 and raw != raw.title():
            return True
    return False


def is_inactive_status(status: str) -> bool:
    return clean_text(status).lower() == "inactif"


def is_active_for_erp(status: str) -> bool:
    """Accounts eligible for ERP load when clean."""
    return clean_text(status).lower() == "actif"


def needs_phone_follow_up(row: pd.Series, issues: list[str]) -> bool:
    """
    Phone is required when outreach is needed (missing/invalid email or contact,
    or fiscal gaps on a non-inactive account).
    """
    contact_gaps = {
        "missing_billing_email", "invalid_billing_email",
        "missing_billing_contact",
    }
    fiscal_gaps = {
        "missing_siret", "invalid_siret", "missing_vat_number",
        "invalid_vat_number", "missing_siren", "invalid_siren",
    }
    if contact_gaps & set(issues):
        return True
    if fiscal_gaps & set(issues) and not is_inactive_status(row["account_status"]):
        return True
    return False


def detect_issues(df: pd.DataFrame) -> pd.DataFrame:
    """Run all quality rules and return issue flags per row."""
    out = df.copy()

    issue_lists: list[list[str]] = []

    for _, row in out.iterrows():
        issues: list[str] = []

        siren = clean_text(row["siren"])
        siret = clean_text(row["siret"])
        vat = clean_text(row["vat_number"])
        email = clean_text(row["billing_email"])
        contact = clean_text(row["billing_contact_name"])
        phone = clean_text(row["phone"])

        # --- SIREN ---
        if not siren:
            issues.append("missing_siren")
        elif not is_valid_siren(siren):
            issues.append("invalid_siren")

        # --- SIRET ---
        if is_missing_siret(siret):
            issues.append("missing_siret")
        elif not is_valid_siret(siret):
            issues.append("invalid_siret")

        # --- VAT ---
        if is_missing_vat(vat):
            issues.append("missing_vat_number")
        elif not is_valid_vat(vat):
            issues.append("invalid_vat_number")

        # --- Email ---
        if not email:
            issues.append("missing_billing_email")
        elif not is_valid_email(email):
            issues.append("invalid_billing_email")

        # --- Contact & phone ---
        if not contact:
            issues.append("missing_billing_contact")
        if not phone:
            issues.append("missing_phone")

        # --- Formatting ---
        if has_formatting_issue(row["company_name"]):
            issues.append("formatting_company_name")

        issue_lists.append(issues)

    out["issue_list"] = issue_lists

    # --- Duplicate detection (strict rules) ---
    dup_flags = mark_potential_duplicates(out, issue_lists)
    out["is_duplicate"] = dup_flags
    out["issue_list"] = issue_lists
    out["data_issues"] = out["issue_list"].apply(
        lambda lst: "; ".join(sorted(set(lst))) if lst else ""
    )

    return out


def has_fiscal_issue(issues: list[str]) -> bool:
    fiscal = {
        "missing_siren", "invalid_siren", "missing_siret", "invalid_siret",
        "missing_vat_number", "invalid_vat_number",
    }
    return bool(fiscal & set(issues))


def has_contact_issue(issues: list[str]) -> bool:
    return bool(
        {"missing_billing_email", "invalid_billing_email", "missing_billing_contact"}
        & set(issues)
    )


def has_only_formatting_issues(issues: list[str]) -> bool:
    """True when formatting is the only problem (no fiscal/contact/duplicate)."""
    if not issues:
        return False
    allowed = {"formatting_company_name"}
    return set(issues).issubset(allowed)


def qualify_row(row: pd.Series) -> tuple[str, str, str, str]:
    """
    Apply business priority rules to set qualification_status, priority,
    and action_to_take for one client row.
    Returns (qualification_status, priority, action_to_take, data_issues unchanged).
    """
    issues = row["issue_list"]
    duplicate = row["is_duplicate"]

    # 1) Inactive accounts — exclude from ERP scope
    if is_inactive_status(row["account_status"]):
        return (
            STATUS_NEEDS_CORRECTION,
            PRIORITY_LOW,
            ACTION_EXCLUDE_INACTIVE,
        )

    # 2) Duplicates — high priority manual review
    if duplicate:
        return (
            STATUS_DUPLICATE,
            PRIORITY_HIGH,
            ACTION_REVIEW_DUPLICATE,
        )

    # 3) Fiscal identifiers — correction before invoicing reform
    if has_fiscal_issue(issues):
        priority = PRIORITY_HIGH if (
            "missing_siret" in issues or "missing_vat_number" in issues
        ) else PRIORITY_MEDIUM
        return (
            STATUS_NEEDS_CORRECTION,
            priority,
            ACTION_REQUEST_FISCAL,
        )

    # 4) Missing contact channels — phone outreach
    if has_contact_issue(issues):
        status = STATUS_PHONE_FOLLOWUP
        priority = PRIORITY_HIGH if "missing_phone" in issues else PRIORITY_MEDIUM
        return (
            status,
            priority,
            ACTION_CALL_CONTACT,
        )

    # 5) Phone required but empty when follow-up is needed
    if needs_phone_follow_up(row, issues) and "missing_phone" in issues:
        return (
            STATUS_PHONE_FOLLOWUP,
            PRIORITY_HIGH,
            ACTION_CALL_CONTACT,
        )

    # 6) Formatting-only cleanup
    if has_only_formatting_issues(issues):
        return (
            STATUS_NEEDS_CORRECTION,
            PRIORITY_MEDIUM,
            ACTION_CORRECT_FORMAT,
        )

    # 7) Clean active accounts — ready for ERP
    if is_active_for_erp(row["account_status"]) and not issues:
        return (
            STATUS_READY,
            PRIORITY_LOW,
            ACTION_UPDATE_ERP,
        )

    # Default: minor gaps on non-actif accounts (En attente, Suspendu)
    if issues:
        return (
            STATUS_NEEDS_CORRECTION,
            PRIORITY_MEDIUM,
            ACTION_CORRECT_FORMAT,
        )

    return (
        STATUS_NEEDS_CORRECTION,
        PRIORITY_LOW,
        ACTION_CORRECT_FORMAT,
    )


def apply_qualification(df: pd.DataFrame) -> pd.DataFrame:
    """Add qualification_status, priority, and action_to_take columns."""
    qualified = df.copy()
    results = qualified.apply(qualify_row, axis=1, result_type="expand")
    qualified["qualification_status"] = results[0]
    qualified["priority"] = results[1]
    qualified["action_to_take"] = results[2]
    return qualified


def build_summary_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Build the main summary table requested for management reporting."""
    issue_set = df["issue_list"]

    rows = [
        ("total_clients", len(df)),
        ("ready_for_erp_update", (df["qualification_status"] == STATUS_READY).sum()),
        ("needs_correction", (df["qualification_status"] == STATUS_NEEDS_CORRECTION).sum()),
        ("needs_phone_follow_up", (df["qualification_status"] == STATUS_PHONE_FOLLOWUP).sum()),
        ("potential_duplicates", (df["qualification_status"] == STATUS_DUPLICATE).sum()),
        ("inactive_accounts", df["account_status"].map(is_inactive_status).sum()),
        ("missing_siret", issue_set.map(lambda x: "missing_siret" in x).sum()),
        ("invalid_siret", issue_set.map(lambda x: "invalid_siret" in x).sum()),
        ("missing_vat_number", issue_set.map(lambda x: "missing_vat_number" in x).sum()),
        ("invalid_vat_number", issue_set.map(lambda x: "invalid_vat_number" in x).sum()),
        ("missing_billing_email", issue_set.map(lambda x: "missing_billing_email" in x).sum()),
        ("invalid_billing_email", issue_set.map(lambda x: "invalid_billing_email" in x).sum()),
        ("missing_billing_contact", issue_set.map(lambda x: "missing_billing_contact" in x).sum()),
        ("missing_phone", issue_set.map(lambda x: "missing_phone" in x).sum()),
    ]

    return pd.DataFrame(rows, columns=["metric", "count"])


def build_priority_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Counts by priority level for the 'by_priority' sheet."""
    summary = (
        df.groupby("priority", dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values(
            "priority",
            key=lambda s: s.map({PRIORITY_HIGH: 0, PRIORITY_MEDIUM: 1, PRIORITY_LOW: 2}),
        )
    )
    return summary


def export_qualified_excel(df: pd.DataFrame, path: Path) -> None:
    """Write qualified clients; keep identifier columns as text in Excel."""
    export_cols = TEXT_COLUMNS + [
        "data_issues", "qualification_status", "priority", "action_to_take",
    ]
    export_df = df[export_cols].copy()

    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="qualified_clients")
        ws = writer.sheets["qualified_clients"]
        text_col_idx = {
            "siren": 3, "siret": 4, "vat_number": 5,
            "postal_code": 10, "phone": 8,
        }
        for col_name, col_idx in text_col_idx.items():
            for row in range(2, len(export_df) + 2):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value not in (None, ""):
                    cell.value = str(cell.value)
                    cell.number_format = "@"


def export_summary_excel(summary: pd.DataFrame, by_priority: pd.DataFrame, path: Path) -> None:
    """Write summary metrics and by_priority breakdown."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="summary")
        by_priority.to_excel(writer, index=False, sheet_name="by_priority")


def load_raw_data(path: Path) -> pd.DataFrame:
    """Load raw Excel file with string dtypes for identifiers."""
    df = pd.read_excel(path, dtype=str)
    df = df.fillna("")
    for col in TEXT_COLUMNS:
        if col not in df.columns:
            raise ValueError(f"Missing expected column: {col}")
        df[col] = df[col].map(clean_text)
    return df


def main() -> None:
    print(f"Reading: {INPUT_FILE}")
    df = load_raw_data(INPUT_FILE)

    # Compare duplicate detection before applying new rules
    duplicates_before = count_duplicates_legacy(df)

    # Quality assessment
    assessed = detect_issues(df)
    duplicates_after = int(assessed["is_duplicate"].sum())
    qualified = apply_qualification(assessed)

    # Drop internal helper columns from export
    helper_cols = ["issue_list", "is_duplicate"]
    export_base = qualified.drop(columns=helper_cols, errors="ignore")

    summary = build_summary_metrics(qualified)
    by_priority = build_priority_summary(qualified)

    export_qualified_excel(export_base, OUTPUT_QUALIFIED)
    export_summary_excel(summary, by_priority, OUTPUT_SUMMARY)

    # --- Duplicate detection impact ---
    if duplicates_before > 0:
        pct_reduction = (duplicates_before - duplicates_after) / duplicates_before * 100
    else:
        pct_reduction = 0.0

    print("\n--- Duplicate detection ---")
    print(f"Duplicates before:  {duplicates_before}")
    print(f"Duplicates after:   {duplicates_after}")
    print(f"Reduction:          {pct_reduction:.1f}%")

    # --- Console summary ---
    print("\n--- Quality check summary ---")
    print(f"Total clients:           {len(qualified)}")
    print(f"Ready for ERP update:    {(qualified['qualification_status'] == STATUS_READY).sum()}")
    print(f"Needs correction:        {(qualified['qualification_status'] == STATUS_NEEDS_CORRECTION).sum()}")
    print(f"Needs phone follow-up:   {(qualified['qualification_status'] == STATUS_PHONE_FOLLOWUP).sum()}")
    print(f"Potential duplicates:    {(qualified['qualification_status'] == STATUS_DUPLICATE).sum()}")
    print(f"\nBy priority:")
    for _, row in by_priority.iterrows():
        print(f"  {row['priority']}: {row['count']}")
    print(f"\nOutput files:")
    print(f"  {OUTPUT_QUALIFIED}")
    print(f"  {OUTPUT_SUMMARY}")


if __name__ == "__main__":
    main()
