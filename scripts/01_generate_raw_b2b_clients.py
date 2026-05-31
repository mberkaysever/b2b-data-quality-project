#!/usr/bin/env python3
"""
Generate a synthetic B2B customer dataset for data quality exercises.

This script creates fictional French B2B accounts with intentional defects
(missing SIRET, invalid VAT, duplicates, etc.) for electronic invoicing
reform preparation workflows. Not real Europcar data.
"""

from __future__ import annotations

import random
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# Reproducible generation
RANDOM_SEED = 42
N_ROWS = 1000

# Project paths (script lives in scripts/)
PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_XLSX = DATA_DIR / "b2b_clients_raw.xlsx"
OUTPUT_CSV = DATA_DIR / "b2b_clients_raw.csv"

# --- Reference data (synthetic French B2B context) ---

COMPANY_PREFIXES = [
    "Transports", "Logistique", "Services", "Industrie", "Groupe",
    "Solutions", "Maintenance", "Distribution", "Équipements", "Négoce",
]
COMPANY_CORES = [
    "Dupont", "Martin", "Bernard", "Moreau", "Laurent", "Simon",
    "Lefebvre", "Roux", "Girard", "Mercier", "Blanc", "Garnier",
    "Faure", "Perrin", "Marchand", "Renault", "Dubois", "Lambert",
]
COMPANY_SUFFIXES = ["SARL", "SAS", "SA", "EURL", "SASU", "& Fils", "France"]

CITIES = [
    ("Paris", "75001"), ("Lyon", "69001"), ("Marseille", "13001"),
    ("Toulouse", "31000"), ("Nice", "06000"), ("Nantes", "44000"),
    ("Strasbourg", "67000"), ("Bordeaux", "33000"), ("Lille", "59000"),
    ("Rennes", "35000"), ("Montpellier", "34000"), ("Grenoble", "38000"),
    ("Reims", "51100"), ("Le Havre", "76600"), ("Saint-Étienne", "42000"),
]

FIRST_NAMES = [
    "Jean", "Marie", "Pierre", "Sophie", "Nicolas", "Isabelle",
    "Philippe", "Catherine", "Thomas", "Anne", "François", "Julie",
]
LAST_NAMES = [
    "Durand", "Petit", "Robert", "Richard", "Michel", "Leroy",
    "Moreau", "Fournier", "Giraud", "Bonnet", "Dupuis", "Lemaire",
]

ACCOUNT_STATUSES = ["Actif", "Actif", "Actif", "Inactif", "En attente", "Suspendu"]
SOURCE_FILES = [
    "import_crm_2024_Q1.xlsx",
    "legacy_erp_export.csv",
    "manual_onboarding_2023.xlsx",
    "partner_feed_b2b.csv",
    "migration_batch_12.xlsx",
]
NOTES_SAMPLES = [
    "", "", "",
    "Client prioritaire - relance facturation",
    "Migration ERP en cours",
    "Doublon suspecté",
    "SIRET à vérifier",
    "Email facturation manquant",
    "Compte créé avant réforme facturation électronique",
]


def luhn_checksum(digits: list[int]) -> int:
    """French SIREN/SIRET uses Luhn algorithm on the 9-digit SIREN."""
    total = 0
    for i, d in enumerate(reversed(digits)):
        if i % 2 == 1:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return (10 - (total % 10)) % 10


def generate_valid_siren(rng: random.Random) -> str:
    """Generate a syntactically valid 9-digit SIREN."""
    base = [rng.randint(0, 9) for _ in range(8)]
    base.append(luhn_checksum(base))
    return "".join(str(d) for d in base)


def generate_valid_siret(rng: random.Random, siren: str | None = None) -> str:
    """SIRET = 9-digit SIREN + 5-digit NIC (establishment)."""
    if siren is None:
        siren = generate_valid_siren(rng)
    nic = f"{rng.randint(0, 99999):05d}"
    return siren + nic


def compute_vat_key(siren: str) -> str:
    """Compute the 2-digit VAT key for FR format: key = (12 + 3 * (SIREN mod 97)) mod 97."""
    siren_int = int(siren)
    key = (12 + 3 * (siren_int % 97)) % 97
    return f"{key:02d}"


def format_valid_vat(siren: str) -> str:
    """Standard French VAT: FR + 2 key digits + 9 SIREN digits."""
    return f"FR{compute_vat_key(siren)}{siren}"


def base_company_name(rng: random.Random) -> str:
    """Clean canonical company name."""
    parts = [
        rng.choice(COMPANY_PREFIXES),
        rng.choice(COMPANY_CORES),
        rng.choice(COMPANY_SUFFIXES),
    ]
    return " ".join(parts)


def apply_name_formatting_issue(rng: random.Random, name: str, issue: str) -> str:
    """Introduce inconsistent capitalization or spacing."""
    if issue == "lower":
        return name.lower()
    if issue == "upper":
        return name.upper()
    if issue == "extra_spaces":
        return re.sub(r"\s+", "  ", name)
    if issue == "leading_trailing":
        return f"  {name}  "
    if issue == "mixed":
        return "".join(
            c.upper() if rng.random() > 0.5 else c.lower() for c in name
        )
    return name


def generate_phone(rng: random.Random) -> str:
    """French mobile or landline style (synthetic)."""
    if rng.random() < 0.6:
        # Mobile 06/07
        prefix = rng.choice(["06", "07"])
        return f"+33 {prefix[1]} " + " ".join(
            f"{rng.randint(10, 99)}" for _ in range(4)
        )
    # Landline with area code
    codes = ["1", "4", "5", "9"]  # simplified Paris/regions
    return f"+33 {rng.choice(codes)} " + " ".join(
        f"{rng.randint(10, 99)}" for _ in range(4)
    )


def generate_email(company_slug: str, rng: random.Random) -> str:
    domains = ["gmail.com", "orange.fr", "laposte.net", "entreprise.fr", "societe.com"]
    return f"facturation.{company_slug}@{rng.choice(domains)}"


def slugify(name: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "", name.lower())[:20]
    return slug or "client"


def random_invoice_date(rng: random.Random) -> str:
    start = datetime(2022, 1, 1)
    end = datetime(2025, 5, 1)
    delta = (end - start).days
    day = start + timedelta(days=rng.randint(0, delta))
    return day.strftime("%Y-%m-%d")


def build_duplicate_pool(rng: random.Random, n_duplicates: int) -> list[dict]:
    """
    Pre-build a pool of 'same company' records that will appear multiple times
    with different client_id and slight field variations.
    """
    pool = []
    for _ in range(n_duplicates):
        siren = generate_valid_siren(rng)
        siret = generate_valid_siret(rng, siren)
        canonical_name = base_company_name(rng)
        city, postal = rng.choice(CITIES)
        pool.append({
            "canonical_name": canonical_name,
            "siren": siren,
            "siret": siret,
            "vat_number": format_valid_vat(siren),
            "city": city,
            "postal_code": postal,
        })
    return pool


def count_duplicated_company_names(df: pd.DataFrame) -> int:
    """
    Count rows whose company_name appears more than once (after trim + lower case).
    Formatting variants of the same legal name are treated as duplicates.
    """
    key = df["company_name"].str.strip().str.lower()
    return int(key.duplicated(keep=False).sum())


def main() -> None:
    rng = random.Random(RANDOM_SEED)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    # Intentional defect rates (approximate, driven by RANDOM_SEED):
    # - missing / invalid SIRET, VAT, email
    # - missing contact, phone
    # - duplicate accounts (~120 rows from shared company pool)
    # - inconsistent company_name capitalization or spacing (~10%)
    duplicate_pool = build_duplicate_pool(rng, n_duplicates=45)
    # Map row indices that should be duplicates
    duplicate_row_indices = set(rng.sample(range(N_ROWS), k=120))
    dup_iter = iter(duplicate_pool)

    rows: list[dict] = []

    for i in range(N_ROWS):
        client_id = f"B2B-{rng.randint(10000, 99999):05d}-{i + 1:04d}"

        # --- Base entity ---
        if i in duplicate_row_indices:
            try:
                dup_base = next(dup_iter)
            except StopIteration:
                dup_base = duplicate_pool[rng.randint(0, len(duplicate_pool) - 1)]
            siren = dup_base["siren"]
            canonical_name = dup_base["canonical_name"]
            city = dup_base["city"]
            postal_code = dup_base["postal_code"]
            # Duplicate accounts may share SIREN but differ on SIRET/email
            siret_base = dup_base["siret"]
            vat_base = dup_base["vat_number"]
        else:
            siren = generate_valid_siren(rng)
            canonical_name = base_company_name(rng)
            city, postal_code = rng.choice(CITIES)
            siret_base = generate_valid_siret(rng, siren)
            vat_base = format_valid_vat(siren)

        company_name = canonical_name
        siret = siret_base
        vat_number = vat_base
        billing_email = generate_email(slugify(canonical_name), rng)
        billing_contact_name = f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
        phone = generate_phone(rng)

        # --- Inject data quality issues (probabilistic + guaranteed minimums) ---
        issue_roll = rng.random()

        # Missing SIRET (~8%)
        if issue_roll < 0.08:
            siret = ""
        # Invalid SIRET length (~5%)
        elif issue_roll < 0.13:
            bad_lengths = [9, 11, 13, 15, 20]
            siret = siret[: rng.choice(bad_lengths)]
        # Wrong SIRET (NIC only) (~2%)
        elif issue_roll < 0.15:
            siret = siret[-5:]

        # Missing VAT (~7%)
        vat_roll = rng.random()
        if vat_roll < 0.07:
            vat_number = ""
        # Invalid VAT format (~6%)
        elif vat_roll < 0.13:
            invalid_patterns = [
                f"FR{siren}",  # missing key
                f"FRXX{siren}",  # non-numeric key
                f"DE{compute_vat_key(siren)}{siren}",  # wrong country
                f"FR{compute_vat_key(siren)}{siren[:5]}",  # truncated
                "INVALID",
                f"fr{compute_vat_key(siren)}{siren}".lower(),  # lowercase
            ]
            vat_number = rng.choice(invalid_patterns)

        # Missing billing email (~6%)
        email_roll = rng.random()
        if email_roll < 0.06:
            billing_email = ""
        # Invalid email format (~5%)
        elif email_roll < 0.11:
            invalid_emails = [
                "facturation@",
                "facturation.societe",
                "@domaine.fr",
                "facturation @entreprise.fr",
                "facturation..dupont@mail.fr",
                "facturation@.fr",
            ]
            billing_email = rng.choice(invalid_emails)

        # Missing billing contact name (~4%)
        if rng.random() < 0.04:
            billing_contact_name = ""

        # Missing phone (~5%)
        if rng.random() < 0.05:
            phone = ""

        # Company name formatting issues (~10%)
        if rng.random() < 0.10:
            fmt_issue = rng.choice(
                ["lower", "upper", "extra_spaces", "leading_trailing", "mixed"]
            )
            company_name = apply_name_formatting_issue(rng, canonical_name, fmt_issue)

        # For duplicates: vary email/phone/contact slightly to mimic real CRM noise
        if i in duplicate_row_indices and rng.random() < 0.5:
            if billing_email and rng.random() < 0.4:
                billing_email = billing_email.replace(
                    "facturation", rng.choice(["compta", "billing", "facture"])
                )
            if rng.random() < 0.3:
                client_id = f"B2B-{rng.randint(10000, 99999):05d}-{i + 1:04d}"

        rows.append({
            "client_id": client_id,
            "company_name": company_name,
            "siren": siren,
            "siret": siret,
            "vat_number": vat_number,
            "billing_email": billing_email,
            "billing_contact_name": billing_contact_name,
            "phone": phone,
            "city": city,
            "postal_code": postal_code,
            "country": "FR",
            "account_status": rng.choice(ACCOUNT_STATUSES),
            "last_invoice_date": random_invoice_date(rng),
            "source_file": rng.choice(SOURCE_FILES),
            "notes": rng.choice(NOTES_SAMPLES),
        })

    df = pd.DataFrame(rows)

    # Keep French identifiers as text (avoid Excel/CSV scientific notation)
    for col in ("siren", "siret", "vat_number", "postal_code", "phone"):
        df[col] = df[col].fillna("").astype(str).replace("nan", "")
    df["postal_code"] = df["postal_code"].apply(
        lambda x: x.zfill(5) if x.isdigit() else x
    )

    # Column order as specified
    columns = [
        "client_id", "company_name", "siren", "siret", "vat_number",
        "billing_email", "billing_contact_name", "phone", "city",
        "postal_code", "country", "account_status", "last_invoice_date",
        "source_file", "notes",
    ]
    df = df[columns]

    # Persist outputs (string dtypes preserved for SIRET/SIREN)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig", na_rep="")
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="clients")
        # Force text format on identifier columns in Excel
        ws = writer.sheets["clients"]
        id_cols = {"siren": 3, "siret": 4, "vat_number": 5, "postal_code": 10, "phone": 8}
        for _name, col_idx in id_cols.items():
            for row in range(2, len(df) + 2):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value not in (None, ""):
                    cell.value = str(cell.value)
                    cell.number_format = "@"

    # --- Summary (portfolio verification) ---
    print("\n--- Generation summary ---")
    print(f"Rows created:              {len(df)}")
    print(f"Missing SIRET:             {(df['siret'] == '').sum()}")
    print(f"Missing VAT numbers:       {(df['vat_number'] == '').sum()}")
    print(f"Missing billing emails:    {(df['billing_email'] == '').sum()}")
    print(f"Duplicated company names:  {count_duplicated_company_names(df)}")
    print(f"\nOutput files:")
    print(f"  {OUTPUT_CSV}")
    print(f"  {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
