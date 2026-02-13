#!/usr/bin/env python3
"""
Convert scraped contacts CSV to Apollo.io import format.

Usage:
    python3 convert_to_apollo.py input.csv output_apollo.csv
"""

import csv
import sys
from pathlib import Path


def split_name(full_name: str) -> tuple[str, str]:
    """Split full name into first and last name."""
    if not full_name or full_name.strip() == "-":
        return "", ""

    parts = full_name.strip().split()
    if len(parts) == 0:
        return "", ""
    elif len(parts) == 1:
        return parts[0], ""
    else:
        # First word is first name, rest is last name
        return parts[0], " ".join(parts[1:])


def clean_value(val: str) -> str:
    """Clean up placeholder values."""
    if val in ["-", "None", "none", ""]:
        return ""
    return val.strip()


def convert_to_apollo(input_file: str, output_file: str) -> int:
    """
    Convert scraped contacts CSV to Apollo format.

    Returns number of contacts converted.
    """
    input_path = Path(input_file)
    output_path = Path(output_file)

    if not input_path.exists():
        print(f"Error: Input file '{input_file}' not found")
        sys.exit(1)

    # Apollo column mapping
    apollo_fields = [
        "First Name",
        "Last Name",
        "Email",
        "Title",
        "Company",
        "Phone",
        "LinkedIn URL",  # Empty for now, Apollo can enrich
    ]

    converted = 0
    skipped = 0
    seen_emails = set()

    with open(input_path, 'r', encoding='utf-8') as infile:
        reader = csv.DictReader(infile)

        with open(output_path, 'w', newline='', encoding='utf-8') as outfile:
            writer = csv.DictWriter(outfile, fieldnames=apollo_fields)
            writer.writeheader()

            for row in reader:
                email = clean_value(row.get('email', ''))

                # Skip rows without email or duplicates
                if not email or '@' not in email:
                    skipped += 1
                    continue

                email_lower = email.lower()
                if email_lower in seen_emails:
                    skipped += 1
                    continue
                seen_emails.add(email_lower)

                # Skip obvious non-contacts
                skip_patterns = ['noreply', 'no-reply', 'donotreply', 'mailer-daemon',
                               'postmaster', 'unsubscribe', 'notifications', 'support@',
                               'info@', 'hello@', 'team@', 'news@', 'newsletter@']
                if any(pattern in email_lower for pattern in skip_patterns):
                    skipped += 1
                    continue

                # Parse name
                full_name = clean_value(row.get('name', ''))
                first_name, last_name = split_name(full_name)

                # Build Apollo row
                apollo_row = {
                    "First Name": first_name,
                    "Last Name": last_name,
                    "Email": email,
                    "Title": clean_value(row.get('title', '')),
                    "Company": clean_value(row.get('company', '')),
                    "Phone": clean_value(row.get('phone', '')),
                    "LinkedIn URL": "",  # Apollo will enrich this
                }

                writer.writerow(apollo_row)
                converted += 1

    print(f"Converted: {converted} contacts")
    print(f"Skipped: {skipped} (duplicates, no email, or generic addresses)")
    print(f"Output saved to: {output_path}")

    return converted


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python3 convert_to_apollo.py <input.csv> <output_apollo.csv>")
        print("\nExample:")
        print("  python3 convert_to_apollo.py output/contacts.csv output/apollo_import.csv")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    convert_to_apollo(input_file, output_file)
