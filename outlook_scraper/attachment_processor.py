"""
Attachment Processor Module

Processes email attachments to extract contact information.
Supports CSV, Excel, vCard, and PDF files.
"""

import io
import re
from typing import List, Optional

import pandas as pd
import pdfplumber
import vobject
from rich.console import Console

from .contact_extractor import Contact
from .email_fetcher import Attachment

console = Console()


class AttachmentProcessor:
    """Processes various attachment types to extract contacts."""

    # Common column names that might contain contact info
    EMAIL_COLUMNS = ['email', 'e-mail', 'email address', 'emailaddress', 'mail', 'email_address']
    NAME_COLUMNS = ['name', 'full name', 'fullname', 'contact name', 'contactname', 'first name', 'firstname', 'full_name']
    FIRST_NAME_COLUMNS = ['first name', 'firstname', 'first_name', 'fname', 'given name']
    LAST_NAME_COLUMNS = ['last name', 'lastname', 'last_name', 'lname', 'surname', 'family name']
    TITLE_COLUMNS = ['title', 'job title', 'jobtitle', 'position', 'role', 'job_title']
    COMPANY_COLUMNS = ['company', 'organization', 'org', 'employer', 'company name', 'companyname', 'organisation']
    PHONE_COLUMNS = ['phone', 'telephone', 'tel', 'mobile', 'cell', 'phone number', 'phonenumber', 'phone_number']

    def __init__(self):
        """Initialize the attachment processor."""
        self.email_pattern = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')

    def process_attachment(self, attachment: Attachment) -> List[Contact]:
        """
        Process an attachment and extract contacts.

        Args:
            attachment: Attachment object to process

        Returns:
            List of extracted Contact objects
        """
        content_type = attachment.content_type.lower()
        name_lower = attachment.name.lower()

        try:
            # CSV files
            if 'csv' in content_type or name_lower.endswith('.csv'):
                return self._process_csv(attachment)

            # Excel files
            elif 'excel' in content_type or 'spreadsheet' in content_type or \
                    name_lower.endswith(('.xlsx', '.xls', '.xlsm')):
                return self._process_excel(attachment)

            # vCard files
            elif 'vcard' in content_type or name_lower.endswith(('.vcf', '.vcard')):
                return self._process_vcard(attachment)

            # PDF files
            elif 'pdf' in content_type or name_lower.endswith('.pdf'):
                return self._process_pdf(attachment)

            # Text files (might contain contact lists)
            elif 'text' in content_type or name_lower.endswith('.txt'):
                return self._process_text(attachment)

        except Exception as e:
            console.print(f"[yellow]Warning: Could not process attachment '{attachment.name}': {e}[/yellow]")

        return []

    def _find_column(self, columns: List[str], possible_names: List[str]) -> Optional[str]:
        """Find a column matching one of the possible names."""
        columns_lower = {c.lower().strip(): c for c in columns}
        for name in possible_names:
            if name in columns_lower:
                return columns_lower[name]
        return None

    def _process_dataframe(self, df: pd.DataFrame, source: str) -> List[Contact]:
        """
        Extract contacts from a pandas DataFrame.

        Args:
            df: DataFrame to process
            source: Source identifier for the contacts

        Returns:
            List of Contact objects
        """
        contacts = []

        if df.empty:
            return contacts

        # Normalize column names
        df.columns = [str(c).strip() for c in df.columns]

        # Find relevant columns
        email_col = self._find_column(df.columns.tolist(), self.EMAIL_COLUMNS)
        name_col = self._find_column(df.columns.tolist(), self.NAME_COLUMNS)
        first_name_col = self._find_column(df.columns.tolist(), self.FIRST_NAME_COLUMNS)
        last_name_col = self._find_column(df.columns.tolist(), self.LAST_NAME_COLUMNS)
        title_col = self._find_column(df.columns.tolist(), self.TITLE_COLUMNS)
        company_col = self._find_column(df.columns.tolist(), self.COMPANY_COLUMNS)
        phone_col = self._find_column(df.columns.tolist(), self.PHONE_COLUMNS)

        # If no email column found, search all columns for email patterns
        if not email_col:
            for col in df.columns:
                sample = df[col].dropna().astype(str).head(10)
                if sample.str.contains(self.email_pattern).any():
                    email_col = col
                    break

        if not email_col:
            # Try to extract emails from all text
            all_text = df.to_string()
            emails = self.email_pattern.findall(all_text)
            for email in set(emails):
                contacts.append(Contact(
                    email=email.lower(),
                    source=source,
                ))
            return contacts

        # Process each row
        for _, row in df.iterrows():
            email_value = row.get(email_col)
            if pd.isna(email_value):
                continue

            email_str = str(email_value).strip().lower()

            # Validate email
            if not self.email_pattern.match(email_str):
                # Try to extract email from the value
                match = self.email_pattern.search(email_str)
                if match:
                    email_str = match.group(0).lower()
                else:
                    continue

            # Build name
            name = None
            if name_col and not pd.isna(row.get(name_col)):
                name = str(row[name_col]).strip()
            elif first_name_col or last_name_col:
                parts = []
                if first_name_col and not pd.isna(row.get(first_name_col)):
                    parts.append(str(row[first_name_col]).strip())
                if last_name_col and not pd.isna(row.get(last_name_col)):
                    parts.append(str(row[last_name_col]).strip())
                if parts:
                    name = ' '.join(parts)

            # Get other fields
            title = None
            if title_col and not pd.isna(row.get(title_col)):
                title = str(row[title_col]).strip()

            company = None
            if company_col and not pd.isna(row.get(company_col)):
                company = str(row[company_col]).strip()

            phone = None
            if phone_col and not pd.isna(row.get(phone_col)):
                phone = str(row[phone_col]).strip()

            contact = Contact(
                email=email_str,
                name=name,
                title=title,
                company=company,
                phone=phone,
                source=source,
            )
            contacts.append(contact)

        return contacts

    def _process_csv(self, attachment: Attachment) -> List[Contact]:
        """Process a CSV file."""
        content = attachment.content.decode('utf-8', errors='ignore')

        # Try different delimiters
        for delimiter in [',', ';', '\t', '|']:
            try:
                df = pd.read_csv(io.StringIO(content), delimiter=delimiter)
                if len(df.columns) > 1:
                    contacts = self._process_dataframe(df, f"csv:{attachment.name}")
                    if contacts:
                        return contacts
            except Exception:
                continue

        return []

    def _process_excel(self, attachment: Attachment) -> List[Contact]:
        """Process an Excel file."""
        contacts = []

        try:
            # Read all sheets
            excel_file = pd.ExcelFile(io.BytesIO(attachment.content))

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                sheet_contacts = self._process_dataframe(df, f"excel:{attachment.name}:{sheet_name}")
                contacts.extend(sheet_contacts)

        except Exception as e:
            console.print(f"[yellow]Warning: Could not read Excel file '{attachment.name}': {e}[/yellow]")

        return contacts

    def _process_vcard(self, attachment: Attachment) -> List[Contact]:
        """Process a vCard file."""
        contacts = []

        try:
            content = attachment.content.decode('utf-8', errors='ignore')

            # Parse vCard(s)
            for vcard in vobject.readComponents(content):
                email = None
                name = None
                title = None
                company = None
                phone = None

                # Extract email
                if hasattr(vcard, 'email'):
                    email = str(vcard.email.value).lower()

                # Extract name
                if hasattr(vcard, 'fn'):
                    name = str(vcard.fn.value)
                elif hasattr(vcard, 'n'):
                    n = vcard.n.value
                    name_parts = [n.given, n.family]
                    name = ' '.join(p for p in name_parts if p)

                # Extract title
                if hasattr(vcard, 'title'):
                    title = str(vcard.title.value)

                # Extract organization
                if hasattr(vcard, 'org'):
                    org = vcard.org.value
                    if isinstance(org, list):
                        company = org[0] if org else None
                    else:
                        company = str(org)

                # Extract phone
                if hasattr(vcard, 'tel'):
                    phone = str(vcard.tel.value)

                if email:
                    contact = Contact(
                        email=email,
                        name=name,
                        title=title,
                        company=company,
                        phone=phone,
                        source=f"vcard:{attachment.name}",
                    )
                    contacts.append(contact)

        except Exception as e:
            console.print(f"[yellow]Warning: Could not parse vCard '{attachment.name}': {e}[/yellow]")

        return contacts

    def _process_pdf(self, attachment: Attachment) -> List[Contact]:
        """Process a PDF file to extract email addresses."""
        contacts = []

        try:
            with pdfplumber.open(io.BytesIO(attachment.content)) as pdf:
                all_text = ""
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        all_text += text + "\n"

                    # Also try to extract tables
                    tables = page.extract_tables()
                    for table in tables:
                        if table:
                            try:
                                df = pd.DataFrame(table[1:], columns=table[0] if table[0] else None)
                                table_contacts = self._process_dataframe(df, f"pdf:{attachment.name}")
                                contacts.extend(table_contacts)
                            except Exception:
                                pass

                # Extract emails from text
                emails = set(self.email_pattern.findall(all_text))
                for email in emails:
                    # Check if we already have this email from table extraction
                    if not any(c.email == email.lower() for c in contacts):
                        contacts.append(Contact(
                            email=email.lower(),
                            source=f"pdf:{attachment.name}",
                        ))

        except Exception as e:
            console.print(f"[yellow]Warning: Could not process PDF '{attachment.name}': {e}[/yellow]")

        return contacts

    def _process_text(self, attachment: Attachment) -> List[Contact]:
        """Process a text file to extract email addresses."""
        contacts = []

        try:
            content = attachment.content.decode('utf-8', errors='ignore')
            emails = set(self.email_pattern.findall(content))

            for email in emails:
                contacts.append(Contact(
                    email=email.lower(),
                    source=f"text:{attachment.name}",
                ))

        except Exception as e:
            console.print(f"[yellow]Warning: Could not process text file '{attachment.name}': {e}[/yellow]")

        return contacts
