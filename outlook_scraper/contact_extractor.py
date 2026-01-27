"""
Contact Extraction Module

Extracts contact information (names, emails, titles) from emails.
Parses email signatures and headers for job titles.
"""

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Set

from .email_fetcher import Email, EmailAddress


@dataclass
class Contact:
    """Represents an extracted contact with all available information."""
    email: str
    name: Optional[str] = None
    title: Optional[str] = None
    company: Optional[str] = None
    phone: Optional[str] = None
    last_contact_date: Optional[datetime] = None
    contact_count: int = 0
    source: str = "email"  # email, attachment, signature

    def __hash__(self):
        return hash(self.email.lower())

    def __eq__(self, other):
        if isinstance(other, Contact):
            return self.email.lower() == other.email.lower()
        return False


class ContactExtractor:
    """Extracts contact information from emails."""

    # Common job title patterns
    TITLE_PATTERNS = [
        # C-level
        r'\b(CEO|CFO|CTO|COO|CMO|CIO|CISO|Chief\s+\w+\s+Officer)\b',
        # VP/Director level
        r'\b(Vice\s+President|VP|Director|Head\s+of|Managing\s+Director)\s+(?:of\s+)?[\w\s&]+',
        # Manager level
        r'\b(Senior|Jr\.?|Junior|Lead|Principal|Staff)?\s*(Manager|Engineer|Developer|Designer|Analyst|Consultant|Architect|Administrator|Coordinator|Specialist|Executive|Associate|Representative)\b',
        # Specific roles
        r'\b(Software|Hardware|Sales|Marketing|Product|Project|Program|Account|Business|Data|DevOps|Cloud|Security|Network|Systems|IT|HR|Human\s+Resources|Finance|Legal|Operations)\s+(Manager|Engineer|Developer|Designer|Analyst|Consultant|Architect|Administrator|Coordinator|Specialist|Director|Lead)\b',
        # Academic/Medical
        r'\b(Professor|Dr\.|Doctor|MD|PhD|Researcher|Scientist)\b',
        # Other common titles
        r'\b(Founder|Co-Founder|Owner|Partner|President|Chairman|Board\s+Member)\b',
    ]

    # Company indicators in signatures
    COMPANY_PATTERNS = [
        r'(?:^|\n)([A-Z][\w\s&,\.]+(?:Inc\.?|LLC|Ltd\.?|Corp\.?|Corporation|Company|Co\.|Group|Partners|Solutions|Services|Technologies|Tech))\b',
        r'@([\w-]+)\.\w+',  # Extract company from email domain
    ]

    # Phone patterns
    PHONE_PATTERNS = [
        r'(?:Phone|Tel|Mobile|Cell|Office|Direct|Fax)?[:\s]*([+]?[\d\s\-\(\)\.]{10,20})',
        r'\b(\+?\d{1,3}[-.\s]?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})\b',
    ]

    def __init__(self):
        """Initialize the contact extractor."""
        self.title_regex = [re.compile(p, re.IGNORECASE) for p in self.TITLE_PATTERNS]
        self.company_regex = [re.compile(p, re.IGNORECASE | re.MULTILINE) for p in self.COMPANY_PATTERNS]
        self.phone_regex = [re.compile(p, re.IGNORECASE) for p in self.PHONE_PATTERNS]

    def _extract_title_from_text(self, text: str) -> Optional[str]:
        """
        Extract job title from text (usually email signature).

        Args:
            text: Text to search for title

        Returns:
            Extracted title or None
        """
        for regex in self.title_regex:
            match = regex.search(text)
            if match:
                title = match.group(0).strip()
                # Clean up the title
                title = re.sub(r'\s+', ' ', title)
                return title
        return None

    def _extract_company_from_text(self, text: str, email: str = "") -> Optional[str]:
        """
        Extract company name from text or email domain.

        Args:
            text: Text to search for company
            email: Email address to extract domain from

        Returns:
            Extracted company name or None
        """
        # First try to extract from text
        for regex in self.company_regex[:1]:  # Only the first pattern for text
            match = regex.search(text)
            if match:
                return match.group(1).strip()

        # Fall back to email domain
        if email:
            domain_match = re.search(r'@([\w-]+)\.\w+', email)
            if domain_match:
                domain = domain_match.group(1)
                # Skip common email providers
                common_providers = {'gmail', 'yahoo', 'hotmail', 'outlook', 'aol', 'icloud', 'live', 'msn'}
                if domain.lower() not in common_providers:
                    return domain.title()

        return None

    def _extract_phone_from_text(self, text: str) -> Optional[str]:
        """
        Extract phone number from text.

        Args:
            text: Text to search for phone number

        Returns:
            Extracted phone number or None
        """
        for regex in self.phone_regex:
            match = regex.search(text)
            if match:
                phone = match.group(1).strip()
                # Clean up the phone number
                phone = re.sub(r'[\s\-\(\)\.]+', '', phone)
                if len(phone) >= 10:
                    return phone
        return None

    def _parse_name(self, name: Optional[str], email: str) -> str:
        """
        Parse or infer name from available information.

        Args:
            name: Provided name (may be None)
            email: Email address to infer name from

        Returns:
            Best guess at the person's name
        """
        if name and name.strip():
            # Clean up the name
            name = name.strip()
            # Remove quotes
            name = name.strip('"\'')
            # Remove email if it appears in the name field
            name = re.sub(r'<[^>]+>', '', name).strip()
            if name:
                return name

        # Infer from email address
        local_part = email.split('@')[0]
        # Common patterns: john.doe, john_doe, johndoe
        name_parts = re.split(r'[._]', local_part)
        if len(name_parts) >= 2:
            return ' '.join(part.title() for part in name_parts if part)
        return local_part.title()

    def extract_from_email(self, email: Email) -> List[Contact]:
        """
        Extract all contacts from an email.

        Args:
            email: Email object to extract contacts from

        Returns:
            List of Contact objects
        """
        contacts = []

        # Extract sender
        if email.sender and email.sender.email:
            sender_contact = Contact(
                email=email.sender.email.lower(),
                name=self._parse_name(email.sender.name, email.sender.email),
                last_contact_date=email.received_datetime,
                contact_count=1,
                source="email_sender",
            )

            # Try to extract title and company from email body preview
            if email.body_preview:
                sender_contact.title = self._extract_title_from_text(email.body_preview)
                sender_contact.company = self._extract_company_from_text(
                    email.body_preview, email.sender.email
                )
                sender_contact.phone = self._extract_phone_from_text(email.body_preview)

            # If no company from signature, try email domain
            if not sender_contact.company:
                sender_contact.company = self._extract_company_from_text("", email.sender.email)

            contacts.append(sender_contact)

        # Extract To recipients
        for recipient in email.to_recipients:
            if recipient.email:
                contact = Contact(
                    email=recipient.email.lower(),
                    name=self._parse_name(recipient.name, recipient.email),
                    company=self._extract_company_from_text("", recipient.email),
                    last_contact_date=email.received_datetime,
                    contact_count=1,
                    source="email_recipient",
                )
                contacts.append(contact)

        # Extract CC recipients
        for recipient in email.cc_recipients:
            if recipient.email:
                contact = Contact(
                    email=recipient.email.lower(),
                    name=self._parse_name(recipient.name, recipient.email),
                    company=self._extract_company_from_text("", recipient.email),
                    last_contact_date=email.received_datetime,
                    contact_count=1,
                    source="email_cc",
                )
                contacts.append(contact)

        return contacts

    def extract_from_signature(self, signature_text: str, email_address: str = "") -> Optional[Contact]:
        """
        Extract contact from an email signature block.

        Args:
            signature_text: Email signature text
            email_address: Associated email address

        Returns:
            Contact object or None
        """
        # Try to find email in signature if not provided
        if not email_address:
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', signature_text)
            if email_match:
                email_address = email_match.group(0)
            else:
                return None

        # Extract name from first line (common signature format)
        lines = [l.strip() for l in signature_text.split('\n') if l.strip()]
        name = lines[0] if lines else None

        # Clean up name
        if name:
            # Remove common prefixes
            name = re.sub(r'^(Best|Regards|Thanks|Sincerely|Cheers),?\s*', '', name, flags=re.IGNORECASE)
            name = name.strip()
            if not name or '@' in name:
                name = None

        return Contact(
            email=email_address.lower(),
            name=name or self._parse_name(None, email_address),
            title=self._extract_title_from_text(signature_text),
            company=self._extract_company_from_text(signature_text, email_address),
            phone=self._extract_phone_from_text(signature_text),
            source="signature",
        )


class ContactAggregator:
    """Aggregates and deduplicates contacts from multiple sources."""

    def __init__(self):
        """Initialize the aggregator."""
        self.contacts: Dict[str, Contact] = {}

    def add_contact(self, contact: Contact) -> None:
        """
        Add a contact, merging with existing if duplicate.

        Args:
            contact: Contact to add
        """
        email_key = contact.email.lower()

        if email_key in self.contacts:
            existing = self.contacts[email_key]

            # Update with new information (prefer non-None values)
            if contact.name and (not existing.name or existing.name == self._parse_email_name(email_key)):
                existing.name = contact.name

            if contact.title and not existing.title:
                existing.title = contact.title

            if contact.company and not existing.company:
                existing.company = contact.company

            if contact.phone and not existing.phone:
                existing.phone = contact.phone

            # Update last contact date (keep the most recent)
            if contact.last_contact_date:
                if not existing.last_contact_date or contact.last_contact_date > existing.last_contact_date:
                    existing.last_contact_date = contact.last_contact_date

            existing.contact_count += contact.contact_count
        else:
            self.contacts[email_key] = contact

    def _parse_email_name(self, email: str) -> str:
        """Parse name from email address."""
        local_part = email.split('@')[0]
        name_parts = re.split(r'[._]', local_part)
        if len(name_parts) >= 2:
            return ' '.join(part.title() for part in name_parts if part)
        return local_part.title()

    def add_contacts(self, contacts: List[Contact]) -> None:
        """
        Add multiple contacts.

        Args:
            contacts: List of contacts to add
        """
        for contact in contacts:
            self.add_contact(contact)

    def get_all_contacts(self) -> List[Contact]:
        """
        Get all aggregated contacts.

        Returns:
            List of all contacts
        """
        return list(self.contacts.values())

    def get_contacts_sorted_by_last_contact(self) -> List[Contact]:
        """
        Get contacts sorted by last contact date (most recent first).

        Returns:
            Sorted list of contacts
        """
        return sorted(
            self.contacts.values(),
            key=lambda c: c.last_contact_date or datetime.min,
            reverse=True,
        )

    def get_contacts_sorted_by_frequency(self) -> List[Contact]:
        """
        Get contacts sorted by contact frequency (most frequent first).

        Returns:
            Sorted list of contacts
        """
        return sorted(
            self.contacts.values(),
            key=lambda c: c.contact_count,
            reverse=True,
        )
