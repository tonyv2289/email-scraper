"""
Email Fetcher Module

Fetches emails from Microsoft Outlook using Graph API.
Retrieves sender info, recipients, and attachments.
"""

import base64
from dataclasses import dataclass, field
from datetime import datetime
from typing import Generator, List, Optional

import requests
from dateutil.parser import parse as parse_date
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn

console = Console()

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"


@dataclass
class EmailAddress:
    """Represents an email address with optional name."""
    email: str
    name: Optional[str] = None


@dataclass
class Attachment:
    """Represents an email attachment."""
    name: str
    content_type: str
    content: bytes
    size: int


@dataclass
class Email:
    """Represents a fetched email with all relevant information."""
    id: str
    subject: str
    sender: EmailAddress
    to_recipients: List[EmailAddress] = field(default_factory=list)
    cc_recipients: List[EmailAddress] = field(default_factory=list)
    received_datetime: Optional[datetime] = None
    body_preview: str = ""
    has_attachments: bool = False
    attachments: List[Attachment] = field(default_factory=list)


class EmailFetcher:
    """Fetches emails from Outlook via Microsoft Graph API."""

    def __init__(self, access_token: str):
        """
        Initialize the email fetcher.

        Args:
            access_token: Microsoft Graph API access token
        """
        self.access_token = access_token
        self.headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }

    def _make_request(self, url: str, params: Optional[dict] = None) -> Optional[dict]:
        """
        Make a GET request to the Graph API.

        Args:
            url: API endpoint URL
            params: Query parameters

        Returns:
            JSON response or None on error
        """
        try:
            response = requests.get(url, headers=self.headers, params=params)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            if response.status_code == 401:
                console.print("[red]Authentication expired. Please re-authenticate.[/red]")
            else:
                console.print(f"[red]API error: {e}[/red]")
            return None
        except requests.exceptions.RequestException as e:
            console.print(f"[red]Request failed: {e}[/red]")
            return None

    def get_user_info(self) -> Optional[dict]:
        """Get information about the authenticated user."""
        url = f"{GRAPH_API_BASE}/me"
        return self._make_request(url)

    def _parse_email_address(self, data: dict) -> EmailAddress:
        """Parse an email address from Graph API response."""
        email_data = data.get("emailAddress", {})
        return EmailAddress(
            email=email_data.get("address", ""),
            name=email_data.get("name"),
        )

    def _parse_email(self, data: dict) -> Email:
        """Parse an email from Graph API response."""
        sender_data = data.get("sender", {})
        sender = self._parse_email_address(sender_data) if sender_data else EmailAddress(email="")

        to_recipients = [
            self._parse_email_address(r)
            for r in data.get("toRecipients", [])
        ]

        cc_recipients = [
            self._parse_email_address(r)
            for r in data.get("ccRecipients", [])
        ]

        received_datetime = None
        if data.get("receivedDateTime"):
            try:
                received_datetime = parse_date(data["receivedDateTime"])
            except (ValueError, TypeError):
                pass

        return Email(
            id=data.get("id", ""),
            subject=data.get("subject", ""),
            sender=sender,
            to_recipients=to_recipients,
            cc_recipients=cc_recipients,
            received_datetime=received_datetime,
            body_preview=data.get("bodyPreview", ""),
            has_attachments=data.get("hasAttachments", False),
        )

    def fetch_emails(
        self,
        folder: str = "inbox",
        max_emails: Optional[int] = None,
        include_sent: bool = True,
        since_date: Optional[datetime] = None,
    ) -> Generator[Email, None, None]:
        """
        Fetch emails from Outlook.

        Args:
            folder: Mail folder to fetch from ("inbox", "sentitems", etc.)
            max_emails: Maximum number of emails to fetch (None for all)
            include_sent: Also fetch sent emails
            since_date: Only fetch emails after this date

        Yields:
            Email objects
        """
        folders_to_fetch = [folder]
        if include_sent and folder.lower() != "sentitems":
            folders_to_fetch.append("sentitems")

        total_fetched = 0

        for mail_folder in folders_to_fetch:
            if max_emails and total_fetched >= max_emails:
                break

            console.print(f"\n[cyan]Fetching from {mail_folder}...[/cyan]")

            url = f"{GRAPH_API_BASE}/me/mailFolders/{mail_folder}/messages"
            params = {
                "$select": "id,subject,sender,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments",
                "$orderby": "receivedDateTime desc",
                "$top": 100,  # Fetch in batches of 100
            }

            # Add date filter if specified
            if since_date:
                date_str = since_date.strftime("%Y-%m-%dT%H:%M:%SZ")
                params["$filter"] = f"receivedDateTime ge {date_str}"

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TaskProgressColumn(),
                console=console,
            ) as progress:
                task = progress.add_task(f"[cyan]Fetching {mail_folder}...", total=None)

                while url:
                    if max_emails and total_fetched >= max_emails:
                        break

                    response = self._make_request(url, params if url.startswith(GRAPH_API_BASE) else None)
                    if not response:
                        break

                    messages = response.get("value", [])
                    for msg_data in messages:
                        if max_emails and total_fetched >= max_emails:
                            break

                        email = self._parse_email(msg_data)
                        total_fetched += 1
                        progress.update(task, description=f"[cyan]Fetched {total_fetched} emails...")
                        yield email

                    # Get next page URL
                    url = response.get("@odata.nextLink")
                    params = None  # params are included in nextLink

        console.print(f"[green]Total emails fetched: {total_fetched}[/green]")

    def fetch_attachment(self, email_id: str, attachment_id: str) -> Optional[Attachment]:
        """
        Fetch a specific attachment.

        Args:
            email_id: ID of the email
            attachment_id: ID of the attachment

        Returns:
            Attachment object or None
        """
        url = f"{GRAPH_API_BASE}/me/messages/{email_id}/attachments/{attachment_id}"
        response = self._make_request(url)

        if not response:
            return None

        content_bytes = base64.b64decode(response.get("contentBytes", ""))

        return Attachment(
            name=response.get("name", "unknown"),
            content_type=response.get("contentType", "application/octet-stream"),
            content=content_bytes,
            size=response.get("size", 0),
        )

    def fetch_email_attachments(self, email: Email) -> List[Attachment]:
        """
        Fetch all attachments for an email.

        Args:
            email: Email object

        Returns:
            List of Attachment objects
        """
        if not email.has_attachments:
            return []

        url = f"{GRAPH_API_BASE}/me/messages/{email.id}/attachments"
        response = self._make_request(url)

        if not response:
            return []

        attachments = []
        for att_data in response.get("value", []):
            # Only process file attachments (not item attachments)
            if att_data.get("@odata.type") == "#microsoft.graph.fileAttachment":
                content_bytes = base64.b64decode(att_data.get("contentBytes", ""))
                attachment = Attachment(
                    name=att_data.get("name", "unknown"),
                    content_type=att_data.get("contentType", "application/octet-stream"),
                    content=content_bytes,
                    size=att_data.get("size", 0),
                )
                attachments.append(attachment)

        return attachments
