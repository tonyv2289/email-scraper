"""
Local Outlook Module

Reads emails directly from the local Outlook desktop application on Windows.
No Azure AD or admin permissions required - just needs Outlook installed.
"""

import sys
from dataclasses import dataclass
from datetime import datetime
from typing import Generator, List, Optional

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn

console = Console()


def check_windows():
    """Check if running on Windows."""
    if sys.platform != 'win32':
        console.print("[red]Local Outlook mode is only available on Windows.[/red]")
        console.print("[yellow]On macOS/Linux, you'll need to use the Graph API method.[/yellow]")
        return False
    return True


@dataclass
class LocalEmailAddress:
    """Represents an email address."""
    email: str
    name: Optional[str] = None


@dataclass
class LocalAttachment:
    """Represents an email attachment."""
    name: str
    content: bytes
    size: int


@dataclass
class LocalEmail:
    """Represents an email from local Outlook."""
    subject: str
    sender_email: str
    sender_name: Optional[str]
    to_recipients: List[LocalEmailAddress]
    cc_recipients: List[LocalEmailAddress]
    received_time: Optional[datetime]
    body: str
    attachments: List[LocalAttachment]


class LocalOutlookReader:
    """Reads emails from local Outlook installation via COM."""

    def __init__(self):
        """Initialize the Outlook COM connection."""
        if not check_windows():
            raise RuntimeError("Local Outlook mode requires Windows")

        try:
            import win32com.client
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except ImportError:
            console.print("[red]pywin32 is required for local Outlook access.[/red]")
            console.print("[yellow]Install it with: pip install pywin32[/yellow]")
            raise
        except Exception as e:
            console.print(f"[red]Could not connect to Outlook: {e}[/red]")
            console.print("[yellow]Make sure Outlook is installed and has been opened at least once.[/yellow]")
            raise

    def _get_folder(self, folder_name: str):
        """Get a mail folder by name."""
        # Folder constants
        FOLDER_MAP = {
            'inbox': 6,      # olFolderInbox
            'sent': 5,       # olFolderSentMail
            'sentitems': 5,
            'drafts': 16,    # olFolderDrafts
            'deleted': 3,    # olFolderDeletedItems
            'outbox': 4,     # olFolderOutbox
        }

        folder_const = FOLDER_MAP.get(folder_name.lower())
        if folder_const:
            return self.namespace.GetDefaultFolder(folder_const)

        # Try to find by name
        try:
            root = self.namespace.Folders
            for folder in root:
                try:
                    return folder.Folders[folder_name]
                except:
                    pass
        except:
            pass

        return None

    def _parse_recipients(self, recipients) -> List[LocalEmailAddress]:
        """Parse recipients from Outlook recipients collection."""
        result = []
        try:
            for i in range(1, recipients.Count + 1):
                recip = recipients.Item(i)
                try:
                    # Try to get SMTP address
                    email = recip.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    )
                except:
                    email = recip.Address

                result.append(LocalEmailAddress(
                    email=email.lower() if email else "",
                    name=recip.Name
                ))
        except Exception as e:
            pass
        return result

    def _get_sender_email(self, mail_item) -> str:
        """Extract sender email address."""
        try:
            # For Exchange users
            if mail_item.SenderEmailType == "EX":
                try:
                    sender = mail_item.Sender
                    return sender.GetExchangeUser().PrimarySmtpAddress.lower()
                except:
                    try:
                        return mail_item.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
                        ).lower()
                    except:
                        pass
            return mail_item.SenderEmailAddress.lower()
        except:
            return ""

    def fetch_emails(
        self,
        folders: List[str] = None,
        max_emails: Optional[int] = None,
        since_date: Optional[datetime] = None,
        include_attachments: bool = True,
    ) -> Generator[LocalEmail, None, None]:
        """
        Fetch emails from local Outlook.

        Args:
            folders: List of folder names to fetch from (default: inbox, sent)
            max_emails: Maximum emails to fetch
            since_date: Only fetch emails after this date
            include_attachments: Whether to include attachment content

        Yields:
            LocalEmail objects
        """
        if folders is None:
            folders = ['inbox', 'sent']

        total_fetched = 0

        for folder_name in folders:
            if max_emails and total_fetched >= max_emails:
                break

            folder = self._get_folder(folder_name)
            if not folder:
                console.print(f"[yellow]Warning: Could not access folder '{folder_name}'[/yellow]")
                continue

            console.print(f"\n[cyan]Reading from {folder_name}...[/cyan]")

            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)  # Sort by date, newest first

                # Apply date filter if specified
                if since_date:
                    date_str = since_date.strftime("%m/%d/%Y")
                    items = items.Restrict(f"[ReceivedTime] >= '{date_str}'")

                item_count = items.Count
                console.print(f"[dim]Found {item_count} emails in {folder_name}[/dim]")

                with Progress(
                    SpinnerColumn(),
                    TextColumn("[progress.description]{task.description}"),
                    BarColumn(),
                    TaskProgressColumn(),
                    console=console,
                ) as progress:
                    task = progress.add_task(f"[cyan]Processing {folder_name}...", total=min(item_count, max_emails - total_fetched if max_emails else item_count))

                    for i in range(1, item_count + 1):
                        if max_emails and total_fetched >= max_emails:
                            break

                        try:
                            mail = items.Item(i)

                            # Skip non-mail items
                            if mail.Class != 43:  # olMail
                                continue

                            # Get sender info
                            sender_email = self._get_sender_email(mail)
                            sender_name = mail.SenderName if hasattr(mail, 'SenderName') else None

                            # Get recipients
                            to_recipients = self._parse_recipients(mail.Recipients)
                            # Filter to TO recipients (type 1)
                            to_list = []
                            cc_list = []
                            try:
                                for j in range(1, mail.Recipients.Count + 1):
                                    recip = mail.Recipients.Item(j)
                                    addr = LocalEmailAddress(
                                        email=self._get_recipient_email(recip),
                                        name=recip.Name
                                    )
                                    if recip.Type == 1:  # TO
                                        to_list.append(addr)
                                    elif recip.Type == 2:  # CC
                                        cc_list.append(addr)
                            except:
                                pass

                            # Get received time
                            received_time = None
                            try:
                                received_time = mail.ReceivedTime
                                if hasattr(received_time, 'replace'):
                                    # It's already a datetime
                                    pass
                                else:
                                    # Convert from COM date
                                    received_time = datetime.fromtimestamp(received_time)
                            except:
                                pass

                            # Get body
                            body = ""
                            try:
                                body = mail.Body[:2000] if mail.Body else ""  # Limit body size
                            except:
                                pass

                            # Get attachments
                            attachments = []
                            if include_attachments and mail.Attachments.Count > 0:
                                attachments = self._get_attachments(mail)

                            email = LocalEmail(
                                subject=mail.Subject or "",
                                sender_email=sender_email,
                                sender_name=sender_name,
                                to_recipients=to_list,
                                cc_recipients=cc_list,
                                received_time=received_time,
                                body=body,
                                attachments=attachments,
                            )

                            total_fetched += 1
                            progress.update(task, advance=1, description=f"[cyan]Processed {total_fetched} emails...")
                            yield email

                        except Exception as e:
                            # Skip problematic emails
                            continue

            except Exception as e:
                console.print(f"[red]Error reading folder {folder_name}: {e}[/red]")

        console.print(f"\n[green]Total emails processed: {total_fetched}[/green]")

    def _get_recipient_email(self, recipient) -> str:
        """Get email address from a recipient."""
        try:
            if recipient.AddressEntry.Type == "EX":
                try:
                    return recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
                except:
                    pass
            return recipient.Address.lower() if recipient.Address else ""
        except:
            return ""

    def _get_attachments(self, mail_item) -> List[LocalAttachment]:
        """Extract attachments from an email."""
        attachments = []
        try:
            import tempfile
            import os

            for i in range(1, mail_item.Attachments.Count + 1):
                try:
                    att = mail_item.Attachments.Item(i)

                    # Skip inline images and very large files
                    if att.Size > 10 * 1024 * 1024:  # Skip files > 10MB
                        continue

                    # Save to temp file to get content
                    temp_path = os.path.join(tempfile.gettempdir(), att.FileName)
                    att.SaveAsFile(temp_path)

                    with open(temp_path, 'rb') as f:
                        content = f.read()

                    os.unlink(temp_path)  # Clean up temp file

                    attachments.append(LocalAttachment(
                        name=att.FileName,
                        content=content,
                        size=att.Size,
                    ))
                except Exception as e:
                    continue

        except Exception as e:
            pass

        return attachments

    def get_account_info(self) -> dict:
        """Get information about the current Outlook account."""
        try:
            accounts = self.namespace.Accounts
            if accounts.Count > 0:
                account = accounts.Item(1)
                return {
                    "display_name": account.DisplayName,
                    "email": account.SmtpAddress,
                }
        except:
            pass

        return {"display_name": "Unknown", "email": "Unknown"}
