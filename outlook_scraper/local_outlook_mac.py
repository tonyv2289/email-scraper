"""
Local Outlook Module for macOS

Reads emails directly from Outlook for Mac using AppleScript.
No Azure AD or admin permissions required - just needs Outlook for Mac installed.
"""

import subprocess
import json
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Generator, List, Optional

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TaskProgressColumn

console = Console()


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


def _run_applescript(script: str) -> str:
    """Run an AppleScript and return the result."""
    try:
        result = subprocess.run(
            ['osascript', '-e', script],
            capture_output=True,
            text=True,
            timeout=120
        )
        if result.returncode != 0:
            raise RuntimeError(f"AppleScript error: {result.stderr}")
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        raise RuntimeError("AppleScript timed out")
    except FileNotFoundError:
        raise RuntimeError("osascript not found - are you on macOS?")


class MacOutlookReader:
    """Reads emails from Outlook for Mac via AppleScript."""

    def __init__(self):
        """Initialize and verify Outlook is available."""
        # Check if Outlook is installed and accessible
        try:
            script = '''
            tell application "System Events"
                return exists application process "Microsoft Outlook"
            end tell
            '''
            # Just check if we can talk to Outlook
            test_script = '''
            tell application "Microsoft Outlook"
                return name
            end tell
            '''
            _run_applescript(test_script)
        except Exception as e:
            raise RuntimeError(f"Could not connect to Outlook for Mac: {e}")

    def get_account_info(self) -> dict:
        """Get information about the default Outlook account."""
        try:
            script = '''
            tell application "Microsoft Outlook"
                set defaultAccount to default account
                set accountName to name of defaultAccount
                set accountEmail to email address of defaultAccount
                return accountName & "|" & accountEmail
            end tell
            '''
            result = _run_applescript(script)
            parts = result.split("|")
            return {
                "display_name": parts[0] if len(parts) > 0 else "Unknown",
                "email": parts[1] if len(parts) > 1 else "Unknown",
            }
        except Exception:
            return {"display_name": "Unknown", "email": "Unknown"}

    def _parse_email_list(self, raw: str) -> List[LocalEmailAddress]:
        """Parse a comma-separated list of email addresses."""
        addresses = []
        if not raw or raw == "missing value":
            return addresses

        # Split by comma and parse each
        for item in raw.split(","):
            item = item.strip()
            if not item:
                continue

            # Try to extract email from various formats
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', item)
            if email_match:
                email = email_match.group(0).lower()
                # Try to get name (text before the email)
                name = item.replace(email_match.group(0), "").strip(" <>()")
                addresses.append(LocalEmailAddress(email=email, name=name if name else None))

        return addresses

    def fetch_emails(
        self,
        folders: List[str] = None,
        max_emails: Optional[int] = None,
        since_date: Optional[datetime] = None,
        include_attachments: bool = False,  # Attachment extraction is slower on Mac
    ) -> Generator[LocalEmail, None, None]:
        """
        Fetch emails from Outlook for Mac.

        Args:
            folders: List of folder names (default: inbox, sent)
            max_emails: Maximum emails to fetch
            since_date: Only fetch emails after this date
            include_attachments: Whether to include attachments (slower)

        Yields:
            LocalEmail objects
        """
        if folders is None:
            folders = ['inbox', 'sent items']

        total_fetched = 0

        # Map common folder names (but allow custom folder names to pass through)
        folder_map = {
            'inbox': 'inbox',
            'sent': 'sent items',
            'sent items': 'sent items',
            'sentitems': 'sent items',
            'drafts': 'drafts',
        }

        for folder_name in folders:
            if max_emails and total_fetched >= max_emails:
                break

            # Use mapped name if exists, otherwise use as-is (for custom folders like "EXPORT")
            outlook_folder = folder_map.get(folder_name.lower(), folder_name)
            console.print(f"\n[cyan]Reading from {outlook_folder}...[/cyan]")

            # Get message count - search through folder hierarchy (3 levels deep)
            count_script = f'''
            tell application "Microsoft Outlook"
                set targetName to "{outlook_folder}"

                -- Try direct access first (for standard folders like inbox)
                try
                    set theFolder to mail folder targetName
                    return count of messages of theFolder
                end try

                -- Search in all accounts - check 3 levels deep
                set allAccounts to {{}}
                try
                    set allAccounts to allAccounts & exchange accounts
                end try
                try
                    set allAccounts to allAccounts & imap accounts
                end try
                try
                    set allAccounts to allAccounts & pop accounts
                end try

                repeat with acct in allAccounts
                    try
                        set rootF to root folder of acct
                        -- Level 1: direct children of root
                        repeat with f1 in mail folders of rootF
                            if name of f1 is targetName then
                                return count of messages of f1
                            end if
                            -- Level 2: subfolders
                            try
                                repeat with f2 in mail folders of f1
                                    if name of f2 is targetName then
                                        return count of messages of f2
                                    end if
                                    -- Level 3: sub-subfolders
                                    try
                                        repeat with f3 in mail folders of f2
                                            if name of f3 is targetName then
                                                return count of messages of f3
                                            end if
                                        end repeat
                                    end try
                                end repeat
                            end try
                        end repeat
                    end try
                end repeat

                return 0
            end tell
            '''

            try:
                msg_count = int(_run_applescript(count_script))
                if msg_count == 0:
                    console.print(f"[yellow]Folder '{outlook_folder}' is empty or not found[/yellow]")
                    continue
            except Exception as e:
                console.print(f"[yellow]Could not access folder '{outlook_folder}': {e}[/yellow]")
                continue

            console.print(f"[dim]Found {msg_count} emails in {outlook_folder}[/dim]")

            # Calculate how many to fetch from this folder
            remaining = max_emails - total_fetched if max_emails else msg_count
            to_fetch = min(msg_count, remaining)

            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TaskProgressColumn(),
                console=console,
            ) as progress:
                task = progress.add_task(f"[cyan]Processing {outlook_folder}...", total=to_fetch)

                # Fetch in batches to avoid AppleScript timeouts
                batch_size = 50
                for batch_start in range(1, to_fetch + 1, batch_size):
                    batch_end = min(batch_start + batch_size - 1, to_fetch)

                    # Fetch batch of emails (no date filter - we filter in Python)
                    fetch_script = f'''
                    tell application "Microsoft Outlook"
                        set targetName to "{outlook_folder}"
                        set theFolder to missing value

                        -- Try direct access first
                        try
                            set theFolder to mail folder targetName
                        end try

                        -- If not found, search in all accounts (3 levels deep)
                        if theFolder is missing value then
                            set allAccounts to {{}}
                            try
                                set allAccounts to allAccounts & exchange accounts
                            end try
                            try
                                set allAccounts to allAccounts & imap accounts
                            end try
                            try
                                set allAccounts to allAccounts & pop accounts
                            end try

                            repeat with acct in allAccounts
                                if theFolder is not missing value then exit repeat
                                try
                                    set rootF to root folder of acct
                                    repeat with f1 in mail folders of rootF
                                        if theFolder is not missing value then exit repeat
                                        if name of f1 is targetName then
                                            set theFolder to f1
                                            exit repeat
                                        end if
                                        try
                                            repeat with f2 in mail folders of f1
                                                if theFolder is not missing value then exit repeat
                                                if name of f2 is targetName then
                                                    set theFolder to f2
                                                    exit repeat
                                                end if
                                                try
                                                    repeat with f3 in mail folders of f2
                                                        if name of f3 is targetName then
                                                            set theFolder to f3
                                                            exit repeat
                                                        end if
                                                    end repeat
                                                end try
                                            end repeat
                                        end try
                                    end repeat
                                end try
                            end repeat
                        end if

                        if theFolder is missing value then
                            return ""
                        end if

                        set allMessages to messages of theFolder
                        set output to ""

                        repeat with i from {batch_start} to {batch_end}
                            try
                                set theMessage to item i of allMessages
                                set msgSubject to subject of theMessage
                                set msgSender to sender of theMessage
                                set senderEmail to address of msgSender
                                set senderName to name of msgSender
                                set msgTime to time received of theMessage
                                set msgBody to ""
                                try
                                    set msgBody to plain text content of theMessage
                                end try

                                -- Get recipients
                                set toList to ""
                                repeat with r in to recipients of theMessage
                                    set toList to toList & address of r & ","
                                end repeat

                                set ccList to ""
                                repeat with r in cc recipients of theMessage
                                    set ccList to ccList & address of r & ","
                                end repeat

                                -- Format: subject|||senderEmail|||senderName|||time|||body|||toList|||ccList
                                set bodyText to ""
                                if length of msgBody > 0 then
                                    if length of msgBody > 2000 then
                                        set bodyText to text 1 thru 2000 of msgBody
                                    else
                                        set bodyText to msgBody
                                    end if
                                end if
                                set msgLine to msgSubject & "|||" & senderEmail & "|||" & senderName & "|||" & (msgTime as string) & "|||" & bodyText & "|||" & toList & "|||" & ccList
                                set output to output & msgLine & "<<<MSGSEP>>>"
                            end try
                        end repeat

                        return output
                    end tell
                    '''

                    try:
                        result = _run_applescript(fetch_script)
                        messages = result.split("<<<MSGSEP>>>")

                        for msg_data in messages:
                            if not msg_data.strip():
                                continue

                            parts = msg_data.split("|||")
                            if len(parts) < 7:
                                continue

                            subject = parts[0]
                            sender_email = parts[1].lower() if parts[1] else ""
                            sender_name = parts[2] if parts[2] and parts[2] != "missing value" else None
                            time_str = parts[3]
                            body = parts[4] if parts[4] != "missing value" else ""
                            to_raw = parts[5]
                            cc_raw = parts[6]

                            # Parse received time
                            received_time = None
                            if time_str and time_str != "missing value":
                                try:
                                    # macOS date format varies, try common formats
                                    for fmt in ["%A, %B %d, %Y at %I:%M:%S %p", "%m/%d/%Y %I:%M:%S %p", "%Y-%m-%d %H:%M:%S"]:
                                        try:
                                            received_time = datetime.strptime(time_str, fmt)
                                            break
                                        except ValueError:
                                            continue
                                except Exception:
                                    pass

                            # Parse recipients
                            to_recipients = self._parse_email_list(to_raw)
                            cc_recipients = self._parse_email_list(cc_raw)

                            # Filter by date in Python (more reliable than AppleScript date filtering)
                            if since_date and received_time and received_time < since_date:
                                progress.update(task, advance=1)
                                continue

                            email = LocalEmail(
                                subject=subject,
                                sender_email=sender_email,
                                sender_name=sender_name,
                                to_recipients=to_recipients,
                                cc_recipients=cc_recipients,
                                received_time=received_time,
                                body=body,
                                attachments=[],  # Attachment extraction is complex on Mac
                            )

                            total_fetched += 1
                            progress.update(task, advance=1, description=f"[cyan]Processed {total_fetched} emails...")
                            yield email

                            if max_emails and total_fetched >= max_emails:
                                break

                    except Exception as e:
                        console.print(f"[yellow]Warning: Error fetching batch: {e}[/yellow]")
                        continue

                    if max_emails and total_fetched >= max_emails:
                        break

        console.print(f"\n[green]Total emails processed: {total_fetched}[/green]")
