"""
Local Outlook Module for macOS

Reads emails directly from Outlook for Mac using AppleScript.
No Azure AD or admin permissions required - just needs Outlook for Mac installed.
"""

import os
import subprocess
import json
import re
import tempfile
import shutil
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
            timeout=600  # 10 minutes timeout
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
            # Just check if we can talk to Outlook
            test_script = '''
            tell application "Microsoft Outlook"
                return name
            end tell
            '''
            _run_applescript(test_script)
        except Exception as e:
            raise RuntimeError(f"Could not connect to Outlook for Mac: {e}")

    def list_all_folders(self) -> str:
        """List all folders from all accounts for debugging."""
        script = '''
        tell application "Microsoft Outlook"
            set output to ""

            -- Try to list all accounts with their folders
            try
                set output to output & "=== All Accounts ===" & linefeed
                repeat with acct in (every account)
                    try
                        set acctName to name of acct
                        set output to output & "Account: " & acctName & linefeed

                        -- Try to get folders from this account
                        try
                            repeat with f in (mail folders of acct)
                                try
                                    set output to output & "  - " & (name of f) & linefeed
                                    -- Check subfolders
                                    try
                                        repeat with f2 in (mail folders of f)
                                            set output to output & "    - " & (name of f2) & linefeed
                                            try
                                                repeat with f3 in (mail folders of f2)
                                                    set output to output & "      - " & (name of f3) & linefeed
                                                end repeat
                                            end try
                                        end repeat
                                    end try
                                end try
                            end repeat
                        end try
                    end try
                end repeat
            end try

            -- Also try the selected folder in UI
            try
                set output to output & linefeed & "=== Currently Selected ===" & linefeed
                set selFolder to selected folder
                set output to output & "Selected: " & (name of selFolder) & linefeed
            end try

            return output
        end tell
        '''
        try:
            return _run_applescript(script)
        except Exception as e:
            return f"Error listing folders: {e}"

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

        # Create temp directory for attachments if needed
        attachment_dir = None
        if include_attachments:
            attachment_dir = tempfile.mkdtemp(prefix="outlook_attachments_")
            console.print(f"[dim]Saving attachments to temp dir: {attachment_dir}[/dim]")

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

            # Special case: use currently selected folder in Outlook UI
            if outlook_folder.lower() == "selected":
                count_script = '''
                tell application "Microsoft Outlook"
                    try
                        set selFolder to selected folder
                        return count of messages of selFolder
                    end try
                    return 0
                end tell
                '''
            else:
                # Get message count - search through all mail folders directly
                count_script = f'''
                tell application "Microsoft Outlook"
                    set targetName to "{outlook_folder}"

                    -- Search through all mail folders and their subfolders
                    repeat with f in (every mail folder)
                        try
                            if name of f is targetName then
                                return count of messages of f
                            end if
                            -- Search subfolders (level 2)
                            try
                                repeat with f2 in (mail folders of f)
                                    if name of f2 is targetName then
                                        return count of messages of f2
                                    end if
                                    -- Search sub-subfolders (level 3)
                                    try
                                        repeat with f3 in (mail folders of f2)
                                            if name of f3 is targetName then
                                                return count of messages of f3
                                            end if
                                        end repeat
                                    end try
                                end repeat
                            end try
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
                batch_size = 200  # Larger batches with iterator approach
                for batch_start in range(1, to_fetch + 1, batch_size):
                    batch_end = min(batch_start + batch_size - 1, to_fetch)

                    # Fetch batch of emails (no date filter - we filter in Python)
                    if outlook_folder.lower() == "selected":
                        fetch_script = f'''
                        tell application "Microsoft Outlook"
                            set theFolder to selected folder
                            if theFolder is missing value then
                                return ""
                            end if'''
                    else:
                        fetch_script = f'''
                        tell application "Microsoft Outlook"
                            set targetName to "{outlook_folder}"
                            set theFolder to missing value

                            -- Search through all mail folders and subfolders
                            repeat with f in (every mail folder)
                                if theFolder is not missing value then exit repeat
                                try
                                    if name of f is targetName then
                                        set theFolder to f
                                        exit repeat
                                    end if
                                    -- Search subfolders (level 2)
                                    try
                                        repeat with f2 in (mail folders of f)
                                            if theFolder is not missing value then exit repeat
                                            if name of f2 is targetName then
                                                set theFolder to f2
                                                exit repeat
                                            end if
                                            -- Search sub-subfolders (level 3)
                                            try
                                                repeat with f3 in (mail folders of f2)
                                                    if name of f3 is targetName then
                                                        set theFolder to f3
                                                        exit repeat
                                                    end if
                                                end repeat
                                            end try
                                        end repeat
                                    end try
                                end try
                            end repeat

                            if theFolder is missing value then
                                return ""
                            end if'''

                    fetch_script = fetch_script + f'''

                        set output to ""
                        set counter to 0

                        repeat with theMessage in (messages of theFolder)
                            set counter to counter + 1
                            if counter < {batch_start} then
                                -- skip to batch start
                            else if counter > {batch_end} then
                                exit repeat
                            else
                                -- Robust property access: each property in its own try block
                                set msgSubject to ""
                                try
                                    set msgSubject to subject of theMessage
                                end try

                                set senderEmail to ""
                                set senderName to ""
                                try
                                    set msgSender to sender of theMessage
                                    try
                                        set senderEmail to address of msgSender
                                    end try
                                    try
                                        set senderName to name of msgSender
                                    end try
                                end try

                                set msgTime to ""
                                try
                                    set msgTime to (time received of theMessage) as string
                                end try

                                set msgBody to ""
                                try
                                    set msgBody to plain text content of theMessage
                                end try

                                -- Get recipients - try multiple property access methods
                                set toList to ""
                                set toCount to 0
                                try
                                    set toRecips to every to recipient of theMessage
                                    set toCount to count of toRecips
                                    repeat with r in toRecips
                                        try
                                            -- Try method 1: direct email address property
                                            set recipAddr to email address of r
                                            if recipAddr is not missing value then
                                                -- email address might be an object, try to get address from it
                                                try
                                                    set actualAddr to address of recipAddr
                                                    if actualAddr is not missing value and actualAddr is not "" then
                                                        set toList to toList & actualAddr & ","
                                                    end if
                                                on error
                                                    -- email address was already a string
                                                    if recipAddr is not "" then
                                                        set toList to toList & recipAddr & ","
                                                    end if
                                                end try
                                            end if
                                        on error
                                            -- Try method 2: direct address property
                                            try
                                                set recipAddr to address of r
                                                if recipAddr is not missing value and recipAddr is not "" then
                                                    set toList to toList & recipAddr & ","
                                                end if
                                            end try
                                        end try
                                    end repeat
                                end try

                                set ccList to ""
                                set ccCount to 0
                                try
                                    set ccRecips to every cc recipient of theMessage
                                    set ccCount to count of ccRecips
                                    repeat with r in ccRecips
                                        try
                                            set recipAddr to email address of r
                                            if recipAddr is not missing value then
                                                try
                                                    set actualAddr to address of recipAddr
                                                    if actualAddr is not missing value and actualAddr is not "" then
                                                        set ccList to ccList & actualAddr & ","
                                                    end if
                                                on error
                                                    if recipAddr is not "" then
                                                        set ccList to ccList & recipAddr & ","
                                                    end if
                                                end try
                                            end if
                                        on error
                                            try
                                                set recipAddr to address of r
                                                if recipAddr is not missing value and recipAddr is not "" then
                                                    set ccList to ccList & recipAddr & ","
                                                end if
                                            end try
                                        end try
                                    end repeat
                                end try

                                -- Format: subject|||senderEmail|||senderName|||time|||body|||toList|||ccList|||attachmentPaths
                                set bodyText to ""
                                try
                                    if length of msgBody > 0 then
                                        if length of msgBody > 2000 then
                                            set bodyText to text 1 thru 2000 of msgBody
                                        else
                                            set bodyText to msgBody
                                        end if
                                    end if
                                end try

                                -- Handle attachments if requested
                                set attachmentPaths to ""
                                '''

                    # Add attachment handling if enabled
                    if include_attachments and attachment_dir:
                        fetch_script = fetch_script + f'''
                                try
                                    set attList to every attachment of theMessage
                                    repeat with att in attList
                                        try
                                            set attName to name of att
                                            -- Create unique filename: counter_originalname
                                            set savePath to "{attachment_dir}/" & counter & "_" & attName
                                            -- Save the attachment
                                            save att in savePath
                                            set attachmentPaths to attachmentPaths & savePath & ";;;"
                                        end try
                                    end repeat
                                end try
                                '''

                    fetch_script = fetch_script + '''
                                set msgLine to msgSubject & "|||" & senderEmail & "|||" & senderName & "|||" & msgTime & "|||" & bodyText & "|||" & toList & "|||" & ccList & "|||" & attachmentPaths
                                set output to output & msgLine & "<<<MSGSEP>>>"
                            end if
                        end repeat

                        return output
                    end tell
                    '''

                    try:
                        console.print(f"[dim]Fetching batch {batch_start}-{batch_end}...[/dim]")
                        result = _run_applescript(fetch_script)
                        messages = result.split("<<<MSGSEP>>>")
                        console.print(f"[dim]Got {len([m for m in messages if m.strip()])} messages in batch[/dim]")

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
                            if to_recipients or cc_recipients:
                                console.print(f"[dim]  Found {len(to_recipients)} to, {len(cc_recipients)} cc recipients[/dim]")

                            # Parse attachments if available
                            attachments = []
                            if len(parts) > 7 and parts[7].strip():
                                attachment_paths = parts[7].split(";;;")
                                for att_path in attachment_paths:
                                    att_path = att_path.strip()
                                    if att_path and os.path.exists(att_path):
                                        try:
                                            with open(att_path, 'rb') as f:
                                                content = f.read()
                                            att_name = os.path.basename(att_path)
                                            # Remove the counter prefix (e.g., "123_filename.pdf" -> "filename.pdf")
                                            if '_' in att_name:
                                                att_name = att_name.split('_', 1)[1]
                                            attachments.append(LocalAttachment(
                                                name=att_name,
                                                content=content,
                                                size=len(content),
                                            ))
                                        except Exception as e:
                                            console.print(f"[yellow]Warning: Could not read attachment {att_path}: {e}[/yellow]")
                                if attachments:
                                    console.print(f"[dim]  Found {len(attachments)} attachments[/dim]")

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
                                attachments=attachments,
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

        # Clean up temp attachment directory
        if attachment_dir and os.path.exists(attachment_dir):
            try:
                shutil.rmtree(attachment_dir)
                console.print(f"[dim]Cleaned up temp attachment directory[/dim]")
            except Exception as e:
                console.print(f"[yellow]Warning: Could not clean up temp directory {attachment_dir}: {e}[/yellow]")
