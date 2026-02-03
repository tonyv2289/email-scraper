"""
Command Line Interface

Provides CLI commands for the Outlook contact scraper.
"""

import os
import sys
from datetime import datetime, timedelta
from typing import Optional

import click
from dotenv import load_dotenv
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn

from .attachment_processor import AttachmentProcessor
from .auth import get_authenticator
from .contact_extractor import ContactAggregator, ContactExtractor
from .email_fetcher import EmailFetcher
from .storage import ContactStorage

console = Console()


def _get_local_reader():
    """Get the appropriate local Outlook reader for the current platform."""
    if sys.platform == 'darwin':
        # macOS - use AppleScript
        from .local_outlook_mac import MacOutlookReader
        return MacOutlookReader()
    elif sys.platform == 'win32':
        # Windows - use COM automation
        from .local_outlook import LocalOutlookReader
        return LocalOutlookReader()
    else:
        raise RuntimeError(f"Local mode not supported on {sys.platform}. Use API mode instead.")


def _scrape_local(extractor, aggregator, attachment_processor, max_emails, since_date, include_sent, process_attachments, custom_folder=None):
    """Scrape using local Outlook application (Windows or macOS)."""
    try:
        reader = _get_local_reader()
    except ImportError as e:
        console.print(f"[red]Could not import local Outlook module: {e}[/red]")
        if sys.platform == 'win32':
            console.print("[yellow]Install pywin32: pip install pywin32[/yellow]")
        return 0, 0
    except Exception as e:
        console.print(f"[red]Failed to connect to Outlook: {e}[/red]")
        console.print("\n[yellow]Make sure:[/yellow]")
        if sys.platform == 'darwin':
            console.print("  1. Outlook for Mac is installed")
            console.print("  2. You've opened Outlook at least once")
            console.print("  3. Grant Terminal/Python permission to control Outlook when prompted")
        else:
            console.print("  1. You're running on Windows or macOS")
            console.print("  2. Outlook desktop app is installed")
            console.print("  3. You've opened Outlook at least once")
            console.print("  4. (Windows) pywin32 is installed: pip install pywin32")
        return 0, 0

    # Get account info
    account_info = reader.get_account_info()
    console.print(f"\n[green]Reading from:[/green] {account_info.get('display_name', 'Unknown')} ({account_info.get('email', 'Unknown')})")

    # Set up folders
    if custom_folder:
        # Use only the specified custom folder
        folders = [custom_folder]
        console.print(f"[cyan]Using custom folder: {custom_folder}[/cyan]")
    else:
        folders = ['inbox']
        if include_sent:
            folders.append('sent')

    console.print("\n[bold]Fetching emails from local Outlook...[/bold]")

    email_count = 0
    attachment_count = 0

    from .contact_extractor import Contact
    from .email_fetcher import Attachment

    for local_email in reader.fetch_emails(
        folders=folders,
        max_emails=max_emails,
        since_date=since_date,
        include_attachments=process_attachments,
    ):
        email_count += 1

        # Create contacts from the local email
        if local_email.sender_email:
            sender_contact = Contact(
                email=local_email.sender_email,
                name=local_email.sender_name,
                last_contact_date=local_email.received_time,
                contact_count=1,
                source="email_sender",
            )
            # Extract company from email domain (more reliable than signature parsing)
            sender_contact.company = extractor._extract_company_from_text("", local_email.sender_email)
            aggregator.add_contact(sender_contact)

        # Add recipients
        for recip in local_email.to_recipients:
            if recip.email:
                contact = Contact(
                    email=recip.email,
                    name=recip.name,
                    company=extractor._extract_company_from_text("", recip.email),
                    last_contact_date=local_email.received_time,
                    contact_count=1,
                    source="email_recipient",
                )
                aggregator.add_contact(contact)

        for recip in local_email.cc_recipients:
            if recip.email:
                contact = Contact(
                    email=recip.email,
                    name=recip.name,
                    company=extractor._extract_company_from_text("", recip.email),
                    last_contact_date=local_email.received_time,
                    contact_count=1,
                    source="email_cc",
                )
                aggregator.add_contact(contact)

        # Process attachments
        if process_attachments and local_email.attachments:
            for local_att in local_email.attachments:
                attachment_count += 1
                # Convert to our Attachment type
                import mimetypes
                content_type, _ = mimetypes.guess_type(local_att.name)
                attachment = Attachment(
                    name=local_att.name,
                    content_type=content_type or 'application/octet-stream',
                    content=local_att.content,
                    size=local_att.size,
                )
                att_contacts = attachment_processor.process_attachment(attachment)
                aggregator.add_contacts(att_contacts)

    return email_count, attachment_count


def _scrape_api(extractor, aggregator, attachment_processor, max_emails, since_date, include_sent, process_attachments):
    """Scrape using Microsoft Graph API."""
    # Authenticate
    try:
        auth = get_authenticator()
        access_token = auth.get_access_token()
        if not access_token:
            console.print("[red]Authentication failed. Exiting.[/red]")
            return 0, 0
    except ValueError as e:
        console.print(f"[red]{e}[/red]")
        console.print("\n[yellow]Setup instructions:[/yellow]")
        console.print("1. Create an Azure AD app at https://portal.azure.com")
        console.print("2. Add API permissions: Mail.Read, Mail.ReadBasic, User.Read")
        console.print("3. Copy .env.example to .env and fill in your credentials")
        console.print("\n[cyan]Or use --local flag on Windows to read from Outlook desktop app (no admin required)[/cyan]")
        return 0, 0

    fetcher = EmailFetcher(access_token)

    # Get user info
    user_info = fetcher.get_user_info()
    if user_info:
        console.print(f"\n[green]Authenticated as:[/green] {user_info.get('displayName', 'Unknown')} ({user_info.get('mail', user_info.get('userPrincipalName', 'Unknown'))})")

    console.print("\n[bold]Fetching emails...[/bold]")

    email_count = 0
    attachment_count = 0

    for email in fetcher.fetch_emails(
        max_emails=max_emails,
        include_sent=include_sent,
        since_date=since_date,
    ):
        email_count += 1

        # Extract contacts from email
        contacts = extractor.extract_from_email(email)
        aggregator.add_contacts(contacts)

        # Process attachments if enabled
        if process_attachments and email.has_attachments:
            attachments = fetcher.fetch_email_attachments(email)
            for attachment in attachments:
                attachment_count += 1
                attachment_contacts = attachment_processor.process_attachment(attachment)
                aggregator.add_contacts(attachment_contacts)

    return email_count, attachment_count


@click.group()
@click.version_option(version="1.0.0")
def cli():
    """
    Outlook Contact Scraper

    Extract contact information from your Outlook emails including names,
    email addresses, job titles, and last contact dates.
    """
    # Load environment variables
    load_dotenv()


@cli.command()
@click.option('--local', '-l', is_flag=True, help='Use local Outlook app (Windows/macOS, no admin required)')
@click.option('--folder', default=None, type=str, help='Specific folder to scrape (e.g., "EXPORT", "Projects")')
@click.option('--max-emails', '-n', default=None, type=int, help='Maximum number of emails to process')
@click.option('--days', '-d', default=None, type=int, help='Only process emails from the last N days')
@click.option('--include-sent/--no-sent', default=True, help='Include sent emails (default: yes)')
@click.option('--process-attachments/--no-attachments', default=True, help='Process attachments for contacts (default: yes)')
@click.option('--output-dir', '-o', default='./output', help='Output directory for exported files')
@click.option('--format', '-f', 'output_format', type=click.Choice(['csv', 'json', 'both']), default='both', help='Output format')
@click.option('--show-table/--no-table', default=True, help='Show results table (default: yes)')
def scrape(
    local: bool,
    folder: Optional[str],
    max_emails: Optional[int],
    days: Optional[int],
    include_sent: bool,
    process_attachments: bool,
    output_dir: str,
    output_format: str,
    show_table: bool,
):
    """
    Scrape contacts from Outlook emails.

    This command will:
    1. Connect to Outlook (via API or local app)
    2. Fetch emails from your inbox (and optionally sent folder)
    3. Extract contact information from email headers and signatures
    4. Optionally process attachments for contact lists
    5. Export the results to CSV and/or JSON

    Use --local flag on Windows to read from Outlook desktop app (no admin required).
    """
    console.print("\n[bold blue]Outlook Contact Scraper[/bold blue]")
    console.print("=" * 40)

    # Initialize components
    extractor = ContactExtractor()
    aggregator = ContactAggregator()
    attachment_processor = AttachmentProcessor()
    storage = ContactStorage(output_dir)

    # Calculate since_date if days specified
    since_date = None
    if days:
        since_date = datetime.now() - timedelta(days=days)
        console.print(f"[cyan]Filtering emails from the last {days} days[/cyan]")

    email_count = 0
    attachment_count = 0

    if local:
        # Use local Outlook mode (Windows/macOS, no admin required)
        email_count, attachment_count = _scrape_local(
            extractor, aggregator, attachment_processor,
            max_emails, since_date, include_sent, process_attachments, folder
        )
    else:
        # Use Graph API mode
        email_count, attachment_count = _scrape_api(
            extractor, aggregator, attachment_processor,
            max_emails, since_date, include_sent, process_attachments
        )

    if email_count == 0:
        return

    console.print(f"\n[green]Processed {email_count} emails and {attachment_count} attachments[/green]")

    # Get all contacts
    contacts = aggregator.get_contacts_sorted_by_last_contact()

    if not contacts:
        console.print("[yellow]No contacts found.[/yellow]")
        return

    # Print summary
    storage.print_summary(contacts)

    # Show table
    if show_table:
        console.print()
        storage.print_contacts_table(contacts, max_rows=30)

    # Export results
    console.print("\n[bold]Exporting results...[/bold]")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if output_format in ('csv', 'both'):
        storage.save_to_csv(contacts, f"contacts_{timestamp}.csv")

    if output_format in ('json', 'both'):
        storage.save_to_json(contacts, f"contacts_{timestamp}.json")

    console.print("\n[bold green]Done![/bold green]")


@cli.command()
def list_folders():
    """
    List all available mail folders in Outlook (macOS only).

    Use this to find the exact folder names for the --folder option.
    """
    import sys
    if sys.platform != 'darwin':
        console.print("[red]This command is only available on macOS[/red]")
        return

    console.print("\n[bold blue]Listing Outlook Folders[/bold blue]")
    console.print("=" * 40)

    import subprocess

    # Try multiple approaches
    scripts = [
        # Approach 1: List all mail folders directly
        ('Method 1: All mail folders', '''
tell application "Microsoft Outlook"
    set output to ""
    try
        repeat with f in (every mail folder)
            try
                set output to output & name of f & " (" & (count of messages of f) & ")" & linefeed
            end try
        end repeat
    end try
    return output
end tell
        '''),
        # Approach 2: Check inbox/sent directly
        ('Method 2: Standard folders', '''
tell application "Microsoft Outlook"
    set output to ""
    try
        set output to output & "inbox (" & (count of messages of inbox) & ")" & linefeed
    end try
    try
        set output to output & "sent items (" & (count of messages of sent items) & ")" & linefeed
    end try
    try
        set output to output & "drafts (" & (count of messages of drafts) & ")" & linefeed
    end try
    return output
end tell
        '''),
        # Approach 3: Account info
        ('Method 3: Account info', '''
tell application "Microsoft Outlook"
    set output to ""
    try
        set output to output & "Default account: " & (name of default account) & linefeed
        set output to output & "Email: " & (email address of default account) & linefeed
    end try
    try
        set output to output & "Exchange accounts: " & (count of exchange accounts) & linefeed
    end try
    try
        set output to output & "IMAP accounts: " & (count of imap accounts) & linefeed
    end try
    return output
end tell
        '''),
    ]

    for name, script in scripts:
        console.print(f"\n[cyan]{name}:[/cyan]")
        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True,
                timeout=30
            )
            if result.returncode == 0 and result.stdout.strip():
                console.print(result.stdout)
            else:
                console.print(f"[dim](no output)[/dim]")
                if result.stderr:
                    console.print(f"[dim]Error: {result.stderr.strip()}[/dim]")
        except Exception as e:
            console.print(f"[dim]Failed: {e}[/dim]")


@cli.command()
def login():
    """
    Authenticate with Microsoft Graph API.

    This will open a browser window for you to sign in with your Microsoft account.
    Your authentication will be cached for future use.
    """
    console.print("\n[bold blue]Microsoft Graph API Authentication[/bold blue]")

    try:
        auth = get_authenticator()
        access_token = auth.get_access_token()

        if access_token:
            fetcher = EmailFetcher(access_token)
            user_info = fetcher.get_user_info()

            if user_info:
                console.print(f"\n[green]Successfully authenticated as:[/green]")
                console.print(f"  Name: {user_info.get('displayName', 'Unknown')}")
                console.print(f"  Email: {user_info.get('mail', user_info.get('userPrincipalName', 'Unknown'))}")
        else:
            console.print("[red]Authentication failed.[/red]")

    except ValueError as e:
        console.print(f"[red]{e}[/red]")


@cli.command()
def logout():
    """
    Clear cached authentication.

    This will remove your saved login credentials.
    """
    console.print("\n[bold blue]Clearing authentication...[/bold blue]")

    try:
        auth = get_authenticator()
        auth.logout()
    except ValueError:
        # Still try to clear the cache even without credentials
        from pathlib import Path
        cache_file = Path.home() / ".outlook_scraper_token_cache.json"
        if cache_file.exists():
            cache_file.unlink()
            console.print("[green]Logged out successfully.[/green]")
        else:
            console.print("[yellow]No cached authentication found.[/yellow]")


@cli.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('--output-dir', '-o', default='./output', help='Output directory')
@click.option('--format', '-f', 'output_format', type=click.Choice(['csv', 'json', 'both']), default='both', help='Output format')
def convert(input_file: str, output_dir: str, output_format: str):
    """
    Convert a previously exported contacts file to another format.

    INPUT_FILE should be a JSON or CSV file from a previous export.
    """
    import json
    import csv
    from pathlib import Path

    storage = ContactStorage(output_dir)
    filepath = Path(input_file)

    contacts = []

    if filepath.suffix == '.json':
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
            contacts = [storage._dict_to_contact(c) for c in data.get('contacts', [])]
    elif filepath.suffix == '.csv':
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            contacts = [storage._dict_to_contact(row) for row in reader]
    else:
        console.print(f"[red]Unsupported file format: {filepath.suffix}[/red]")
        return

    console.print(f"[green]Loaded {len(contacts)} contacts from {filepath}[/green]")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if output_format in ('csv', 'both'):
        storage.save_to_csv(contacts, f"contacts_{timestamp}.csv")

    if output_format in ('json', 'both'):
        storage.save_to_json(contacts, f"contacts_{timestamp}.json")


@cli.command()
@click.argument('attachment_path', type=click.Path(exists=True))
@click.option('--output-dir', '-o', default='./output', help='Output directory')
def parse_file(attachment_path: str, output_dir: str):
    """
    Parse a local file (CSV, Excel, vCard, PDF) for contacts.

    This is useful for extracting contacts from files you already have.
    """
    from pathlib import Path

    filepath = Path(attachment_path)

    with open(filepath, 'rb') as f:
        content = f.read()

    # Create a fake attachment object
    from .email_fetcher import Attachment
    import mimetypes

    content_type, _ = mimetypes.guess_type(str(filepath))
    attachment = Attachment(
        name=filepath.name,
        content_type=content_type or 'application/octet-stream',
        content=content,
        size=len(content),
    )

    processor = AttachmentProcessor()
    contacts = processor.process_attachment(attachment)

    if not contacts:
        console.print("[yellow]No contacts found in file.[/yellow]")
        return

    console.print(f"[green]Found {len(contacts)} contacts[/green]")

    storage = ContactStorage(output_dir)
    storage.print_contacts_table(contacts, max_rows=50, show_all_columns=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    storage.save_to_csv(contacts, f"contacts_{timestamp}.csv")
    storage.save_to_json(contacts, f"contacts_{timestamp}.json")


def main():
    """Entry point for the CLI."""
    cli()


if __name__ == "__main__":
    main()
