"""
Command Line Interface

Provides CLI commands for the Outlook contact scraper.
"""

import os
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
@click.option('--max-emails', '-n', default=None, type=int, help='Maximum number of emails to process')
@click.option('--days', '-d', default=None, type=int, help='Only process emails from the last N days')
@click.option('--include-sent/--no-sent', default=True, help='Include sent emails (default: yes)')
@click.option('--process-attachments/--no-attachments', default=True, help='Process attachments for contacts (default: yes)')
@click.option('--output-dir', '-o', default='./output', help='Output directory for exported files')
@click.option('--format', '-f', 'output_format', type=click.Choice(['csv', 'json', 'both']), default='both', help='Output format')
@click.option('--show-table/--no-table', default=True, help='Show results table (default: yes)')
def scrape(
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
    1. Authenticate with Microsoft Graph API
    2. Fetch emails from your inbox (and optionally sent folder)
    3. Extract contact information from email headers and signatures
    4. Optionally process attachments for contact lists
    5. Export the results to CSV and/or JSON
    """
    console.print("\n[bold blue]Outlook Contact Scraper[/bold blue]")
    console.print("=" * 40)

    # Authenticate
    try:
        auth = get_authenticator()
        access_token = auth.get_access_token()
        if not access_token:
            console.print("[red]Authentication failed. Exiting.[/red]")
            return
    except ValueError as e:
        console.print(f"[red]{e}[/red]")
        console.print("\n[yellow]Setup instructions:[/yellow]")
        console.print("1. Create an Azure AD app at https://portal.azure.com")
        console.print("2. Add API permissions: Mail.Read, Mail.ReadBasic, User.Read")
        console.print("3. Copy .env.example to .env and fill in your credentials")
        return

    # Initialize components
    fetcher = EmailFetcher(access_token)
    extractor = ContactExtractor()
    aggregator = ContactAggregator()
    attachment_processor = AttachmentProcessor()
    storage = ContactStorage(output_dir)

    # Get user info
    user_info = fetcher.get_user_info()
    if user_info:
        console.print(f"\n[green]Authenticated as:[/green] {user_info.get('displayName', 'Unknown')} ({user_info.get('mail', user_info.get('userPrincipalName', 'Unknown'))})")

    # Calculate since_date if days specified
    since_date = None
    if days:
        since_date = datetime.now() - timedelta(days=days)
        console.print(f"[cyan]Filtering emails from the last {days} days[/cyan]")

    # Fetch and process emails
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
