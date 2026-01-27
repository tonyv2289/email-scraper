"""
Storage Module

Handles persistent storage and export of extracted contacts.
Supports JSON and CSV formats.
"""

import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

from rich.console import Console
from rich.table import Table

from .contact_extractor import Contact

console = Console()


class ContactStorage:
    """Handles storage and retrieval of contacts."""

    def __init__(self, output_dir: str = "./output"):
        """
        Initialize the storage.

        Args:
            output_dir: Directory to store output files
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _contact_to_dict(self, contact: Contact) -> Dict:
        """Convert a Contact to a dictionary."""
        return {
            "email": contact.email,
            "name": contact.name,
            "title": contact.title,
            "company": contact.company,
            "phone": contact.phone,
            "last_contact_date": contact.last_contact_date.isoformat() if contact.last_contact_date else None,
            "contact_count": contact.contact_count,
            "source": contact.source,
        }

    def _dict_to_contact(self, data: Dict) -> Contact:
        """Convert a dictionary to a Contact."""
        last_contact = None
        if data.get("last_contact_date"):
            try:
                last_contact = datetime.fromisoformat(data["last_contact_date"])
            except (ValueError, TypeError):
                pass

        return Contact(
            email=data["email"],
            name=data.get("name"),
            title=data.get("title"),
            company=data.get("company"),
            phone=data.get("phone"),
            last_contact_date=last_contact,
            contact_count=data.get("contact_count", 1),
            source=data.get("source", "unknown"),
        )

    def save_to_json(self, contacts: List[Contact], filename: str = "contacts.json") -> Path:
        """
        Save contacts to a JSON file.

        Args:
            contacts: List of contacts to save
            filename: Output filename

        Returns:
            Path to the saved file
        """
        filepath = self.output_dir / filename
        data = {
            "exported_at": datetime.now().isoformat(),
            "total_contacts": len(contacts),
            "contacts": [self._contact_to_dict(c) for c in contacts],
        }

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        console.print(f"[green]Saved {len(contacts)} contacts to {filepath}[/green]")
        return filepath

    def save_to_csv(self, contacts: List[Contact], filename: str = "contacts.csv") -> Path:
        """
        Save contacts to a CSV file.

        Args:
            contacts: List of contacts to save
            filename: Output filename

        Returns:
            Path to the saved file
        """
        filepath = self.output_dir / filename
        fieldnames = ["email", "name", "title", "company", "phone", "last_contact_date", "contact_count", "source"]

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

            for contact in contacts:
                row = self._contact_to_dict(contact)
                writer.writerow(row)

        console.print(f"[green]Saved {len(contacts)} contacts to {filepath}[/green]")
        return filepath

    def load_from_json(self, filename: str = "contacts.json") -> List[Contact]:
        """
        Load contacts from a JSON file.

        Args:
            filename: Input filename

        Returns:
            List of Contact objects
        """
        filepath = self.output_dir / filename

        if not filepath.exists():
            console.print(f"[yellow]Warning: File {filepath} does not exist[/yellow]")
            return []

        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        contacts = [self._dict_to_contact(c) for c in data.get("contacts", [])]
        console.print(f"[green]Loaded {len(contacts)} contacts from {filepath}[/green]")
        return contacts

    def print_contacts_table(
        self,
        contacts: List[Contact],
        max_rows: int = 50,
        show_all_columns: bool = False,
    ) -> None:
        """
        Print contacts in a formatted table.

        Args:
            contacts: List of contacts to display
            max_rows: Maximum number of rows to display
            show_all_columns: Show all columns including source
        """
        table = Table(title=f"Extracted Contacts ({len(contacts)} total)")

        table.add_column("Email", style="cyan", no_wrap=True)
        table.add_column("Name", style="green")
        table.add_column("Title", style="yellow")
        table.add_column("Company", style="magenta")
        table.add_column("Last Contact", style="blue")
        table.add_column("Count", justify="right", style="white")

        if show_all_columns:
            table.add_column("Phone", style="white")
            table.add_column("Source", style="dim")

        for i, contact in enumerate(contacts[:max_rows]):
            last_contact = ""
            if contact.last_contact_date:
                last_contact = contact.last_contact_date.strftime("%Y-%m-%d")

            row = [
                contact.email,
                contact.name or "-",
                contact.title or "-",
                contact.company or "-",
                last_contact or "-",
                str(contact.contact_count),
            ]

            if show_all_columns:
                row.extend([
                    contact.phone or "-",
                    contact.source,
                ])

            table.add_row(*row)

        if len(contacts) > max_rows:
            table.add_row("...", "...", "...", "...", "...", "...")
            console.print(f"\n[dim]Showing first {max_rows} of {len(contacts)} contacts[/dim]")

        console.print(table)

    def print_summary(self, contacts: List[Contact]) -> None:
        """
        Print a summary of extracted contacts.

        Args:
            contacts: List of contacts
        """
        console.print("\n[bold]Contact Extraction Summary[/bold]")
        console.print(f"  Total contacts: [cyan]{len(contacts)}[/cyan]")

        # Count by source
        sources = {}
        for c in contacts:
            source_type = c.source.split(":")[0] if ":" in c.source else c.source
            sources[source_type] = sources.get(source_type, 0) + 1

        console.print("\n  [bold]By source:[/bold]")
        for source, count in sorted(sources.items(), key=lambda x: -x[1]):
            console.print(f"    {source}: [cyan]{count}[/cyan]")

        # Count with titles
        with_title = sum(1 for c in contacts if c.title)
        with_company = sum(1 for c in contacts if c.company)
        with_phone = sum(1 for c in contacts if c.phone)

        console.print("\n  [bold]Data completeness:[/bold]")
        console.print(f"    With title: [cyan]{with_title}[/cyan] ({100*with_title//max(len(contacts),1)}%)")
        console.print(f"    With company: [cyan]{with_company}[/cyan] ({100*with_company//max(len(contacts),1)}%)")
        console.print(f"    With phone: [cyan]{with_phone}[/cyan] ({100*with_phone//max(len(contacts),1)}%)")

        # Top companies
        companies = {}
        for c in contacts:
            if c.company:
                companies[c.company] = companies.get(c.company, 0) + 1

        if companies:
            console.print("\n  [bold]Top companies:[/bold]")
            for company, count in sorted(companies.items(), key=lambda x: -x[1])[:10]:
                console.print(f"    {company}: [cyan]{count}[/cyan]")
