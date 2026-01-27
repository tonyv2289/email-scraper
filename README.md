# Outlook Contact Scraper

Extract contact information (names, emails, job titles) from your Outlook emails, including processing attachments that contain contact lists.

## Features

- **Email Extraction**: Scrapes contacts from email senders, recipients, and CC fields
- **Signature Parsing**: Extracts job titles, companies, and phone numbers from email signatures
- **Attachment Processing**: Parses contact lists from CSV, Excel, vCard, and PDF attachments
- **Last Contact Tracking**: Records the most recent date you communicated with each contact
- **Contact Frequency**: Tracks how many times you've emailed each contact
- **Deduplication**: Automatically merges duplicate contacts and enriches data
- **Export Options**: Outputs to both CSV and JSON formats

## Prerequisites

- Python 3.8+
- Microsoft 365 / Outlook account
- **Either:**
  - Outlook desktop app installed (Windows) - **no admin required**, or
  - Azure AD app registration (for API access)

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd email-scraper
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. **For Windows local mode (no admin required):**
   ```bash
   pip install pywin32
   ```

4. **For API mode:** Set up Azure AD app registration (see below) and configure:
   ```bash
   cp .env.example .env
   # Edit .env with your Azure AD credentials
   ```

## Quick Start (No Admin Required - Windows Only)

If you're on Windows with Outlook desktop installed, you can skip all Azure setup:

```bash
# Install dependencies
pip install -r requirements.txt
pip install pywin32

# Run with local mode
python main.py scrape --local

# Limit to recent emails
python main.py scrape --local --days 30
```

This reads directly from your Outlook desktop application - no Azure AD registration or admin permissions needed.

## Azure AD Setup (API Mode)

To access your Outlook emails via the Microsoft Graph API, you need to register an application in Azure AD:

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure:
   - Name: "Outlook Contact Scraper" (or your choice)
   - Supported account types: Choose based on your needs
   - Redirect URI: Select "Public client/native" and enter `http://localhost`
5. After registration, note the **Application (client) ID** and **Directory (tenant) ID**
6. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
7. Add these permissions:
   - `Mail.Read`
   - `Mail.ReadBasic`
   - `User.Read`
8. Click **Grant admin consent** (if you're an admin) or request consent from your admin

## Usage

### Basic Scraping

```bash
# LOCAL MODE (Windows, no admin required)
python main.py scrape --local

# API MODE (requires Azure setup)
python main.py scrape

# Common options (work with both modes)
python main.py scrape --local --days 30        # Last 30 days only
python main.py scrape --local --max-emails 100 # Limit to 100 emails
python main.py scrape --local --no-attachments # Skip attachments (faster)
python main.py scrape --local --format csv     # Export only to CSV
```

### Authentication Commands

```bash
# Login and cache credentials
python main.py login

# Clear cached credentials
python main.py logout
```

### File Processing

```bash
# Parse a local file for contacts
python main.py parse-file contacts.xlsx

# Convert between formats
python main.py convert contacts.json --format csv
```

### All Options

```
python main.py scrape --help

Options:
  -l, --local                 Use local Outlook app (Windows only, no admin required)
  -n, --max-emails INTEGER    Maximum number of emails to process
  -d, --days INTEGER          Only process emails from the last N days
  --include-sent/--no-sent    Include sent emails (default: yes)
  --process-attachments/--no-attachments
                              Process attachments for contacts (default: yes)
  -o, --output-dir TEXT       Output directory for exported files
  -f, --format [csv|json|both]
                              Output format
  --show-table/--no-table     Show results table (default: yes)
```

## Output

The scraper exports contacts with the following fields:

| Field | Description |
|-------|-------------|
| email | Email address |
| name | Full name |
| title | Job title (extracted from signatures) |
| company | Company name |
| phone | Phone number |
| last_contact_date | Date of most recent email |
| contact_count | Number of emails with this contact |
| source | Where the contact was found |

### Example Output

```
┏━━━━━━━━━━━━━━━━━━━━━━━┳━━━━━━━━━━━━━━━━┳━━━━━━━━━━━━━━━━━━━━┳━━━━━━━━━━━━━┳━━━━━━━━━━━━━━┳━━━━━━━┓
┃ Email                 ┃ Name           ┃ Title              ┃ Company     ┃ Last Contact ┃ Count ┃
┡━━━━━━━━━━━━━━━━━━━━━━━╇━━━━━━━━━━━━━━━━╇━━━━━━━━━━━━━━━━━━━━╇━━━━━━━━━━━━━╇━━━━━━━━━━━━━━╇━━━━━━━┩
│ john.doe@company.com  │ John Doe       │ Senior Engineer    │ Company     │ 2024-01-15   │    42 │
│ jane.smith@corp.com   │ Jane Smith     │ Product Manager    │ Corp        │ 2024-01-14   │    28 │
│ bob@startup.io        │ Bob Johnson    │ CEO                │ Startup     │ 2024-01-10   │    15 │
└───────────────────────┴────────────────┴────────────────────┴─────────────┴──────────────┴───────┘
```

## Supported Attachment Types

- **CSV files** (`.csv`)
- **Excel files** (`.xlsx`, `.xls`, `.xlsm`)
- **vCard files** (`.vcf`, `.vcard`)
- **PDF files** (`.pdf`) - extracts emails from text and tables
- **Text files** (`.txt`)

## Privacy & Security

- Your credentials are cached locally in `~/.outlook_scraper_token_cache.json`
- No data is sent to any third-party servers
- All processing happens locally on your machine
- Use `python main.py logout` to clear cached credentials

## Troubleshooting

### "CLIENT_ID environment variable is required"
You're trying to use API mode without Azure credentials. Either:
- Set up Azure AD and configure `.env` file, or
- Use `--local` flag on Windows to skip Azure entirely

### "Authentication failed"
1. Verify your CLIENT_ID and TENANT_ID are correct
2. Check that you've granted the required API permissions
3. Try running `python main.py logout` and then `python main.py login` again
4. **Alternative:** Use `--local` flag on Windows (no Azure required)

### "Access denied" or permission errors
Your Azure AD admin may need to grant consent for the application permissions.
**Alternative:** Use `--local` flag on Windows - it reads from your local Outlook app and doesn't require admin permissions.

### Local mode issues (Windows)
1. Make sure Outlook desktop is installed and has been opened at least once
2. Install pywin32: `pip install pywin32`
3. If using a corporate account, your emails should already be synced locally

## License

MIT License
