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

3. Set up Azure AD app registration (see below)

4. Configure environment variables:
   ```bash
   cp .env.example .env
   # Edit .env with your Azure AD credentials
   ```

## Azure AD Setup

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
# Run the scraper (will prompt for authentication)
python main.py scrape

# Limit to last 30 days of emails
python main.py scrape --days 30

# Process only 100 emails maximum
python main.py scrape --max-emails 100

# Skip attachment processing (faster)
python main.py scrape --no-attachments

# Export only to CSV
python main.py scrape --format csv
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
Make sure you've copied `.env.example` to `.env` and filled in your Azure AD credentials.

### "Authentication failed"
1. Verify your CLIENT_ID and TENANT_ID are correct
2. Check that you've granted the required API permissions
3. Try running `python main.py logout` and then `python main.py login` again

### "Access denied" or permission errors
Your Azure AD admin may need to grant consent for the application permissions.

## License

MIT License
