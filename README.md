# Google Sheets to CSV Exporter

A configurable, shareable tool to export Google Sheets from your Google Drive to local CSV files. Perfect for data analysis, backups, or sharing data with others.

## Features

- 🔐 **Secure Authentication**: Uses OAuth 2.0 with Google APIs
- 📊 **Flexible Export**: Export single sheets, multiple sheets, or all worksheets
- ⚙️ **Configurable**: Customizable file naming, output directories, and export options
- 🖥️ **Multiple Usage Modes**: Interactive, command-line, or batch processing
- 🔄 **Shareable**: Easy setup for different users with their own credentials
- 📱 **Cross-Platform**: Works on Windows, macOS, and Linux

## Quick Start

### 1. Install Dependencies

First, set up a virtual environment to keep your project dependencies isolated.

```bash
# Create a virtual environment
python3 -m venv venv

# Activate it (on macOS/Linux)
source venv/bin/activate

# On Windows, use:
# venv\\Scripts\\activate
```

Now, install the required packages:

```bash
# Install dependencies
python3 -m pip install -r requirements.txt
```

### 2. Set Up Google API Credentials

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Sheets API** and **Google Drive API**:
   - Navigate to "APIs & Services" > "Library"
   - Search for and enable both APIs
4. Create credentials:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth 2.0 Client ID"
   - Choose "Desktop application"
   - Download the JSON file
5. Save the downloaded file as `credentials.json` in this directory

### 3. Run Setup

```bash
python setup.py
```

This will guide you through configuration and test your authentication.

### 4. Start Exporting

```bash
# Interactive mode - browse and select sheets
python sheets_to_csv.py

# List all your sheets
python sheets_to_csv.py --list

# Export specific sheets by name
python sheets_to_csv.py --name "My Budget" --name "Sales Data"

# Export by sheet ID
python sheets_to_csv.py --id "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
```

## Usage Examples

### Interactive Mode
```bash
python sheets_to_csv.py
```
Browse your sheets and select which ones to export interactively.

### Export by Name
```bash
python sheets_to_csv.py --name "Budget 2024" --name "Inventory"
```

### Export by Google Sheets ID
```bash
python sheets_to_csv.py --id "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
```

### List Available Sheets
```bash
python sheets_to_csv.py --list
```

### Use Custom Configuration
```bash
python sheets_to_csv.py --config my_config.json
```

## Configuration

The tool uses a `config.json` file for settings. You can customize:

### File Naming
```json
{
  "file_naming": {
    "include_sheet_name": true,
    "include_timestamp": false,
    "custom_prefix": "export_",
    "custom_suffix": ""
  }
}
```

### Export Options
```json
{
  "export_options": {
    "include_headers": true,
    "export_all_worksheets": true,
    "worksheet_separator": "_",
    "date_format": "%Y%m%d_%H%M%S"
  }
}
```

### Output Settings
```json
{
  "output_directory": "./exports",
  "display_options": {
    "show_progress": true,
    "list_limit": 50,
    "confirm_before_export": true
  }
}
```

## Sharing with Others

This tool is designed to be easily shared. To set it up for someone else:

1. **Share the code**: Send them this entire directory (or zip it up)
2. **Exclude your credentials**: The `.gitignore` file ensures your personal credentials aren't shared
3. **They follow setup**: They run the same setup process with their own Google account

### What gets shared:
- ✅ All code files
- ✅ Configuration templates
- ✅ Documentation

### What doesn't get shared:
- ❌ Your `credentials.json`
- ❌ Your `token.json`
- ❌ Your personal `config.json`
- ❌ Exported CSV files

## File Structure

The `.gitignore` file is configured to exclude the `venv` directory.

```
google-sheets-csv-exporter/
├── venv/                     # Your virtual environment (auto-generated)
├── sheets_to_csv.py          # Main application
├── auth.py                   # Authentication utilities
├── config.py                 # Configuration management
├── setup.py                  # Setup wizard
├── requirements.txt          # Python dependencies
├── config.json.template      # Configuration template
├── .gitignore               # Git ignore rules
├── README.md                # This file
├── credentials.json         # Your Google API credentials (you create this)
├── token.json              # OAuth token (auto-generated)
├── config.json             # Your configuration (auto-generated)
└── exports/                # Default output directory
```

## Troubleshooting

### "Credentials file not found"
- Make sure you've downloaded the credentials JSON from Google Cloud Console
- Save it as `credentials.json` in the project directory
- Run `python setup.py` to verify

### "Authentication failed" 
- Check that you've enabled both Google Sheets API and Google Drive API
- Try deleting `token.json` and re-authenticating
- Ensure your Google Cloud project has the correct OAuth consent screen configured
- **If you see an "Access Denied" or "App isn't verified" error during login:**
  - In the Google Cloud Console, go to the **OAuth consent screen**.
  - Check the **Publishing status**. If it is "In production", you must click the **"Back to Testing"** button. You cannot add test users while in production mode.
  - Once the status is "Testing", find the **Test users** section and click **+ ADD USERS**.
  - Add your own Google email address to authorize it for use.

### "No sheets found"
- Verify you have Google Sheets in your Drive
- Check that the APIs have the correct permissions
- Try running `python sheets_to_csv.py --list` to debug

### "Permission denied" errors
- Make sure the `exports/` directory is writable
- Check your configured output directory in `config.json`

### Import errors
- Make sure you have activated the virtual environment (`source venv/bin/activate`)
- Run `python3 -m pip install -r requirements.txt` to install dependencies
- Make sure you're using Python 3.7 or later

## Advanced Usage

### Batch Processing
Create a script to export multiple sheets automatically:

```bash
#!/bin/bash
python sheets_to_csv.py --name "Sales Q1"
python sheets_to_csv.py --name "Sales Q2"
python sheets_to_csv.py --name "Sales Q3"
python sheets_to_csv.py --name "Sales Q4"
```

### Custom Configuration
Create different configurations for different use cases:

```bash
# Personal exports
python sheets_to_csv.py --config personal_config.json

# Work exports
python sheets_to_csv.py --config work_config.json
```

### Automated Backups
Set up scheduled exports using cron (Linux/macOS) or Task Scheduler (Windows):

```bash
# Daily backup at 2 AM
0 2 * * * /usr/bin/python3 /path/to/sheets_to_csv.py --name "Important Data"
```

## Security Notes

- Your `credentials.json` and `token.json` files contain sensitive authentication information
- Never commit these files to version control
- Don't share these files with others
- The tool only requests read-only access to your Google Sheets and Drive

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## License

This project is provided as-is for personal and educational use. 