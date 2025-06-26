#!/usr/bin/env python3
"""
Setup wizard for Google Sheets to CSV Exporter
Helps users configure the tool for first-time use.
"""

import os
import json
import sys
from pathlib import Path
from config import Config

def check_dependencies():
    """Check if required dependencies are installed"""
    required_packages = {
        'google': 'google-auth',
        'google_auth_oauthlib': 'google-auth-oauthlib',
        'googleapiclient': 'google-api-python-client'
    }
    
    missing_packages = []
    for import_name, package_name in required_packages.items():
        try:
            __import__(import_name)
        except ImportError:
            missing_packages.append(package_name)
    
    return missing_packages

def create_directories():
    """Create necessary directories"""
    directories = ['exports']
    
    for directory in directories:
        Path(directory).mkdir(exist_ok=True)
        print(f"✓ Created directory: {directory}")

def setup_config():
    """Setup configuration file"""
    config = Config()
    
    print("\n" + "="*60)
    print("CONFIGURATION SETUP")
    print("="*60)
    
    # Output directory
    default_output = config.get('output_directory')
    output_dir = input(f"Output directory for CSV files [{default_output}]: ").strip()
    if output_dir:
        config.set('output_directory', output_dir)
    
    # File naming preferences
    print("\n--- File Naming Options ---")
    
    include_timestamp = input("Include timestamp in filenames? (y/N): ").strip().lower() == 'y'
    config.set('file_naming.include_timestamp', include_timestamp)
    
    custom_prefix = input("Custom prefix for filenames (optional): ").strip()
    if custom_prefix:
        config.set('file_naming.custom_prefix', custom_prefix)
    
    # Export options
    print("\n--- Export Options ---")
    
    export_all = input("Export all worksheets by default? (Y/n): ").strip().lower()
    export_all = export_all != 'n'
    config.set('export_options.export_all_worksheets', export_all)
    
    # Save configuration
    if config.save_config():
        print("✓ Configuration saved successfully")
    else:
        print("✗ Failed to save configuration")
    
    return True

def check_credentials():
    """Check for Google API credentials"""
    print("\n" + "="*60)
    print("GOOGLE API CREDENTIALS CHECK")
    print("="*60)
    
    creds_file = 'credentials.json'
    
    if os.path.exists(creds_file):
        print(f"✓ Found credentials file: {creds_file}")
        return True
    else:
        print(f"✗ Credentials file not found: {creds_file}")
        print("\nTo use this tool, you need to:")
        print("1. Go to the Google Cloud Console (https://console.cloud.google.com/)")
        print("2. Create a new project or select an existing one")
        print("3. Enable the Google Sheets API and Google Drive API")
        print("4. Create credentials (OAuth 2.0 Client ID)")
        print("5. Download the credentials JSON file")
        print(f"6. Save it as '{creds_file}' in this directory")
        print("\nSee README.md for detailed instructions.")
        return False

def test_authentication():
    """Test Google API authentication"""
    print("\n" + "="*60)
    print("AUTHENTICATION TEST")
    print("="*60)
    
    try:
        from auth import GoogleSheetsAuth
        
        auth = GoogleSheetsAuth()
        print("Testing authentication...")
        
        # This will trigger the OAuth flow if needed
        creds = auth.authenticate()
        
        if creds and creds.valid:
            print("✓ Authentication successful!")
            
            # Test API access
            sheets_service = auth.get_sheets_service()
            drive_service = auth.get_drive_service()
            
            # Try to list a few sheets
            results = drive_service.files().list(
                q="mimeType='application/vnd.google-apps.spreadsheet'",
                pageSize=5,
                fields="files(name)"
            ).execute()
            
            sheets = results.get('files', [])
            if sheets:
                print(f"✓ Found {len(sheets)} Google Sheets in your Drive")
                print("Authentication and API access working correctly!")
            else:
                print("✓ API access working, but no Google Sheets found in your Drive")
            
            return True
        else:
            print("✗ Authentication failed")
            return False
            
    except FileNotFoundError:
        print("✗ Credentials file not found")
        return False
    except Exception as e:
        print(f"✗ Authentication error: {e}")
        return False

def run_setup():
    """Run the complete setup wizard"""
    print("Google Sheets to CSV Exporter - Setup Wizard")
    print("=" * 60)
    
    # Check dependencies
    print("\n1. Checking dependencies...")
    missing = check_dependencies()
    if missing:
        print(f"✗ Missing required packages: {', '.join(missing)}")
        print("Please make sure you have activated your virtual environment and run:")
        print("  python3 -m pip install -r requirements.txt")
        return False
    else:
        print("✓ All required dependencies are installed")
    
    # Create directories
    print("\n2. Creating directories...")
    create_directories()
    
    # Setup configuration
    print("\n3. Setting up configuration...")
    setup_config()
    
    # Check credentials
    print("\n4. Checking credentials...")
    if not check_credentials():
        print("\n⚠️  Setup incomplete. Please follow the credential setup instructions.")
        return False
    
    # Test authentication
    print("\n5. Testing authentication...")
    if not test_authentication():
        print("\n⚠️  Authentication test failed. Please check your credentials.")
        return False
    
    print("\n" + "="*60)
    print("🎉 SETUP COMPLETE!")
    print("="*60)
    print("\nYou can now use the Google Sheets to CSV Exporter:")
    print("  python sheets_to_csv.py                    # Interactive mode")
    print("  python sheets_to_csv.py --list             # List your sheets")
    print("  python sheets_to_csv.py --name 'My Sheet'  # Export by name")
    print("\nFor more options, run: python sheets_to_csv.py --help")
    
    return True

def main():
    """Main entry point for setup script"""
    if len(sys.argv) > 1 and sys.argv[1] == '--quick':
        # Quick setup - minimal prompts
        print("Running quick setup...")
        create_directories()
        Config().save_config()  # Save default config
        print("Quick setup complete. Run with --setup for full configuration.")
    else:
        run_setup()

if __name__ == '__main__':
    main()