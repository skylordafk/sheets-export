#!/usr/bin/env python3
"""
Google Sheets to CSV Exporter
A configurable tool to export Google Sheets from your Drive to local CSV files.
"""

import argparse
import sys
import os
import csv
from pathlib import Path
from typing import List, Dict, Optional

from auth import GoogleSheetsAuth
from config import Config

class SheetsToCSVExporter:
    """Main class for exporting Google Sheets to CSV"""
    
    def __init__(self):
        self.config = Config()
        self.auth = GoogleSheetsAuth()
        self.sheets_service = None
        self.drive_service = None
    
    def initialize_services(self):
        """Initialize Google API services"""
        try:
            print("Authenticating with Google APIs...")
            self.sheets_service = self.auth.get_sheets_service()
            self.drive_service = self.auth.get_drive_service()
            print("✓ Authentication successful")
            return True
        except Exception as e:
            print(f"✗ Authentication failed: {e}")
            return False
    
    def list_sheets(self, limit: Optional[int] = None) -> List[Dict]:
        """List all Google Sheets in user's Drive"""
        if not self.drive_service:
            raise Exception("Drive service not initialized")
        
        limit = limit or self.config.get('display_options.list_limit', 50)
        
        try:
            query = "mimeType='application/vnd.google-apps.spreadsheet'"
            results = self.drive_service.files().list(
                q=query,
                pageSize=limit,
                fields="files(id, name, modifiedTime, webViewLink)"
            ).execute()
            
            sheets = results.get('files', [])
            return sheets
        except Exception as e:
            print(f"Error listing sheets: {e}")
            return []
    
    def display_sheets(self, sheets: List[Dict]):
        """Display available sheets in a user-friendly format"""
        if not sheets:
            print("No Google Sheets found in your Drive.")
            return
        
        print(f"\nFound {len(sheets)} Google Sheets:")
        print("-" * 80)
        
        for i, sheet in enumerate(sheets, 1):
            name = sheet['name']
            modified = sheet.get('modifiedTime', 'Unknown')[:10]  # Just date part
            print(f"{i:2d}. {name}")
            print(f"    Modified: {modified}")
            print(f"    ID: {sheet['id']}")
            print()
    
    def get_sheet_worksheets(self, sheet_id: str) -> List[Dict]:
        """Get all worksheets within a Google Sheet"""
        if not self.sheets_service:
            raise Exception("Sheets service not initialized")
        
        try:
            spreadsheet = self.sheets_service.spreadsheets().get(
                spreadsheetId=sheet_id
            ).execute()
            
            worksheets = []
            for sheet in spreadsheet.get('sheets', []):
                properties = sheet.get('properties', {})
                worksheets.append({
                    'title': properties.get('title', 'Untitled'),
                    'id': properties.get('sheetId'),
                    'index': properties.get('index', 0)
                })
            
            return sorted(worksheets, key=lambda x: x['index'])
        except Exception as e:
            print(f"Error getting worksheets: {e}")
            return []
    
    def export_worksheet_to_csv(self, sheet_id: str, worksheet_name: str, 
                               output_path: str) -> bool:
        """Export a single worksheet to CSV"""
        if not self.sheets_service:
            raise Exception("Sheets service not initialized")
        
        try:
            # Get worksheet data
            range_name = f"'{worksheet_name}'"  # Quote sheet name in case of spaces
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            if not values:
                print(f"Warning: '{worksheet_name}' appears to be empty")
                return False
            
            # Write to CSV
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                for row in values:
                    # Pad row to ensure consistent column count
                    max_cols = max(len(row) for row in values) if values else 0
                    padded_row = row + [''] * (max_cols - len(row))
                    writer.writerow(padded_row)
            
            print(f"✓ Exported '{worksheet_name}' to {output_path}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to export '{worksheet_name}': {e}")
            return False
    
    def export_sheet(self, sheet_id: str, sheet_name: str) -> bool:
        """Export a Google Sheet (all worksheets) to CSV files"""
        worksheets = self.get_sheet_worksheets(sheet_id)
        if not worksheets:
            print(f"No worksheets found in '{sheet_name}'")
            return False
        
        output_dir = self.config.get_output_directory()
        export_all = self.config.get('export_options.export_all_worksheets', True)
        success_count = 0
        
        if len(worksheets) == 1 or not export_all:
            # Single worksheet or user preference
            worksheet = worksheets[0]
            filename = self.config.get_filename(sheet_name)
            output_path = os.path.join(output_dir, filename)
            
            if self.export_worksheet_to_csv(sheet_id, worksheet['title'], output_path):
                success_count += 1
        else:
            # Multiple worksheets
            for worksheet in worksheets:
                filename = self.config.get_filename(sheet_name, worksheet['title'])
                output_path = os.path.join(output_dir, filename)
                
                if self.export_worksheet_to_csv(sheet_id, worksheet['title'], output_path):
                    success_count += 1
        
        print(f"Successfully exported {success_count}/{len(worksheets)} worksheets from '{sheet_name}'")
        return success_count > 0
    
    def interactive_mode(self):
        """Run in interactive mode for sheet selection"""
        if not self.initialize_services():
            return False
        
        print("🔍 Fetching your Google Sheets...")
        sheets = self.list_sheets()
        
        if not sheets:
            return False
        
        self.display_sheets(sheets)
        
        while True:
            try:
                choice = input("\nEnter sheet number to export (or 'q' to quit): ").strip()
                
                if choice.lower() == 'q':
                    print("Goodbye!")
                    return True
                
                sheet_index = int(choice) - 1
                if 0 <= sheet_index < len(sheets):
                    sheet = sheets[sheet_index]
                    
                    if self.config.get('display_options.confirm_before_export', True):
                        confirm = input(f"Export '{sheet['name']}'? (y/N): ").strip().lower()
                        if confirm != 'y':
                            continue
                    
                    print(f"\n📤 Exporting '{sheet['name']}'...")
                    self.export_sheet(sheet['id'], sheet['name'])
                    
                    another = input("\nExport another sheet? (y/N): ").strip().lower()
                    if another != 'y':
                        break
                else:
                    print("Invalid selection. Please try again.")
                    
            except ValueError:
                print("Please enter a valid number.")
            except KeyboardInterrupt:
                print("\n\nOperation cancelled by user.")
                return True
        
        return True
    
    def export_by_name(self, sheet_names: List[str]) -> bool:
        """Export sheets by name"""
        if not self.initialize_services():
            return False
        
        print("🔍 Searching for sheets...")
        all_sheets = self.list_sheets(limit=1000)  # Get more sheets for name matching
        
        success_count = 0
        for sheet_name in sheet_names:
            # Find sheet by name (case-insensitive)
            matching_sheet = None
            for sheet in all_sheets:
                if sheet['name'].lower() == sheet_name.lower():
                    matching_sheet = sheet
                    break
            
            if matching_sheet:
                print(f"\n📤 Exporting '{matching_sheet['name']}'...")
                if self.export_sheet(matching_sheet['id'], matching_sheet['name']):
                    success_count += 1
            else:
                print(f"✗ Sheet '{sheet_name}' not found")
        
        print(f"\nExported {success_count}/{len(sheet_names)} sheets successfully")
        return success_count > 0
    
    def export_by_id(self, sheet_ids: List[str]) -> bool:
        """Export sheets by ID"""
        if not self.initialize_services():
            return False
        
        success_count = 0
        for sheet_id in sheet_ids:
            try:
                # Get sheet name
                spreadsheet = self.sheets_service.spreadsheets().get(
                    spreadsheetId=sheet_id,
                    fields='properties.title'
                ).execute()
                
                sheet_name = spreadsheet['properties']['title']
                print(f"\n📤 Exporting '{sheet_name}'...")
                
                if self.export_sheet(sheet_id, sheet_name):
                    success_count += 1
                    
            except Exception as e:
                print(f"✗ Failed to export sheet {sheet_id}: {e}")
        
        print(f"\nExported {success_count}/{len(sheet_ids)} sheets successfully")
        return success_count > 0

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Export Google Sheets to CSV files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sheets_to_csv.py                          # Interactive mode
  python sheets_to_csv.py --name "My Sheet"        # Export by name
  python sheets_to_csv.py --id "1BxiMVs0..."      # Export by ID
  python sheets_to_csv.py --list                   # List available sheets
  python sheets_to_csv.py --setup                  # Run setup wizard
        """
    )
    
    parser.add_argument('--name', '-n', action='append',
                       help='Export sheet by name (can be used multiple times)')
    parser.add_argument('--id', '-i', action='append',
                       help='Export sheet by ID (can be used multiple times)')
    parser.add_argument('--list', '-l', action='store_true',
                       help='List available sheets and exit')
    parser.add_argument('--setup', action='store_true',
                       help='Run initial setup wizard')
    parser.add_argument('--config', help='Use custom config file')
    
    args = parser.parse_args()
    
    # Initialize exporter
    if args.config:
        exporter = SheetsToCSVExporter()
        exporter.config = Config(args.config)
    else:
        exporter = SheetsToCSVExporter()
    
    # Handle setup
    if args.setup:
        from setup import run_setup
        return run_setup()
    
    # Handle list command
    if args.list:
        if not exporter.initialize_services():
            return 1
        sheets = exporter.list_sheets()
        exporter.display_sheets(sheets)
        return 0
    
    # Handle export by name
    if args.name:
        success = exporter.export_by_name(args.name)
        return 0 if success else 1
    
    # Handle export by ID
    if args.id:
        success = exporter.export_by_id(args.id)
        return 0 if success else 1
    
    # Default to interactive mode
    success = exporter.interactive_mode()
    return 0 if success else 1

if __name__ == '__main__':
    sys.exit(main())