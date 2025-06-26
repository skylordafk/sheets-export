import os
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

class GoogleSheetsAuth:
    """Handle Google Sheets API authentication"""
    
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets.readonly',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
    
    def __init__(self, credentials_file='credentials.json', token_file='token.json'):
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.creds = None
        
    def authenticate(self):
        """Authenticate and return valid credentials"""
        # Load existing token if available
        if os.path.exists(self.token_file):
            self.creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)
        
        # If there are no valid credentials available, request authorization
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                try:
                    self.creds.refresh(Request())
                except Exception as e:
                    print(f"Error refreshing token: {e}")
                    return self._get_new_credentials()
            else:
                return self._get_new_credentials()
            
            # Save the credentials for the next run
            with open(self.token_file, 'w') as token:
                token.write(self.creds.to_json())
                
        return self.creds
    
    def _get_new_credentials(self):
        """Get new credentials via OAuth flow"""
        if not os.path.exists(self.credentials_file):
            raise FileNotFoundError(
                f"Credentials file '{self.credentials_file}' not found. "
                "Please follow the setup instructions to create this file."
            )
        
        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                self.credentials_file, self.SCOPES
            )
            self.creds = flow.run_local_server(port=0)
            
            # Save the credentials
            with open(self.token_file, 'w') as token:
                token.write(self.creds.to_json())
                
            return self.creds
        except Exception as e:
            raise Exception(f"Authentication failed: {e}")
    
    def get_sheets_service(self):
        """Get authenticated Google Sheets service"""
        creds = self.authenticate()
        return build('sheets', 'v4', credentials=creds)
    
    def get_drive_service(self):
        """Get authenticated Google Drive service"""
        creds = self.authenticate()
        return build('drive', 'v3', credentials=creds)
    
    def is_authenticated(self):
        """Check if user is authenticated"""
        return os.path.exists(self.token_file) and self.creds and self.creds.valid
    
    def clear_credentials(self):
        """Clear stored credentials (for re-authentication)"""
        if os.path.exists(self.token_file):
            os.remove(self.token_file)
        self.creds = None