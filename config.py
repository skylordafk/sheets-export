import json
import os
from pathlib import Path

class Config:
    """Configuration management for Google Sheets to CSV exporter"""
    
    DEFAULT_CONFIG = {
        "output_directory": "./exports",
        "file_naming": {
            "include_sheet_name": True,
            "include_timestamp": False,
            "custom_prefix": "",
            "custom_suffix": ""
        },
        "export_options": {
            "include_headers": True,
            "export_all_worksheets": True,
            "worksheet_separator": "_",
            "date_format": "%Y%m%d_%H%M%S"
        },
        "display_options": {
            "show_progress": True,
            "list_limit": 50,
            "confirm_before_export": True
        }
    }
    
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self):
        """Load configuration from file or create default"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    user_config = json.load(f)
                # Merge with defaults to ensure all keys exist
                config = self.DEFAULT_CONFIG.copy()
                self._deep_update(config, user_config)
                return config
            except (json.JSONDecodeError, Exception) as e:
                print(f"Error loading config file: {e}")
                print("Using default configuration...")
                return self.DEFAULT_CONFIG.copy()
        else:
            return self.DEFAULT_CONFIG.copy()
    
    def save_config(self):
        """Save current configuration to file"""
        try:
            with open(self.config_file, 'w') as f:
                json.dump(self.config, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    def create_template(self, template_file='config.json.template'):
        """Create a configuration template file"""
        try:
            with open(template_file, 'w') as f:
                json.dump(self.DEFAULT_CONFIG, f, indent=2)
            return True
        except Exception as e:
            print(f"Error creating config template: {e}")
            return False
    
    def get(self, key, default=None):
        """Get configuration value using dot notation"""
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key, value):
        """Set configuration value using dot notation"""
        keys = key.split('.')
        target = self.config
        
        for k in keys[:-1]:
            if k not in target:
                target[k] = {}
            target = target[k]
        
        target[keys[-1]] = value
    
    def get_output_directory(self):
        """Get and ensure output directory exists"""
        output_dir = self.get('output_directory')
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        return output_dir
    
    def get_filename(self, sheet_name, worksheet_name=None):
        """Generate filename based on configuration"""
        filename_parts = []
        
        # Add custom prefix
        prefix = self.get('file_naming.custom_prefix')
        if prefix:
            filename_parts.append(prefix)
        
        # Add sheet name
        if self.get('file_naming.include_sheet_name'):
            filename_parts.append(self._sanitize_filename(sheet_name))
        
        # Add worksheet name if specified
        if worksheet_name:
            separator = self.get('export_options.worksheet_separator', '_')
            filename_parts.append(f"{separator}{self._sanitize_filename(worksheet_name)}")
        
        # Add timestamp if requested
        if self.get('file_naming.include_timestamp'):
            from datetime import datetime
            date_format = self.get('export_options.date_format')
            timestamp = datetime.now().strftime(date_format)
            filename_parts.append(timestamp)
        
        # Add custom suffix
        suffix = self.get('file_naming.custom_suffix')
        if suffix:
            filename_parts.append(suffix)
        
        # Join parts and add extension
        filename = '_'.join(filter(None, filename_parts))
        return f"{filename}.csv"
    
    def _sanitize_filename(self, name):
        """Remove or replace invalid filename characters"""
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            name = name.replace(char, '_')
        return name.strip()
    
    def _deep_update(self, target, source):
        """Deep update dictionary"""
        for key, value in source.items():
            if key in target and isinstance(target[key], dict) and isinstance(value, dict):
                self._deep_update(target[key], value)
            else:
                target[key] = value
    
    def print_config(self):
        """Print current configuration"""
        print("Current Configuration:")
        print(json.dumps(self.config, indent=2))