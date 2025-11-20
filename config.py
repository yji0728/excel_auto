"""
Configuration Module
Manages application configuration and settings.
"""
import json
import os
from typing import Dict, Any


class Config:
    """Configuration management class"""
    
    DEFAULT_CONFIG = {
        'app_name': 'Excel Auto',
        'version': '1.0.0',
        'window_size': '1000x700',
        'default_workflow_dir': './workflows',
        'recent_files': [],
        'max_recent_files': 10,
        'theme': 'default',
        'language': 'ko',
        'auto_save': True,
        'auto_save_interval': 300  # seconds
    }
    
    def __init__(self, config_file: str = 'config.json'):
        self.config_file = config_file
        self.config = self.DEFAULT_CONFIG.copy()
        self.load()
    
    def load(self) -> bool:
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                    self.config.update(loaded_config)
                return True
            return False
        except Exception as e:
            print(f"Error loading config: {e}")
            return False
    
    def save(self) -> bool:
        """Save configuration to file"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving config: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value"""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """Set a configuration value"""
        self.config[key] = value
    
    def add_recent_file(self, file_path: str) -> None:
        """Add a file to recent files list"""
        recent = self.config.get('recent_files', [])
        if file_path in recent:
            recent.remove(file_path)
        recent.insert(0, file_path)
        recent = recent[:self.config.get('max_recent_files', 10)]
        self.config['recent_files'] = recent
        self.save()
    
    def get_recent_files(self) -> list:
        """Get list of recent files"""
        return self.config.get('recent_files', [])
    
    def clear_recent_files(self) -> None:
        """Clear recent files list"""
        self.config['recent_files'] = []
        self.save()
