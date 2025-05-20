import configparser
import os

class ConfigManager:
    def __init__(self, ini_path=None):
        if ini_path is None:
            ini_path = os.path.join(os.path.dirname(__file__), "..", "setup.ini")
        self.ini_path = os.path.abspath(ini_path)
        self.config = configparser.ConfigParser()
        self.load()
    
    def load(self):
        self.config.read(self.ini_path, encoding='utf-8')
    
    def save(self):
        with open(self.ini_path, 'w', encoding='utf-8') as f:
            self.config.write(f)
    
    def get(self, section, option, fallback=None):
        return self.config.get(section, option, fallback=fallback)
    
    def set(self, section, option, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, value)
        self.save()
    
    def remove_option(self, section, option):
        if self.config.has_section(section):
            self.config.remove_option(section, option)
            self.save()
