import os
from datetime import datetime

class Logger:
    def __init__(self, log_path=None):
        if log_path is None:
            log_path = os.path.join(os.path.dirname(__file__), "..", "message.log")
        self.log_path = os.path.abspath(log_path)
    
    def log(self, message):
        with open(self.log_path, "a", encoding="utf-8") as f:
            f.write("\n----------\n")
            f.write(f"Start {self.log_path} {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(message + "\n")
