import os
from datetime import *

class log_file():
    def debug(self, log_file, message):
        if os.path.isfile(log_file):
            file = open(log_file, "a+")
        else:
            file = open(log_file, "w+")
    
        time_now = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        file.write(time_now + " - DEBUG: " + message + "\n")
    
        file.close()
        
    def info(self, log_file, message):
        if os.path.isfile(log_file):
            file = open(log_file, "a+")
        else:
            file = open(log_file, "w+")
    
        time_now = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        file.write(time_now + " - INFO: " + message + "\n")
    
        file.close()
        
    def warning(self, log_file, message):
        if os.path.isfile(log_file):
            file = open(log_file, "a+")
        else:
            file = open(log_file, "w+")
    
        time_now = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        file.write(time_now + " - WARNING: " + message + "\n")
    
        file.close()

    def error(self, log_file, message):
        if os.path.isfile(log_file):
            file = open(log_file, "a+")
        else:
            file = open(log_file, "w+")
    
        time_now = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        file.write(time_now + " - ERROR: " + message + "\n")
    
        file.close()
        
    def force_restart_info(self, log_file, message):
        file = open(log_file, "w+")

        time_now = str(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        file.write(time_now + " - INFO: " + message + "\n")

        file.close()        