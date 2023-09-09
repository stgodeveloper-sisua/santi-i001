# file_name:   monitoring.py
# created_on:  2023-06-30 ; vicente.diaz
# modified_on: 2023-06-30 ; vicente.diaz

# This is a dummy file only for dev environment!
# This file must be replaced by the production monitoring.py file that uploads the execution logs.

class Monitoring:
    def __init__(self, config:dict) -> None:
        self.config= config

    def uplog(self):
        pass

if __name__ == "__main__":
    
    monitor = Monitoring(config="")
    monitor.uplog()

