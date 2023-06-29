import platform
from enum import Enum

class System(Enum):
    Linux = 0
    Mac = 1
    Windows = 2

def detect_os():
    sys = platform.system()
    if sys == 'Windows':
        response = System.Windows
    elif sys == 'Darwin':
        response = System.Darwin
    else:
        response = System.Linux
    return response
