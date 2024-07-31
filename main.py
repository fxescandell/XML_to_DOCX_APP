import sys
import os
from gui import start_gui

def main():
    log_file = os.path.join(os.path.expanduser("~"), "app_log.txt")
    with open(log_file, "w") as f:
        sys.stdout = f
        sys.stderr = f
        print(f"Log file created at: {log_file}")
        start_gui()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        with open(os.path.join(os.path.expanduser("~"), "app_log.txt"), "a") as f:
            f.write(f"Unhandled exception: {e}\n")
        raise