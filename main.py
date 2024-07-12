from erp_manager import ERPManager, CreateOperation, OverwriteOperation, DeleteOperation
from application import Application
import sys
import os


def print_debug_info():
    print("Python version:", sys.version)
    print("Executable:", sys.executable)
    print("Current working directory:", os.getcwd())
    print("Contents of current directory:")
    for item in os.listdir():
        print(f"  - {item}")
    print("\nSystem path:")
    for path in sys.path:
        print(f"  - {path}")


# Your existing imports and code here...
if __name__ == "__main__":
    try:
        print_debug_info()
        # Start the program
        erp_manager = ERPManager(
            CreateOperation(),
            OverwriteOperation(),
            DeleteOperation()
        )
        app = Application(erp_manager)
        app.run()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        input("Press Enter to exit...")






