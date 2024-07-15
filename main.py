from erp_manager import ERPManager, CreateOperation, OverwriteOperation, DeleteOperation
from application import Application


if __name__ == "__main__":
    try:
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






