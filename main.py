from erp_manager import ERPManager, CreateOperation, OverwriteOperation, DeleteOperation
from application import Application

if __name__ == "__main__":
    erp_manager = ERPManager(
        CreateOperation(),
        OverwriteOperation(),
        DeleteOperation()
    )
    app = Application(erp_manager)
    app.run()
