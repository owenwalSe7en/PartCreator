from enum import Enum
from abc import ABC, abstractmethod


class OperationType(Enum):
    # Shared dictionaries
    file_data = {}
    label_data = {}

    OVERWRITE = 1
    DELETE = 2
    CREATE = 3


class Operation(ABC):
    @abstractmethod
    def execute(self):
        pass


class OverwriteOperation(Operation):
    def execute(self):
        print(f"Overwriting part: ")


class DeleteOperation(Operation):
    def execute(self):
        print(f"Deleting part: ")


class CreateOperation(Operation):
    def execute(self):
        print(f"Creating part: ")


class ERPManager:
    def __init__(self, overwrite_op: Operation, delete_op: Operation, create_op: Operation):
        self.operations = {
            OperationType.OVERWRITE: overwrite_op,
            OperationType.DELETE: delete_op,
            OperationType.CREATE: create_op
        }

    def perform_operation(self, op_type: OperationType):
        operation = self.operations.get(op_type)
        if operation:
            operation.execute()
        else:
            raise ValueError("Invalid operation type")
