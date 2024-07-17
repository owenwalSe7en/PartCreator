# PartCreator - ERP Automation Tool

The ERP Part Management System is a Python-based application designed to interact with an Enterprise Resource Planning (ERP)
system for managing part information. It provides a graphical user interface for creating, overwriting, and deleting part
information in bulk using data from Excel files.

## Features
 - Create new parts in ERP system
 - Overwrite existing part information
 - Delete parts from the ERP system
 - User-friendly graphical interface
 - Validation of input data

## Installation
1. **Clone the repository**
    ```bash
    git clone https://github.com/owenwalSe7en/PartCreator.git
    cd PartCreator
    ```

2. **Create a virtual environment**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. **Install dependencies**
    ```bash
    pip install -r requirements.txt
    ```

## Usage
1. **Run the application**
    ```bash
    python main.py
    ```
2. **Configure the Excel spreadsheet**
   - **Select Excel File**: Use the provided UI to navigate and select the Excel file containing the part numbers.
   - **Specify Parameters**: Follow the usage procedure and enter the necessary details for the program to connect to the Excel file and the ERP system.
  
3. **Start Automation**
    - Click the "Submit" button in the UI to begin the automation process. The tool will read the part numbers from the selected Excel file and input them into the ERP system.
  
## Dependencies
- **pywinauto**
- **openpyxl**
- **tkinter**

## File Structure
- `main.py` - The main scripts running the program 
- `erp_manager.py` - The managing class for Epicor access and operation functionality
- `forms.py` - UI/UX main file controlling the flow of `tkinter` forms
- `application.py` - Initial operation selection and general code flow manager
- `combobox_options.py` - Contains global variabled for the combobox options
- `requirements.txt` - Lists the Python dependencies required for the project

## Contributing

Contributions are welcome! Please follow these steps to contribute:

1. **Fork the repository**
2. **Create a new branch** for your feature or bugfix
    ```bash
    git checkout -b feature/your-feature-name
    ```
3. **Commit your changes**
    ```bash
    git commit -m 'Add some feature'
    ```
4. **Push to the branch**
    ```bash
    git push origin feature/your-feature-name
    ```
5. **Create a new Pull Request**

## Contact

For any questions or suggestions, please feel free to open an issue or contact me at [wallaceowenh45@gmail.com](mailto:your-email@example.com).
