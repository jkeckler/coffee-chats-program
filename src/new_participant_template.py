import os
import sys

# Add the src directory to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from src.excel_handler import ExcelTemplateHandler

def main():
    handler = ExcelTemplateHandler()
    handler.create_template()
    print("Template created successfully")

if __name__ == "__main__":
    main()