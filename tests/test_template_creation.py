import sys
import os

# Add the project root directory to the Python path
project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, project_root)

from src.new_participant_template import ExcelTemplateHandler

def test_create_template():
    handler = ExcelTemplateHandler()
    output_path = "data/test_participant_template.xlsx"
    handler.create_template(output_path)
    print(f"Template created at: {output_path}")

if __name__ == "__main__":
    test_create_template()