import openpyxl
from openpyxl.styles import Font, PatternFill
from typing import List, Dict
import os

class ExcelTemplateHandler:
    def create_template(self, output_path: str = "participant_template.xlsx") -> None:
        """
        Create a new Excel template file for participant data.
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Participants"
        
        # Define headers
        headers = [
            "Employee ID",
            "Name",
            "Email",
            "Department",
            "Available Hours",
            "Captain Eligible",
            "Status"
        ]
        
        # Style for headers
        header_font = Font(bold=True)
        header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        
        # Add headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
        
        # Add example row
        example_data = [
            "EMP001",
            "John Doe",
            "john.doe@company.com",
            "Engineering",
            "9-17",  # 9 AM to 5 PM
            "Yes",
            "Active"
        ]
        
        for col, value in enumerate(example_data, 1):
            ws.cell(row=2, column=col, value=value)
        
        # Adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            
            for cell in col:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            ws.column_dimensions[column].width = max_length + 2

        wb.save(output_path)
        print(f"Template created successfully at: {output_path}")

    def read_participants(self, file_path: str) -> List[Dict]:
        """
        Read participant data from Excel file.
        Returns a list of dictionaries containing participant data.
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found at: {file_path}")
        
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        
        # Get headers
        headers = [cell.value for cell in ws[1]]
        
        # Read data
        participants = []
        for row in ws.iter_rows(min_row=2):  # Skip header row
            # Skip completely empty rows
            if all(cell.value is None for cell in row):
                continue
                
            participant = {}
            for header, cell in zip(headers, row):
                participant[header] = cell.value
            
            participants.append(participant)
            
        return participants

if __name__ == "__main__":
    # Example usage
    handler = ExcelTemplateHandler()
    
    # Create template
    handler.create_template()
    
    # You can test reading it back
    # participants = handler.read_participants("participant_template.xlsx")
    # print("Read participants:", participants)