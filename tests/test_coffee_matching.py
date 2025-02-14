from src.excel_handler import ExcelTemplateHandler
from src.coffee_matching import GroupMatcher, run_matching_program
import random

def generate_test_data(num_participants=20):
    handler = ExcelTemplateHandler()
    template_path = "data/test_participant_template.xlsx"
    handler.create_template(template_path)
    
    # Generate some random test data
    departments = ["Engineering", "Marketing", "Sales", "Finance", "HR"]
    countries = ["United States", "India", "United Kingdom", "Germany", "Japan"]
    timezones = ["UTC-5", "UTC+5:30", "UTC+0", "UTC+1", "UTC+9"]
    
    wb = handler.read_excel(template_path)
    ws = wb['Participants']
    
    for i in range(2, num_participants + 2):
        ws.cell(row=i, column=1, value=f"EMP{i-1}")  # Employee ID
        ws.cell(row=i, column=2, value=f"Employee {i-1}")  # Name
        ws.cell(row=i, column=3, value=f"employee{i-1}@example.com")  # Email
        ws.cell(row=i, column=4, value=random.choice(departments))
        ws.cell(row=i, column=5, value=random.choice(countries))
        ws.cell(row=i, column=6, value=random.choice(timezones))
        ws.cell(row=i, column=7, value="2023-01-01")  # Start Date
        ws.cell(row=i, column=8, value="Active")  # Status
        
        # Set random availability
        for j in range(9, 33):
            ws.cell(row=i, column=j, value=random.choice(['Y', 'N']))
    
    wb.save(template_path)
    print(f"Test data generated and saved to {template_path}")
    return template_path

def test_matching_process():
    # Generate test data
    input_path = generate_test_data()
    
    # Run the matching program
    output_path = "data/test_coffee_chat_matches.xlsx"
    run_matching_program(input_path, output_path)
    
    # Verify results
    handler = ExcelTemplateHandler()
    results = handler.read_excel(output_path)
    if results:
        print("Matching process completed successfully.")
        print(f"Results saved to {output_path}")
    else:
        print("Matching process failed or produced no results.")

if __name__ == "__main__":
    test_matching_process()