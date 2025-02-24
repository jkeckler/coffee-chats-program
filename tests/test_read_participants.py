import os
import sys

# Add the project root directory to the Python path
project_root = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, project_root)

from src.excel_handler import ExcelTemplateHandler

def test_read_participants():
    # Specify the path to your test Excel file
    filepath = "data/Coffee_Chats_ParticipantsList.xlsx"
    
    print(f"Reading participants from: {filepath}")
    
    handler = ExcelTemplateHandler()
    participants = handler.read_participants_from_excel(filepath)
    
    print(f"\nTotal participants read: {len(participants)}")
    
    if participants:
        print("\nSample participant data:")
        sample = participants[0]
        print(f"Name: {sample['name']}")
        print(f"Department: {sample['department']}")
        print(f"Timezone: {sample['timezone']}")
        print("Availability:")
        for hour, available in sample['availability'].items():
            if available == 'Y':
                print(f"  {hour}:00")

if __name__ == "__main__":
    test_read_participants()