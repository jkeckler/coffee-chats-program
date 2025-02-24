from src.excel_handler import ExcelTemplateHandler

def test_excel_handler():
    # Create instance of handler
    handler = ExcelTemplateHandler()
    
    # Test department
    print("\nTesting Department:")
    print(f"Expected: ['Business Intelligence']")
    print(f"Got: {handler.departments}")
    
    # Test some timezone formats
    print("\nTesting Sample Timezones:")
    print("US Eastern timezone format:")
    print(f"Expected: 'UTC-05:00'")
    print(f"Found: {any('UTC-05:00' in tz for tz in handler.timezones)}")
    
    # Test India timezone
    print("\nIndia timezone format:")
    print(f"Expected: 'UTC+05:30'")
    print(f"Found: {any('UTC+05:30' in tz for tz in handler.timezones)}")
    
    # Print all available timezones
    print("\nAll available timezones:")
    for tz in sorted(handler.timezones):
        print(f"  {tz}")

if __name__ == "__main__":
    test_excel_handler()