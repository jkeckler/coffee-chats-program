from src.utils import convert_timezone_to_float

# Test cases
test_cases = [
    ("UTC-05:00", -5.0),    # Your format (US Eastern)
    ("UTC+05:30", 5.5),     # India
    ("UTC+00:00", 0.0),     # UK
    ("UTC-08:00", -8.0),    # US Pacific
]

# Run tests
for timezone, expected in test_cases:
    result = convert_timezone_to_float(timezone)
    print(f"Testing {timezone}:")
    print(f"  Expected: {expected}")
    print(f"  Got: {result}")
    print(f"  Pass: {result == expected}\n")