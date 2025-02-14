"""
Utility functions for the coffee chat matching system.
Contains timezone conversion and other helper functions used across the application.
"""


from typing import Dict

def convert_timezone_to_float(timezone_str: str) -> float:
    """
    Convert timezone string (e.g., 'UTC+5:30') to decimal hours (e.g., 5.5)
    
    Args:
        timezone_str: String in format 'UTC±H' or 'UTC±H:MM'
    
    Returns:
        float: Timezone offset in decimal hours
    
    Example:
        >>> convert_timezone_to_float('UTC+5:30')
        5.5
        >>> convert_timezone_to_float('UTC-8')
        -8.0
    """
    timezone_str = timezone_str.replace('UTC', '')
    if ':' in timezone_str:
        hours, minutes = map(int, timezone_str.split(':'))
        # Handle negative timezones correctly by applying the sign to the minutes conversion
        return hours + (minutes/60) * (1 if hours >= 0 else -1)
    return float(timezone_str)

def convert_local_to_utc(local_hour: str, timezone: str) -> int:
    """
    Convert a local hour to UTC, handling fractional timezones
    
    Args:
        local_hour: Hour in local time (0-23 as string)
        timezone: Timezone string (e.g., 'UTC+5:30')
    
    Returns:
        int: Hour in UTC (0-23)
    
    Example:
        >>> convert_local_to_utc('9', 'UTC+5:30')
        3  # 9:00 AM in UTC+5:30 is 3:30 AM UTC
    """
    try:
        timezone_float = convert_timezone_to_float(timezone)
        return (int(local_hour) - int(timezone_float)) % 24
    except (ValueError, TypeError) as e:
        print(f"Error converting timezone {timezone}: {str(e)}")
        return None

def generate_time_ranges(hours: list) -> str:
    """
    Convert a list of hours into a formatted string of time ranges
    
    Args:
        hours: List of hours (0-23)
    
    Returns:
        str: Formatted string of time ranges (e.g., "09:00-11:00, 14:00-16:00")
    
    Example:
        >>> generate_time_ranges([9, 10, 11, 14, 15])
        '09:00-12:00, 14:00-16:00'
    """
    if not hours:
        return ""
    
    sorted_hours = sorted(hours)
    ranges = []
    start = sorted_hours[0]
    prev = start
    
    for hour in sorted_hours[1:]:
        if hour != prev + 1:
            ranges.append(f"{start:02d}:00-{prev + 1:02d}:00")
            start = hour
        prev = hour
    ranges.append(f"{start:02d}:00-{prev + 1:02d}:00")
    
    return ", ".join(ranges)