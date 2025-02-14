"""
Core data models for the coffee chat matching system.
Contains TimeRegion and CoffeeGroup classes that define the fundamental structures.
"""

from dataclasses import dataclass
from typing import List, Dict, Set
import random
from src.utils import convert_local_to_utc, generate_time_ranges # Fixed during review if import statments

@dataclass
class TimeRegion:
    """
    Represents a geographic region with its timezone boundaries and preferred meeting times.
    
    Attributes:
        name: Region name (e.g., 'APAC', 'EMEA', 'AMER')
        start_utc: Start of timezone range (inclusive)
        end_utc: End of timezone range (inclusive)
        preferred_meeting_times: List of preferred meeting hours in UTC
    """
    name: str
    start_utc: float  # Start of timezone range (inclusive)
    end_utc: float    # End of timezone range (inclusive)
    preferred_meeting_times: List[int]  # Preferred meeting hours in UTC

# Define global regions with fractional timezone boundaries
REGIONS = {
    'APAC': TimeRegion('APAC', 4.5, 12, [1, 2, 3]),     # Asia-Pacific: early UTC hours
    'EMEA': TimeRegion('EMEA', -2, 4.5, [8, 9, 10]),    # Europe/Middle East: mid UTC hours
    'AMER': TimeRegion('AMER', -12, -2, [14, 15, 16])   # Americas: later UTC hours
}

class CoffeeGroup:
    """
    Represents a group of participants who will meet for coffee chat.
    Handles availability matching, captain assignment, and meeting time selection.
    """
    
    def __init__(self, members: List[dict], min_overlap_hours: int = 1):
        """
        Initialize a coffee chat group.
        
        Args:
            members: List of participant dictionaries
            min_overlap_hours: Minimum required overlapping hours for valid group
        """
        self.members = members
        self.min_overlap_hours = min_overlap_hours
        self.common_hours = self.find_common_hours()
        self.captain = None
        if self.is_valid_group():
            self.optimal_meeting_time = self.find_optimal_meeting_time()
            self.captain = self.assign_captain()
            if self.captain:
                self.generate_captain_message()

    def find_common_hours(self) -> Set[int]:
        """Find hours when all group members are available"""
        all_available_hours = []
        for member in self.members:
            member_utc_hours = set()
            for hour, status in member['availability'].items():
                if status == 'Y':
                    utc_hour = convert_local_to_utc(hour, member['timezone'])
                    if utc_hour is not None:
                        member_utc_hours.add(utc_hour)
                        # Add next day's hour for edge cases
                        member_utc_hours.add((utc_hour + 24) % 24)
            all_available_hours.append(member_utc_hours)
        
        return set.intersection(*all_available_hours) if all_available_hours else set()

    def is_valid_group(self) -> bool:
        """Check if group has enough overlapping hours"""
        return len(self.common_hours) >= self.min_overlap_hours

    def find_optimal_meeting_time(self) -> int:
        """Find best meeting time based on common availability"""
        if not self.common_hours:
            return None
        return min(self.common_hours)  # Start with earliest common hour

    def assign_captain(self):
        """Assign a captain, preferring those who haven't been captain before"""
        if not self.is_valid_group():
            return None
        
        # Select captain with lowest captain_count
        min_count = min(m['captain_count'] for m in self.members)
        eligible_captains = [m for m in self.members if m['captain_count'] == min_count]
        captain = random.choice(eligible_captains)
        captain['captain_count'] += 1
        return captain

    def generate_captain_message(self) -> str:
        """Generate welcome message for captain with meeting details"""
        if not self.captain or not self.common_hours:
            return ""
        
        message = [f"\nHello {self.captain['name']}, you are the coffee chat captain!\n"]
        message.append("Your group members are:")
        
        # Add member information
        for member in self.members:
            if member != self.captain:
                message.append(f"- {member['name']} ({member['timezone']})")
        
        # Calculate and add available time ranges
        message.append("\nAvailable time ranges (local time):")
        for member in self.members:
            local_hours = []
            for utc_hour in self.common_hours:
                # Convert UTC back to local time
                offset = -convert_local_to_utc('0', member['timezone'])
                local_hour = (utc_hour + offset) % 24
                local_hours.append(local_hour)
            
            time_ranges = generate_time_ranges(local_hours)
            message.append(f"{member['name']}: {time_ranges}")
        
        return "\n".join(message)