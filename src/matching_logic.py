from typing import List, Dict, Set
import random

class CoffeeGroup:
    def __init__(self, members, min_overlap_hours=1, previous_captains=None):
        self.members = members
        self.min_overlap_hours = min_overlap_hours
        self.previous_captains = previous_captains or set()
        self.common_hours = self.find_common_hours()
        
        if self.is_valid_group():
            self.captain = self.assign_captain()
            if self.captain:
                self.generate_captain_message()
        else:
            self.captain = None
            print("No captain assigned - group has insufficient overlapping hours")

    def find_common_hours(self):
        """Find hours when all group members are available in UTC"""
        if not self.members:
            return set()
        
        all_available_hours = []
        for member in self.members:
            # Convert each member's available hours to UTC
            member_utc_hours = set()
            for hour, status in member['availability'].items():
                if status == 'Y':
                    utc_hour = self.convert_to_utc(hour, member['timezone'])
                    member_utc_hours.add(utc_hour)
            all_available_hours.append(member_utc_hours)
        
        # Find intersection of all members' UTC hours
        common_hours = all_available_hours[0]
        for hours in all_available_hours[1:]:
            common_hours = common_hours.intersection(hours)
        
        return common_hours

    def convert_to_utc(self, local_hour: str, timezone: str) -> int:
        """Convert a local hour to UTC"""
        offset = int(timezone.replace('UTC', ''))
        return (int(local_hour) - offset) % 24

    def is_valid_group(self):
        """Check if group has enough overlapping hours"""
        return len(self.common_hours) >= self.min_overlap_hours

    def assign_captain(self):
        """Assign a captain, avoiding those who were captain last round"""
        eligible_captains = [m for m in self.members 
                           if m['name'] not in self.previous_captains]
        
        if eligible_captains:
            return random.choice(eligible_captains)
        else:
            return random.choice(self.members)

    def generate_captain_message(self):
        """Generate a helpful message for the captain about meeting times"""
        if not self.captain or not self.common_hours:
            return
        
        message = f"\nHello {self.captain['name']}, you are the coffee chat captain!\n"
        message += "Your group members are:\n"
        
        for member in self.members:
            if member != self.captain:
                message += f"- {member['name']} ({member['timezone']})\n"
        
        message += "\nYour group can meet at the following times (in each person's local time):\n"
        
        for utc_hour in self.common_hours:
            message += "\n"
            for member in self.members:
                local_hour = (utc_hour + int(member['timezone'].replace('UTC', ''))) % 24
                message += f"{member['name']}: {local_hour}:00\n"
        
        print(message)

# Test data
test_members = [
    {
        'name': 'Alice',
        'timezone': 'UTC+1',
        'department': 'Engineering',
        'availability': {'13': 'Y', '14': 'Y', '15': 'Y'},
        'has_been_captain': False
    },
    {
        'name': 'Bob',
        'timezone': 'UTC+1',
        'department': 'Marketing',
        'availability': {'13': 'Y', '14': 'Y', '15': 'N'},
        'has_been_captain': True
    },
    {
        'name': 'Carol',
        'timezone': 'UTC+2',
        'department': 'Design',
        'availability': {'12': 'Y', '13': 'Y', '14': 'Y'},
        'has_been_captain': False
    }
]

# Test with previous captains
last_round_captains = {'Bob'}
group1 = CoffeeGroup(test_members, previous_captains=last_round_captains)