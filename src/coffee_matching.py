"""
Main script for the coffee chat matching system.
Orchestrates the entire matching process using the modular components.
"""
from typing import List, Dict, Set
import random

# Change relative imports (.) to absolute imports (src.)
from src.models import CoffeeGroup, REGIONS
from src.utils import convert_timezone_to_float
from src.excel_handler import quick_read_participants as read_participants_from_excel, export_matches_to_excel
from src.captain_history import CaptainHistoryDB

class GroupMatcher:
    """Main class for matching participants into compatible coffee chat groups"""
    
    def __init__(self, participants: List[dict], min_group_size: int = 3, max_group_size: int = 3):
        self.participants = participants
        self.min_group_size = min_group_size
        self.max_group_size = max_group_size
        self.groups = []
        self.used_participants = set()
        self.db = CaptainHistoryDB()  # Initialize database connection
        
        # Update captain counts from database
        for participant in self.participants:
            stats = self.db.get_captain_stats(participant['id'])
            if stats:
                participant['captain_count'] = stats[0]
            # Register new participants in the database
            self.db.update_or_create_employee(participant)

    def get_member_region(self, participant: dict) -> str:
        """Determine which region a participant belongs to based on timezone"""
        try:
            timezone = convert_timezone_to_float(participant['timezone'])
            for region_name, region in REGIONS.items():
                if region.start_utc <= timezone <= region.end_utc:
                    return region_name
            return 'EMEA'  # Default to EMEA if no match
        except ValueError:
            return 'EMEA'

    def find_compatible_group(self, seed: dict, candidates: List[dict], size: int) -> List[dict]:
        """Find compatible participants to form a group with the seed participant"""
        compatible = []
        for candidate in candidates:
            if candidate == seed or candidate['id'] in self.used_participants:
                continue
            test_group = CoffeeGroup([seed, candidate])
            if test_group.is_valid_group():
                compatible.append(candidate)
                
        if len(compatible) >= size - 1:
            selected = random.sample(compatible, size - 1)
            group = CoffeeGroup([seed] + selected)
            if group.is_valid_group():
                return [seed] + selected
        return []

    def create_groups(self) -> List[CoffeeGroup]:
        """Create groups of participants with flexible group sizes"""
        available = [p for p in self.participants if p['id'] not in self.used_participants]
        formed_groups = []
        
        while len(available) >= 2:
            seed = random.choice(available)
            
            # First try to form a group of 3
            if len(available) >= 3:
                group_members = self.find_compatible_group(seed, available, 3)
                if not group_members:  # If can't form trio, try pair
                    group_members = self.find_compatible_group(seed, available, 2)
            else:  # Only enough people left for a pair
                group_members = self.find_compatible_group(seed, available, 2)
            
            if group_members:
                group = CoffeeGroup(group_members)
                if group.is_valid_group():
                    formed_groups.append(group)
                    # Update database with new group
                    if group.captain:
                        member_ids = [m['id'] for m in group.members]
                        self.db.record_meeting(
                            group.captain['id'],
                            member_ids,
                            f"{group.optimal_meeting_time:02d}:00 UTC"
                        )
                    for member in group_members:
                        self.used_participants.add(member['id'])
                        available.remove(member)
                else:
                    available.remove(seed)
                    self.used_participants.add(seed['id'])
            else:
                available.remove(seed)
                self.used_participants.add(seed['id'])
                
            print(f"Remaining participants: {len(available)}")
        
        self.groups = formed_groups
        return formed_groups

    def print_group_summary(self):
        """Print summary of all formed groups"""
        print(f"\nCreated {len(self.groups)} groups:")
        for i, group in enumerate(self.groups, 1):
            if group.captain:
                print(f"\nGroup {i}:")
                print(f"Captain: {group.captain['name']} ({group.captain['timezone']})")
                other_members = [m for m in group.members if m != group.captain]
                members_str = ", ".join([f"{m['name']} ({m['timezone']})" for m in other_members])
                print(f"Members: {members_str}")
                print(f"Meeting time (UTC): {group.optimal_meeting_time}:00")

    def get_unmatched_participants(self) -> List[dict]:
        """Return list of participants who couldn't be matched into groups"""
        matched = {member['id'] for group in self.groups for member in group.members}
        unmatched = [p for p in self.participants if p['id'] not in matched]
        if unmatched:
            print("\nUnmatched participants:")
            for p in unmatched:
                print(f"- {p['name']} ({p['timezone']})")
        return unmatched
    
    def analyze_unmatched_participants(self) -> List[Dict]:
        """Analyze why participants couldn't be matched and provide recommendations"""
        unmatched = self.get_unmatched_participants()
        analysis = []
        
        for participant in unmatched:
            available_hours = sum(1 for status in participant['availability'].values() if status == 'Y')
            
            issue = {
                'name': participant['name'],
                'timezone': participant['timezone'],
                'department': participant['department'],
                'available_hours': available_hours,
                'recommendation': ''
            }
            
            if available_hours < 3:
                issue['recommendation'] = "Consider expanding availability window"
            elif abs(convert_timezone_to_float(participant['timezone'])) > 8:
                issue['recommendation'] = "Time zone significantly different from others"
            else:
                issue['recommendation'] = "Consider for next round's priority matching"
                
            analysis.append(issue)
        
        return analysis

def run_matching_program(input_path: str = "data/participant_template.xlsx",
                        output_path: str = "data/coffee_chat_matches_round2.xlsx"):
    """Main function to run the coffee chat matching program"""
    print("Starting coffee chat matching program...")
    
    # Read participants
    participants = read_participants_from_excel(input_path)
    if not participants:
        print("No participants found in Excel file!")
        return
    
    # Create and run matcher
    matcher = GroupMatcher(participants)
    groups = matcher.create_groups()
    matcher.print_group_summary()
    
    # Export results
    export_matches_to_excel(groups, output_path)
    print(f"\nExport complete! Check {output_path} for the full schedule.")
    
    # Analyze unmatched participants
    unmatched_analysis = matcher.analyze_unmatched_participants()
    if unmatched_analysis:
        print("\nAnalysis of unmatched participants:")
        for analysis in unmatched_analysis:
            print(f"\n{analysis['name']} ({analysis['timezone']}):")
            print(f"Available hours: {analysis['available_hours']}")
            print(f"Recommendation: {analysis['recommendation']}")

if __name__ == "__main__":
    run_matching_program()
