## Edge Case Handling (Added January 29, 2025)

The system now includes intelligent analysis of unmatched participants:

### Features
- Analyzes available hours
- Identifies timezone challenges
- Provides specific recommendations

### Recommendations Include:
1. Expanding availability window (for limited hours)
2. Manual matching consideration (for extreme time zones)
3. Priority matching in next round (for other cases)

### Usage
```python
# Get analysis of unmatched participants
analysis = matcher.analyze_unmatched_participants()

# Results include:
# - Name, timezone, department
# - Number of available hours
# - Specific recommendations





## Latest Updates (January 14, 2025)
- Implemented core matching logic
- Added time zone conversion functionality
- Created captain assignment system
- Developed captain notification messages
## Current Features
- Group validation based on overlapping hours
- Time zone conversion for global participants
- Prevention of back-to-back captain assignments
- Automated meeting time suggestions in local time zones

## Technical Implementation
- CoffeeGroup class handles group formation and validation
- Time conversion system accounts for different UTC offsets
- Captain selection considers previous assignments
- Clear output format for meeting coordination



# Coffee Chats Program - Project Notes - Last Updated: January 13, 2025


## Current Features
- Excel template with data validation
- Sample data generator for testing
- Global time zone handling
- Department distribution weighting
- Availability patterns based on local time

## Data Structure
- Participant data includes: ID, Name, Email, Department, Country, Time Zone, NIQ Tenure, Status
- Availability tracked in 24-hour format (Y/N for each hour)
- Data validations prevent invalid entries

## Technical Decisions
- Using openpyxl for Excel operations
- Conda environment for dependency management
- Sample data weighted towards Engineering (25%) and newer employees
- Availability adjusted automatically for time zones

## Future Improvements
- Add data validations for:
  - Time zones (matching with countries)
  - Countries (standardized list)
  - Email format consistency