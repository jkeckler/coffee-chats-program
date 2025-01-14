# Coffee Chats Program - Project Notes
Last Updated: January 13, 2025

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