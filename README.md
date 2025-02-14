# Coffee Chats Program

A Python-based program for facilitating workplace connections through randomized coffee chat matching. This tool helps create meaningful connections across departments and time zones by automatically generating groups for casual coffee conversations.

## Features
- Random group matching (2-3 people per group)
- Group captain assignment
- Excel-based interface for easy data input and output
- Consideration of time zones and availability for optimal matching

## Prerequisites
- Python 3.7 or higher
- Git (for cloning the repository)

## Setup
1. Clone the repository:
git clone https://github.com/jkeckler/coffee-chats-program.git cd coffee-chats-program

2. Create and activate a virtual environment:
python -m venv venv source venv/bin/activate # On Windows use venv\Scripts\activate

3. Install required packages:
pip install -r requirements.txt

## Usage
1. Generate a new participant template:
python src/new_participant_template.py

This creates a new Excel template file in the `data` directory.

2. Fill out the generated Excel template with participant information, including their availability.

3. Run the matching program:
python src/coffee_matching.py

This generates coffee chat groups based on the participant data.

4. View the results:
The program creates an Excel file with the matched groups in the `data` directory.

## Project Structure
- `src/`: Core Python scripts
- `data/`: Input and output Excel files (not tracked in Git)
- `tests/`: Unit tests
- `docs/`: Project documentation

## Development Notes

### Git and GitHub Setup
1. Ensure Git is installed on your machine.
2. Initialize Git in the project directory:

git init

3. Create a .gitignore file:

4. Add files to Git:
git add .

5. Make your first commit:
git commit -m "Initial commit"

6. Link to your GitHub repository:
git remote add origin https://github.com/yourusername/coffee-chats-program.git

7. Push to GitHub:
git push -u origin main

### Updating GitHub
After making changes:
1. Check status: `git status`
2. Add changes: `git add .`
3. Commit: `git commit -m "Description of changes"`
4. Push: `git push origin main`

### Setting Up on a New Machine
1. Clone the repository
2. Set up a virtual environment
3. Install dependencies from requirements.txt
4. Run the program as described in the Usage section

## Contributing (PLACEHOLDER - DOES NOT EXIST YET)
Please read [CONTRIBUTING.md](CONTRIBUTING.md) for details on our code of conduct and the process for submitting pull requests.

## About
This project aims to foster workplace connections and break down silos in global organizations by facilitating casual coffee chats among employees across different departments and time zones.

## License (PLACEHOLDER - LICENSE DOES NOT EXIST YET)
This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.