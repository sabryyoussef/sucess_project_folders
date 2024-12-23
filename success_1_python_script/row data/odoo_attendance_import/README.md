# Odoo Attendance Import Script

This script processes attendance records from an Excel file, provides visualizations and analysis, and optionally imports the data into Odoo.

## Features

- Processes Excel file with attendance records
- Extracts first check-in and last check-out for each day
- Generates visualizations of working hours
- Creates detailed statistical analysis
- Optional import to Odoo system

## Prerequisites

- Python 3.10 or higher
- Access to Odoo instance
- Excel file with attendance records

## Installation

1. Clone or download this repository to your local machine

2. Create a virtual environment:
```bash
cd "/path/to/odoo_attendance_import"
python3 -m venv ../venv
```

3. Activate the virtual environment:
```bash
# On Linux/Mac
source ../venv/bin/activate
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Configuration

1. Copy the `.env.example` file to `.env`:
```bash
cp .env.example .env
```

2. Edit the `.env` file with your Odoo credentials:
```
ODOO_URL=your_odoo_url
ODOO_DB=your_database_name
ODOO_USERNAME=your_username
ODOO_PASSWORD=your_password
api_key=your_api_key
```

## Usage

1. Place your Excel file in the project directory or update the file path in `import_attendance.py`

2. Run the script:
```bash
# Make sure you're in the project directory
cd "/path/to/odoo_attendance_import"

# Activate virtual environment if not already activated
source ../venv/bin/activate

# Run the script
python import_attendance.py
```

3. The script will:
   - Process the Excel file
   - Create visualizations in `attendance_analysis` directory
   - Show sample of processed records
   - Ask if you want to import to Odoo

## Output Files

The script creates an `attendance_analysis` directory containing:

- `daily_hours.png`: Graph showing daily hours worked by each employee
- `avg_hours.png`: Bar chart of average hours worked by employee
- `attendance_summary.csv`: Detailed statistics including:
  - Average, minimum, and maximum hours worked
  - Most common check-in and check-out times
  - Number of attendance records

## Excel File Format

The Excel file should contain the following columns:
- `AC-No.`: Employee ID/Badge number
- `Time`: Date and time of the record
- `State`: Check-in (C/In) or Check-out (C/Out) status

## Support

For any issues or questions, please contact your system administrator.
