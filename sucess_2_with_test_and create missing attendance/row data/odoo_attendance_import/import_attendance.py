import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
import requests
from dotenv import load_dotenv
from collections import defaultdict

# Load environment variables
load_dotenv()

class OdooAPI:
    def __init__(self):
        self.url = os.getenv('ODOO_URL')
        self.db = os.getenv('ODOO_DB')
        self.username = os.getenv('ODOO_USERNAME')
        self.password = os.getenv('ODOO_PASSWORD')
        self.api_key = os.getenv('api_key')
        self.session = requests.Session()
        self.uid = None
        self.login()

    def login(self):
        """Login to Odoo and get user ID"""
        login_url = f"{self.url}/web/session/authenticate"
        login_data = {
            "jsonrpc": "2.0",
            "params": {
                "db": self.db,
                "login": self.username,
                "password": self.password,
                "api_key": self.api_key
            }
        }
        try:
            response = self.session.post(login_url, json=login_data)
            result = response.json()
            if 'error' in result:
                raise Exception(f"Login failed: {result['error']['data']['message']}")
            self.uid = result.get('result', {}).get('uid')
            if not self.uid:
                raise Exception("Login failed: Could not get user ID")
            return self.uid
        except requests.exceptions.RequestException as e:
            raise Exception(f"Connection error: {str(e)}")
        except Exception as e:
            raise Exception(f"Login error: {str(e)}")

    def get_employee_id(self, badge_id):
        """Get Odoo employee ID from badge ID"""
        endpoint = f"{self.url}/web/dataset/call_kw"
        params = {
            "model": "hr.employee",
            "method": "search_read",
            "args": [[["barcode", "=", str(badge_id)]]],
            "kwargs": {
                "fields": ["id", "name"],
                "limit": 1
            }
        }
        data = {
            "jsonrpc": "2.0",
            "params": params
        }
        try:
            response = self.session.post(endpoint, json=data)
            result = response.json()
            if 'error' in result:
                raise Exception(f"Error getting employee: {result['error']['data']['message']}")
            employees = result.get('result', [])
            return employees[0]['id'] if employees else None
        except Exception as e:
            raise Exception(f"Error getting employee: {str(e)}")

    def check_missing_employees(self, badge_ids):
        """Check which employees need to be created in Odoo"""
        missing_employees = []
        existing_employees = []
        
        for badge_id in badge_ids:
            if not self.get_employee_id(badge_id):
                missing_employees.append(badge_id)
            else:
                existing_employees.append(badge_id)
        
        return missing_employees, existing_employees

    def create_attendance(self, employee_id, check_in, check_out=None):
        """Create attendance record in Odoo"""
        endpoint = f"{self.url}/web/dataset/call_kw"
        
        # Format datetime to match Odoo's expected format
        def format_datetime(dt):
            return dt.strftime('%Y-%m-%d %H:%M:%S')
        
        attendance_data = {
            "employee_id": employee_id,
            "check_in": format_datetime(check_in),
        }
        if check_out:
            attendance_data["check_out"] = format_datetime(check_out)

        params = {
            "model": "hr.attendance",
            "method": "create",
            "args": [attendance_data],
            "kwargs": {}
        }
        data = {
            "jsonrpc": "2.0",
            "params": params
        }
        try:
            response = self.session.post(endpoint, json=data)
            result = response.json()
            if 'error' in result:
                raise Exception(f"Error creating attendance: {result['error']['data']['message']}")
            return result.get('result')
        except Exception as e:
            raise Exception(f"Error creating attendance: {str(e)}")

    def create_employee(self, badge_id, name):
        """Create a new employee in Odoo"""
        endpoint = f"{self.url}/web/dataset/call_kw"
        params = {
            "model": "hr.employee",
            "method": "create",
            "args": [{
                "name": name,
                "barcode": str(badge_id),
                "pin": str(badge_id),  # Using badge_id as PIN for simplicity
            }],
            "kwargs": {}
        }
        data = {
            "jsonrpc": "2.0",
            "params": params
        }
        try:
            response = self.session.post(endpoint, json=data)
            result = response.json()
            if 'error' in result:
                raise Exception(f"Error creating employee: {result['error']['data']['message']}")
            return result.get('result')
        except Exception as e:
            raise Exception(f"Error creating employee: {str(e)}")

def process_excel_file(file_path):
    """Process the Excel file and return attendance data"""
    df = pd.read_excel(file_path)
    print("Available columns:", df.columns.tolist())
    print("\nFirst few rows of data:")
    print(df.head())
    
    # Convert Time column to datetime
    df['Time'] = pd.to_datetime(df['Time'])
    
    # Extract date from timestamp for grouping
    df['Date'] = df['Time'].dt.date
    
    # Process records by employee and date
    attendance_records = []
    
    # Group by AC-No. and Date
    for (employee_id, date), group in df.groupby(['AC-No.', 'Date']):
        # Get first check-in and last check-out for the day
        check_ins = group[group['State'] == 'C/In']['Time']
        check_outs = group[group['State'] == 'C/Out']['Time']
        
        if not check_ins.empty and not check_outs.empty:
            first_check_in = check_ins.min()
            last_check_out = check_outs.max()
            
            # Only add if we have both check-in and check-out
            if first_check_in < last_check_out:
                attendance_records.append({
                    'employee_id': str(employee_id),
                    'date': date,
                    'check_in': first_check_in,
                    'check_out': last_check_out,
                    'total_hours': (last_check_out - first_check_in).total_seconds() / 3600
                })
    
    # Convert to DataFrame for easier analysis
    attendance_df = pd.DataFrame(attendance_records)
    return attendance_df

def visualize_attendance(attendance_df):
    """Create visualizations of the attendance data"""
    if attendance_df.empty:
        print("No attendance records to visualize")
        return
    
    # Create a directory for the visualizations
    os.makedirs('attendance_analysis', exist_ok=True)
    
    # 1. Daily hours worked by employee
    plt.figure(figsize=(12, 6))
    for employee in attendance_df['employee_id'].unique():
        employee_data = attendance_df[attendance_df['employee_id'] == employee]
        plt.plot(employee_data['date'], employee_data['total_hours'], 
                marker='o', label=f'Employee {employee}')
    
    plt.title('Daily Hours Worked by Employee')
    plt.xlabel('Date')
    plt.ylabel('Hours Worked')
    plt.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('attendance_analysis/daily_hours.png')
    plt.close()
    
    # 2. Average hours worked by employee
    plt.figure(figsize=(10, 6))
    avg_hours = attendance_df.groupby('employee_id')['total_hours'].mean()
    avg_hours.plot(kind='bar')
    plt.title('Average Hours Worked by Employee')
    plt.xlabel('Employee ID')
    plt.ylabel('Average Hours')
    plt.tight_layout()
    plt.savefig('attendance_analysis/avg_hours.png')
    plt.close()
    
    # Generate summary statistics
    summary = attendance_df.groupby('employee_id').agg({
        'total_hours': ['mean', 'min', 'max', 'count'],
        'check_in': lambda x: x.dt.strftime('%H:%M:%S').mode().iloc[0],
        'check_out': lambda x: x.dt.strftime('%H:%M:%S').mode().iloc[0]
    }).round(2)
    
    # Save summary to CSV
    summary.to_csv('attendance_analysis/attendance_summary.csv')
    
    print("\nAttendance Analysis:")
    print(f"Total number of records: {len(attendance_df)}")
    print(f"Date range: {attendance_df['date'].min()} to {attendance_df['date'].max()}")
    print(f"\nSummary statistics saved to 'attendance_analysis/attendance_summary.csv'")
    print("Visualizations saved to 'attendance_analysis' directory")
    
    # Display first few processed records
    print("\nSample of processed records:")
    print(attendance_df.head().to_string())
    
    return attendance_df

def create_missing_employees(odoo, missing_employees):
    """Create missing employees in Odoo"""
    print("\nCreating missing employees in Odoo...")
    success_count = 0
    error_count = 0
    error_details = defaultdict(list)

    for badge_id in missing_employees:
        try:
            # Create employee name based on badge ID
            employee_name = f"Employee {badge_id}"
            
            # Ask for employee name
            name = input(f"\nEnter name for employee with Badge ID {badge_id} (or press Enter to use '{employee_name}'): ").strip()
            if not name:
                name = employee_name

            # Create employee
            odoo.create_employee(badge_id, name)
            success_count += 1
            print(f"✓ Created employee: {name} (Badge ID: {badge_id})")
        
        except Exception as e:
            error_count += 1
            error_details[str(e)].append(badge_id)
            print(f"✗ Error creating employee with Badge ID {badge_id}: {str(e)}")
    
    # Print summary
    print("\nEmployee Creation Summary:")
    print(f"Successfully created: {success_count} employees")
    print(f"Failed to create: {error_count} employees")
    
    if error_details:
        print("\nError Details:")
        for error, employees in error_details.items():
            print(f"\nError: {error}")
            print("Affected Badge IDs:", ", ".join(map(str, employees)))

def main():
    # Path to your Excel file
    excel_file = '/home/sabry/Downloads/row data/acnLog12.xls'
    
    try:
        # Process Excel file
        attendance_df = process_excel_file(excel_file)
        
        # Visualize and analyze the data
        attendance_df = visualize_attendance(attendance_df)
        
        # Initialize Odoo API
        print("\nChecking employee records in Odoo...")
        odoo = OdooAPI()
        
        # Get unique employee IDs from the data
        unique_employees = attendance_df['employee_id'].unique()
        
        # Check for missing employees
        missing_employees, existing_employees = odoo.check_missing_employees(unique_employees)
        
        if missing_employees:
            print("\n⚠️ WARNING: The following employees need to be created in Odoo first:")
            print("\nMissing Employees (Badge IDs):")
            for emp_id in missing_employees:
                print(f"- {emp_id}")
            
            while True:
                response = input("\nWould you like to create these employees now? (yes/no): ").lower()
                if response in ['yes', 'no']:
                    break
                print("Please enter 'yes' or 'no'")
            
            if response == 'yes':
                create_missing_employees(odoo, missing_employees)
                # Recheck for missing employees
                missing_employees, existing_employees = odoo.check_missing_employees(unique_employees)
                if missing_employees:
                    print("\nSome employees could not be created. Please create them manually in Odoo.")
                    return
            else:
                print("\nPlease create the employees manually in Odoo:")
                print("1. Go to Employees > Employees")
                print("2. Click 'Create'")
                print("3. Fill in the employee details")
                print("4. Set the 'Badge ID' field to match the AC-No. from your Excel file")
                print("\nAfter creating the employees, run this script again.")
                return
        
        # Proceed with import for existing employees
        print("\nStarting import for existing employees...")
        success_count = 0
        error_count = 0
        error_details = defaultdict(list)
        
        for index, row in attendance_df.iterrows():
            if row['employee_id'] not in missing_employees:
                try:
                    employee_id = odoo.get_employee_id(row['employee_id'])
                    odoo.create_attendance(
                        employee_id=employee_id,
                        check_in=row['check_in'],
                        check_out=row['check_out']
                    )
                    success_count += 1
                    print(f"✓ Created attendance record for employee {row['employee_id']}")
                except Exception as e:
                    error_count += 1
                    error_details[str(e)].append(row['employee_id'])
                    print(f"✗ Error for employee {row['employee_id']}: {str(e)}")
        
        # Print summary
        print("\nImport Summary:")
        print(f"Successfully imported: {success_count} records")
        print(f"Failed to import: {error_count} records")
        
        if error_details:
            print("\nError Details:")
            for error, employees in error_details.items():
                print(f"\nError: {error}")
                print("Affected employees:", ", ".join(employees))
        
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
