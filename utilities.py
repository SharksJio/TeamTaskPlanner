#!/usr/bin/env python3
"""
Team Task Planner Utilities
Additional utilities for customizing and working with the Team Task Planner Excel file.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
import sys

def add_new_team_sheet(workbook_path, team_name):
    """Add a new team sheet to an existing workbook."""
    try:
        wb = openpyxl.load_workbook(workbook_path)
        
        # Check if team already exists
        if team_name in wb.sheetnames:
            print(f"Team '{team_name}' already exists!")
            return False
        
        # Create new sheet based on existing team template
        template_sheet = wb["Frontend Team"]  # Use as template
        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = team_name
        
        # Update the title in the new sheet
        new_sheet['A1'] = f"{team_name.upper()} - TASK MANAGEMENT"
        
        # Clear sample data (keep headers)
        for row in range(7, 10):  # Clear rows 7-9 (sample tasks)
            for col in range(1, 14):  # Clear columns A-M
                new_sheet.cell(row=row, column=col).value = ""
        
        wb.save(workbook_path)
        print(f"Successfully added team '{team_name}' to {workbook_path}")
        return True
        
    except Exception as e:
        print(f"Error adding team: {e}")
        return False

def update_master_dashboard(workbook_path, teams_list):
    """Update master dashboard with new team information."""
    try:
        wb = openpyxl.load_workbook(workbook_path)
        master_sheet = wb["Master Dashboard"]
        
        # Clear existing team data (rows 7-10)
        for row in range(7, 11):
            for col in range(1, 9):
                master_sheet.cell(row=row, column=col).value = ""
        
        # Add new teams
        for idx, team in enumerate(teams_list, 7):
            master_sheet.cell(row=idx, column=1).value = team
            # Add placeholder formulas that can be updated manually
            master_sheet.cell(row=idx, column=2).value = 0  # Total Tasks
            master_sheet.cell(row=idx, column=3).value = 0  # Completed
            master_sheet.cell(row=idx, column=4).value = 0  # In Progress
            master_sheet.cell(row=idx, column=5).value = 0  # Pending
            master_sheet.cell(row=idx, column=6).value = 0  # Overdue
            master_sheet.cell(row=idx, column=7).value = 0  # Team Size
            master_sheet.cell(row=idx, column=8).value = 0  # Workload %
        
        wb.save(workbook_path)
        print(f"Successfully updated master dashboard in {workbook_path}")
        return True
        
    except Exception as e:
        print(f"Error updating master dashboard: {e}")
        return False

def export_team_tasks_to_csv(workbook_path, team_name, output_file):
    """Export a team's tasks to CSV format."""
    try:
        wb = openpyxl.load_workbook(workbook_path)
        
        if team_name not in wb.sheetnames:
            print(f"Team '{team_name}' not found!")
            return False
        
        sheet = wb[team_name]
        
        # Read headers and data
        headers = []
        for col in range(1, 14):  # Columns A-M
            headers.append(sheet.cell(row=6, column=col).value)
        
        tasks = []
        row = 7
        while sheet.cell(row=row, column=1).value:  # While there's a Task ID
            task = []
            for col in range(1, 14):
                task.append(sheet.cell(row=row, column=col).value)
            tasks.append(task)
            row += 1
        
        # Write to CSV
        import csv
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            writer.writerows(tasks)
        
        print(f"Successfully exported {team_name} tasks to {output_file}")
        return True
        
    except Exception as e:
        print(f"Error exporting tasks: {e}")
        return False

def generate_weekly_report(workbook_path, output_file):
    """Generate a weekly progress report."""
    try:
        wb = openpyxl.load_workbook(workbook_path)
        
        report_data = []
        report_data.append("WEEKLY PROGRESS REPORT")
        report_data.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        report_data.append("="*50)
        
        # Get master dashboard data
        master_sheet = wb["Master Dashboard"]
        report_data.append("\nOVERALL SUMMARY:")
        
        for row in range(7, 11):  # Team data rows
            team = master_sheet.cell(row=row, column=1).value
            if team:
                total_tasks = master_sheet.cell(row=row, column=2).value or 0
                completed = master_sheet.cell(row=row, column=3).value or 0
                in_progress = master_sheet.cell(row=row, column=4).value or 0
                workload = master_sheet.cell(row=row, column=8).value or 0
                
                report_data.append(f"\n{team}:")
                report_data.append(f"  Total Tasks: {total_tasks}")
                report_data.append(f"  Completed: {completed}")
                report_data.append(f"  In Progress: {in_progress}")
                report_data.append(f"  Workload: {workload}%")
        
        # Get team allocation summary
        allocation_sheet = wb["Team Allocation"]
        report_data.append("\nTEAM AVAILABILITY:")
        
        row = 5
        while allocation_sheet.cell(row=row, column=1).value:
            name = allocation_sheet.cell(row=row, column=1).value
            team = allocation_sheet.cell(row=row, column=2).value
            availability = allocation_sheet.cell(row=row, column=6).value or 0
            status = allocation_sheet.cell(row=row, column=9).value
            
            report_data.append(f"  {name} ({team}): {availability}% - {status}")
            row += 1
        
        # Write report
        with open(output_file, 'w', encoding='utf-8') as f:
            for line in report_data:
                f.write(line + '\n')
        
        print(f"Weekly report generated: {output_file}")
        return True
        
    except Exception as e:
        print(f"Error generating report: {e}")
        return False

def validate_workbook(workbook_path):
    """Validate the workbook structure and data."""
    try:
        wb = openpyxl.load_workbook(workbook_path)
        
        required_sheets = ["Master Dashboard", "Team Allocation", "Weekly Planning", "Templates"]
        issues = []
        
        # Check required sheets
        for sheet_name in required_sheets:
            if sheet_name not in wb.sheetnames:
                issues.append(f"Missing required sheet: {sheet_name}")
        
        # Check master dashboard
        if "Master Dashboard" in wb.sheetnames:
            master_sheet = wb["Master Dashboard"]
            if not master_sheet['A1'].value or "MASTER DASHBOARD" not in master_sheet['A1'].value:
                issues.append("Master Dashboard title is missing or incorrect")
        
        # Check team allocation
        if "Team Allocation" in wb.sheetnames:
            allocation_sheet = wb["Team Allocation"]
            expected_headers = ["Employee Name", "Team", "Role", "Current Tasks"]
            for col, header in enumerate(expected_headers, 1):
                cell_value = allocation_sheet.cell(row=4, column=col).value
                if cell_value != header:
                    issues.append(f"Team Allocation header mismatch: expected '{header}', got '{cell_value}'")
        
        if issues:
            print("Validation Issues Found:")
            for issue in issues:
                print(f"  - {issue}")
            return False
        else:
            print("Workbook validation passed!")
            return True
            
    except Exception as e:
        print(f"Error validating workbook: {e}")
        return False

def main():
    """Main utility function with command-line interface."""
    if len(sys.argv) < 2:
        print("Team Task Planner Utilities")
        print("Usage:")
        print("  python3 utilities.py add_team <workbook_path> <team_name>")
        print("  python3 utilities.py export_csv <workbook_path> <team_name> <output_file>")
        print("  python3 utilities.py weekly_report <workbook_path> <output_file>")
        print("  python3 utilities.py validate <workbook_path>")
        return
    
    command = sys.argv[1]
    
    if command == "add_team" and len(sys.argv) == 4:
        workbook_path = sys.argv[2]
        team_name = sys.argv[3]
        add_new_team_sheet(workbook_path, team_name)
        
    elif command == "export_csv" and len(sys.argv) == 5:
        workbook_path = sys.argv[2]
        team_name = sys.argv[3]
        output_file = sys.argv[4]
        export_team_tasks_to_csv(workbook_path, team_name, output_file)
        
    elif command == "weekly_report" and len(sys.argv) == 4:
        workbook_path = sys.argv[2]
        output_file = sys.argv[3]
        generate_weekly_report(workbook_path, output_file)
        
    elif command == "validate" and len(sys.argv) == 3:
        workbook_path = sys.argv[2]
        validate_workbook(workbook_path)
        
    else:
        print("Invalid command or arguments. Use without arguments to see usage.")

if __name__ == "__main__":
    main()