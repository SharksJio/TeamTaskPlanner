# Dynamic Configuration Guide

## Overview

The Team Task Planner now supports dynamic configuration, allowing you to customize teams, organizations, and project details without modifying Python code. This guide explains how to use the new dynamic features.

## Configuration System

### Configuration File Structure

The system uses JSON configuration files with the following structure:

```json
{
  "organization": {
    "name": "Your Organization",
    "project_name": "Team Task Planner"
  },
  "teams": [
    {
      "name": "Development Team",
      "description": "Software development tasks",
      "sample_tasks": [
        {
          "task_id": "DEV001",
          "task_name": "Feature Implementation",
          "assigned_to": "Developer Name",
          "priority": "High",
          "status": "In Progress",
          "start_date": "2024-01-01",
          "due_date": "2024-01-15",
          "progress": 50,
          "hours_est": 20,
          "hours_actual": 10,
          "dependencies": "",
          "subtasks": "Analysis, Code, Test",
          "notes": "Sample task"
        }
      ]
    }
  ],
  "task_headers": [
    "Task ID", "Task Name", "Assigned To", "Priority", "Status",
    "Start Date", "Due Date", "Progress %", "Hours Est.", "Hours Actual",
    "Dependencies", "Subtasks", "Notes"
  ],
  "status_options": ["Not Started", "Planning", "In Progress", "Completed", "On Hold"],
  "priority_options": ["Low", "Medium", "High", "Critical"]
}
```

## Usage Methods

### Method 1: Create with Custom Configuration

1. **Create a configuration file** (e.g., `my_config.json`)
2. **Generate the Excel file:**
   ```bash
   python3 create_team_task_planner.py my_config.json
   ```

### Method 2: Auto-Detect from Existing Workbook

If you have an existing Team Task Planner Excel file:

1. **Extract configuration:**
   ```bash
   python3 utilities.py auto_config existing_file.xlsx extracted_config.json
   ```

2. **Use the extracted configuration:**
   ```bash
   python3 create_team_task_planner.py extracted_config.json
   ```

### Method 3: Interactive Setup

Use the enhanced setup script:

```bash
python3 setup.py
```

This will:
- Ask for organization name
- Collect team names interactively
- Create both configuration and Excel files
- Generate a quick start guide

## Adding Teams Dynamically

### Add New Team to Existing Workbook

```bash
python3 utilities.py add_team workbook.xlsx "New Team Name" [config_file.json]
```

Features:
- Automatically detects existing configuration
- Creates configuration file if none exists
- Updates master dashboard with new team
- Adds sample task for the new team

### Auto-Configuration Detection

```bash
python3 utilities.py auto_config workbook.xlsx [config_file.json]
```

This command:
- Extracts organization name from Master Dashboard
- Identifies all team sheets
- Creates a configuration file matching the workbook structure
- Preserves existing team structure

## Advanced Features

### Configuration Validation

The system automatically validates configuration files and provides helpful error messages for:
- Missing required fields
- Invalid team structures
- Malformed JSON

### Dynamic Formula Updates

When adding teams, the system:
- Updates master dashboard formulas automatically
- Adjusts cell ranges for new team count
- Maintains Excel formula integrity

### Template Matching

The system intelligently:
- Uses existing team sheets as templates
- Preserves formatting and structure
- Adapts to custom header configurations

## Examples

### Example 1: Software Company

```json
{
  "organization": {
    "name": "TechCorp",
    "project_name": "Software Development Tracker"
  },
  "teams": [
    {
      "name": "Frontend",
      "description": "React and Angular development"
    },
    {
      "name": "Backend", 
      "description": "API and server development"
    },
    {
      "name": "Mobile",
      "description": "iOS and Android apps"
    }
  ]
}
```

### Example 2: Marketing Agency

```json
{
  "organization": {
    "name": "Creative Agency",
    "project_name": "Campaign Management"
  },
  "teams": [
    {
      "name": "Creative",
      "description": "Design and content creation"
    },
    {
      "name": "Strategy",
      "description": "Campaign planning and analysis"
    },
    {
      "name": "Digital",
      "description": "Social media and digital marketing"
    }
  ]
}
```

## Best Practices

1. **Backup Before Changes:** Always backup your Excel files before adding teams
2. **Use Descriptive Names:** Team names become sheet names, so avoid special characters
3. **Configuration Management:** Keep configuration files in version control
4. **Regular Updates:** Use `auto_config` to keep configuration synchronized with Excel changes

## Troubleshooting

### Common Issues

1. **Invalid Sheet Names:** Avoid characters like `/`, `\`, `?`, `*` in team names
2. **Missing Configuration:** The system will create defaults but may not match your needs
3. **Formula Errors:** Run utilities after manual Excel changes to update formulas

### Error Messages

- `"Team already exists"` - Team name is already used in workbook
- `"Invalid character found"` - Team name contains Excel-forbidden characters
- `"Missing required configuration key"` - Configuration file is incomplete

## Migration from Static Templates

If you have existing static templates:

1. **Extract configuration:**
   ```bash
   python3 utilities.py auto_config old_template.xlsx
   ```

2. **Customize as needed** by editing the generated configuration file

3. **Generate new template:**
   ```bash
   python3 create_team_task_planner.py config.json
   ```

This approach maintains your existing structure while enabling dynamic management.