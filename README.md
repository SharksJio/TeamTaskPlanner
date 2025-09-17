# Team Task Planner 📊

A comprehensive Excel-based team task management system designed for multiple technical teams working on different projects with varying deadlines and resource requirements.

## 🌟 NEW: Dynamic Configuration System

**Major Update**: The Team Task Planner now supports dynamic configuration! Create and manage teams without regenerating Python templates.

### ✨ Key Dynamic Features:
- **JSON-based Configuration**: Define teams, organization, and task structures in config files
- **Auto-Detection**: Extract configuration from existing Excel workbooks
- **Dynamic Team Addition**: Add new teams to existing workbooks without coding
- **Template Matching**: Automatically adapts to your existing structure
- **Backward Compatible**: Works with existing templates

## 🚀 Quick Start (Dynamic)

### Option 1: Auto-Configure from Existing File
```bash
# Extract configuration from existing workbook
python3 utilities.py auto_config Team_Task_Planner.xlsx

# Add new teams dynamically
python3 utilities.py add_team Team_Task_Planner.xlsx "New Team Name"
```

### Option 2: Custom Configuration
```bash
# Create with custom configuration
python3 create_team_task_planner.py my_config.json
```

### Option 3: Interactive Setup
```bash
# Guided setup with custom teams
python3 setup.py
```

## 🌟 Overview

This tool provides a complete solution for:
- Managing multiple technical teams (Frontend, Backend, DevOps, QA)
- Tracking tasks, deadlines, and progress across projects
- Monitoring resource allocation and team member availability
- Weekly task planning and sprint coordination
- Automatic master plan updates when individual tasks change
- **Dynamic team and project management without code changes**

## 📦 What's Included

- **Team_Task_Planner.xlsx** - Ready-to-use Excel workbook
- **create_team_task_planner.py** - Python script to generate/regenerate Excel templates
- **config_manager.py** - Dynamic configuration management system
- **utilities.py** - Enhanced utilities for dynamic team management
- **setup.py** - Interactive setup wizard with custom configuration
- **config.json** - Default configuration file
- **DYNAMIC_CONFIG_GUIDE.md** - Comprehensive guide for dynamic features
- **DOCUMENTATION.md** - Comprehensive usage guide

## 🚀 Traditional Quick Start

1. **Download the Excel file**: `Team_Task_Planner.xlsx`
2. **Open in Microsoft Excel**
3. **Start with Master Dashboard** for overall project view
4. **Navigate to team sheets** for detailed task management
5. **Use Team Allocation** to check who's available for new tasks

## 📋 Key Features

### ✅ Master Dashboard
- Real-time overview of all teams and projects
- Automatic progress calculations
- Key performance metrics
- Color-coded team workload indicators

### ✅ Resource Management
- Team member availability tracking
- Skills matrix for each employee
- Workload distribution analysis
- Next available dates for task assignment

### ✅ Task Management
- Individual sheets for each team (Frontend, Backend, DevOps, QA)
- Priority and status tracking
- Progress monitoring
- Dependency management
- Subtask creation and tracking

### ✅ Weekly Planning
- Cross-team coordination
- Sprint planning support
- Daily task allocation
- Timeline visualization

### ✅ Automated Updates
- Master plan automatically reflects individual task changes
- Progress calculations update in real-time
- Resource allocation updates automatically
- Color-coded status indicators

## 🎯 Perfect For

- **Technical Teams**: Frontend, Backend, DevOps, QA teams
- **Project Managers**: Overseeing multiple projects and deadlines
- **Resource Managers**: Optimizing team allocation
- **Scrum Masters**: Sprint planning and progress tracking
- **Team Leads**: Managing individual team workloads

## 📊 Color Coding System

- 🟢 **Green**: Available, Completed, Low Priority
- 🟡 **Yellow**: Busy, In Progress, Medium Priority
- 🔴 **Red**: Overloaded, High Priority, Overdue
- 🔵 **Blue**: In Progress tasks
- ⚪ **Gray**: Not Started tasks

## 🛠️ Requirements

- Microsoft Excel 2016 or later
- Python 3.x (only for regenerating templates)

## 📚 Documentation

See [DOCUMENTATION.md](DOCUMENTATION.md) for detailed usage instructions, best practices, and customization options.

## 🔄 Working with Templates

### Dynamic Approach (Recommended)
```bash
# Create configuration from existing workbook
python3 utilities.py auto_config workbook.xlsx config.json

# Add teams dynamically without regeneration
python3 utilities.py add_team workbook.xlsx "New Team"

# Create new workbook with custom configuration
python3 create_team_task_planner.py config.json
```

### Traditional Regeneration
To create a fresh template or modify the structure:

```bash
python3 create_team_task_planner.py
```

This will generate a new `Team_Task_Planner.xlsx` with sample data and the latest structure.

## 📚 Documentation

- **[DYNAMIC_CONFIG_GUIDE.md](DYNAMIC_CONFIG_GUIDE.md)** - Complete guide for dynamic configuration features
- **[DOCUMENTATION.md](DOCUMENTATION.md)** - Detailed usage instructions and best practices

---

*Transform your team management with this comprehensive, Excel-based solution that grows with your organization's needs - now with dynamic configuration support!*
