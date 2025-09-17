# Team Task Planner - Excel-Based Solution

A comprehensive Excel-based team task management system for multiple technical teams working on different projects with different deadlines.

## 🎯 Features

- **Master Dashboard**: Overall project tracking and key metrics
- **Team Allocation**: Resource management and availability tracking
- **Individual Team Sheets**: Detailed task management for each team
- **Weekly Planning**: Cross-team coordination and sprint planning
- **Templates & Instructions**: Usage guidelines and best practices

## 📊 Excel Workbook Structure

### 1. Master Dashboard
- Real-time overview of all teams and projects
- Automatic calculations of progress and workload
- Color-coded team performance indicators
- Key metrics including total tasks, completion rates, and resource utilization

### 2. Team Allocation
- Complete resource management system
- Employee availability tracking (Available/Busy/Overloaded)
- Skills matrix for each team member
- Next available dates for task assignment
- Team summary with availability statistics

### 3. Team Sheets (Frontend, Backend, DevOps, QA)
Each team has a dedicated sheet with:
- Task ID and naming system
- Assigned team members
- Priority levels (High/Medium/Low)
- Status tracking (Not Started/Planning/In Progress/Completed)
- Progress percentage
- Estimated vs actual hours
- Dependencies between tasks
- Subtask management
- Notes and comments

### 4. Weekly Planning
- Week-by-week task scheduling
- Cross-team coordination view
- Sprint planning support
- Daily task allocation

### 5. Templates & Instructions
- Complete usage guide
- Color coding system explanation
- Formula documentation
- Best practices for task management

## 🚀 Getting Started

### Prerequisites
- Microsoft Excel 2016 or later (recommended)
- Python 3.x with openpyxl (for regenerating the template)

### Quick Start
1. Download or generate the `Team_Task_Planner.xlsx` file
2. Open in Microsoft Excel
3. Start with the **Master Dashboard** for an overview
4. Navigate to individual team sheets to manage specific tasks
5. Use **Team Allocation** to check resource availability
6. Plan weekly activities in the **Weekly Planning** sheet

## 📈 How to Use

### Adding New Tasks
1. Go to the appropriate team sheet
2. Add a new row below existing tasks
3. Fill in all required columns:
   - Task ID (follow team naming convention)
   - Task Name (descriptive title)
   - Assigned To (team member name)
   - Priority (High/Medium/Low)
   - Status (current state)
   - Start and Due Dates
   - Progress percentage
   - Estimated hours
   - Dependencies (if any)
   - Subtasks (if applicable)
   - Notes

### Managing Team Members
1. Open the **Team Allocation** sheet
2. Add new team members in empty rows
3. Update their skills, availability, and current task load
4. The system will automatically color-code availability

### Tracking Progress
1. Update task progress percentages regularly
2. Change status as tasks move through workflow
3. Update actual hours spent
4. The Master Dashboard will automatically reflect changes

### Weekly Planning
1. Use the **Weekly Planning** sheet for sprint planning
2. Allocate tasks across days of the week
3. Coordinate cross-team dependencies
4. Plan team meetings and reviews

## 🎨 Color Coding System

### Priority Levels
- 🔴 **Red**: High Priority
- 🟡 **Yellow**: Medium Priority  
- 🟢 **Green**: Low Priority

### Task Status
- 🟢 **Green**: Completed
- 🔵 **Blue**: In Progress
- 🟡 **Yellow**: Planning
- ⚪ **Gray**: Not Started

### Resource Availability
- 🟢 **Green**: Available (<80% loaded)
- 🟡 **Yellow**: Busy (80-90% loaded)
- 🔴 **Red**: Overloaded (>90% loaded)

## 📊 Automated Calculations

The workbook includes several automated formulas:

### Master Dashboard
- Total task counts across all teams
- Overall progress percentages
- Team workload averages
- Critical issue tracking

### Team Sheets
- Task completion percentages
- Team progress summaries
- Individual task progress tracking

### Team Allocation
- Availability calculations
- Team size counts
- Workload distributions

## 🔧 Customization

### Adding New Teams
1. Copy an existing team sheet
2. Rename it with the new team name
3. Update the Master Dashboard to include the new team
4. Add team members to the Team Allocation sheet

### Modifying Task Fields
- Add new columns as needed
- Update formulas in Master Dashboard accordingly
- Maintain consistent formatting

### Changing Time Periods
- Modify the Weekly Planning sheet for different time frames
- Update date ranges in team sheets
- Adjust progress tracking periods

## 📝 Best Practices

1. **Regular Updates**: Update task progress weekly
2. **Consistent Naming**: Use standardized task IDs and naming conventions
3. **Dependency Tracking**: Clearly mark task dependencies
4. **Realistic Estimates**: Provide accurate time estimates
5. **Status Updates**: Keep task statuses current
6. **Resource Planning**: Monitor team availability regularly
7. **Communication**: Use notes field for important updates

## 🔄 Workflow Integration

### Sprint Planning
1. Start with Weekly Planning sheet
2. Identify available team members in Team Allocation
3. Assign tasks based on priority and dependencies
4. Monitor progress through individual team sheets
5. Review overall status in Master Dashboard

### Daily Standups
- Review current day's tasks in Weekly Planning
- Check individual progress in team sheets
- Identify blockers and dependencies
- Update task statuses

### Sprint Reviews
- Use Master Dashboard for overall progress review
- Analyze team performance metrics
- Plan next sprint based on completion rates
- Update resource allocation as needed

## 🛠️ Regenerating the Template

To recreate or modify the Excel template:

```bash
python3 create_team_task_planner.py
```

This will generate a new `Team_Task_Planner.xlsx` file with the latest structure and sample data.

## 📞 Support

For questions or issues:
1. Check the **Templates & Instructions** sheet in the Excel file
2. Review this documentation
3. Modify the Python script for custom requirements

## 🚀 Future Enhancements

Potential additions:
- Gantt chart integration
- Burndown chart templates
- Time tracking formulas
- Resource cost calculations
- Automated reporting macros
- Integration with project management tools

---

*This Team Task Planner provides a comprehensive solution for managing multiple technical teams, tracking progress, and optimizing resource allocation through an easy-to-use Excel interface.*