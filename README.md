# Team Productivity Tracker

## Overview
The Team Productivity Tracker is an Excel VBA solution designed to analyze and visualize team productivity metrics. It processes daily work logs and generates comprehensive dashboards for both weekly and monthly performance analysis.

## Features

### 1. Productivity Dashboard
- **Weekly Summary**: Tracks team performance on a weekly basis
  - Week start and end dates
  - Number of team members active each week
  - Count of team members who met/exceeded weekly targets
  - Weekly productivity percentage

- **Monthly Summary**: Aggregates weekly data into monthly insights
  - Month and year
  - Monthly performance metrics
  - Monthly productivity percentage

### 2. Monthly Breakdown
- Detailed individual performance metrics by month
- Tracks hours worked per team member
- Compares actual hours against target (140 hours/month)
- Color-coded productivity indicators:
  - 🟢 100%+ of target
  - 🟡 90-99% of target
  - 🔴 Below 90% of target

## Prerequisites
- Microsoft Excel 2010 or later
- Macros must be enabled

## Setup Instructions

1. **Prepare Your Data**
   - Ensure you have two worksheets in your workbook:
     - `Output`: Contains daily work entries with columns for date, name, and hours
     - `OutputNE`: Contains non-entry related tasks with similar structure

2. **Import the Macro**
   - Open the VBA Editor (Alt + F11)
   - Import the `TeamProductivity.bas` module
   - Save the workbook as a macro-enabled workbook (.xlsm)

3. **Run the Macro**
   - Press `Alt + F8`, select `CalculateProductivityMetrics`, and click "Run"
   - Or assign the macro to a button for one-click access

## Output
After running the macro, two new worksheets will be created/updated:

1. **ProductivityDashboard**
   - Left side: Weekly productivity metrics
   - Right side: Monthly productivity metrics
   - Automatic formatting and calculations

2. **MonthlyBreakdown**
   - Detailed view of each team member's monthly performance
   - Sorted chronologically by month and alphabetically by name
   - Visual indicators for at-a-glance performance assessment

## Usage Notes
- The macro automatically detects and processes all available data
- Weekly target: 32.5 hours per team member
- Monthly target: 140 hours per team member (32.5 hours/week × 4.3 weeks)
- The dashboard updates the "Last Updated" timestamp automatically

## Troubleshooting
- **Missing Worksheets**: Ensure both `Output` and `OutputNE` worksheets exist
- **Macro Not Running**: Check Excel's macro security settings and enable macros
- **Incorrect Data**: Verify that date and hour columns contain valid data

## Support
For assistance or feature requests, please contact your system administrator or the development team.

---
*Last Updated: May 2025*
