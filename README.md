# Harvest Team Time Report Automation

This Google Apps Script automates the generation of time and budget reports for your team using data from both the Harvest and LeaveDates APIs. It creates individual Google Sheets for each team member and emails links to the recipients. The script can be run manually or scheduled to run automatically at a time of your choosing.

---
Example of a 'Client Billable Time' sheet output from the app:
![image](https://github.com/user-attachments/assets/594c642d-7a48-4f71-8f9b-cb56bcf973aa)

---

## Table of Contents

- [Features](#features)
- [How It Works](#how-it-works)
- [Configuration](#configuration)
- [Manual Execution](#manual-execution)
- [Automated Scheduling](#automated-scheduling)
- [Code Structure](#code-structure)
- [Adding/Removing Team Members](#addingremoving-team-members)
- [API Integration](#api-integration)
- [Error Handling & Logging](#error-handling--logging)
- [Extending the Script](#extending-the-script)
- [Support](#support)

---

## Features

- Fetches time entries, client/project data, and leave/holiday information from Harvest and LeaveDates APIs.
- Aggregates and categorizes hours into:
  - Client Billable
  - Client Unbillable
  - Employer Time (billed to the employer, e.g., Annual Leave and Other Leave)
- Generates a Google Sheet per team member with four tabs which each inclue pie charts: 
  1. Summary - A summary of total of the three other tabs
  2. Client Billable TIme - All billable client hours logged by team member
  3. CLient Unbillable Time - All non-billable client hours logged by team member 
  4. Employer Time - All logged 'Employer' hours. Includes Annual leave, internal meetings, training, etc.
- Emails links to the generated reports to configured recipients (set in the recipientEmails array).
- Can be run manually from Google Apps Script or on a schedule (e.g., monthly).

---

## How It Works

1. **Fetches Data:**  
   - Retrieves time entries for each team member for the previous month from Harvest.
   - Fetches leave and holiday data for the same period from Leave Dates.

2. **Processes Data:**  
   - Categorizes hours by client, task, and leave type.
   - Calculates totals and percentages for each category.

3. **Generates Reports:**  
   - Creates a Google Sheet for each team member.
   - Populates sheets with detailed breakdowns and summary charts.

4. **Sends Email:**  
   - Emails the links to the generated reports to the configured recipients.

---

## Configuration

### Script Properties

Set the following script properties in Google Apps Script (`File > Project properties > Script properties`):

- `harvestAccessToken`: Your Harvest API access token.
- `harvestAccountID`: Your Harvest account ID.
- `leaveDatesAccessToken`: Your LeaveDates API access token.

### Folder and Recipients

- **Production/Development:**  
  Set the correct `folderId` and `recipientEmails` at the top of the script for production or development use.
  Please comment out the 'Production settings' and update the 'Development Settings' to use your details.
  If you receive an error informing you that you don't own the folder set at `folderId`, you will need to set it use the id of a Google Drive folder that you own. Failure to do so can result in a permissions error.

---

## Manual Execution

To run the script manually for the first time:

1. Open [Google Apps Script](https://script.google.com/) and paste the code into a new project, just be sure to update the secrets in the project settings, and set the `folderId` at the top of the script to the id of a google Drive folder that you own.
2. Select the `main` function from the 'select function to run' dropdown from the toolbar in the editor.
3. Click the **Run** button.

The script will fetch data, generate reports, and send emails as described above.

---

## Automated Scheduling

The script is designed to run automatically at the first of each month, to deliver a report for the previous month.
To set up a trigger yourself on a new project:

1. In Google Apps Script, go to **Triggers** (clock icon in the left sidebar).
2. Click **Add Trigger**.
3. Choose the `main` function.
4. Set the event source to **Time-driven**.
5. Choose a schedule (e.g., "Month timer" > "1st").
6. Save the trigger.

---

## Code Structure

- **API URLs & Headers:**  
  Defined at the top for Harvest and LeaveDates.

- **Team Members:**  
  The `teamObj` object contains the names and Harvest user IDs for each team member.

- **Main Logic:**  
  - `main()`: Orchestrates fetching, processing, report generation, and emailing.
  - `fetchAllTimeEntries()`: Fetches all time entries for a user.
  - `fetchEmploymentHolidays()`: Fetches holiday data from LeaveDates.
  - `createSpreadsheet()`: Creates and populates Google Sheets.
  - `sendEmail()`: Sends the report links to recipients.

- **Sheet Writers:**  
  - `writeClientBillableSheet()`
  - `writeClientUnbillableSheet()`
  - `writeEmployerSheet()`
  - `writeSummarySheet()`

- **Utilities:**  
  - `fetchFromApi()`, `fetchAllPages()`: Handle API requests and pagination.
  - `autoResizeColumnsWithPadding()`, `addPieChart()`, etc.: Formatting and chart helpers.

---

## Adding/Removing Team Members

It is easy to edit the teamObj to add or remove anyone you like, regardless of the team they're on. Simply add their name and their Harvest ID to the object.

To add or remove team members, update the `teamObj` object:

```javascript
const teamObj = {
  "Randy Savage": { id: 4248562 },
  // Add or remove members here
};
```

---

## API Integration

- **Harvest:**  
  Used for time entries, users, roles, projects, and invoices.
- **LeaveDates:**  
  Used for employment holidays and non-working days.

All API calls use secure tokens stored in script properties.

---

## Error Handling & Logging

- The script uses `Logger.log()` and `console.log()` for debugging and error reporting.
- API calls have retry logic and log failures.
- All errors are logged to the Apps Script log, accessible via **View > Logs** in the Apps Script editor.

---

## Extending the Script

- **To add new report categories or sheets:**  
  Create new functions similar to the existing `write*Sheet` functions.
- **To change email recipients or folder:**  
  Update the `recipientEmails` array and `folderId` variable at the top of the script.
- **To change the reporting period:**  
  Adjust the date logic near the top of the script.

---

## Support

For questions or maintenance, review the code comments and refer to this documentation.  
If you encounter issues with API access, check your script properties and permissions.

---

**Maintainer:**  
Danny Ryan  
[dannykryan@marketingpod.com](mailto:dannykryan@gmail.com)

