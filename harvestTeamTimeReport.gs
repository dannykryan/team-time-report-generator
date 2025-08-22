// Development Settings
let folderId = "ABCDLoEkqY4YIxxqj0CIE1vQTsBGK1234"; // replace with your test Google Drive folder ID
let recipientEmails = [
  "myEmail@gmail.com", // replace with your email for testing
];

let employer = "ACME Corp"; // Replace with the employer name you want to filter by

// Api urls
const budgetReportApi =
  "https://api.harvestapp.com/v2/reports/project_budget?&per_page=2000";
const usersApiUrl = "https://api.harvestapp.com/v2/users";
const rolesApiUrl = "https://api.harvestapp.com/v2/roles";
const invoicesApiUrl = `https://api.harvestapp.com/v2/invoices`;
// projectApiUrl and timeEntriesApiUrl will require a project ID appended
const projectApiUrl = "https://api.harvestapp.com/v2/projects/";
const timeEntriesApiUrl =
  "https://api.harvestapp.com/v2/time_entries?project_id=";
const invoiceApiUrl = "https://api.harvestapp.com/v2/invoices?project_id=";

// Script Properties include Harvest token and ID needed for API calls
const scriptProperties = PropertiesService.getScriptProperties();
const accessToken = scriptProperties.getProperty("harvestAccessToken");
const accountId = scriptProperties.getProperty("harvestAccountID");
const leaveDatesAccessToken = scriptProperties.getProperty("leaveDatesAccessToken");
const leaveDatesDepartmentId = "80f513c6-1b58-48e4-8bc6-db69ab624804"

// Create API headers object using Harvest token and ID
const headers = {
  "Harvest-Account-Id": accountId,
  Authorization: `Bearer ${accessToken}`,
  "User-Agent": "Budget Report AppScript Integration",
};

// Get and format the date
const currentDate = new Date();
const timeZone = "Europe/London"; // London timezone
const formattedDate = Utilities.formatDate(
  currentDate,
  timeZone,
  "yyyy/MM/dd HH:mm"
);

// This script will be run on the first day of the month
// and will fetch the previous month's data
const currentYear = currentDate.getFullYear();
const currentMonth = currentDate.getMonth();

// Determine the start and end dates of the previous month
const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;
const prevMonthYear = currentMonth === 0 ? currentYear - 1 : currentYear;

const firstDayOfPrevMonth = new Date(prevMonthYear, prevMonth, 1);
const lastDayOfPrevMonth = new Date(prevMonthYear, prevMonth + 1, 0);

const fromFormatted = Utilities.formatDate(firstDayOfPrevMonth, timeZone, "yyyy-MM-dd");
const toFormatted = Utilities.formatDate(lastDayOfPrevMonth, timeZone, "yyyy-MM-dd");

const fromFormattedDisplay = Utilities.formatDate(firstDayOfPrevMonth, timeZone, "dd-MM-yyyy");
const toFormattedDisplay = Utilities.formatDate(lastDayOfPrevMonth, timeZone, "dd-MM-yyyy");

// Reusable function to fetch data from an API - includes automatic retry logic
async function fetchFromApi(url, options, retries = 3, delay = 1000) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      // Wrap the synchronous UrlFetchApp.fetch in a Promise
      const response = await new Promise((resolve, reject) => {
        try {
          const result = UrlFetchApp.fetch(url, options);
          resolve(result);
        } catch (error) {
          reject(error);
        }
      });

      // Check response code
      if (response.getResponseCode() === 200) {
        return JSON.parse(response.getContentText());
      } else {
        throw new Error(
          `Attempt ${attempt} failed: ${response.getResponseCode()} ${response.getContentText()}`
        );
      }
    } catch (error) {
      Logger.log(`Attempt ${attempt} failed: ${error.message}`);
      if (attempt < retries) {
        Logger.log(`Retrying in ${delay} ms...`);
        Utilities.sleep(delay); // Apps Script utility to add delay
      } else {
        throw new Error(`All ${retries} attempts failed.`);
      }
    }
  }
}

// Fetch all pages
async function fetchAllPages({
  url,
  method = "GET",
  muteHttpExceptions = true,
} = {}) {
  let pages = [];
  let page = 1;

  Logger.log(`Fetching Data from Harvest. Please wait...`);
  try {
    while (url) {
      const response = await fetchFromApi(url, {
        method,
        headers,
        muteHttpExceptions,
      });

      pages.push(response);
      if (response.links && response.links.next) {
        url = response.links.next;
        page++;
      } else {
        url = null; // Exit the loop if no more pages
      }
    }
  } catch (error) {
    Logger.log("Error fetching all time entries: " + error.message);
  }

  return pages;
}

// Fetch time entries
async function fetchAllTimeEntries(userId) {
  try {
    let allTimeEntriesData = [];
    let nextPageUrl = `https://api.harvestapp.com/v2/time_entries?from=${fromFormatted}&to=${toFormatted}&user_id=${userId}&page=1`;

    const pages = await fetchAllPages({ url: nextPageUrl }); // Fetch all pages

    if (pages && pages.length > 0) {
      allTimeEntriesData = pages.reduce(
        (previous, current) => previous.concat(current.time_entries),
        []
      );
    }

    return allTimeEntriesData;
  } catch (error) {
    Logger.log("Error fetching all time entries: " + error.message);
    return [];
  }
}

async function fetchEmploymentHolidays(departmentId, fromDate, toDate) {
  console.log(`Fetching employment holidays for department ID: ${departmentId} from ${fromDate} to ${toDate}`);
  console.log(`Using LeaveDates API token: ${leaveDatesAccessToken}`);
  const fromIso = Utilities.formatDate(fromDate, "GMT", "yyyy-MM-dd'T'00:00:00XXX");
  const toIso = Utilities.formatDate(toDate, "GMT", "yyyy-MM-dd'T'23:59:59XXX");
  
  const url = `https://api.leavedates.com/reports/employment-holidays?company=075ddb37-246c-43fe-9a1d-97c41af946d6&page=1&report_type=detail-report&department=${departmentId}&within=${encodeURIComponent(fromIso)},${encodeURIComponent(toIso)}`;

  console.log(`LeaveDates API URL: ${url}`);
  const options = {
    method: "get",
    headers: {
      // Add your Authorization header here if needed
      Authorization: `Bearer ${leaveDatesAccessToken}`,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true,
  };
  const response = await fetchFromApi(url, options);
  Logger.log("LeaveDates API response: " + JSON.stringify(response));
  return response.data;
}


function getUniqueHolidayDates(holidays) {
  const dateSet = new Set();
  if (!holidays || !Array.isArray(holidays)) {
    return [];
  }
  holidays.forEach(entry => {
    if (entry.holiday_type === "non_working_day") {
      dateSet.add(entry.holiday_date);
    }
  });
  return Array.from(dateSet);
}

// Define team members and their Harvest user IDs
// Harvest user IDs must be correct for this to work
const teamObj = {
  "Randy Savage": {
    id: 4248562,
  },
  "Dougal McGuire": {
    id: 4248564,
  },
  "David Liebe Hart": {
    id: 4645739,
  },
  "Miles Prower": {
    id: 4323387,
  },
};

async function main() {
  const holidays = await fetchEmploymentHolidays(leaveDatesDepartmentId, firstDayOfPrevMonth, lastDayOfPrevMonth);
  const uniqueHolidayDates = getUniqueHolidayDates(holidays);

  console.log("Unique Holiday Dates: ", uniqueHolidayDates);
  
  const teamMemberTimeEntries = {};

  for (const userName in team) {
    const user = team[userName];
    const userId = user.id;
    const timeEntries = await fetchAllTimeEntries(userId);

    Logger.log("Please Wait... Fetching time entries for " + userName);

    if (timeEntries && timeEntries.length > 0) {
      const clientBillable = timeEntries.filter(
        (entry) => entry.billable && entry.client.name !== employer && entry.client.name !== employer + "Marketing" && !entry.task.name.startsWith("Leave")
      );
      const clientUnbillable = timeEntries.filter(
        (entry) => !entry.billable && entry.client.name !== employer && entry.client.name !== employer + "Marketing" && !entry.task.name.startsWith("Leave")
      );
      const employer = timeEntries.filter(
        (entry) => entry.client.name === employer && !entry.task.name.startsWith("Leave")
      );
      const employerMarketing = timeEntries.filter(
        (entry) => entry.client.name === employer + "Marketing" && !entry.task.name.startsWith("Leave")
      );
      const pitchProposal = timeEntries.filter(
        (entry) => entry.task.name === "Pitch / Proposal" && !entry.task.name.startsWith("Leave")
      );
      const annualLeave = timeEntries.filter(
        (entry) => entry.task.name == "Leave - Annual Leave" && 
        !uniqueHolidayDates.includes(entry.spent_date)
      );
      const leaveOther = timeEntries.filter(
        (entry) => entry.task.name == "Leave - Other"
      );

      teamMemberTimeEntries[userName] = {
        clientBillableHours: {
          total: clientBillable.reduce((total, entry) => total + entry.hours, 0),
          clients: clientBillable.reduce((clientHours, entry) => {
            const clientName = entry.client.name;
            clientHours[clientName] = (clientHours[clientName] || 0) + entry.hours;
            return clientHours;
          }, {}),
        },
        clientUnbillableHours: {
          total: clientUnbillable.reduce((total, entry) => total + entry.hours, 0),
          clients: clientUnbillable.reduce((clientHours, entry) => {
            const clientName = entry.client.name;
            clientHours[clientName] = (clientHours[clientName] || 0) + entry.hours;
            return clientHours;
          }, {}),
          employerMarketingHours: employerMarketing.reduce((total, entry) => total + entry.hours, 0),
          pitchProposalHours: pitchProposal.reduce((total, entry) => total + entry.hours, 0),
        },
        employerHours: {
          total: employer.reduce((total, entry) => total + entry.hours, 0),
          tasks: employer.reduce((taskHours, entry) => {
            const taskName = entry.task.name;
            taskHours[taskName] = (taskHours[taskName] || 0) + entry.hours;
            return taskHours;
          }, {}),
          annualLeaveHours: annualLeave.reduce((total, entry) => total + entry.hours, 0),
          leaveOtherHours: leaveOther.reduce((total, entry) => total + entry.hours, 0),
        },
      };
    } else {
      teamMemberTimeEntries[userName] = {
        clientBillableHours: { total: 0, clients: {} },
        clientUnbillableHours: { 
          total: 0, 
          clients: {}, 
          employerMarketingHours: 0,
          pitchProposalHours: 0,
        },
        employerHours: { 
          total: 0, 
          tasks: {},
          annualLeaveHours: 0,
          leaveOtherHours: 0,
        },
      };
    }
  }

  // Create the spreadsheet
  const spreadsheetUrls = createSpreadsheet(teamMemberTimeEntries);
  Logger.log(`Spreadsheet URLs: ${JSON.stringify(spreadsheetUrls, null, 2)}`); // Log all URLs

  sendEmail(spreadsheetUrls); // Send the email with all URLs

  Logger.log(`An email has now been sent to the recipient email(s) with links to all of the sheets created by this script.`); // Log all URLs

  return spreadsheetUrls; // Return the URLs
}

function createSpreadsheet(data) {
  const folder = DriveApp.getFolderById(folderId);
  const spreadsheetUrls = {}; // Object to store URLs for each user

  for (const userName in data) {
    const ss = SpreadsheetApp.create(
      "Harvest CST Report - " + userName + " (" + fromFormattedDisplay + " to " + toFormattedDisplay + ")"
    );
    const ssId = ss.getId();
    const ssFile = DriveApp.getFileById(ssId);
    folder.addFile(ssFile);
    DriveApp.getRootFolder().removeFile(ssFile);

    const userData = data[userName];

    writeClientBillableSheet(ss, userData);
    writeClientUnbillableSheet(ss, userData);
    writeEmployerSheet(ss, userData);
    writeSummarySheet(ss, userData);

    spreadsheetUrls[userName] = ss.getUrl();
  }

  return spreadsheetUrls; // Return all URLs
}

function writeClientBillableSheet(ss, data) {
  const sheet = ss.getSheetByName("Client Billable Time") || ss.insertSheet("Client Billable Time");
  sheet.clearContents(); // Clear previous data

  const headers = []; // Array to store all header ranges for bolding
  let currentRow = 1; // Keep track of the current row
  let dataStartRow = 0; // Track where the data starts for coloring

  // Write Client Billable Time
  sheet.appendRow(["Client", "Percentage", "Total Hours"]);
  const clientBillableColumnHeaderRange = sheet.getRange(currentRow, 1, 1, 3);
  headers.push(clientBillableColumnHeaderRange);
  clientBillableColumnHeaderRange.setFontWeight("bold");
  currentRow++;
  dataStartRow = currentRow; // Data starts after headers

  const clientBillableData = [];
  let totalHours = 0; // Initialize total hours counter
  
  if (data.clientBillableHours && Object.keys(data.clientBillableHours.clients).length > 0) {
    const clientBillableTotal = data.clientBillableHours.total;
    for (const client in data.clientBillableHours.clients) {
      const hours = data.clientBillableHours.clients[client];
      totalHours += hours; // Accumulate total hours
      const percentage = clientBillableTotal > 0 ? (hours / clientBillableTotal) * 100 : 0;
      clientBillableData.push([client, percentage.toFixed(2) + "%", hours]);
    }
    clientBillableData.forEach(row => sheet.appendRow(row));
    const dataRange = sheet.getRange(dataStartRow, 1, clientBillableData.length, 3);
    dataRange.setBackgroundColor("#E0F2F7");

    // Add a Total row
    sheet.appendRow(["Total", "", totalHours]); // Empty string for alignment
    const totalRowRange = sheet.getRange(currentRow + clientBillableData.length, 1, 1, 3);
    totalRowRange.setFontWeight("bold");
    totalRowRange.setBackgroundColor("#B3E5FC"); // Light blue for distinction

    // Add a pie chart for Client Billable Hours
    const chartRange = sheet.getRange(dataStartRow - 1, 1, clientBillableData.length + 1, 2); // Include headers
    addPieChart(sheet, "Client Billable Time", chartRange);
  } else {
    sheet.appendRow(["No Client Billable Time Data Available"]);
    const noDataRange = sheet.getRange(currentRow, 1, 1, 3);
    noDataRange.setBackgroundColor("#E0F2F7");
  }

  autoResizeColumnsWithPadding(sheet, 3);
}

function writeClientUnbillableSheet(ss, data) {
  const sheet = ss.getSheetByName("Client Unbillable Time") || ss.insertSheet("Client Unbillable Time");
  sheet.clearContents();

  const headers = [];
  let currentRow = 1;
  let dataStartRow = 0;

  // Write Client Unbillable Time
  sheet.appendRow(["Client", "Percentage", "Total Hours"]);
  const clientUnbillableColumnHeaderRange = sheet.getRange(currentRow, 1, 1, 3);
  headers.push(clientUnbillableColumnHeaderRange);
  clientUnbillableColumnHeaderRange.setFontWeight("bold");
  currentRow++;
  dataStartRow = currentRow;

  const clientUnbillableData = [];
  let totalHours = 0;
  
  // Add regular client unbillable time
  if (data.clientUnbillableHours && Object.keys(data.clientUnbillableHours.clients).length > 0) {
    for (const client in data.clientUnbillableHours.clients) {
      const hours = data.clientUnbillableHours.clients[client];
      totalHours += hours;
      clientUnbillableData.push([client, 0, hours]); // Initialize percentage as 0, will calculate later
    }
  }
  
  // Add Employer Marketing hours
  if (data.clientUnbillableHours.employerMarketingHours > 0) {
    totalHours += data.clientUnbillableHours.employerMarketingHours;
    clientUnbillableData.push(["Employer Marketing", 0, data.clientUnbillableHours.employerMarketingHours]);
  }
  
  // Add Pitch/Proposal hours
  if (data.clientUnbillableHours.pitchProposalHours > 0) {
    totalHours += data.clientUnbillableHours.pitchProposalHours;
    clientUnbillableData.push(["Pitch / Proposal", 0, data.clientUnbillableHours.pitchProposalHours]);
  }
  
  // Calculate percentages
  if (totalHours > 0) {
    for (let i = 0; i < clientUnbillableData.length; i++) {
      const percentage = (clientUnbillableData[i][2] / totalHours) * 100;
      clientUnbillableData[i][1] = percentage.toFixed(2) + "%";
    }
  }
  
  if (clientUnbillableData.length > 0) {
    clientUnbillableData.forEach(row => sheet.appendRow(row));
    const dataRange = sheet.getRange(dataStartRow, 1, clientUnbillableData.length, 3);
    dataRange.setBackgroundColor("#F8CECC");

    // Add a Total row
    sheet.appendRow(["Total", "", totalHours]);
    const totalRowRange = sheet.getRange(dataStartRow + clientUnbillableData.length, 1, 1, 3);
    totalRowRange.setFontWeight("bold");
    totalRowRange.setBackgroundColor("#ea9999");

    // Add a pie chart for Client Unbillable Hours
    const chartRange = sheet.getRange(dataStartRow - 1, 1, clientUnbillableData.length + 1, 2); // Include headers
    addPieChart(sheet, "Client Unbillable Time", chartRange);
  } else {
    sheet.appendRow(["No Client Unbillable Time Data Available"]);
    const noDataRange = sheet.getRange(currentRow, 1, 1, 3);
    noDataRange.setBackgroundColor("#F8CECC");
  }

  autoResizeColumnsWithPadding(sheet, 3);
}

function writeEmployerSheet(ss, data) {
  const sheet = ss.getSheetByName("Employer Time") || ss.insertSheet("Employer Time");
  sheet.clearContents();

  const headers = [];
  let currentRow = 1;
  let dataStartRow = 0;

  // Write Employer Time
  sheet.appendRow(["Task", "Percentage", "Total Hours"]);
  const employerColumnHeaderRange = sheet.getRange(currentRow, 1, 1, 3);
  headers.push(employerColumnHeaderRange);
  employerColumnHeaderRange.setFontWeight("bold");
  currentRow++;
  dataStartRow = currentRow;

  const employerData = [];
  let totalHours = 0;
  
  // Add regular Employer tasks
  if (data.employerHours && Object.keys(data.employerHours.tasks).length > 0) {
    for (const task in data.employerHours.tasks) {
      const hours = data.employerHours.tasks[task];
      totalHours += hours;
      employerData.push([task, 0, hours]); // Initialize percentage as 0, will calculate later
    }
  }
  
  // Add Annual Leave hours
  if (data.employerHours.annualLeaveHours > 0) {
    totalHours += data.employerHours.annualLeaveHours;
    employerData.push(["Annual Leave", 0, data.employerHours.annualLeaveHours]);
  }

  // Add Leave - Other hours
  if (data.employerHours.leaveOtherHours > 0) {
    totalHours += data.employerHours.leaveOtherHours;
    employerData.push(["Leave - Other", 0, data.employerHours.leaveOtherHours]);
  }
  
  // Calculate percentages
  if (totalHours > 0) {
    for (let i = 0; i < employerData.length; i++) {
      const percentage = (employerData[i][2] / totalHours) * 100;
      employerData[i][1] = percentage.toFixed(2) + "%";
    }
  }
  
  if (employerData.length > 0) {
    employerData.forEach(row => sheet.appendRow(row));
    const dataRange = sheet.getRange(dataStartRow, 1, employerData.length, 3);
    dataRange.setBackgroundColor("#D5F5E3");

    // Add a Total row
    sheet.appendRow(["Total", "", totalHours]);
    const totalRowRange = sheet.getRange(dataStartRow + employerData.length, 1, 1, 3);
    totalRowRange.setFontWeight("bold");
    totalRowRange.setBackgroundColor("#b6d7a8");

    // Add a pie chart for Employer Time
    const chartRange = sheet.getRange(dataStartRow - 1, 1, employerData.length + 1, 2); // Include headers
    addPieChart(sheet, "Employer Time", chartRange);
  } else {
    sheet.appendRow(["No Employer Time Data Available"]);
    const noDataRange = sheet.getRange(currentRow, 1, 1, 3);
    noDataRange.setBackgroundColor("#D5F5E3");
  }

  autoResizeColumnsWithPadding(sheet, 3);
  
  // Delete the default "Sheet1"
  const sheets = ss.getSheets();
  if (sheets.length > 1) {
    ss.deleteSheet(sheets[0]);
  }
}

function writeSummarySheet(ss, data) {
  const sheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");
  sheet.clearContents();

  let currentRow = 1;
  let dataStartRow = 0;

  // Write Summary header
  sheet.appendRow(["Category", "Percentage", "Total Hours"]);
  const summaryHeaderRange = sheet.getRange(currentRow, 1, 1, 3);
  summaryHeaderRange.setFontWeight("bold");
  currentRow++;
  dataStartRow = currentRow;

  // Calculate total hours for each category
  const clientBillableTotal = data.clientBillableHours.total || 0;
  
  // Calculate client unbillable total (including Employer Marketing and Pitch/Proposal)
  let clientUnbillableTotal = data.clientUnbillableHours.total || 0;
  clientUnbillableTotal += data.clientUnbillableHours.employerMarketingHours || 0;
  clientUnbillableTotal += data.clientUnbillableHours.pitchProposalHours || 0;
  
  // Calculate employer total (including Annual Leave)
  let employerTotal = data.employerHours.total || 0;
  employerTotal += data.employerHours.annualLeaveHours += data.employerHours.leaveOtherHours || 0;
  
  // Calculate grand total for percentage calculation
  const grandTotal = clientBillableTotal + clientUnbillableTotal + employerTotal;
  
  // Prepare data rows
  const summaryData = [
    ["Client Billable Time", ((clientBillableTotal / grandTotal) * 100).toFixed(2) + "%", clientBillableTotal],
    ["Client Unbillable Time", ((clientUnbillableTotal / grandTotal) * 100).toFixed(2) + "%", clientUnbillableTotal],
    ["Employer Time", ((employerTotal / grandTotal) * 100).toFixed(2) + "%", employerTotal]
  ];
  
  // Add data rows with appropriate color coding
  summaryData.forEach((row, index) => {
    sheet.appendRow(row);
    
    // Apply appropriate background colors to match the other sheets
    const rowRange = sheet.getRange(currentRow, 1, 1, 3);
    rowRange.setBackgroundColor("#faf4aa"); // Client Billable color
    
    currentRow++;
  });
  
  // Add Total row
  sheet.appendRow(["Total", "100.00%", grandTotal]);
  const totalRowRange = sheet.getRange(currentRow, 1, 1, 3);
  totalRowRange.setFontWeight("bold");
  totalRowRange.setBackgroundColor("#fff251");
  
  // Add pie chart for Summary
  const chartRange = sheet.getRange(dataStartRow - 1, 1, summaryData.length + 1, 2); // Include headers
  addPieChart(sheet, "Time Distribution Summary", chartRange);
  
  autoResizeColumnsWithPadding(sheet, 3);
  
  // Move Summary to be the first sheet
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
}

function cropSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();

  if (lastRow < maxRows) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
  if (lastColumn < maxColumns) {
    sheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }
}

function autoResizeColumnsWithPadding(sheet, columns) {
  for (let i = 0; i < columns; i++) {
    let maxWidth = 0;

    // Iterate through all rows in the column to find the longest text
    for (let j = 1; j <= sheet.getLastRow(); j++) {
      const cell = sheet.getRange(j, i + 1);
      const value = cell.getValue();
      if (value) {
        const textLength = String(value).length;
        // Calculate width with a bit of padding (adjust as needed)
        const cellWidth = textLength * 8 + 10; // 8 pixels per character + 10 pixels padding
        maxWidth = Math.max(maxWidth, cellWidth);
      }
    }
    
    sheet.setColumnWidth(i + 1, Math.max(maxWidth, 50));
  }
}

function addBarChart(sheet, title, range) {
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(range)
    .setPosition(2, 5, 0, 0)
    .setOption('title', title)
    .setOption('vAxis.title', 'Percentage')
    .setOption('hAxis.title', 'Client')
    .build();

  sheet.insertChart(chart);
}

function addPieChart(sheet, title, range) {
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(2, 5, 0, 0)
    .setOption('title', title)
    .build();

  sheet.insertChart(chart);
}

function sendEmail(sheetUrls) {
  const subject = "Automated Harvest Team Time Report " + formattedDate;
  let body = `Hello,\n\nPlease find the Team Time Reports from ${fromFormatted} to ${toFormatted} in the following Google Sheets:\n\n`;

  for (const userName in sheetUrls) {
    body += `${userName}: ${sheetUrls[userName]}\n`;
  }

  body += "\nAs this is an automated report.";

  MailApp.sendEmail({
    to: recipientEmails.join(","),
    subject: subject,
    body: body,
  });
}