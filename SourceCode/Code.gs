function doGet(e) {
  // Get the email of the active user
  var userEmail = Session.getActiveUser().getEmail();
  var template;

  // Check if the user is authorized and get their role
  var userInfo = isAuthorizedUser(userEmail);
  if (!userInfo.isAuthorized) {
    return HtmlService.createHtmlOutput('Access denied: Your email is not authorized.');
  }

  // Get the employee ID
  var loggedId = getEmployeeId(userEmail);

  // Check if there are parameters indicating the page to load
  if (e.parameter.page == 'applydetails') {
    // Only allow managers (HR) to access applydetails
    if (userInfo.isManager) {
      template = HtmlService.createTemplateFromFile('HR_applydetails');
      template.employeeId = e.parameter.employeeId;
      template.timestamp = e.parameter.timestamp;
      template.loggedId = loggedId;
      return template.evaluate();
    } else {
      return HtmlService.createHtmlOutput('Access denied: You do not have permission to view this page.');
    }
  }

  if (e.parameter.page == 'calendar') {
    template = HtmlService.createTemplateFromFile('Calendar');
    template.userRole = userInfo.isManager ? 'manager' : 'staff'; // Pass the role to the HTML template
    template.loggedId = loggedId; // Pass the logged ID to the template
    return template.evaluate();
  }

  // Default case: Load 'HR_main' for managers, 'E_main' for non-managers
  if (userInfo.isManager) {
    template = HtmlService.createTemplateFromFile('HR_main');
  } else {
    template = HtmlService.createTemplateFromFile('E_main');
  }

  template.loggedId = loggedId; // Pass the logged ID to the template
  return template.evaluate();
}

function isAuthorizedUser(email) {
  var sheet = SpreadsheetApp.openById('1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI').getSheetByName('EmployeeData');

  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === email) {
      return {
        isAuthorized: true,
        isManager: data[i][4] === 'Manager'
      };
    }
  }

  return {
    isAuthorized: false,
    isManager: false
  };
}

function getEmployeeId(email) {
  var sheet = SpreadsheetApp.openById('1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI').getSheetByName('EmployeeData');
  
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][3] === email) {
      return data[i][1]; // Assuming employee ID is in the first column
    }
  }
  
  return 'Unknown ID';
}

function getTodayLeaves() {
  var sheetId = '1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc'; // Replace with your Google Sheet ID
  var sheetName = 'Form Responses 1'; // Replace with your sheet name

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  var today = new Date().toDateString();
  var todaysLeaves = [];

  // Assuming headers are in the first row
  var headers = data[0];

  // Find the column indices for relevant data
  var startDateIndex = headers.indexOf('Start Date');
  var endDateIndex = headers.indexOf('End Date');
  var statusIndex = headers.indexOf('Status');

  // Process each row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var startDate = new Date(row[startDateIndex]).toDateString();
    var endDate = new Date(row[endDateIndex]).toDateString();

    if (startDate <= today && endDate >= today && row[statusIndex] === 'Approved') {
      todaysLeaves.push(row);
    }
  }

  // Return data as JSON
  return JSON.stringify({ headers: headers, leaves: todaysLeaves });
}

function getMonthlyLeaves() {
  var sheetId = '1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc'; // Replace with your Google Sheet ID
  var sheetName = 'Form Responses 1'; // Replace with your sheet name

  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var data = sheet.getDataRange().getValues();

  // Determine the start and end dates of the current month
  var today = new Date();
  var startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  var endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  var leavesByEmployee = {};

  // Assuming headers are in the first row
  var headers = data[0];

  // Find the column indices for relevant data
  var fullNameIndex = headers.indexOf('Employee Name');  // Index for Full Name (adjust as necessary)
  var startDateIndex = headers.indexOf('Start Date');
  var endDateIndex = headers.indexOf('End Date');
  var statusIndex = headers.indexOf('Status');

  // Process each row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var fullName = row[fullNameIndex];
    var startDate = new Date(row[startDateIndex]);
    var endDate = new Date(row[endDateIndex]);

    // Check if the leave falls within the current month and is approved
    if ((startDate <= endOfMonth && endDate >= startOfMonth) && row[statusIndex] === 'Approved') {
      if (!leavesByEmployee[fullName]) {
        leavesByEmployee[fullName] = [];
      }

      // Collect leave dates
      var date = new Date(startDate);
      while (date <= endDate) {
        if (date >= startOfMonth && date <= endOfMonth) {
          leavesByEmployee[fullName].push(date.toDateString());
        }
        date.setDate(date.getDate() + 1);
      }
    }
  }

  // Return data as JSON
  return JSON.stringify({ headers: headers, leavesByEmployee: leavesByEmployee });
}

function redirectToCalendar() {
  var scriptUrl = 'https://script.google.com/macros/s/AKfycbzQViWRtIWQwZHabBnth7WIAhDdZWEaJOtpsYm3fhc/exec';
  return scriptUrl + "?page=calendar";
}

function redirectToApplication() {
  var scriptUrl = 'https://script.google.com/macros/s/AKfycbzQViWRtIWQwZHabBnth7WIAhDdZWEaJOtpsYm3fhc/exec';
  return scriptUrl + "?page=HR_main";
}

function redirectToHistory() {
  var scriptUrl = 'https://script.google.com/macros/s/AKfycbzQViWRtIWQwZHabBnth7WIAhDdZWEaJOtpsYm3fhc/exec';
  return scriptUrl + "?page=E_main";
}

function getDetailsPageUrl(employeeId, timestamp) {
  const baseUrl = 'https://script.google.com/macros/s/AKfycbzQViWRtIWQwZHabBnth7WIAhDdZWEaJOtpsYm3fhc/exec';
  const url = `${baseUrl}?page=applydetails&employeeId=${encodeURIComponent(employeeId)}&timestamp=${encodeURIComponent(timestamp)}`;
  return url;
}

function getMainPageUrl() {
  return ScriptApp.getService().getUrl();
}

// Function to format date and time
function formatDateTime(timestamp) {
  const date = new Date(timestamp);
  const options = {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  };
  return date.toLocaleString('en-US', options);
}

// Function to calculate leave days
function calculateLeaveDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const oneDay = 24 * 60 * 60 * 1000; // hours*minutes*seconds*milliseconds
  return Math.round(Math.abs((end - start) / oneDay)) + 1; // +1 to include the end date
}

// Function to retrieve leave data, count approved leave applications for this month, and return as JSON
function getLeaveApplications() {
  const leaveSheet = SpreadsheetApp.openById('1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc').getSheetByName('Form Responses 1');
  const employeeSheet = SpreadsheetApp.openById('1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI').getSheetByName('EmployeeData');

  const leaveData = leaveSheet.getDataRange().getValues();
  const employeeData = employeeSheet.getDataRange().getValues();

  // Remove header row from leave data
  leaveData.shift();
  employeeData.shift();

  // Create a map to store employee details
  const employeeMap = new Map();
  employeeData.forEach(row => {
    const [fullName, employeeId] = row; // Extract only the necessary details
    employeeMap.set(employeeId, { fullName });
  });

  // Get current month and year
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-based index (0 = January)
  const currentYear = now.getFullYear();

  // Create a map to store total approved leave days for each employee this month
  const leaveDaysMap = new Map();

  // Calculate total approved leave days for each employee
  leaveData.forEach(row => {
    const [timestamp, leaveType, startDate, endDate, reason, email, employeeId, status] = row;
    const startDateObj = new Date(startDate);

    if (status === 'Approved' && startDateObj.getMonth() === currentMonth && startDateObj.getFullYear() === currentYear) {
      const leaveDays = calculateLeaveDays(startDate, endDate); // Calculate leave days using the helper function
      const currentDays = leaveDaysMap.get(employeeId) || 0;
      leaveDaysMap.set(employeeId, currentDays + leaveDays); // Accumulate leave days
    }
  });

  // Append employee details to leave data and filter by status 'Pending'
  const result = leaveData.filter(row => row[7] === 'Pending').map(row => {
    const [timestamp, leaveType, startDate, endDate, reason, email, employeeId, status] = row;
    const employeeDetails = employeeMap.get(employeeId) || {};
    const approvedLeaveDays = leaveDaysMap.get(employeeId) || 0;

    return [
      timestamp,
      startDate,
      endDate,
      reason,
      leaveType,
      status,
      employeeDetails.fullName || 'Unknown',
      approvedLeaveDays, // Show approved leave days here
      employeeId
    ];
  });

  // Sort data: emergency leave first, then by timestamp
  result.sort(function (a, b) {
    if (a[4] === 'Emergency' && b[4] !== 'Emergency') return -1;
    if (a[4] !== 'Emergency' && b[4] === 'Emergency') return 1;
    return new Date(b[0]) - new Date(a[0]);
  });

  // Return data as JSON
  return JSON.stringify(result);
}

// Function to update leave application status and send an email notification
function updateLeaveApplicationStatus(employeeId, timestamp, status) {
  try {
    const leaveSheet = SpreadsheetApp.openById('1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc').getSheetByName('Form Responses 1');
    const employeeSheet = SpreadsheetApp.openById('1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI').getSheetByName('EmployeeData');

    const leaveData = leaveSheet.getDataRange().getValues();
    const employeeData = employeeSheet.getDataRange().getValues();

    // Remove header row from employee data
    employeeData.shift();

    // Find the row in the leave sheet that matches the employeeId and timestamp
    let leaveRow;
    for (let i = 1; i < leaveData.length; i++) {
      const rowTimestamp = leaveData[i][0];
      const rowEmployeeId = leaveData[i][6];
      if (rowEmployeeId.toString() === employeeId.toString() &&
        new Date(rowTimestamp).getTime() === new Date(timestamp).getTime()) {
        leaveRow = i + 1; // account for header row
        break;
      }
    }

    if (leaveRow) {
      leaveSheet.getRange(leaveRow, 8).setValue(status); // Update status

      const employee = employeeData.find(row => row[1].toString() === employeeId.toString());
      if (employee) {
        if (status === 'Approved') {
          const leaveDays = calculateLeaveDays(leaveData[leaveRow - 1][2], leaveData[leaveRow - 1][3]); // Adjusted index

          const employeeRow = employeeData.findIndex(row => row[1].toString() === employeeId.toString()) + 2; // +2 for header row and 0-based index
          const currentBalance = employeeSheet.getRange(employeeRow, 12).getValue(); // Assuming balance is in column 11
          const updatedBalance = Math.max(currentBalance - leaveDays, 0);

          employeeSheet.getRange(employeeRow, 12).setValue(updatedBalance); // Update the leave balance
        }

        // Send email notification
        const email = employee[3]; // Ensure correct index for email
        const fullName = employee[0];

        // Format the timestamp
        const formattedTimestamp = formatDateTime(timestamp);

        // Send email notification
        const subject = 'Leave Application Status Update';
        const message = `Dear ${fullName},\n\nYour leave application submitted on ${formattedTimestamp} has been ${status}.\n\nBest regards,\nHR Department`;
        MailApp.sendEmail(email, subject, message);

      } else {
        console.log('Employee not found for ID:', employeeId);
      }
      return 'Success: Leave application updated';
    } else {
      return 'Error: Leave application not found';
    }
  } catch (error) {
    console.error('Error updating leave application status:', error);
    throw error;  // Re-throw the error so it's caught by the failure handler
  }
}

// Function to calculate leave days
function calculateLeaveDays(startDate, endDate) {
  const start = new Date(startDate);
  const end = new Date(endDate);
  const oneDay = 24 * 60 * 60 * 1000; // hours*minutes*seconds*milliseconds
  return Math.round(Math.abs((end - start) / oneDay)) + 1; // +1 to include the end date
}

function getLeaveDetailById(leaveId, timestamp) {
  const leaveResponsesSheetId = '1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc';
  const employeeDataSheetId = '1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI';

  const leaveResponsesSheet = SpreadsheetApp.openById(leaveResponsesSheetId).getSheetByName('Form Responses 1');
  const employeeDataSheet = SpreadsheetApp.openById(employeeDataSheetId).getSheetByName('EmployeeData');

  if (!leaveResponsesSheet || !employeeDataSheet) {
    throw new Error('One or more sheets could not be found.');
  }

  // Get data from 'leave responses' sheet
  const leaveResponsesData = leaveResponsesSheet.getDataRange().getValues();
  const employeeData = employeeDataSheet.getDataRange().getValues();

  // Find the leave record
  let leaveRecord = null;
  for (let i = 1; i < leaveResponsesData.length; i++) { // Assuming the first row is header
    if (leaveResponsesData[i][6] == leaveId &&
      new Date(leaveResponsesData[i][0]).getTime() == new Date(timestamp).getTime()) {
      leaveRecord = leaveResponsesData[i];
      break;
    }
  }

  if (!leaveRecord) {
    throw new Error('Leave record not found.');
  }

  // Find the employee record
  let employeeRecord = null;
  for (let j = 1; j < employeeData.length; j++) {
    if (employeeData[j][1] == leaveRecord[6]) { // Assuming column index for Employee ID in leave responses
      employeeRecord = employeeData[j];
      break;
    }
  }

  if (!employeeRecord) {
    throw new Error('Employee record not found.');
  }

  // Get current month and year
  const today = new Date();
  const currentMonth = today.getMonth();
  const currentYear = today.getFullYear();

  // Count approved leave applications for the current month
  let leaveTakenThisMonth = 0;
  for (let k = 1; k < leaveResponsesData.length; k++) {
    const leave = leaveResponsesData[k];
    const leaveStartDate = new Date(leave[2]);
    const leaveEndDate = new Date(leave[3]);
    const leaveStartMonth = leaveStartDate.getMonth();
    const leaveStartYear = leaveStartDate.getFullYear();
    const leaveEndMonth = leaveEndDate.getMonth();
    const leaveEndYear = leaveEndDate.getFullYear();

    // Check if the leave is for the current employee, approved, and intersects with the current month
    if (leave[6] == leaveId && leave[7] === 'Approved') {
      // Define the start and end dates of the leave period within the current month
      const startOfMonth = new Date(currentYear, currentMonth, 1);
      const endOfMonth = new Date(currentYear, currentMonth + 1, 0);
      const startDate = new Date(Math.max(leaveStartDate, startOfMonth));
      const endDate = new Date(Math.min(leaveEndDate, endOfMonth));

      // Calculate leave days only if the leave period falls within the current month
      if (startDate <= endDate) {
        const leaveDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        leaveTakenThisMonth += leaveDays;
      }
    }
  }

  const result = {
    fullName: employeeRecord[0],
    employeeId: employeeRecord[1],
    email: employeeRecord[3],
    department: employeeRecord[5],
    title: employeeRecord[4],
    employmentStatus: employeeRecord[6],
    dateOfHire: employeeRecord[7],
    manager: employeeRecord[8],
    salary: employeeRecord[9],
    timestamp: leaveRecord[0],
    typeOfLeave: leaveRecord[1],
    leaveTakenThisMonth: leaveTakenThisMonth,
    startDate: leaveRecord[2],
    endDate: leaveRecord[3],
    reason: leaveRecord[4],
    status: leaveRecord[7]
  };
  
  return JSON.stringify(result);
}

function getLeaveData(employeeEmail) {
  const leaveSheetId = '1aysd7NbGzfPwwAZaVvMtBxbMyfxvZLV180VdG3PEhIc';
  const leaveSheetName = 'Form Responses 1';
  const employeeSheetId = '1WLcwxpfmKbb9fHDFRfweXSo-Lm4dnwoacYSXgZKihHI';
  const employeeSheetName = 'EmployeeData';

  const leaveResponsesSheet = SpreadsheetApp.openById(leaveSheetId).getSheetByName(leaveSheetName);
  const employeeSheet = SpreadsheetApp.openById(employeeSheetId).getSheetByName(employeeSheetName);

  if (!leaveResponsesSheet || !employeeSheet) {
    throw new Error('One or both sheets could not be found.');
  }

  // Get data from 'leave responses' sheet
  const leaveResponsesData = leaveResponsesSheet.getDataRange().getValues();
  const headers = leaveResponsesData[0];
  const employeeEmailIndex = headers.indexOf('Email Address');
  const startDateIndex = headers.indexOf('Start Date');
  const endDateIndex = headers.indexOf('End Date');
  const timestampIndex = headers.indexOf('Timestamp');

  // Get data from 'employee' sheet
  const employeeData = employeeSheet.getDataRange().getValues();
  const balanceHeaders = employeeData[0];
  const balanceEmailIndex = balanceHeaders.indexOf('Email');
  const balanceAmountIndex = balanceHeaders.indexOf('Leave Balances');

  // Collect leave records for the given employeeEmail
  const leaveRecords = [];
  for (let i = 1; i < leaveResponsesData.length; i++) {
    if (leaveResponsesData[i][employeeEmailIndex] === employeeEmail) {
      const startDate = leaveResponsesData[i][startDateIndex];
      const endDate = leaveResponsesData[i][endDateIndex];
      const timestampStr = leaveResponsesData[i][timestampIndex];

      leaveRecords.push({
        typeOfLeave: leaveResponsesData[i][headers.indexOf('Leave Type')],
        startDate: startDate,  // Keep original startDate
        endDate: endDate,      // Keep original endDate
        reason: leaveResponsesData[i][headers.indexOf('Reason')],
        status: leaveResponsesData[i][headers.indexOf('Status')],
        timestamp: timestampStr // Keep original timestamp
      });
    }
  }

  if (leaveRecords.length === 0) {
    throw new Error('No leave records found for this employee.');
  }

  // Sort records by startDate in descending order
  leaveRecords.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));

  // Separate recent leave and history
  const recentLeave = leaveRecords.slice(0, 1); // Adjust as needed for recent data
  const history = leaveRecords
    .filter(record => record.status === 'Approved') // Filter approved leaves for history
    .sort((a, b) => new Date(b.startDate) - new Date(a.startDate)); // Sort history by startDate

  // Find leave balance for the employee
  let leaveBalance = 0;
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][balanceEmailIndex] === employeeEmail) {
      leaveBalance = employeeData[i][balanceAmountIndex];
      break;
    }
  }

  // Return the filtered and sorted records along with leave balance
  return JSON.stringify({
    recentLeave,
    history,
    leaveBalance
  });
}
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}
