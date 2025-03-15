// Leave Management System - Core Functions

// Constants and Configuration
const SHEET_NAMES = {
  LEAVE_REQUESTS: 'LeaveRequests',
  USER_ROLES: 'UserRoles',
  LEAVE_TYPES: 'LeaveTypes',
  AUDIT_LOG: 'AuditLog',
  SETTINGS: 'Settings',
  SYSTEM_STATE: 'SystemState'
};

const STATUS = {
  PENDING: 'Pending',
  IN_REVIEW: 'In Review',
  APPROVED: 'Approved',
  REJECTED: 'Rejected'
};

const SYSTEM_STATE = {
  UNINITIALIZED: 'UNINITIALIZED',
  INITIALIZING: 'INITIALIZING',
  READY: 'READY',
  ERROR: 'ERROR'
};

// Error handling utility
function handleError(error, context) {
  console.error(`Error in ${context}:`, error);
  logAction('system', 'error', `${context}: ${error.toString()}`);
  return {
    success: false,
    error: error.toString(),
    context: context
  };
}

// System state management
function getSystemState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const stateSheet = ss.getSheetByName(SHEET_NAMES.SYSTEM_STATE);
    if (!stateSheet) {
      return SYSTEM_STATE.UNINITIALIZED;
    }
    return stateSheet.getRange('A1').getValue() || SYSTEM_STATE.UNINITIALIZED;
  } catch (error) {
    return SYSTEM_STATE.ERROR;
  }
}

function setSystemState(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let stateSheet = ss.getSheetByName(SHEET_NAMES.SYSTEM_STATE);
    if (!stateSheet) {
      stateSheet = ss.insertSheet(SHEET_NAMES.SYSTEM_STATE);
      stateSheet.hideSheet();
    }
    stateSheet.getRange('A1').setValue(state);
    return true;
  } catch (error) {
    console.error('Error setting system state:', error);
    return false;
  }
}

// Role Management
function getUserRole() {
  try {
    const currentUserEmail = Session.getEffectiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roleSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    
    if (!roleSheet) {
      throw new Error('User roles sheet not found');
    }
    
    const data = roleSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('User roles data not initialized');
    }
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === currentUserEmail) {
        return data[i][1].toLowerCase();
      }
    }
    
    // Log unknown user attempt
    logAction('system', 'access_attempt', `Unknown user access attempt: ${currentUserEmail}`);
    return "unknown";
  } catch (error) {
    handleError(error, 'getUserRole');
    return "error";
  }
}

// Leave ID Generation with validation
function generateLeaveId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    
    if (!sheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const timestamp = new Date().getTime();
    const newId = `LR${timestamp}${lastRow + 1}`;
    
    // Verify ID uniqueness
    const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    if (existingIds.includes(newId)) {
      // In the unlikely case of collision, add random suffix
      return `${newId}_${Math.random().toString(36).substr(2, 5)}`;
    }
    
    return newId;
  } catch (error) {
    handleError(error, 'generateLeaveId');
    throw error; // Propagate error as this is a critical function
  }
}

// Enhanced User Access Management
function getAllowedLeaveTypes(userEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    const leaveTypesSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_TYPES);
    
    if (!userRolesSheet || !leaveTypesSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Get user data with validation
    const userData = userRolesSheet.getDataRange().getValues();
    if (userData.length < 2) {
      throw new Error('User roles data not initialized');
    }
    
    let userAllowedTypes = [];
    let userRole = '';
    let userDepartment = '';
    let userFound = false;
    
    // Find user's role and department
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === userEmail) {
        userRole = userData[i][1];
        userDepartment = userData[i][3];
        userAllowedTypes = userData[i][4] ? userData[i][4].split(',').map(type => type.trim()) : [];
        userFound = true;
        break;
      }
    }
    
    if (!userFound) {
      throw new Error(`User ${userEmail} not found in role assignments`);
    }
    
    // Get leave types data with validation
    const leaveTypesData = leaveTypesSheet.getDataRange().getValues();
    if (leaveTypesData.length < 2) {
      throw new Error('Leave types data not initialized');
    }
    
    const allowedTypes = [];
    
    // Filter leave types with enhanced validation
    for (let i = 1; i < leaveTypesData.length; i++) {
      const typeId = leaveTypesData[i][0];
      const typeName = leaveTypesData[i][1];
      const maxDays = leaveTypesData[i][3];
      const requiresApproval = leaveTypesData[i][4];
      const applicableRoles = leaveTypesData[i][5] ? 
        leaveTypesData[i][5].split(',').map(role => role.trim()) : [];
      
      if (applicableRoles.includes(userRole) || userAllowedTypes.includes(typeId)) {
        if (!typeId || !typeName || !maxDays) {
          logAction('system', 'data_error', 
            `Invalid leave type data found for type ${typeId}`);
          continue;
        }
        
        allowedTypes.push({
          id: typeId,
          name: typeName,
          maxDays: maxDays,
          requiresApproval: requiresApproval,
          department: userDepartment
        });
      }
    }
    
    if (allowedTypes.length === 0) {
      logAction('system', 'access_warning', 
        `No leave types available for user ${userEmail}`);
    }
    
    return allowedTypes;
  } catch (error) {
    handleError(error, 'getAllowedLeaveTypes');
    return [];
  }
}

// Email Notifications
function sendNotification(recipientEmail, subject, htmlBody) {
  GmailApp.sendEmail(
    recipientEmail,
    subject,
    "This email requires HTML to display properly.",
    { htmlBody: htmlBody }
  );
}

function notifyLeaveRequestUpdate(leaveId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
  const requestData = requestsSheet.getDataRange().getValues();
  
  // Find the request
  let requestRow;
  let employeeEmail;
  let managerEmail;
  let startDate;
  let endDate;
  let leaveType;
  
  for (let i = 1; i < requestData.length; i++) {
    if (requestData[i][0] === leaveId) {
      requestRow = requestData[i];
      employeeEmail = requestRow[1];
      managerEmail = requestRow[3];
      leaveType = requestRow[4];
      startDate = requestRow[5];
      endDate = requestRow[6];
      break;
    }
  }
  
  if (!requestRow) return;
  
  // Prepare email content
  const dateFormat = { year: 'numeric', month: 'long', day: 'numeric' };
  const formattedStartDate = startDate.toLocaleDateString(undefined, dateFormat);
  const formattedEndDate = endDate.toLocaleDateString(undefined, dateFormat);
  
  let emailSubject = `Leave Request ${status} - ${leaveId}`;
  let emailBody = `
    <h2>Leave Request Update</h2>
    <p>Leave request ${leaveId} has been <strong>${status}</strong>.</p>
    <h3>Details:</h3>
    <ul>
      <li>Leave Type: ${leaveType}</li>
      <li>Start Date: ${formattedStartDate}</li>
      <li>End Date: ${formattedEndDate}</li>
      <li>Status: ${status}</li>
    </ul>
  `;
  
  // Send notifications
  sendNotification(employeeEmail, emailSubject, emailBody);
  if (status === STATUS.PENDING) {
    sendNotification(managerEmail, `New Leave Request to Review - ${leaveId}`, emailBody);
  }
}

// Audit Logging
function logAction(userEmail, action, details, leaveId = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  
  auditSheet.appendRow([
    new Date(),
    action,
    userEmail,
    details,
    leaveId
  ]);
}

// UI Functions
function onOpen() {
  const ss = SpreadsheetApp.getActive();
  
  // Check if system is initialized
  if (!isSystemInitialized()) {
    ss.addMenu('Admin Actions', [
      {name: 'Initialize System', functionName: 'initializeSystem'}
    ]);
    return;
  }
  
  // Get user role and create appropriate menu
  const userRole = getUserRole();
  if (userRole === 'admin') {
    createAdminMenu(ss);
  } else if (userRole === 'manager') {
    createManagerMenu(ss);
  } else if (userRole === 'employee') {
    createEmployeeMenu(ss);
  }
}

// Check if system is initialized
function isSystemInitialized() {
  const ss = SpreadsheetApp.getActive();
  try {
    const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    return userRolesSheet !== null;
  } catch (e) {
    return false;
  }
}

// Core initialization without UI dependencies
function initializeSystemCore() {
  try {
    // First initialize the test data
    initializeTestData();
    
    // Then setup all interfaces
    const interfaces = setupAllInterfaces();
    
    // Share the interfaces with users
    shareInterfaceUrls();
    
    // Refresh menu
    onOpen();
    
    return {
      success: true,
      interfaces: interfaces,
      message: 'System initialized successfully'
    };
  } catch (error) {
    console.error('Initialization error:', error);
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to initialize system'
    };
  }
}

// UI version of initialization (for menu triggers)
function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const result = initializeSystemCore();
    
    if (result.success) {
      ui.alert(
        'System Initialized',
        'The leave management system has been set up successfully. ' +
        'All users will receive emails with their interface links.\n\n' +
        'Admin interface: ' + result.interfaces.adminUrl + '\n' +
        'Manager interface: ' + result.interfaces.managerUrl + '\n' +
        'Employee interface: ' + result.interfaces.employeeUrl,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', result.message + ': ' + result.error, ui.ButtonSet.OK);
    }
    
    return result;
  } catch (error) {
    ui.alert('Error', 'Failed to initialize system: ' + error.toString(), ui.ButtonSet.OK);
    return {
      success: false,
      error: error.toString(),
      message: 'Failed to initialize system'
    };
  }
}

// Direct initialization function (for API calls)
function initializeSystemDirect() {
  return initializeSystemCore();
}

// Share interfaces with new users
function shareNewUserAccess(email, role, managerEmail = '') {
  const ss = SpreadsheetApp.getActive();
  
  try {
    // Get interface URLs from properties
    const interfaces = {
      adminUrl: ss.getUrl(),
      managerUrl: PropertiesService.getDocumentProperties().getProperty('managerInterfaceUrl'),
      employeeUrl: PropertiesService.getDocumentProperties().getProperty('employeeInterfaceUrl')
    };
    
    // Prepare email content
    let subject = 'Welcome to Leave Management System';
    let body = `Dear ${role},\n\n`;
    body += 'Welcome to the Leave Management System! ';
    body += 'You have been granted access with the following role: ' + role + '\n\n';
    
    let interfaceUrl = '';
    switch(role.toLowerCase()) {
      case 'admin':
        interfaceUrl = interfaces.adminUrl;
        body += 'Access the admin interface here: ' + interfaceUrl + '\n';
        // Add admin as editor
        ss.addEditor(email);
        break;
        
      case 'manager':
        interfaceUrl = interfaces.managerUrl;
        body += 'Access the manager interface here: ' + interfaceUrl + '\n';
        // Add manager as viewer to manager interface
        const managerSheet = SpreadsheetApp.openByUrl(interfaceUrl);
        managerSheet.addViewer(email);
        // Add manager email to UserRoles sheet
        const managerUserRole = managerSheet.getSheetByName('UserRoles');
        if (!managerUserRole) {
          managerSheet.insertSheet('UserRoles', 0).hideSheet();
        }
        managerSheet.getSheetByName('UserRoles').getRange('A1').setValue(email);
        break;
        
      case 'employee':
        interfaceUrl = interfaces.employeeUrl;
        body += 'Access the employee interface here: ' + interfaceUrl + '\n';
        // Add employee as viewer to employee interface
        const employeeSheet = SpreadsheetApp.openByUrl(interfaceUrl);
        employeeSheet.addViewer(email);
        // Add employee and manager email to UserRoles sheet
        const employeeUserRole = employeeSheet.getSheetByName('UserRoles');
        if (!employeeUserRole) {
          employeeSheet.insertSheet('UserRoles', 0).hideSheet();
        }
        employeeSheet.getSheetByName('UserRoles').getRange('A1').setValue(email);
        employeeSheet.getSheetByName('UserRoles').getRange('B1').setValue(managerEmail);
        break;
    }
    
    body += '\nPlease note that you will need to be logged into your Google account to access the system.\n\n';
    body += 'Best regards,\nLeave Management System';
    
    // Send email
    GmailApp.sendEmail(email, subject, body);
    
    // Log action
    logAction('system', 'new_user_access', `Access granted to ${email} as ${role}`);
    
    return true;
  } catch (error) {
    console.error('Error sharing access:', error);
    return false;
  }
}

// Initialize the spreadsheet with test data
function initializeTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create sheets if they don't exist
  [SHEET_NAMES.LEAVE_REQUESTS, SHEET_NAMES.USER_ROLES, 
   SHEET_NAMES.LEAVE_TYPES, SHEET_NAMES.AUDIT_LOG, 
   SHEET_NAMES.SETTINGS].forEach(sheetName => {
    if (!ss.getSheetByName(sheetName)) {
      ss.insertSheet(sheetName);
    }
  });
  
  // Initialize UserRoles
  const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
  userRolesSheet.clear();
  userRolesSheet.getRange('A1:E1').setValues([['Email', 'Role', 'Manager Email', 'Department', 'Allowed Leave Types']]);
  userRolesSheet.getRange('A2:E7').setValues([
    ['admin@company.com', 'Admin', '', 'Administration', 'ALL'],
    ['manager1@company.com', 'Manager', '', 'Engineering', 'SL,VAC,PL,TL'],
    ['manager2@company.com', 'Manager', '', 'Marketing', 'SL,VAC,PL'],
    ['emp1@company.com', 'Employee', 'manager1@company.com', 'Engineering', 'SL,VAC,TL'],
    ['emp2@company.com', 'Employee', 'manager1@company.com', 'Engineering', 'SL,VAC,PL'],
    ['emp3@company.com', 'Employee', 'manager2@company.com', 'Marketing', 'SL,VAC']
  ]);
  
  // Initialize LeaveTypes
  const leaveTypesSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_TYPES);
  leaveTypesSheet.clear();
  leaveTypesSheet.getRange('A1:F1').setValues([['Type ID', 'Type Name', 'Description', 'Max Days', 'Requires Approval', 'Applicable Roles']]);
  leaveTypesSheet.getRange('A2:F5').setValues([
    ['SL', 'Sick Leave', 'Medical and health-related leave', 12, true, 'Admin,Manager,Employee'],
    ['VAC', 'Vacation', 'Annual vacation leave', 20, true, 'Admin,Manager,Employee'],
    ['PL', 'Personal Leave', 'Personal time off', 5, true, 'Admin,Manager,Employee'],
    ['TL', 'Training Leave', 'Professional development', 10, true, 'Admin,Manager,Employee']
  ]);
  
  // Initialize LeaveRequests
  const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
  requestsSheet.clear();
  requestsSheet.getRange('A1:L1').setValues([[
    'Leave ID', 'Employee Email', 'Employee Name', 'Manager Email', 
    'Leave Type', 'Start Date', 'End Date', 'Reason', 
    'Submission Timestamp', 'Status', 'Manager Comment', 'Decision Timestamp'
  ]]);
  
  // Initialize AuditLog
  const auditSheet = ss.getSheetByName(SHEET_NAMES.AUDIT_LOG);
  auditSheet.clear();
  auditSheet.getRange('A1:E1').setValues([['Timestamp', 'Action', 'User Email', 'Details', 'Leave ID']]);
  
  // Log initialization
  logAction('system', 'initialization', 'Test data initialized');
}

// Function to process user updates from the form
function processUserUpdate(formData) {
  try {
    // For new user
    if (!formData.editIndex) {
      return addNewUser(
        formData.email,
        formData.role,
        formData.manager,
        formData.department,
        formData.leaveTypes
      );
    }
    
    // For editing existing user
    const ss = SpreadsheetApp.getActive();
    const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    const rowIndex = parseInt(formData.editIndex) + 1;
    
    userRolesSheet.getRange(rowIndex, 1, 1, 5).setValues([[
      formData.email,
      formData.role,
      formData.manager,
      formData.department,
      formData.leaveTypes
    ]]);
    
    return {
      success: true,
      message: 'User updated successfully'
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Function to handle user deletion
function deleteUser(email) {
  try {
    return removeUser(email);
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}
