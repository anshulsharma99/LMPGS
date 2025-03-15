// Leave Management System - UI Components

// Employee UI Functions with Enhanced Error Handling
function showLeaveRequestForm() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    if (!userEmail) {
      throw new Error('User email not available');
    }
    
    // Verify user role
    const userRole = getUserRole();
    if (userRole === 'unknown' || userRole === 'error') {
      throw new Error('Unable to verify user role');
    }
    
    // Get allowed leave types with validation
    const allowedTypes = getAllowedLeaveTypes(userEmail);
    if (!allowedTypes || allowedTypes.length === 0) {
      throw new Error('No leave types available for user');
    }
    
    try {
      // Create form HTML
      const template = HtmlService.createTemplateFromFile('LeaveRequestForm');
      template.allowedTypes = allowedTypes;
      
      const html = template.evaluate()
        .setWidth(400)
        .setHeight(600)
        .setTitle('Submit Leave Request');
      
      SpreadsheetApp.getUi().showModalDialog(html, 'Submit Leave Request');
      
      // Log successful form display
      logAction(userEmail, 'view_form', 'Leave request form displayed');
      
    } catch (error) {
      handleError(error, 'create_form');
      throw new Error('Failed to create leave request form');
    }
    
  } catch (error) {
    handleError(error, 'showLeaveRequestForm');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display leave request form: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showMyRequests() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    if (!userEmail) {
      throw new Error('User email not available');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    try {
      const requestsData = requestsSheet.getDataRange().getValues();
      if (!requestsData || requestsData.length < 1) {
        throw new Error('No leave request data available');
      }
      
      // Filter requests for current user with validation
      const myRequests = requestsData.filter((row, index) => {
        return index === 0 || row[1] === userEmail; // Include header row
      });
      
      if (myRequests.length < 1) {
        throw new Error('No leave requests found');
      }
      
      // Create HTML table with error handling
      try {
        const template = HtmlService.createTemplateFromFile('MyRequests');
        template.requests = myRequests;
        
        const html = template.evaluate()
          .setWidth(800)
          .setHeight(600)
          .setTitle('My Leave Requests');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'My Leave Requests');
        
        // Log successful display
        logAction(userEmail, 'view_requests', 'My leave requests displayed');
        
      } catch (error) {
        handleError(error, 'create_requests_view');
        throw new Error('Failed to create requests view');
      }
      
    } catch (error) {
      handleError(error, 'process_requests');
      throw new Error('Failed to process leave requests');
    }
    
  } catch (error) {
    handleError(error, 'showMyRequests');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display leave requests: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Manager UI Functions with Enhanced Error Handling
function showPendingRequests() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    if (!userEmail) {
      throw new Error('User email not available');
    }
    
    // Verify manager role
    const userRole = getUserRole();
    if (userRole !== 'manager' && userRole !== 'admin') {
      throw new Error('Unauthorized to view pending requests');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    try {
      const requestsData = requestsSheet.getDataRange().getValues();
      if (!requestsData || requestsData.length < 1) {
        throw new Error('No leave request data available');
      }
      
      // Filter pending requests with validation
      const pendingRequests = requestsData.filter((row, index) => {
        return index === 0 || (row[3] === userEmail && row[9] === STATUS.PENDING);
      });
      
      // Create HTML table with error handling
      try {
        const template = HtmlService.createTemplateFromFile('PendingRequests');
        template.requests = pendingRequests;
        
        const html = template.evaluate()
          .setWidth(800)
          .setHeight(600)
          .setTitle('Pending Leave Requests');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Pending Leave Requests');
        
        // Log successful display
        logAction(userEmail, 'view_pending', 
          `Pending requests displayed (${pendingRequests.length - 1} requests)`);
        
      } catch (error) {
        handleError(error, 'create_pending_view');
        throw new Error('Failed to create pending requests view');
      }
      
    } catch (error) {
      handleError(error, 'process_pending');
      throw new Error('Failed to process pending requests');
    }
    
  } catch (error) {
    handleError(error, 'showPendingRequests');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display pending requests: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showTeamHistory() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    if (!userEmail) {
      throw new Error('User email not available');
    }
    
    // Verify manager role
    const userRole = getUserRole();
    if (userRole !== 'manager' && userRole !== 'admin') {
      throw new Error('Unauthorized to view team history');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    try {
      const requestsData = requestsSheet.getDataRange().getValues();
      if (!requestsData || requestsData.length < 1) {
        throw new Error('No leave request data available');
      }
      
      // Filter team requests with validation
      const teamRequests = requestsData.filter((row, index) => {
        return index === 0 || row[3] === userEmail;
      });
      
      // Create HTML table with error handling
      try {
        const template = HtmlService.createTemplateFromFile('TeamHistory');
        template.requests = teamRequests;
        
        const html = template.evaluate()
          .setWidth(800)
          .setHeight(600)
          .setTitle('Team Leave History');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'Team Leave History');
        
        // Log successful display
        logAction(userEmail, 'view_history', 
          `Team history displayed (${teamRequests.length - 1} records)`);
        
      } catch (error) {
        handleError(error, 'create_history_view');
        throw new Error('Failed to create team history view');
      }
      
    } catch (error) {
      handleError(error, 'process_history');
      throw new Error('Failed to process team history');
    }
    
  } catch (error) {
    handleError(error, 'showTeamHistory');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display team history: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Admin UI Functions with Enhanced Error Handling
function showAdminSettings() {
  try {
    // Verify admin role
    const userRole = getUserRole();
    if (userRole !== 'admin') {
      throw new Error('Unauthorized to access admin settings');
    }
    
    try {
      const template = HtmlService.createTemplateFromFile('AdminSettings');
      
      const html = template.evaluate()
        .setWidth(800)
        .setHeight(600)
        .setTitle('System Settings');
      
      SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
      
      // Log successful display
      logAction(Session.getEffectiveUser().getEmail(), 
        'view_settings', 'Admin settings displayed');
      
    } catch (error) {
      handleError(error, 'create_settings_view');
      throw new Error('Failed to create admin settings view');
    }
    
  } catch (error) {
    handleError(error, 'showAdminSettings');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display admin settings: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showAllRequests() {
  try {
    // Verify admin role
    const userRole = getUserRole();
    if (userRole !== 'admin') {
      throw new Error('Unauthorized to view all requests');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    try {
      const requestsData = requestsSheet.getDataRange().getValues();
      if (!requestsData || requestsData.length < 1) {
        throw new Error('No leave request data available');
      }
      
      // Create HTML table with error handling
      try {
        const template = HtmlService.createTemplateFromFile('AllRequests');
        template.requests = requestsData;
        
        const html = template.evaluate()
          .setWidth(1000)
          .setHeight(800)
          .setTitle('All Leave Requests');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'All Leave Requests');
        
        // Log successful display
        logAction(Session.getEffectiveUser().getEmail(), 
          'view_all', `All requests displayed (${requestsData.length - 1} records)`);
        
      } catch (error) {
        handleError(error, 'create_all_requests_view');
        throw new Error('Failed to create all requests view');
      }
      
    } catch (error) {
      handleError(error, 'process_all_requests');
      throw new Error('Failed to process all requests');
    }
    
  } catch (error) {
    handleError(error, 'showAllRequests');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display all requests: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showUserManagement() {
  try {
    // Verify admin role
    const userRole = getUserRole();
    if (userRole !== 'admin') {
      throw new Error('Unauthorized to access user management');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    if (!userRolesSheet) {
      throw new Error('User roles sheet not found');
    }
    
    try {
      const userData = userRolesSheet.getDataRange().getValues();
      if (!userData || userData.length < 1) {
        throw new Error('No user data available');
      }
      
      // Create user management interface with error handling
      try {
        const template = HtmlService.createTemplateFromFile('UserManagement');
        template.users = userData;
        
        const html = template.evaluate()
          .setWidth(800)
          .setHeight(600)
          .setTitle('User Management');
        
        SpreadsheetApp.getUi().showModalDialog(html, 'User Management');
        
        // Log successful display
        logAction(Session.getEffectiveUser().getEmail(), 
          'view_users', `User management displayed (${userData.length - 1} users)`);
        
      } catch (error) {
        handleError(error, 'create_user_management_view');
        throw new Error('Failed to create user management view');
      }
      
    } catch (error) {
      handleError(error, 'process_user_data');
      throw new Error('Failed to process user data');
    }
    
  } catch (error) {
    handleError(error, 'showUserManagement');
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to display user management: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Form Processing Functions with Enhanced Validation
function processLeaveRequest(formData) {
  try {
    // Input validation
    if (!formData || !formData.leaveType || !formData.startDate || !formData.endDate || !formData.reason) {
      throw new Error('Missing required fields');
    }
    
    // Date validation
    const startDate = new Date(formData.startDate);
    const endDate = new Date(formData.endDate);
    const today = new Date();
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new Error('Invalid date format');
    }
    
    if (startDate < today) {
      throw new Error('Start date cannot be in the past');
    }
    
    if (endDate < startDate) {
      throw new Error('End date must be after start date');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    const userEmail = Session.getEffectiveUser().getEmail();
    
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    // Get user details with validation
    const userRolesSheet = ss.getSheetByName(SHEET_NAMES.USER_ROLES);
    if (!userRolesSheet) {
      throw new Error('User roles sheet not found');
    }
    
    const userData = userRolesSheet.getDataRange().getValues();
    let userName = '';
    let managerEmail = '';
    let userFound = false;
    
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === userEmail) {
        managerEmail = userData[i][2];
        userName = userData[i][1]; // Get user's name
        userFound = true;
        break;
      }
    }
    
    if (!userFound) {
      throw new Error('User not found in role assignments');
    }
    
    if (!managerEmail) {
      throw new Error('Manager email not found for user');
    }
    
    // Check leave balance
    const allowedTypes = getAllowedLeaveTypes(userEmail);
    const leaveType = allowedTypes.find(type => type.name === formData.leaveType);
    
    if (!leaveType) {
      throw new Error('Invalid leave type');
    }
    
    const daysRequested = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    
    // Check for overlapping leaves
    const existingRequests = requestsSheet.getDataRange().getValues();
    for (let i = 1; i < existingRequests.length; i++) {
      if (existingRequests[i][1] === userEmail && 
          existingRequests[i][9] !== STATUS.REJECTED) {
        const existingStart = new Date(existingRequests[i][5]);
        const existingEnd = new Date(existingRequests[i][6]);
        
        if ((startDate <= existingEnd && endDate >= existingStart)) {
          throw new Error('Leave request overlaps with existing request');
        }
      }
    }
    
    // Generate leave ID with validation
    const leaveId = generateLeaveId();
    if (!leaveId) {
      throw new Error('Failed to generate leave ID');
    }
    
    // Add request to sheet with validation
    try {
      requestsSheet.appendRow([
        leaveId,
        userEmail,
        userName,
        managerEmail,
        formData.leaveType,
        startDate,
        endDate,
        formData.reason,
        new Date(),
        STATUS.PENDING,
        '',
        ''
      ]);
      
      // Verify row was added
      const lastRow = requestsSheet.getLastRow();
      const addedId = requestsSheet.getRange(lastRow, 1).getValue();
      if (addedId !== leaveId) {
        throw new Error('Failed to add request to sheet');
      }
    } catch (error) {
      handleError(error, 'append_request');
      throw new Error('Failed to save leave request');
    }
    
    // Log the action
    logAction(userEmail, 'submit_request', 
      `New leave request submitted: ${daysRequested} days of ${formData.leaveType}`, 
      leaveId);
    
    // Send notifications
    try {
      notifyLeaveRequestUpdate(leaveId, STATUS.PENDING);
    } catch (error) {
      handleError(error, 'send_notification');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Failed to send notification for ' + leaveId);
    }
    
    return { 
      success: true, 
      message: 'Leave request submitted successfully', 
      leaveId: leaveId 
    };
    
  } catch (error) {
    handleError(error, 'processLeaveRequest');
    return { 
      success: false, 
      message: error.message 
    };
  }
}

function processLeaveDecision(leaveId, decision, comment) {
  try {
    // Input validation
    if (!leaveId || !decision) {
      throw new Error('Missing required fields');
    }
    
    if (!Object.values(STATUS).includes(decision)) {
      throw new Error('Invalid decision status');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestsSheet = ss.getSheetByName(SHEET_NAMES.LEAVE_REQUESTS);
    if (!requestsSheet) {
      throw new Error('Leave requests sheet not found');
    }
    
    const requestsData = requestsSheet.getDataRange().getValues();
    const userEmail = Session.getEffectiveUser().getEmail();
    
    // Verify manager authority
    const userRole = getUserRole();
    if (userRole !== 'manager' && userRole !== 'admin') {
      throw new Error('Unauthorized to make leave decisions');
    }
    
    // Find the request with validation
    let rowIndex = -1;
    let employeeEmail = '';
    let currentStatus = '';
    
    for (let i = 1; i < requestsData.length; i++) {
      if (requestsData[i][0] === leaveId) {
        rowIndex = i + 1; // +1 because array is 0-based but sheet is 1-based
        employeeEmail = requestsData[i][1];
        currentStatus = requestsData[i][9];
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('Leave request not found');
    }
    
    // Validate request can be processed
    if (currentStatus !== STATUS.PENDING && currentStatus !== STATUS.IN_REVIEW) {
      throw new Error(`Cannot process request in ${currentStatus} status`);
    }
    
    // Verify manager is authorized for this employee
    if (userRole === 'manager') {
      const managerEmail = requestsData[rowIndex - 1][3];
      if (managerEmail !== userEmail) {
        throw new Error('Not authorized to process this request');
      }
    }
    
    // Update request status with validation
    try {
      requestsSheet.getRange(rowIndex, 10).setValue(decision); // Status
      requestsSheet.getRange(rowIndex, 11).setValue(comment || ''); // Comment
      requestsSheet.getRange(rowIndex, 12).setValue(new Date()); // Decision timestamp
      
      // Verify update was successful
      const updatedStatus = requestsSheet.getRange(rowIndex, 10).getValue();
      if (updatedStatus !== decision) {
        throw new Error('Failed to update request status');
      }
    } catch (error) {
      handleError(error, 'update_request');
      throw new Error('Failed to update leave request');
    }
    
    // Log the action with details
    logAction(userEmail, 'process_decision', 
      `Leave request ${decision} by ${userRole}. Comment: ${comment || 'None'}`, 
      leaveId);
    
    // Send notification with error handling
    try {
      notifyLeaveRequestUpdate(leaveId, decision);
    } catch (error) {
      handleError(error, 'decision_notification');
      // Non-critical error, continue execution
      logAction('system', 'warning', 
        `Failed to send notification for decision on ${leaveId}`);
    }
    
    return { 
      success: true, 
      message: `Leave request ${decision.toLowerCase()} successfully` 
    };
    
  } catch (error) {
    handleError(error, 'processLeaveDecision');
    return { 
      success: false, 
      message: error.message 
    };
  }
}

// Helper Functions
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
