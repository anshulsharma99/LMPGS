// Setup all interfaces and handle sharing permissions with enhanced validation

function setupAllInterfaces() {
  try {
    setSystemState(SYSTEM_STATE.INITIALIZING);
    const ss = SpreadsheetApp.getActive();
    const properties = PropertiesService.getDocumentProperties();
    
    // Verify system requirements
    if (!verifySystemRequirements()) {
      throw new Error('System requirements not met');
    }
    
    // Rename current sheet to Admin interface with validation
    try {
      ss.setName('Leave Management - Admin');
    } catch (error) {
      handleError(error, 'setName');
      throw new Error('Failed to set admin interface name');
    }
    
    // Create interfaces with validation
    let managerUrl, employeeUrl;
    
    try {
      managerUrl = setupManagerInterface();
      if (!managerUrl) throw new Error('Manager interface setup failed');
      
      employeeUrl = setupEmployeeInterface();
      if (!employeeUrl) throw new Error('Employee interface setup failed');
    } catch (error) {
      handleError(error, 'interface_setup');
      throw new Error('Failed to setup interfaces: ' + error.message);
    }
    
    // Store URLs in document properties with validation
    try {
      properties.setProperties({
        'managerInterfaceUrl': managerUrl,
        'employeeInterfaceUrl': employeeUrl
      });
      
      // Verify properties were set
      const savedManagerUrl = properties.getProperty('managerInterfaceUrl');
      const savedEmployeeUrl = properties.getProperty('employeeInterfaceUrl');
      
      if (savedManagerUrl !== managerUrl || savedEmployeeUrl !== employeeUrl) {
        throw new Error('Property verification failed');
      }
    } catch (error) {
      handleError(error, 'property_setup');
      throw new Error('Failed to save interface URLs');
    }
  
    // Get and validate user roles data
    const userRolesSheet = ss.getSheetByName('UserRoles');
    if (!userRolesSheet) {
      throw new Error('UserRoles sheet not found');
    }
    
    const userData = userRolesSheet.getDataRange().getValues();
    if (userData.length < 2) {
      throw new Error('No user data found');
    }
    
    // Share access with validation
    const sharingResults = [];
    for (let i = 1; i < userData.length; i++) {
      const email = userData[i][0];
      const role = userData[i][1];
      const managerEmail = userData[i][2];
      
      if (!email || !role) {
        logAction('system', 'sharing_warning', 
          `Invalid user data found at row ${i + 1}`);
        continue;
      }
      
      try {
        const result = shareNewUserAccess(email, role, managerEmail);
        sharingResults.push({
          email: email,
          success: result
        });
      } catch (error) {
        handleError(error, `share_access_${email}`);
        sharingResults.push({
          email: email,
          success: false,
          error: error.message
        });
      }
    }
    
    // Log sharing results
    logAction('system', 'sharing_summary', 
      JSON.stringify(sharingResults));
  
    // Create menu items with validation
    try {
      createAdminMenu(ss);
      createManagerMenu(SpreadsheetApp.openByUrl(managerUrl));
      createEmployeeMenu(SpreadsheetApp.openByUrl(employeeUrl));
    } catch (error) {
      handleError(error, 'menu_creation');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Menu creation failed but continuing setup');
    }
    
    // Verify final setup
    if (!verifyInterfaceSetup(ss, managerUrl, employeeUrl)) {
      throw new Error('Interface verification failed');
    }
    
    // Set system state to ready
    setSystemState(SYSTEM_STATE.READY);
    
    // Return URLs for all interfaces
    return {
      adminUrl: ss.getUrl(),
      managerUrl: managerUrl,
      employeeUrl: employeeUrl
    };
    
  } catch (error) {
    setSystemState(SYSTEM_STATE.ERROR);
    handleError(error, 'setupAllInterfaces');
    throw error;
  }
}

// Verify system requirements before setup
function verifySystemRequirements() {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Check required sheets exist
    const requiredSheets = Object.values(SHEET_NAMES);
    for (const sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        logAction('system', 'verification_error', 
          `Required sheet ${sheetName} not found`);
        return false;
      }
    }
    
    // Check user has necessary permissions
    try {
      ss.addEditor(Session.getEffectiveUser());
      ss.removeEditor(Session.getEffectiveUser());
    } catch (error) {
      logAction('system', 'verification_error', 
        'User lacks required permissions');
      return false;
    }
    
    return true;
  } catch (error) {
    handleError(error, 'verifySystemRequirements');
    return false;
  }
}

// Verify interface setup completion
function verifyInterfaceSetup(adminSheet, managerUrl, employeeUrl) {
  try {
    // Verify admin interface
    if (!adminSheet || !adminSheet.getName().includes('Leave Management - Admin')) {
      return false;
    }
    
    // Verify manager interface
    const managerSheet = SpreadsheetApp.openByUrl(managerUrl);
    if (!managerSheet || !managerSheet.getName().includes('Leave Management - Manager')) {
      return false;
    }
    
    // Verify employee interface
    const employeeSheet = SpreadsheetApp.openByUrl(employeeUrl);
    if (!employeeSheet || !employeeSheet.getName().includes('Leave Management - Employee')) {
      return false;
    }
    
    return true;
  } catch (error) {
    handleError(error, 'verifyInterfaceSetup');
    return false;
  }
}

function createAdminMenu(ss) {
  ss.addMenu('Admin Actions', [
    {name: 'Initialize System', functionName: 'initializeSystem'},
    {name: 'Manage Users', functionName: 'showUserManagement'},
    {name: 'System Settings', functionName: 'showAdminSettings'},
    {name: 'View All Requests', functionName: 'showAllRequests'},
    {name: 'Share Access', functionName: 'shareInterfaceUrls'},
    {name: 'Archive Old Requests', functionName: 'archiveOldRequests'}
  ]);
}

function createManagerMenu(ss) {
  ss.addMenu('Manager Actions', [
    {name: 'View Pending Requests', functionName: 'showPendingRequests'},
    {name: 'View Team History', functionName: 'showTeamHistory'},
    {name: 'Refresh Calendar', functionName: 'refreshTeamCalendar'}
  ]);
}

function createEmployeeMenu(ss) {
  ss.addMenu('Leave Actions', [
    {name: 'Submit Leave Request', functionName: 'showLeaveRequestForm'},
    {name: 'View My Requests', functionName: 'showMyRequests'},
    {name: 'View Leave Balance', functionName: 'showLeaveBalance'}
  ]);
}

function refreshTeamCalendar() {
  const ss = SpreadsheetApp.getActive();
  const calendarSheet = ss.getSheetByName('Team Calendar');
  if (calendarSheet) {
    const today = new Date();
    createMonthCalendar(calendarSheet, today.getFullYear(), today.getMonth());
  }
}

function showLeaveBalance() {
  const ss = SpreadsheetApp.getActive();
  const balanceSheet = ss.getSheetByName('Leave Balance');
  if (balanceSheet) {
    balanceSheet.activate();
  }
}

// Function to validate interface URLs with graceful fallback
function validateInterfaceUrls(requireAccess = true) {
  const ss = SpreadsheetApp.getActive();
  const properties = PropertiesService.getDocumentProperties();
  
  try {
    const interfaces = {
      adminUrl: ss.getUrl(),
      managerUrl: properties.getProperty('managerInterfaceUrl'),
      employeeUrl: properties.getProperty('employeeInterfaceUrl')
    };
    
    // If we don't need to verify access, return the URLs as is
    if (!requireAccess) {
      return interfaces;
    }
    
    // Try to access existing interfaces first
    let urlsValid = true;
    if (interfaces.managerUrl && interfaces.employeeUrl) {
      try {
        SpreadsheetApp.openByUrl(interfaces.managerUrl);
        SpreadsheetApp.openByUrl(interfaces.employeeUrl);
      } catch (error) {
        urlsValid = false;
        logAction('system', 'url_access_error', 
          `Failed to access existing interfaces: ${error.toString()}`);
      }
    } else {
      urlsValid = false;
    }
    
    // Only try to recreate if URLs are invalid and recreation is required
    if (!urlsValid && requireAccess) {
      logAction('system', 'url_validation', 
        'Attempting to recreate invalid interfaces');
      
      try {
        const setupResult = setupAllInterfaces();
        if (setupResult && setupResult.managerUrl && setupResult.employeeUrl) {
          interfaces.managerUrl = setupResult.managerUrl;
          interfaces.employeeUrl = setupResult.employeeUrl;
        } else {
          throw new Error('Interface recreation failed');
        }
      } catch (error) {
        logAction('system', 'url_recreation_error', 
          `Failed to recreate interfaces: ${error.toString()}`);
        // Return existing URLs even if invalid
        return interfaces;
      }
    }
    
    return interfaces;
  } catch (error) {
    logAction('system', 'url_validation_error', error.toString());
    // Return basic interface object even on error
    return {
      adminUrl: ss.getUrl(),
      managerUrl: properties.getProperty('managerInterfaceUrl'),
      employeeUrl: properties.getProperty('employeeInterfaceUrl')
    };
  }
}

// Function to share interface URLs with users
function shareInterfaceUrls() {
  const ss = SpreadsheetApp.getActive();
  const userRolesSheet = ss.getSheetByName('UserRoles');
  const userData = userRolesSheet.getDataRange().getValues();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Validate interface URLs first
    validateInterfaceUrls();
    
    // Skip header row
    for (let i = 1; i < userData.length; i++) {
      const email = userData[i][0];
      const role = userData[i][1];
      const managerEmail = userData[i][2];
      
      try {
        // Share access with user
        if (shareNewUserAccess(email, role, managerEmail)) {
          logAction('system', 'share_access', `Access shared with ${email} as ${role}`);
        }
      } catch (error) {
        logAction('system', 'share_error', `Failed to share access with ${email}: ${error.toString()}`);
        // Continue with next user instead of throwing
        continue;
      }
    }
    
    ui.alert(
      'Success',
      'Access has been shared with users. Check the audit log for any sharing errors.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Error',
      'Failed to share access: ' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

// Function to add a new user
function addNewUser(email, role, managerEmail = '', department = '', allowedLeaveTypes = '') {
  const ss = SpreadsheetApp.getActive();
  const userRolesSheet = ss.getSheetByName('UserRoles');
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Check if user already exists
    const userData = userRolesSheet.getDataRange().getValues();
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === email) {
        throw new Error('User already exists');
      }
    }
    
    // Validate interface URLs before proceeding
    validateInterfaceUrls();
    
    // Add user to UserRoles sheet first
    userRolesSheet.appendRow([
      email,
      role,
      managerEmail,
      department,
      allowedLeaveTypes
    ]);
    
    // Try to share access, but don't fail if sharing fails
    try {
      shareNewUserAccess(email, role, managerEmail);
      logAction('system', 'add_user', `New user added: ${email} as ${role}`);
    } catch (error) {
      logAction('system', 'share_error', `Added user but failed to share access with ${email}: ${error.toString()}`);
      
      ui.alert(
        'Partial Success',
        `User ${email} has been added but sharing access failed. Please try sharing access manually.`,
        ui.ButtonSet.OK
      );
      
      return true; // Still return true as user was added
    }
    
    ui.alert(
      'Success',
      `User ${email} has been added successfully and will receive an email with access details.`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    ui.alert(
      'Error',
      'Failed to add user: ' + error.toString(),
      ui.ButtonSet.OK
    );
    return false;
  }
}

// Function to remove a user with improved error handling
function removeUser(email) {
  const ss = SpreadsheetApp.getActive();
  const userRolesSheet = ss.getSheetByName('UserRoles');
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Find user row first
    const userData = userRolesSheet.getDataRange().getValues();
    let userRow = -1;
    let userRole = '';
    
    for (let i = 1; i < userData.length; i++) {
      if (userData[i][0] === email) {
        userRow = i + 1; // +1 because array is 0-based but sheet is 1-based
        userRole = userData[i][1];
        break;
      }
    }
    
    if (userRow === -1) {
      throw new Error('User not found');
    }
    
    // Get interface URLs without requiring access validation
    const interfaces = validateInterfaceUrls(false);
    let accessRemovalSuccess = true;
    let accessRemovalError = '';
    
    // Try to remove user's access
    try {
      switch(userRole.toLowerCase()) {
        case 'admin':
          try {
            ss.removeEditor(email);
          } catch (e) {
            accessRemovalSuccess = false;
            accessRemovalError = `Admin access removal failed: ${e.message}`;
          }
          break;
          
        case 'manager':
          try {
            if (interfaces.managerUrl) {
              SpreadsheetApp.openByUrl(interfaces.managerUrl).removeViewer(email);
            } else {
              accessRemovalSuccess = false;
              accessRemovalError = 'Manager interface URL not found';
            }
          } catch (e) {
            accessRemovalSuccess = false;
            accessRemovalError = `Manager access removal failed: ${e.message}`;
          }
          break;
          
        case 'employee':
          try {
            if (interfaces.employeeUrl) {
              SpreadsheetApp.openByUrl(interfaces.employeeUrl).removeViewer(email);
            } else {
              accessRemovalSuccess = false;
              accessRemovalError = 'Employee interface URL not found';
            }
          } catch (e) {
            accessRemovalSuccess = false;
            accessRemovalError = `Employee access removal failed: ${e.message}`;
          }
          break;
      }
    } catch (error) {
      accessRemovalSuccess = false;
      accessRemovalError = error.toString();
    }
    
    // Always delete the user row, regardless of access removal success
    userRolesSheet.deleteRow(userRow);
    
    // Log the removal
    logAction('system', 'remove_user', 
      `User removed: ${email}${!accessRemovalSuccess ? ` (with access removal issues: ${accessRemovalError})` : ''}`);
    
    // Show appropriate message based on success
    if (accessRemovalSuccess) {
      ui.alert(
        'Success',
        `User ${email} has been completely removed from the system.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Partial Success',
        `User ${email} has been removed from the system, but there were issues removing their access: ${accessRemovalError}. ` +
        'You may need to remove their access manually.',
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    ui.alert(
      'Error',
      'Failed to remove user: ' + error.toString(),
      ui.ButtonSet.OK
    );
    return false;
  }
}
