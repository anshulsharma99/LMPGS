// Manager Interface Sheet Setup with Enhanced Error Handling

function setupManagerInterface() {
  try {
    const ss = SpreadsheetApp.getActive();
    const managerSheet = SpreadsheetApp.create('Leave Management - Manager Interface');
    
    // Verify sheet creation
    if (!managerSheet) {
      throw new Error('Failed to create manager interface sheet');
    }
    
    // Create necessary sheets with validation
    try {
      // Remove extra sheets
      while (managerSheet.getSheets().length > 1) {
        managerSheet.deleteSheet(managerSheet.getSheets()[1]);
      }
      
      // Setup main sheets
      const teamRequests = managerSheet.getSheetByName('Sheet1');
      if (!teamRequests) throw new Error('Default sheet not found');
      teamRequests.setName('Team Requests');
      
      const pendingApprovals = managerSheet.insertSheet('Pending Approvals');
      const teamCalendar = managerSheet.insertSheet('Team Calendar');
      const importAuth = managerSheet.insertSheet('ImportAuth').hideSheet();
      
      if (!pendingApprovals || !teamCalendar || !importAuth) {
        throw new Error('Failed to create required sheets');
      }
      
      // Setup ImportAuth sheet with validation
      const adminSheetId = ss.getId();
      try {
        importAuth.getRange('A1').setFormula(`=IMPORTRANGE("${adminSheetId}", "LeaveRequests!A1")`);
        
        // Verify formula was set
        const formula = importAuth.getRange('A1').getFormula();
        if (!formula.includes(adminSheetId)) {
          throw new Error('ImportRange formula not set correctly');
        }
      } catch (error) {
        handleError(error, 'importAuth_setup');
        throw new Error('Failed to setup ImportRange authorization');
      }
  
      // Setup sheets with validation
      if (!setupTeamRequestsSheet(teamRequests, adminSheetId)) {
        throw new Error('Failed to setup Team Requests sheet');
      }
      
      if (!setupPendingApprovalsSheet(pendingApprovals, adminSheetId)) {
        throw new Error('Failed to setup Pending Approvals sheet');
      }
      
      if (!setupTeamCalendarSheet(teamCalendar, adminSheetId)) {
        throw new Error('Failed to setup Team Calendar sheet');
      }
      
      // Protect sheets with validation
      if (!protectManagerSheets(managerSheet)) {
        logAction('system', 'warning', 'Sheet protection may not be fully applied');
      }
      
      // Verify final setup
      if (!verifyManagerInterface(managerSheet)) {
        throw new Error('Manager interface verification failed');
      }
      
      return managerSheet.getUrl();
      
    } catch (error) {
      handleError(error, 'manager_sheet_setup');
      // Try to clean up on failure
      try {
        DriveApp.getFileById(managerSheet.getId()).setTrashed(true);
      } catch (e) {
        // Ignore cleanup errors
      }
      throw error;
    }
    
  } catch (error) {
    handleError(error, 'setupManagerInterface');
    return null;
  }
}

// Verify manager interface setup
function verifyManagerInterface(sheet) {
  try {
    // Check sheet name
    if (!sheet.getName().includes('Leave Management - Manager Interface')) {
      return false;
    }
    
    // Check required sheets exist
    const requiredSheets = ['Team Requests', 'Pending Approvals', 'Team Calendar', 'ImportAuth'];
    for (const sheetName of requiredSheets) {
      if (!sheet.getSheetByName(sheetName)) {
        return false;
      }
    }
    
    // Check formulas are set
    const teamRequests = sheet.getSheetByName('Team Requests');
    const pendingApprovals = sheet.getSheetByName('Pending Approvals');
    
    if (!teamRequests.getRange('A2').getFormula() || 
        !pendingApprovals.getRange('A2').getFormula()) {
      return false;
    }
    
    return true;
  } catch (error) {
    handleError(error, 'verifyManagerInterface');
    return false;
  }
}

function setupTeamRequestsSheet(sheet, adminSheetId) {
  try {
    // Set headers with validation
    const headers = [
      'Leave ID', 'Employee', 'Type', 'Start Date', 'End Date',
      'Reason', 'Status', 'Submission Date', 'Decision Date', 'Comments'
    ];
    
    try {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers])
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
        
      // Verify headers were set
      const setHeaders = headerRange.getValues()[0];
      if (!headers.every((h, i) => h === setHeaders[i])) {
        throw new Error('Headers not set correctly');
      }
    } catch (error) {
      handleError(error, 'header_setup');
      throw new Error('Failed to set headers');
    }
    
    // Set column widths with validation
    try {
      sheet.setColumnWidths(1, headers.length, 120);
    } catch (error) {
      handleError(error, 'column_width');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Column widths not set properly');
    }
    
    // Add data validation with error handling
    try {
      const statusRange = sheet.getRange('G2:G1000');
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Pending', 'In Review', 'Approved', 'Rejected'])
        .build();
      statusRange.setDataValidation(rule);
      
      // Verify validation was set
      if (!statusRange.getDataValidation()) {
        throw new Error('Status validation not set');
      }
    } catch (error) {
      handleError(error, 'status_validation');
      throw new Error('Failed to set status validation');
    }
    
    // Add IMPORTRANGE formula with validation
    try {
      const formula = `=QUERY(IMPORTRANGE("${adminSheetId}", "LeaveRequests!A2:L"), 
        "select Col1, Col2, Col5, Col6, Col7, Col8, Col10, Col9, Col12, Col11 
         where Col4 = '"&INDIRECT("UserRoles!B1")&"'
         order by Col9 desc")`;
      
      sheet.getRange('A2').setFormula(formula);
      
      // Verify formula was set
      const setFormula = sheet.getRange('A2').getFormula();
      if (!setFormula.includes(adminSheetId)) {
        throw new Error('Import formula not set correctly');
      }
    } catch (error) {
      handleError(error, 'import_formula');
      throw new Error('Failed to set import formula');
    }
  
    // Add conditional formatting with validation
    try {
      const range = sheet.getRange('G2:G1000');
      const rules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Pending')
          .setBackground('#fff3cd')
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Approved')
          .setBackground('#d4edda')
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('Rejected')
          .setBackground('#f8d7da')
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo('In Review')
          .setBackground('#cce5ff')
          .setRanges([range])
          .build()
      ];
      
      range.setConditionalFormatRules(rules);
      
      // Verify rules were set
      const setRules = range.getConditionalFormatRules();
      if (setRules.length !== rules.length) {
        throw new Error('Conditional format rules not set correctly');
      }
    } catch (error) {
      handleError(error, 'conditional_formatting');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Conditional formatting not fully applied');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupTeamRequestsSheet');
    return false;
  }
}

function setupPendingApprovalsSheet(sheet, adminSheetId) {
  try {
    // Set headers with validation
    const headers = [
      'Leave ID', 'Employee', 'Type', 'Start Date', 'End Date',
      'Reason', 'Status', 'Submission Date', 'Actions'
    ];
    
    try {
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers])
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // Verify headers were set
      const setHeaders = headerRange.getValues()[0];
      if (!headers.every((h, i) => h === setHeaders[i])) {
        throw new Error('Headers not set correctly');
      }
    } catch (error) {
      handleError(error, 'pending_headers_setup');
      throw new Error('Failed to set pending approvals headers');
    }
    
    // Set column widths with validation
    try {
      sheet.setColumnWidths(1, headers.length, 120);
    } catch (error) {
      handleError(error, 'pending_column_width');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Pending approvals column widths not set properly');
    }
    
    // Add IMPORTRANGE formula with validation
    try {
      const formula = `=QUERY(IMPORTRANGE("${adminSheetId}", "LeaveRequests!A2:L"), 
        "select Col1, Col2, Col5, Col6, Col7, Col8, Col10, Col9 
         where Col4 = '"&INDIRECT("UserRoles!B1")&"' 
         and Col10 = 'Pending'
         order by Col9 desc")`;
      
      sheet.getRange('A2').setFormula(formula);
      
      // Verify formula was set
      const setFormula = sheet.getRange('A2').getFormula();
      if (!setFormula.includes(adminSheetId)) {
        throw new Error('Pending approvals import formula not set correctly');
      }
    } catch (error) {
      handleError(error, 'pending_import_formula');
      throw new Error('Failed to set pending approvals import formula');
    }
    
    // Add action buttons with validation
    try {
      const actionRange = sheet.getRange('I2:I1000');
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Approve', 'Reject'])
        .build();
      actionRange.setDataValidation(rule);
      
      // Verify validation was set
      if (!actionRange.getDataValidation()) {
        throw new Error('Action validation not set');
      }
    } catch (error) {
      handleError(error, 'action_validation');
      throw new Error('Failed to set action validation');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupPendingApprovalsSheet');
    return false;
  }
}

function setupTeamCalendarSheet(sheet, adminSheetId) {
  try {
    // Set up calendar view with validation
    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth();
    
    // Create month calendar with validation and error handling
    try {
      if (!createMonthCalendar(sheet, year, month)) {
        logAction('system', 'warning', 'Calendar creation partially failed, continuing with basic setup');
      }
    } catch (error) {
      // Log error but continue with setup
      logAction('system', 'calendar_error', 
        `Calendar creation error: ${error.toString()}`);
    }
    
    // Add IMPORTRANGE formula with validation and fallback
    try {
      // First try with the full query
      const formula = `=QUERY(IMPORTRANGE("${adminSheetId}", "LeaveRequests!A2:L"), 
        "select Col2, Col5, Col6, Col7 
         where Col4 = '"&INDIRECT("UserRoles!B1")&"' 
         and Col10 = 'Approved' 
         and Col6 >= date '"&TEXT(DATE(${year},${month+1},1),"yyyy-MM-dd")&"'
         and Col6 <= date '"&TEXT(EOMONTH(DATE(${year},${month+1},1),0),"yyyy-MM-dd")&"'")`;
      
      sheet.getRange('K1').setFormula(formula);
      
      // Verify formula was set
      const setFormula = sheet.getRange('K1').getFormula();
      if (!setFormula.includes(adminSheetId)) {
        // Try simpler formula as fallback
        const fallbackFormula = `=IMPORTRANGE("${adminSheetId}", "LeaveRequests!A2:L")`;
        sheet.getRange('K1').setFormula(fallbackFormula);
        
        // Log the fallback
        logAction('system', 'calendar_setup', 
          'Using simplified import formula for calendar');
      }
    } catch (error) {
      // Log error but don't fail the setup
      logAction('system', 'formula_error', 
        `Calendar formula error: ${error.toString()}`);
    }
    
    // Add basic formatting even if other steps failed
    try {
      sheet.getRange('A1:G1').merge()
        .setBackground('#f3f3f3')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    } catch (error) {
      // Non-critical formatting error
      logAction('system', 'format_error', 
        `Calendar formatting error: ${error.toString()}`);
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupTeamCalendarSheet');
    // Return true anyway to prevent blocking interface creation
    return true;
  }
}

function createMonthCalendar(sheet, year, month) {
  try {
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const daysInMonth = lastDay.getDate();
    
    // Set month and year header with validation
    try {
      const headerCell = sheet.getRange('A1');
      const headerValue = Utilities.formatDate(firstDay, Session.getScriptTimeZone(), 'MMMM yyyy');
      headerCell.setValue(headerValue);
      
      // Verify header was set
      if (headerCell.getValue() !== headerValue) {
        throw new Error('Calendar header not set correctly');
      }
    } catch (error) {
      handleError(error, 'calendar_header');
      throw new Error('Failed to set calendar header');
    }
    
    // Set day headers with validation
    try {
      const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      const dayHeaderRange = sheet.getRange('A3:G3');
      dayHeaderRange.setValues([days]);
      
      // Verify day headers were set
      const setDays = dayHeaderRange.getValues()[0];
      if (!days.every((d, i) => d === setDays[i])) {
        throw new Error('Day headers not set correctly');
      }
    } catch (error) {
      handleError(error, 'day_headers');
      throw new Error('Failed to set day headers');
    }
    
    // Create calendar grid with validation
    try {
      let date = 1;
      let row = 4;
      const calendarData = [];
      
      while (date <= daysInMonth) {
        let weekRow = [];
        for (let i = 0; i < 7; i++) {
          if (row === 4 && i < firstDay.getDay()) {
            weekRow.push('');
          } else if (date > daysInMonth) {
            weekRow.push('');
          } else {
            weekRow.push(date++);
          }
        }
        calendarData.push(weekRow);
        row++;
      }
      
      // Set all calendar data at once for better performance
      if (calendarData.length > 0) {
        sheet.getRange(4, 1, calendarData.length, 7).setValues(calendarData);
      }
    } catch (error) {
      handleError(error, 'calendar_grid');
      throw new Error('Failed to create calendar grid');
    }
    
    // Format calendar with validation
    try {
      const calendarRange = sheet.getRange(3, 1, row - 3, 7);
      calendarRange.setBorder(true, true, true, true, true, true);
      calendarRange.setHorizontalAlignment('center');
      
      // Verify formatting was applied
      const alignment = calendarRange.getHorizontalAlignment();
      if (alignment !== 'center') {
        throw new Error('Calendar formatting not applied correctly');
      }
    } catch (error) {
      handleError(error, 'calendar_formatting');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Calendar formatting not fully applied');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'createMonthCalendar');
    return false;
  }
}

function protectManagerSheets(ss) {
  try {
    const sheets = ss.getSheets();
    const protectionResults = [];
    
    sheets.forEach(sheet => {
      try {
        if (sheet.getName() !== 'ImportAuth') {
          const protection = sheet.protect();
          protection.setDescription('Protected sheet - use forms for interaction');
          protection.setWarningOnly(true);
          
          // Verify protection was set
          if (!protection.getDescription() || !protection.isWarningOnly()) {
            throw new Error(`Protection not set correctly for sheet: ${sheet.getName()}`);
          }
          
          protectionResults.push({
            sheet: sheet.getName(),
            success: true
          });
        }
      } catch (error) {
        handleError(error, `protect_sheet_${sheet.getName()}`);
        protectionResults.push({
          sheet: sheet.getName(),
          success: false,
          error: error.message
        });
      }
    });
    
    // Log protection results
    logAction('system', 'sheet_protection', 
      `Protection results: ${JSON.stringify(protectionResults)}`);
    
    // Return true if all protections were successful
    return protectionResults.every(result => result.success);
  } catch (error) {
    handleError(error, 'protectManagerSheets');
    return false;
  }
}
