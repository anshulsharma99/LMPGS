// Employee Interface Sheet Setup with Enhanced Error Handling

function setupEmployeeInterface() {
  try {
    const ss = SpreadsheetApp.getActive();
    const employeeSheet = SpreadsheetApp.create('Leave Management - Employee Interface');
    
    // Verify sheet creation
    if (!employeeSheet) {
      throw new Error('Failed to create employee interface sheet');
    }
    
    // Create necessary sheets with validation
    try {
      // Remove extra sheets
      while (employeeSheet.getSheets().length > 1) {
        employeeSheet.deleteSheet(employeeSheet.getSheets()[1]);
      }
      
      // Setup main sheets
      const myRequests = employeeSheet.getSheetByName('Sheet1');
      if (!myRequests) throw new Error('Default sheet not found');
      myRequests.setName('My Requests');
      
      const requestForm = employeeSheet.insertSheet('Request Form');
      const leaveBalance = employeeSheet.insertSheet('Leave Balance');
      const importAuth = employeeSheet.insertSheet('ImportAuth').hideSheet();
      
      if (!requestForm || !leaveBalance || !importAuth) {
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
        handleError(error, 'employee_importAuth_setup');
        throw new Error('Failed to setup ImportRange authorization');
      }
  
      // Setup sheets with validation
      if (!setupMyRequestsSheet(myRequests, adminSheetId)) {
        throw new Error('Failed to setup My Requests sheet');
      }
      
      if (!setupRequestFormSheet(requestForm)) {
        throw new Error('Failed to setup Request Form sheet');
      }
      
      if (!setupLeaveBalanceSheet(leaveBalance, adminSheetId)) {
        throw new Error('Failed to setup Leave Balance sheet');
      }
      
      // Protect sheets with validation
      if (!protectEmployeeSheets(employeeSheet)) {
        logAction('system', 'warning', 'Employee sheet protection may not be fully applied');
      }
      
      // Verify final setup
      if (!verifyEmployeeInterface(employeeSheet)) {
        throw new Error('Employee interface verification failed');
      }
      
      return employeeSheet.getUrl();
      
    } catch (error) {
      handleError(error, 'employee_sheet_setup');
      // Try to clean up on failure
      try {
        DriveApp.getFileById(employeeSheet.getId()).setTrashed(true);
      } catch (e) {
        // Ignore cleanup errors
      }
      throw error;
    }
    
  } catch (error) {
    handleError(error, 'setupEmployeeInterface');
    return null;
  }
}

// Verify employee interface setup
function verifyEmployeeInterface(sheet) {
  try {
    // Check sheet name
    if (!sheet.getName().includes('Leave Management - Employee Interface')) {
      return false;
    }
    
    // Check required sheets exist
    const requiredSheets = ['My Requests', 'Request Form', 'Leave Balance', 'ImportAuth'];
    for (const sheetName of requiredSheets) {
      if (!sheet.getSheetByName(sheetName)) {
        return false;
      }
    }
    
    // Check formulas are set
    const myRequests = sheet.getSheetByName('My Requests');
    const leaveBalance = sheet.getSheetByName('Leave Balance');
    
    if (!myRequests.getRange('A2').getFormula() || 
        !leaveBalance.getRange('A2').getFormula()) {
      return false;
    }
    
    return true;
  } catch (error) {
    handleError(error, 'verifyEmployeeInterface');
    return false;
  }
}

function setupMyRequestsSheet(sheet, adminSheetId) {
  try {
    // Set headers with validation
    const headers = [
      'Leave ID', 'Type', 'Start Date', 'End Date',
      'Reason', 'Status', 'Submission Date', 'Decision Date', 'Manager Comments'
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
      handleError(error, 'my_requests_headers');
      throw new Error('Failed to set My Requests headers');
    }
    
    // Set column widths with validation
    try {
      sheet.setColumnWidths(1, headers.length, 120);
    } catch (error) {
      handleError(error, 'my_requests_column_width');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'My Requests column widths not set properly');
    }
    
    // Add IMPORTRANGE formula with validation
    try {
      const formula = `=QUERY(IMPORTRANGE("${adminSheetId}", "LeaveRequests!A2:L"), 
        "select Col1, Col5, Col6, Col7, Col8, Col10, Col9, Col12, Col11 
         where Col2 = '"&INDIRECT("UserRoles!A1")&"'
         order by Col9 desc")`;
      
      sheet.getRange('A2').setFormula(formula);
      
      // Verify formula was set
      const setFormula = sheet.getRange('A2').getFormula();
      if (!setFormula.includes(adminSheetId)) {
        throw new Error('My Requests import formula not set correctly');
      }
    } catch (error) {
      handleError(error, 'my_requests_formula');
      throw new Error('Failed to set My Requests import formula');
    }
  
    // Add conditional formatting with validation
    try {
      const range = sheet.getRange('F2:F1000');
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
      handleError(error, 'my_requests_formatting');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'My Requests conditional formatting not fully applied');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupMyRequestsSheet');
    return false;
  }
}

function setupRequestFormSheet(sheet) {
  try {
    // Set form layout with validation
    try {
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 300);
    } catch (error) {
      handleError(error, 'request_form_layout');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Request form layout not set properly');
    }
    
    // Add form title with validation
    try {
      const titleRange = sheet.getRange('A1:B1');
      titleRange.merge()
        .setValue('Leave Request Form')
        .setFontSize(14)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#f3f3f3');
      
      // Verify title was set
      if (titleRange.getValue() !== 'Leave Request Form') {
        throw new Error('Form title not set correctly');
      }
    } catch (error) {
      handleError(error, 'request_form_title');
      throw new Error('Failed to set form title');
    }
    
    // Add form fields with validation
    try {
      const formFields = [
        ['Leave Type:', '=INDIRECT("LeaveTypes!B2:B5")'],
        ['Start Date:', ''],
        ['End Date:', ''],
        ['Reason:', '']
      ];
      
      const fieldsRange = sheet.getRange('A3:B6');
      fieldsRange.setValues(formFields);
      
      // Verify fields were set
      const setFields = fieldsRange.getValues();
      if (!formFields.every((row, i) => row[0] === setFields[i][0])) {
        throw new Error('Form fields not set correctly');
      }
    } catch (error) {
      handleError(error, 'request_form_fields');
      throw new Error('Failed to set form fields');
    }
    
    // Add data validation with error handling
    try {
      // Leave type validation
      const typeRange = sheet.getRange('B3');
      const typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(typeRange.getValues())
        .build();
      typeRange.setDataValidation(typeRule);
      
      // Date validation
      const dateRange = sheet.getRange('B4:B5');
      const dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .build();
      dateRange.setDataValidation(dateRule);
      
      // Verify validations were set
      if (!typeRange.getDataValidation() || !dateRange.getDataValidation()) {
        throw new Error('Form validations not set correctly');
      }
    } catch (error) {
      handleError(error, 'request_form_validation');
      throw new Error('Failed to set form validations');
    }
    
    // Add submit button with validation
    try {
      const buttonRange = sheet.getRange('B8');
      buttonRange
        .setValue('Submit Request')
        .setBackground('#4CAF50')
        .setFontColor('white')
        .setHorizontalAlignment('center');
      
      // Verify button was set
      if (buttonRange.getValue() !== 'Submit Request') {
        throw new Error('Submit button not set correctly');
      }
    } catch (error) {
      handleError(error, 'request_form_button');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Submit button not fully styled');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupRequestFormSheet');
    return false;
  }
}

function setupLeaveBalanceSheet(sheet, adminSheetId) {
  try {
    // Set headers with validation
    const headers = ['Leave Type', 'Total Days', 'Used Days', 'Remaining Days'];
    
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
      handleError(error, 'leave_balance_headers');
      throw new Error('Failed to set leave balance headers');
    }
    
    // Set column widths with validation
    try {
      sheet.setColumnWidths(1, headers.length, 120);
    } catch (error) {
      handleError(error, 'leave_balance_column_width');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Leave balance column widths not set properly');
    }
    
    // Add IMPORTRANGE formula with validation
    try {
      const leaveTypesFormula = `=QUERY(IMPORTRANGE("${adminSheetId}", "LeaveTypes!A2:D"), 
        "select Col2, Col4 
         where Col6 contains '"&INDIRECT("UserRoles!B1")&"'")`;
      
      sheet.getRange('A2').setFormula(leaveTypesFormula);
      
      // Verify formula was set
      const setFormula = sheet.getRange('A2').getFormula();
      if (!setFormula.includes(adminSheetId)) {
        throw new Error('Leave types import formula not set correctly');
      }
    } catch (error) {
      handleError(error, 'leave_types_formula');
      throw new Error('Failed to set leave types formula');
    }
    
    // Add calculation formulas with validation
    try {
      // Used days formula
      const usedDaysFormula = `=SUMIFS(INDIRECT("'My Requests'!D2:D1000") - INDIRECT("'My Requests'!C2:C1000") + 1, 
        INDIRECT("'My Requests'!B2:B1000"), A2, 
        INDIRECT("'My Requests'!F2:F1000"), "Approved")`;
      
      sheet.getRange('C2').setFormula(usedDaysFormula);
      
      // Remaining days formula
      sheet.getRange('D2').setFormula('=B2-C2');
      
      // Verify formulas were set
      if (!sheet.getRange('C2').getFormula() || !sheet.getRange('D2').getFormula()) {
        throw new Error('Calculation formulas not set correctly');
      }
    } catch (error) {
      handleError(error, 'calculation_formulas');
      throw new Error('Failed to set calculation formulas');
    }
    
    // Add conditional formatting with validation
    try {
      const range = sheet.getRange('D2:D1000');
      const rules = [
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(5)
          .setBackground('#f8d7da')
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberBetween(5, 10)
          .setBackground('#fff3cd')
          .setRanges([range])
          .build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(10)
          .setBackground('#d4edda')
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
      handleError(error, 'leave_balance_formatting');
      // Non-critical error, continue execution
      logAction('system', 'warning', 'Leave balance conditional formatting not fully applied');
    }
    
    return true;
  } catch (error) {
    handleError(error, 'setupLeaveBalanceSheet');
    return false;
  }
}

function protectEmployeeSheets(ss) {
  try {
    const sheets = ss.getSheets();
    const protectionResults = [];
    
    sheets.forEach(sheet => {
      try {
        if (sheet.getName() !== 'ImportAuth') {
          const protection = sheet.protect();
          protection.setDescription('Protected sheet - use forms for interaction');
          
          // Configure protection based on sheet type
          if (sheet.getName() === 'Request Form') {
            const range = sheet.getRange('B3:B6');
            protection.setUnprotectedRanges([range]);
            
            // Verify unprotected range was set
            const unprotectedRanges = protection.getUnprotectedRanges();
            if (!unprotectedRanges || unprotectedRanges.length === 0) {
              throw new Error('Unprotected ranges not set correctly');
            }
          } else {
            protection.setWarningOnly(true);
            
            // Verify warning only was set
            if (!protection.isWarningOnly()) {
              throw new Error('Warning only protection not set correctly');
            }
          }
          
          // Verify protection description
          if (!protection.getDescription()) {
            throw new Error('Protection description not set');
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
    logAction('system', 'employee_sheet_protection', 
      `Protection results: ${JSON.stringify(protectionResults)}`);
    
    // Return true if all protections were successful
    return protectionResults.every(result => result.success);
  } catch (error) {
    handleError(error, 'protectEmployeeSheets');
    return false;
  }
}
