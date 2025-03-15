# Technical Context

## Technology Stack

### Primary Platform
- Google Sheets (Core platform)
- Google Apps Script (Automation and logic)
- Google Drive (File storage and permissions)

### Key Technologies

#### Google Sheets Features
- Protected ranges
- Data validation
- Custom formulas
- Named ranges
- Cross-sheet references
- Conditional formatting
- Custom menus
- Sidebar interfaces

#### Google Apps Script
- Custom functions
- Triggers
- UI services
- Spreadsheet service
- Gmail service
- Properties service

## Development Setup

### Sheet Structure
1. **Master Sheet**
   ```
   - Sheets:
     └─ LeaveRequests (Main data store)
     └─ UserRoles (Role mapping)
     └─ LeaveTypes (Leave configuration)
     └─ AuditLog (System events)
     └─ Settings (Configuration)
   ```

2. **Employee Interface**
   ```
   - Sheets:
     └─ RequestForm (Input interface)
     └─ History (Personal requests)
     └─ Instructions
   ```

3. **Manager Interface**
   ```
   - Sheets:
     └─ PendingRequests
     └─ TeamHistory
     └─ DecisionLog
   ```

4. **Admin Interface**
   ```
   - Sheets:
     └─ Dashboard
     └─ Reports
     └─ Settings
   ```

### Data Schema

#### LeaveRequests (Master)
```
- leave_id (Unique identifier)
- employee_email
- employee_name
- manager_email
- leave_type
- start_date
- end_date
- reason
- submission_timestamp
- status (Pending/In Review/Approved/Rejected)
- manager_comment
- decision_timestamp
```

#### UserRoles
```
- email
- role
- manager_email (if employee)
- department
- allowed_leave_types (comma-separated list)
```

#### LeaveTypes
```
- type_id
- type_name
- description
- max_days_per_year
- requires_approval
- applicable_roles (comma-separated list)
```

#### AuditLog
```
- timestamp
- action
- user_email
- details
- leave_id (if applicable)
```

### Test Data Structure
```
Sample Users:
- admin@company.com (Admin)
- manager1@company.com (Manager)
- manager2@company.com (Manager)
- emp1@company.com (Employee)
- emp2@company.com (Employee)
- emp3@company.com (Employee)

Leave Types:
1. Sick Leave
   - All employees
   - Max 12 days/year
2. Vacation
   - All employees
   - Max 20 days/year
3. Personal Leave
   - Based on role/department
   - Max 5 days/year
4. Training Leave
   - Role-specific
   - Max 10 days/year
```

## Technical Constraints

### Google Sheets Limitations
1. **Data Volume**
   - Max 5 million cells
   - Max 18,278 columns
   - Max 200 sheets per workbook

2. **Formula Limitations**
   - Max 400,000 cells with formulas
   - IMPORTRANGE quota limits
   - Query complexity limits

3. **Apps Script Limitations**
   - 6-minute execution time limit
   - Daily quota for email sending
   - Trigger limitations

### Security Considerations
1. **Sheet Protection**
   - Cannot hide data completely
   - Protection is edit-only
   - Limited cell-level security

2. **Data Access**
   - Cross-sheet data exposure risks
   - Limited row-level security
   - Shared file permissions

## Dependencies

### Required Services
1. **Google Workspace**
   - Google Sheets
   - Google Drive
   - Gmail (for notifications)

2. **User Requirements**
   - Google account
   - Appropriate permissions
   - Web browser access

### Integration Points
1. **Cross-Sheet Communication**
   - IMPORTRANGE connections
   - Apps Script data transfer
   - Protected formula chains

2. **Email System**
   - Notification templates for:
     * Leave request submission
     * Status changes (In Review)
     * Approval/Rejection
   - Trigger-based sending
   - HTML email formatting

## Performance Considerations

### Optimization Strategies
1. **Formula Efficiency**
   - Minimize volatile functions
   - Use array formulas judiciously
   - Implement efficient QUERY filters

2. **Script Performance**
   - Batch operations
   - Minimize API calls
   - Cache frequently used data

3. **Data Management**
   - Regular cleanup of old data
   - Archive historical records
   - Optimize sheet structure

### Scalability Factors
1. **Data Growth**
   - Monitor cell count
   - Archive old requests
   - Maintain formula efficiency

2. **User Load**
   - Handle concurrent access
   - Manage trigger conflicts
   - Balance notification load
