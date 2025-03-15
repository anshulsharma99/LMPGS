# Google Sheets Leave Management System

## Project Overview
A comprehensive leave management system built entirely within Google Sheets, supporting three distinct user roles: Employee, Manager, and Admin. The system will handle leave requests, approvals, and tracking without requiring external servers or web hosting.

## Core Requirements

### 1. Data Structure & Master Sheet
- Master spreadsheet for storing all leave requests
- Automatic unique leave ID generation
- Comprehensive tracking of request details and status

### 2. Role-Based Interfaces
- Employee Interface for submitting and tracking requests
- Manager Interface for handling approvals and maintaining audit logs
- Admin Interface for complete system oversight

### 3. User Role Management
- Role-based access control via email mapping
- Secure role identification using Google Apps Script
- Protected data visibility based on user roles

### 4. Security & Data Integrity
- Role-specific sheet access
- Protected sensitive information
- Secure data transfer between sheets

### 5. Automation Features
- Automatic leave ID generation
- Manager assignment
- Email notifications
- Audit logging

## Success Criteria
1. Users can access only their authorized interfaces
2. Leave requests flow smoothly from submission to approval/rejection
3. Data remains secure and properly segregated by role
4. System maintains accurate audit trails
5. All operations work within Google Sheets' constraints

## Constraints
- Must operate entirely within Google Sheets
- No external servers or web hosting
- Limited by Google Sheets' built-in security features
- Must handle concurrent user access gracefully
