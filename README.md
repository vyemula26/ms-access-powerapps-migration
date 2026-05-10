# Project Management System — MS Access → PowerApps Migration

A complete end-to-end **Project Management System** built in **MS Access**, currently being migrated to **Microsoft PowerApps + Dataverse** as part of a Power Platform modernization initiative.

The system manages the full lifecycle of a project — from employee onboarding to invoice generation — with role-based access control, multi-level approval workflows, automated email notifications, attendance tracking, leave and OT management, payroll generation, and live reporting.

> This document is actively being updated as the system evolves.

---

## Table of Contents

- [System Overview](#system-overview)
- [Database Structure](#database-structure)
- [Login System & Role-Based Access](#login-system--role-based-access)
- [Role-Wise Dashboards](#role-wise-dashboards)
- [Core Forms](#core-forms)
- [End-to-End Workflow](#end-to-end-workflow)
- [Key Business Logic](#key-business-logic)
- [Technical Implementation](#technical-implementation)
- [Challenges & Solutions](#challenges--solutions)
- [Leave & OT Management](#leave--ot-management)
- [Payroll Generation](#payroll-generation)
- [Reports](#reports)
- [PowerApps Migration](#powerapps-migration)

---

## System Overview

| Item | Details |
|------|---------|
| Platform | MS Access → PowerApps + Dataverse (in progress) |
| Company | MCD1 |
| Master Tables | 5 |
| User Roles | 8 |
| Workflow Type | Multi-level email-based approval |
| Access Control | Role-based login with dedicated dashboards |
| Technical Stack | MS Access, VBA, Macros, Queries |
| Automation | Auto mail triggers, real-time badge counters, payroll generation, live reports |

---

## Database Structure

### Master Tables & Primary Keys

| Table | Primary Key | Purpose |
|-------|------------|---------|
| `tbl_EmployeeMaster` | Employee ID | Stores all employee information |
| `tbl_ProjectMaster` | Job Code | Stores all project and deliverable details |
| `tbl_InvoiceMaster` | Invoice ID | Stores invoice records per deliverable |
| `tbl_ManpowerUtilization` | ID | Stores team assignments per project |
| `tbl_Attendance` | Attendance ID | Stores daily attendance per employee |

### Table Relationships

```
tbl_EmployeeMaster (Employee ID)
        │
        ├──── tbl_Attendance (Employee ID → FK)
        │
        └──── tbl_ManpowerUtilization (Employee ID → FK)
                        │
                        └──── tbl_ProjectMaster (Job Code → FK)
                                        │
                                        └──── tbl_InvoiceMaster (Job Code → FK)
```

---

## Login System & Role-Based Access

On database launch, all toolbars, menu bars, and MS Access navigation are completely hidden. Only the Login Page is displayed. On successful authentication, the role-specific dashboard opens automatically. Each dashboard exposes only the forms and actions relevant to that role. Real-time badge counters display pending action counts on relevant buttons.

---

## Role-Wise Dashboards

### 1. Admin

| Button | Purpose |
|--------|---------|
| Employee Master Form | Add, edit, and manage all employee records |
| Pending Leave Approval | View and approve pending leave requests |
| Pending OT Approval | View and approve overtime requests |
| Other Requests | Handle miscellaneous employee requests |

---

### 2. Accountant

| Button | Purpose |
|--------|---------|
| Invoice Master Form | Review deliverable completion details and raise invoices |

---

### 3. Commercial Coordinator

| Button | Purpose |
|--------|---------|
| My Attendance | Log personal daily attendance |
| My Leave Form | Submit personal leave requests |
| Project Master Form | Enter details of newly awarded projects |
| JobCode Request Form | Request creation of a new Job Code for a project |

---

### 4. Commercial Lead

| Button | Purpose |
|--------|---------|
| Team Attendance | Monitor attendance records of the entire team |
| Pending Leave Approval | Approve or reject team leave requests |
| Pending OT Approval | Approve or reject overtime requests |
| My Attendance | Log personal daily attendance |
| My Leave Form | Submit personal leave requests |
| Project Master Form | Review project details and select applicable deliverables |
| Global Reserve Pool | View unallocated employees available for assignment |
| Employee Master Form | View and reference employee records |

---

### 5. Project Lead

| Button | Purpose |
|--------|---------|
| Team Attendance | Monitor team attendance records |
| Pending Leave Approval | Approve or reject team leave requests |
| Pending OT Approval | Approve or reject overtime requests |
| My Attendance | Log personal daily attendance |
| My Leave Form | Submit personal leave requests |
| Utilization of Employees | View employee utilization across all assigned deliverables |
| Generalhours Overflow | Monitor employees whose general hours exceeded the allocated limit |
| Project Master Form | Review approved projects and assign manpower |
| Milestonehours Overflow | Monitor milestone hour overflows per deliverable |

---

### 6. Project Coordinator

| Button | Purpose |
|--------|---------|
| Project Master Form | View details of assigned projects |
| Pending Leave Approval | Approve or reject leave requests |
| Pending OT Approval | Approve or reject overtime requests |
| My Attendance | Log personal daily attendance |
| My Leave Form | Submit personal leave requests |
| Utilization of Employees | View utilization of employees on assigned projects |

---

### 7. Manager

| Button | Purpose |
|--------|---------|
| Project Manhours Approval | Review and approve manhour allocations for projects |
| Pending Leave Approval | Approve or reject leave requests |
| Pending OT Approval | Approve or reject overtime requests |
| JobCode Request Form | Approve or reject new Job Code requests |
| Reports Dashboard | Access all live project, employee, and financial reports |
| Project Master Form | View and manage project details |
| Global Reserve Pool | View unallocated employees available for assignment |
| Employee Master Form | View and manage employee records |
| Generate Payroll | One-click payroll generation based on attendance data |

---

### 8. Employee — Self Service Portal

| Button | Purpose |
|--------|---------|
| Add Attendance | Log daily working hours against assigned project and deliverable |
| Leave Application | Submit a leave request |
| Other Requests | Raise miscellaneous requests |

> Employees can only view Job Codes assigned to them. Within each Job Code, only their assigned deliverables are visible.

---

## Core Forms

### Project Master Form

A multi-section form with left-side navigation covering six areas: Project Details, Milestone Details, Client Contact Details, Weightage Details, Milestone Allocation, and Change History.

The Commercial Coordinator enters the core project information. The Commercial Lead selects applicable milestones via checkboxes and assigns weightage percentages — these must total 100% before the record can be saved. Based on the tag count and predefined hour rates, the system auto-calculates the total allocated hours per milestone in the Milestone Allocation section. Every save is tracked in the Change History section, recording the field changed, old value, new value, user, and timestamp.

---

### Invoice Master Form

A dual-role form used by the Commercial Coordinator and the Accountant within the same screen, with role-based field access.

The Commercial Coordinator selects the Job Code — the project number, contract value, and submission date auto-populate. For each applicable milestone, the coordinator enters the invoice raising percentage. The system immediately calculates the invoice amount and updates the total invoiced and balance amounts in the header in real time. Supporting documents are attached per milestone and saved to a centralized file path on the server.

Once saved, an automated email is sent to the Accountant, who logs in, reviews the amounts, uploads the final invoice document, and completes the record.

```
Invoice Amount    =  Contract Value × Invoice Raising %
Total Invoiced    =  Sum of all milestone invoice amounts
Balance Amount    =  Contract Value − Total Invoiced

Example:
  Contract Value  =  QAR 18,000
  FMTL at 50%     =  QAR 9,000
  Balance         =  QAR 9,000
```

---

## End-to-End Workflow

```
Admin
  └── Creates employee records

Commercial Coordinator
  └── Enters new project details → Auto mail → Commercial Lead

Commercial Lead
  └── Selects milestones, assigns weightage → Auto mail → Manager

Manager
  └── Reviews and approves manhours → Auto mail → Project Lead

Project Lead
  └── Assigns coordinator and employee to deliverables
  └── Auto mail → Project Coordinator + Employee

Employee
  └── Logs attendance against assigned Job Code and deliverable
  └── Hours exceeding limit → held for Project Lead approval

Commercial Coordinator
  └── Receives transmittal → raises invoice → Auto mail → Accountant

Accountant
  └── Reviews invoice → uploads final document → Invoice Raised
```

---

## Key Business Logic

```
Allocated Hours    =  Tag Count × Predefined Hour Rate

Remaining Hours    =  Allocated Hours − Total Logged Attendance Hours

Invoice Amount     =  Contract Value × Invoice Raising %

Balance Amount     =  Contract Value − Total Invoiced Amount

Salary             =  (Total Attendance Hours × Hourly Rate) + Allowances − Deductions
```

**Badge Counter:** Increments when a request or record is submitted. Decrements automatically when the responsible role completes the action.

**Overflow Rule:** If logged hours exceed the deliverable limit, the record is held and a notification is sent to the Project Lead. The record saves to the attendance table only after approval.

**Dropdown Filter:** All Job Code dropdowns across the system show only approved projects, filtered by the logged-in employee's assigned records.

---

## Technical Implementation

| Component | Purpose |
|-----------|---------|
| VBA | Auto mail triggers, badge counter updates, invoice real-time calculation, overflow approval flow, attendance validation, change history logging |
| MS Access Queries | Role-based dropdown filtering, utilization reports, remaining hours calculation, payroll aggregation |
| Macros | Startup configuration, hiding toolbars and navigation, form open and close actions |
| Form Events | On Save, On Change, On Load, On Click events used to trigger calculations and workflow steps |
| File Attachment System | Documents uploaded via Attach buttons are saved to a centralized server file path, with file references stored in the database per milestone row |
| Audit Trail | Change history captured on every save using VBA — field name, old value, new value, user, and timestamp |

---

## Challenges & Solutions

**Auto Mail Trigger Not Firing**
Implemented VBA using the CDO (Collaboration Data Objects) library within AfterUpdate and On Save form events. Added error handling to catch SMTP failures without breaking the save flow.

**Role-Based Visibility**
Built parameterized queries that filter records based on the logged-in user's Employee ID stored in a session variable at login. All dropdowns and subforms reference these filtered queries.

**Approved Projects Only in Dropdown**
Modified the record source of all Job Code dropdowns to include a `WHERE ProjectStatus = 'Approved'` filter, ensuring only approved projects appear for attendance and manpower assignment.

**Attendance Deducting from Project Master**
Written a VBA AfterInsert event on the Attendance form that runs an UPDATE query on `tbl_ProjectMaster`, recalculating UsedProjectHours by summing all attendance records for that Job Code and deliverable.

**Overflow Hours Approval Flow**
Added a BeforeUpdate event on the Attendance form that compares logged hours against the pre-defined limit. If exceeded, the save is cancelled, the record is written to a pending overflow table, and an automated email is sent to the Project Lead. On approval, the record moves to the main attendance table.

**Invoice Real-Time Calculation**
Used VBA On Change events on each Invoice Raising % field to recalculate and refresh header totals in real time using `Me.Recalc`, keeping all calculations live within the form session.

---

## Leave & OT Management

| Step | Leave Flow | OT Flow |
|------|-----------|---------|
| 1 | Employee submits Leave Application | Employee submits OT request |
| 2 | Auto mail triggered to approving role | Auto mail triggered to approving role |
| 3 | Badge counter on Pending Leave Approval increments | Badge counter on Pending OT Approval increments |
| 4 | Lead or Manager approves or rejects | Lead or Manager approves or rejects |
| 5 | Leave status updated automatically | OT hours recorded in attendance |

---

## Payroll Generation

Accessible by the Manager via the Generate Payroll button. Single-click execution powered by VBA and aggregation queries. Calculates total hours worked, gross salary, and net salary after allowances and deductions. Payroll report generated instantly for all employees for the selected month.

---

## Reports

All reports are generated from live data. Accessible via Manager → Reports Dashboard.

| Report | Description |
|--------|-------------|
| Employee Utilization | Hours logged vs. allocated per employee |
| Deliverable Progress | Planned vs. actual hours per deliverable |
| Project Status | Completion status across all active projects |
| Invoice Summary | Invoices raised, balance amounts, and status per project |
| Manpower Allocation | Team assignments across all projects |
| Payroll | Monthly net salary breakdown per employee |
| Leave & OT | Leave taken and OT approved per employee |
| Global Reserve Pool | Unallocated employees available for new assignments |

---

## PowerApps Migration

### Component Mapping

| MS Access | PowerApps Equivalent |
|-----------|---------------------|
| Master Tables | Dataverse Tables |
| Login Page | Azure Active Directory Authentication |
| Role-based Dashboards | Canvas App Screens per Role |
| Badge Counters | Power Fx Count formulas with live Dataverse queries |
| Multi-section Project Master Form | Canvas App with tab navigation |
| Checkbox Milestone Selection | Gallery control with checkbox and auto-calculated weightage |
| Invoice Real-Time Calculation | Power Fx OnChange formulas |
| File Attachments | SharePoint Document Library linked to Dataverse |
| Auto Mail Triggers | Power Automate Cloud Flows |
| Leave & OT Approval | Power Automate Approval Flows |
| Overflow Hours Approval | Power Automate with conditional branching |
| Audit Trail | Dataverse audit logging |
| Reports | Power BI Embedded Reports |
| Payroll Generation | Power Automate + Power Fx |
| User Roles & Permissions | Dataverse Security Roles |

### Tech Stack (Post Migration)

- Microsoft PowerApps — Canvas and Model-Driven Apps
- Microsoft Dataverse
- Power Automate
- Power BI
- Azure Active Directory
- Microsoft 365 — Email Integration
- SharePoint — Document storage for attachments

---

## Why Migrate to PowerApps?

| MS Access | PowerApps |
|-----------|-----------|
| Desktop only | Web and mobile accessible |
| Limited concurrent users | Scalable multi-user access |
| VBA-based email via CDO | Automated Power Automate flows |
| Local file path for attachments | SharePoint document library |
| Local file storage | Secure cloud storage via Dataverse |
| Basic reporting | Interactive Power BI dashboards |
| No mobile support | Full mobile support |

---

## Project Status

- [x] System architecture designed and documented
- [x] MS Access version — fully functional
- [x] Role-based login and dashboards — complete
- [x] Real-time badge counter system — complete
- [x] Multi-level approval workflow — complete
- [x] JobCode request and approval flow — complete
- [x] Project Master Form with milestone tracking — complete
- [x] Weightage and auto hour calculation — complete
- [x] Attendance and utilization tracking — complete
- [x] Overflow hours approval workflow — complete
- [x] Invoice auto-calculation and dual-role form — complete
- [x] Centralized document attachment system — complete
- [x] Change history audit trail — complete
- [x] Leave and OT management — complete
- [x] Payroll generation — complete
- [x] Global Reserve Pool — complete
- [ ] PowerApps Canvas App — in progress
- [ ] Dataverse migration — planned
- [ ] Power Automate flows — planned
- [ ] Power BI reports — planned
---

## Author

**Vamshi Yemula**
Power Platform Developer | MS Access → PowerApps Migration
📧 vyemula26@gmail.com
🔗 [LinkedIn Profile URL]

---

*If you found this useful, feel free to star this repository.*
