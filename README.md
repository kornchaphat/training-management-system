# Training Management System (TMS)

A full-stack web application for managing training sessions, tracking attendance, and analyzing training effectiveness across a global organization. Built with Google Apps Script and deployed as a web app integrated with Google Sheets.

> 🟢 **[Live Demo](https://script.google.com/macros/s/AKfycbyyYRrBjAqhnFJX-xsUsUpxPqDeG2DJN1XT4a_gt_aFpaZMahcet7rqILGJzKCKvAR_/exec)** — Try it with sample data, no login required

## Live Demo

The demo runs on a sanitized dataset with fictional companies and employees. Authentication is bypassed so you can explore all features immediately.

| | Production | Demo |
|---|---|---|
| **Authentication** | OTP email verification | Bypassed (auto-login) |
| **Data** | Real employee data across 20 entities | 200 fictional employees, 5 entities |
| **Email notifications** | QR cards, form invitations | Disabled |
| **Access control** | Role-based (6 roles) | Full admin access |

> ⚠️ **Note:** On first visit, Google may show an "unverified app" warning. This is standard for Apps Script web apps deployed from personal accounts. Click **"Advanced" → "Go to TrainingApp (unsafe)"** to proceed.

## Screenshots

<p align="center">
  <img src="https://raw.githubusercontent.com/kornchaphat/training-management-system/main/Screenshot%202026-02-06%20230722.png" width="700" alt="Dashboard">
  <br><em>Training Calendar — Dashboard with session cards and monthly overview</em>
</p>
<p align="center">
  <img src="https://raw.githubusercontent.com/kornchaphat/training-management-system/main/Screenshot%202026-02-06%20230844.png" width="700" alt="Training Analytics">
  <br><em>Training Analytics — KPIs, charts, and filterable breakdowns</em>
</p>
<p align="center">
  <img src="https://raw.githubusercontent.com/kornchaphat/training-management-system/main/Screenshot%202026-02-06%20230950.png" width="700" alt="QR Attendance">
  <br><em>QR Attendance — Camera-based check-in with real-time scanning</em>
</p>
<p align="center">
  <img src="https://raw.githubusercontent.com/kornchaphat/training-management-system/main/Screenshot%202026-02-06%20232336.png" width="700" alt="Session Detail">
  <br><em>Session Detail — Participant list with enrollment management</em>
</p>

## About

This system was built in-house to replace manual training tracking processes and fill functional gaps in an existing HRIS. It is currently in production use, managing training operations across multiple countries and entities.

The production version includes OTP email authentication and role-based access control. The live demo bypasses authentication so you can explore all features freely.

## Features

### Authentication & Access Control (Production)
- OTP-based email authentication (no passwords to manage)
- 6 roles: Developer, Global Talent CoE, Regional Talent CoE, Global HRBP, Country HRBP, Manager
- Entity-scoped access — users only see data for their assigned countries
- Session timeout with automatic expiry

### Session & Program Management
- Create and organize training programs with nested sessions
- Calendar view with interactive session cards across 5-month range
- Status tracking (Open, In Progress, Completed, Cancelled)
- Entity-level session ownership with audit trail

### Attendance Tracking
- QR code generation per participant with embedded employee data
- Real-time QR scanning for check-in via device camera
- Email QR cards directly to participants
- Printable QR card batches for in-person sessions
- Auto-mark absent for no-shows on QR-tracked sessions

### Participant Management
- Add participants individually or via bulk Excel/CSV upload
- Search and filter across the full employee directory
- Track attendance history and training hours per employee
- Per-participant cost tracking with multi-currency support

### Training Analytics
- KPI dashboard: learning hours, completion rate, unique learners, budget utilization
- Breakdown by category, channel, entity, function, and job band
- Year/quarter/month filtering with cascading filters
- Individual development plan tracking
- PDF export for analytics reports

### Forms & Feedback
- Google Forms integration for feedback, surveys, and assessments
- Auto-link forms to sessions with one-click creation
- Poll-based response sync (no trigger dependency)
- Average satisfaction score auto-calculated per session

### Additional
- Built-in onboarding walkthrough for new users
- Client-side caching with cache invalidation
- Skeleton loading states for perceived performance
- Sortable, paginated data tables
- Exchange rate integration for multi-currency cost conversion
- Fully responsive design

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML, CSS, JavaScript (13,000+ lines, single-file) |
| Backend | Google Apps Script (9,600+ lines) |
| Database | Google Sheets (11 sheets) |
| Charts | Chart.js |
| QR | html5-qrcode, qrcode.js |
| PDF Export | Google Docs API (server-side generation) |
| Icons | Remix Icon |
| Fonts | DM Sans, Inter |

## Architecture

```
┌─────────────────────────────────────────────────────┐
│  Browser                                             │
│  ┌─────────────────────────────────────────────────┐ │
│  │  TmsApp.html (13K lines)                        │ │
│  │  - Dashboard, Calendar, Analytics, QR Scanner   │ │
│  │  - Client-side caching + skeleton loading       │ │
│  │  - google.script.run ←→ Backend API calls       │ │
│  └─────────────────────────────────────────────────┘ │
└──────────────────────┬──────────────────────────────┘
                       │ google.script.run
┌──────────────────────▼──────────────────────────────┐
│  Google Apps Script Backend (Code.gs)                │
│  - OTP Auth + Session Management                     │
│  - Role-based Access Control (6 roles)               │
│  - CRUD: Programs, Sessions, Enrollments             │
│  - QR Generation + Attendance Scanning               │
│  - Analytics Engine + PDF Export                      │
│  - Google Forms Integration + Polling                 │
└──────────────────────┬──────────────────────────────┘
                       │ SpreadsheetApp / DriveApp
┌──────────────────────▼──────────────────────────────┐
│  Google Sheets Database                              │
│  Employees │ Programs │ Sessions │ Enrollments       │
│  Users │ TM1 (Budget) │ FormResponses │ Scan_History │
│  FeedbackScores │ Active_Sessions │ OTP_Sessions     │
└─────────────────────────────────────────────────────┘
```

## Project Structure

```
├── Code.gs              # Production backend (9,600+ lines)
│                        #   Full OTP auth, email notifications, RBAC
├── DemoCode.gs          # Demo backend
│                        #   Auth bypassed, emails disabled
├── TmsApp.html          # Main application frontend (13,000+ lines)
├── Login.html           # OTP login page (production only)
├── TMS_Demo_Database.xlsx  # Sample dataset for demo deployment
└── README.md
```

## Deploy Your Own Demo

1. Create a new [Google Apps Script](https://script.google.com) project
2. Create `DemoCode.gs`
3. Create `TmsApp.html` as an HTML file and paste the frontend code
4. Upload `TMS_Demo_Database.xlsx` to Google Sheets (File → Save as Google Sheets)
5. Copy the spreadsheet ID from the URL
6. In Apps Script: **Project Settings → Script Properties → Add**
   - `SPREADSHEET_ID` = your sheet ID
7. **Deploy → New Deployment → Web app**
   - Execute as: Me
   - Who has access: Anyone
8. Open the deployment URL

## Author

Built by **Kornchaphat Piyatakoolkan**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://www.linkedin.com/in/kornchaphat)

