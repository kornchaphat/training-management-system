# Training Management System (TMS)

A full-stack web application for managing training sessions, tracking attendance, and analyzing training effectiveness across a global organization. Built with Google Apps Script and deployed as a web app integrated with Google Sheets.

## About

This system was built in-house to replace manual training tracking processes and fill functional gaps in our existing HRIS. It is currently in production use, managing training operations across multiple countries.

## Features

### Authentication & Access Control
- OTP-based email authentication (no passwords to manage)
- Role-based access control (Admin, Manager, Viewer)
- Session timeout with automatic expiry

### Session Management
- Create, edit, and schedule training sessions
- Bulk session creation for recurring programs
- Calendar view with interactive session cards
- Status tracking (Planned, In Progress, Completed, Cancelled)

### Attendance Tracking
- QR code generation for each participant
- QR scanning for check-in (camera-based)
- Email QR codes directly to participants
- Printable QR cards for in-person sessions

### Participant Management
- Add participants individually or via mass upload (Excel/CSV)
- Search and filter across sessions
- Track attendance history per employee

### Analytics Dashboard
- Training completion rates and trends
- Participation breakdown by department, location, and session type
- Interactive charts powered by Chart.js

### Additional Features
- Google Forms integration for post-training feedback sync
- Built-in onboarding walkthrough for new users
- Client-side caching for improved performance
- Skeleton loading states for better UX
- Sortable and paginated data tables
- Fully responsive design

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML, CSS, JavaScript |
| Backend | Google Apps Script |
| Database | Google Sheets |
| Charts | Chart.js |
| QR Codes | html5-qrcode, qrcode.js |
| Icons | Remix Icon |
| Fonts | DM Sans, Inter |

## Project Structure

```
├── Code.gs          # Backend - API handlers, authentication, data operations
├── TmsApp.html      # Main application frontend (13,000+ lines)
├── Login.html       # OTP login page
└── README.md
```

## Setup

1. Create a new Google Apps Script project
2. Copy `Code.gs` as the script file
3. Create `TmsApp.html` and `Login.html` as HTML files in the project
4. In Project Settings > Script Properties, add:
   - `SPREADSHEET_ID` — your Google Sheets database ID
5. Deploy as a web app (Execute as: Me, Access: Anyone within organization)

## Screenshots

*Coming soon*

## Author

Built by [Kornchaphat Piyatakoolkan](https://www.linkedin.com/in/kornchaphat)

## License

MIT
