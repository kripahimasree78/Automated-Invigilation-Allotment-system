# Automated Invigilation Allotment System

A web-based invigilation duty management system that automates the complete exam duty allocation workflow — uploading faculty and exam timetables, auto-assigning invigilators based on designation, experience, and availability, managing duty swaps, and generating downloadable Excel reports — with role-based access for Admin and Faculty.

---

## Overview

The admin uploads a faculty timetable, FN/AN exam timetables, and a staff joining dates Excel file. The system automatically checks each faculty member's free slots, filters out Professors and Associate Professors, and assigns invigilation duties to eligible staff (Assistant Professors and Non-Teaching Staff) in experience-weighted order — most junior staff get more duties, most senior staff get fewer. Faculty can log in to view their assigned duties and request swaps with the admin or with peers. All state is held server-side; the SQLite database persists assignment counts across sessions.

---

## How It Works

```
1. Admin logs in  →  session created (role = admin)
          ↓
2. Admin uploads 4 files:
     • Faculty Timetable (.xlsx)      — master schedule, all faculty
     • FN Exam Timetable (.xlsx)      — Forenoon exam dates & subjects
     • AN Exam Timetable (.xlsx)      — Afternoon exam dates & subjects
     • Staff Joining Dates (.xlsx)    — DOJ + designation per staff member
          ↓
3. Admin sets FN / AN invigilators-per-slot (default 6 each)
   and clicks Generate Allocations
          ↓
4. System reads timetables  →  builds availability map per faculty per day
   Loads DOJ → computes max-invig limit per staff member:
     ≥25 yrs → 1  |  ≥23 yrs → 2  |  ≥20 yrs → 3
     ≥15 yrs → 4  |  <15 yrs  → 5
          ↓
5. Excludes: Professor & Head, Professor & Dean, Professor,
             Associate Professor / Assoc. Professor
   Eligible: Asst. Professor, Teaching Assistant, Non-Teaching Staff
          ↓
6. Allocation order per slot:
     Group 0 — Non-Teaching Staff   (assigned first, exhausted fully)
     Group 1 — Asst. Professors     (assigned after Group 0 is full)
   Within each group: lowest experience (highest max) → highest experience
          ↓
7. Assignments stored in invigilation.db + displayed in admin dashboard
          ↓
8. Admin can:
     • View assignment counts per faculty  (Assigned Count page)
     • Add extra invigilations manually    (Add Invigilations)
     • Override per-faculty max counts     (Modify Allocations)
     • Finalize and download Excel report  (Download Report)
     • Approve / reject swap requests      (Manage Swaps)
          ↓
9. Faculty logs in with their name (no title prefix) + shared password
   → views their duties and can request swaps with admin or a peer
```

---

## Features

- **Role-Based Login** — Admin (username + password) and Faculty (name + shared password) with separate dashboards
- **Four-File Upload** — Faculty Timetable, FN Exam Timetable, AN Exam Timetable, Staff Joining Dates (.xlsx)
- **Designation-Based Exclusion** — Professors, Associate Professors, and HoD are automatically excluded; fuzzy name-matching handles spelling differences between timetable and staff list
- **Experience-Weighted Allocation** — Max invig count derived from Date of Joining; junior staff carry more duties
- **Priority Ordering** — Non-Teaching Staff are fully loaded before any Asst. Professor receives a duty
- **Add Invigilations** — Admin can manually assign extra slots for any valid exam date / session
- **Modify Allocations** — Admin selects specific faculty, overrides their max count, and regenerates
- **Dual Swap System** — Faculty can request admin swaps or peer-to-peer swaps; admin approves all
- **Assigned Count Report** — Live search + filter (All / Non-Teaching / Asst. Professor) with load-bar visualisation
- **Finalized Schedule** — Date-wise FN / AN table with one-click Excel download
- **SQLite Persistence** — Assignment counts survive server restarts via `invigilation.db`

---

## Allocation Rules

| Years of Experience (from DOJ) | Max Invigilations |
|---|---|
| ≥ 25 years | 1 |
| ≥ 23 years | 2 |
| ≥ 20 years | 3 |
| ≥ 15 years | 4 |
| < 15 years  | 5 |

**Excluded designations (never assigned):**
`Professor & Head` · `Professor & Dean (R&D)` · `Professor & Dean (I&I)` · `Professor` · `Assoc. Professor` · `Associate Professor`

**Eligible designations (assigned in priority order):**

| Priority | Group | Examples |
|---|---|---|
| 1st | Non-Teaching Staff | Lab Technician, Office Staff |
| 2nd | Asst. Professor / Teaching Assistant | Asst. Professor, Teaching Assistant |

---

## System Architecture

| Tier | Components |
|---|---|
| **Presentation Layer** | `login.html`, `admin.html`, `faculty.html`, `finalize.html`, `assigned_count.html`, `add_invigilations.html`, `admin_modify.html`, `admin_swaps.html` — served by Flask + Jinja2 |
| **Application Logic Layer** | Flask routes for upload, allocation, swap management, report generation; Pandas for timetable parsing; openpyxl for Excel output |
| **Data Layer** | `invigilation.db` (SQLite) — faculty assignment counts, max counts, remaining counts; in-memory `STATE` dict — all timetable data, allocations, swap queues |

---

## Tech Stack

| Component | Technology |
|---|---|
| Backend Framework | Python 3.x, Flask |
| Data Processing | Pandas, openpyxl |
| Database | SQLite (`invigilation.db`) |
| Frontend | HTML5, CSS3, JavaScript (ES6) |
| Styling | Custom CSS (no framework) |
| Version Control | Git / GitHub |

---

## Project Structure

```
InvigilationAllotmentSystem/
│
├── app.py                        # Main Flask application — all routes, allocation logic,
│                                 #   swap management, DB sync, timetable parsers
│
├── templates/
│   ├── login.html                # Login page — Admin & Faculty role selector
│   ├── admin.html                # Admin dashboard — file upload, generate, allocations table
│   ├── faculty.html              # Faculty dashboard — duty view, swap request
│   ├── add_invigilations.html    # Admin: add extra invig slots manually
│   ├── admin_modify.html         # Admin: override per-faculty max counts
│   ├── admin_swaps.html          # Admin: approve / reject swap requests
│   ├── assigned_count.html       # Admin: faculty-wise count report with search & filter
│   └── finalize.html             # Admin: finalized date-wise schedule + Excel download
│
├── invigilation.db               # Auto-created on first run — faculty assignment counts
│
└── uploads/  (auto-created)      # Temp directory for uploaded timetable files
```

---

## Input File Formats

### Faculty Timetable (`.xlsx`)
Master schedule for all faculty. Each row is one faculty member.

| Column | Description |
|---|---|
| Col 1 | S.No — serial number |
| Col 2 | Name — full faculty name with title (e.g. `Mrs. Ch. Mandakini`) |
| Col 3+ | Period-wise slots for each day (MON–SAT) — subject codes or free |

### FN / AN Exam Timetables (`.xlsx`)
One row per exam session.

| Column | Description |
|---|---|
| Date | Exam date (DD/MM/YYYY or similar) |
| Day | Day of week (MON, TUE …) |
| Subject | Subject name / code |

### Staff Joining Dates (`.xlsx`)
Used to compute max-invig limits and designation-based exclusions.

| Column | Description |
|---|---|
| S.No | Serial number |
| Name | Full staff name with title |
| Designation | e.g. `Asst. Professor`, `Professor & Head`, `Non-Teaching Staff` |
| Date of Joining | DOJ in DD/MM/YYYY or DD-MM-YYYY format |

> The system auto-detects the header row and skips section dividers such as "Teaching Staff" and "Non-Teaching Staff" automatically.

---

## Installation

**Prerequisites:** Python 3.9+, Windows / Linux / macOS

### Step 1 — Clone the repository

```bash
git clone <your-repository-url>
cd InvigilationAllotmentSystem
```

### Step 2 — (Optional) Create and activate a virtual environment

```bash
python -m venv venv
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate           # Windows
```

### Step 3 — Install Python dependencies

```bash
pip install flask pandas openpyxl python-docx
```

---

## Running the Application

```bash
python app.py
```

Open your browser and navigate to:

```
http://127.0.0.1:5050
```

> The app runs on **port 5050** by default (`app.run(debug=True, port=5050)`).

---

## Default Login Credentials

| Role | Username | Password |
|---|---|---|
| Admin | `admin` | `admin123` |
| Faculty | Name without title prefix | `pass123` |

**Faculty login example:**
If the faculty name in the timetable is `Mrs. Ch. Mandakini`, log in as:
- Username: `Ch. Mandakini`
- Password: `pass123`

> ⚠️ Change `ADMIN_PASS`, `FACULTY_PASS`, and `app.secret_key` before any institutional deployment.

---

## Route Reference

| Route | Method | Description |
|---|---|---|
| `/` | GET | Redirects to login |
| `/login` | GET, POST | Login page — Admin and Faculty |
| `/logout` | GET | Clears session |
| `/admin` | GET | Admin dashboard |
| `/upload` | POST | Upload timetable files |
| `/generate` | POST | Run allocation algorithm |
| `/admin/add_invigilation` | GET, POST | Add manual invigilation slot |
| `/admin/modify` | GET | Select faculty for max-count override |
| `/admin/modify_submit` | POST | Apply overrides and regenerate |
| `/admin/swaps` | GET | View and manage swap requests |
| `/admin/swap_action` | POST | Approve or reject a swap |
| `/admin/swap_candidates/<id>` | GET | Eligible candidates for a swap |
| `/assigned_count` | GET | Faculty-wise assignment count report |
| `/finalize` | GET | Finalized date-wise schedule view |
| `/download_report` | GET | Download Excel report |
| `/faculty` | GET | Faculty duty view dashboard |
| `/faculty/swap_request` | POST | Faculty requests swap with admin |
| `/faculty/peer_swap_request` | POST | Faculty requests swap with a peer |
| `/faculty/peer_swap_respond` | POST | Peer accepts or declines swap |

---

## Database Schema

`invigilation.db` is auto-created on first run via `init_db()`.

```
faculty_counts
├── faculty_name    TEXT  PRIMARY KEY
├── max_count       INTEGER
├── current_count   INTEGER
└── remaining_count INTEGER
```

> All other system state (timetables, allocations, swap queues, joining dates, designation map) is held in the in-memory `STATE` dictionary and reloaded from the uploaded files on each Generate call.

---

## Notes

- **Staff list must be uploaded** before clicking Generate — if omitted, no designations are known and no faculty will be excluded.
- **Name matching is fuzzy** — minor spelling differences between the faculty timetable and the staff list (e.g. `Lalita` vs `Lalitha`, `Sundhar` vs `Sundar`) are resolved automatically by the normaliser in `is_excluded_by_designation()`.
- **Regenerating allocations** always reloads the staff list from disk first, so designation exclusions are never stale across server restarts.
- **Custom max counts** set via Modify Allocations are stored in `STATE['custom_max']` and applied on every subsequent generate.
- **Swap queues** (`STATE['swap_requests']` and `STATE['peer_swap_requests']`) are reset each time allocations are regenerated.
- All SQL queries use parameterised placeholders to prevent SQL injection.
- Change `app.secret_key` from `'invigilation_secret_2026'` before production deployment.

---

## Requirements

```
flask
pandas
openpyxl
python-docx
```
