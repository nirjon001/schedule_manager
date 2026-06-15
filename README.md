# Schedule Manager

A Flask web application that parses university advising Excel files (.xlsx/.xls) and generates an interactive class routine schedule with PDF and PNG export.

## Features

- Upload university advising slips (Excel format)
- Auto-detects courses, sections, time slots, rooms, and days
- Interactive schedule grid with live clock and current class highlighting
- Faculty initials input — displayed inline and embedded in exports
- Download schedule as PDF (ReportLab)
- Download schedule as PNG (Pillow)
- Auto-cleanup of uploaded files (30 min TTL)

## Requirements

- Python 3.9+
- Dependencies listed in `requirements.txt`

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python app.py
```

Open `http://localhost:5000` in a browser. Upload an advising slip Excel file to generate the schedule.

**Live demo:** [https://schedule-manager-19y6.onrender.com](https://schedule-manager-19y6.onrender.com)

For debug mode:

```bash
python app.py --debug
```

## Expected Excel Format

The parser reads a standard university advising slip with columns for:

| Column | Description |
|--------|-------------|
| Course(s) | Course code (e.g. `CSE101`) |
| Sec | Section number |
| Time-WeekDay | Format: `STWRF 8:00AM-10:00AM` (day codes: S=Sun, M=Mon, T=Tue, W=Wed, R=Thu, F=Fri, A=Sat) |
| Room | Room number |

Student name and ID are auto-extracted from cells labeled `Name:` and `ID#`.

## Project Structure

```
schedule_manager/
├── app.py              # Flask application & PDF/PNG generation
├── scheduler.py        # Excel parsing & grid building logic
├── requirements.txt    # Python dependencies
├── static/
│   └── style.css       # Stylesheet
├── templates/
│   ├── index.html      # Upload page
│   └── schedule.html   # Schedule display with live features
└── uploads/            # Temporary uploaded files (auto-cleaned)
```

## Tech Stack

- **Backend:** Flask, openpyxl, ReportLab, Pillow
- **Frontend:** Vanilla JavaScript, CSS
- **Export:** PDF (ReportLab), PNG (Pillow)

## Author

Developed by Ratul Hasan Nirjon
