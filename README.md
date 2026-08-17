<div align="center">

# 🗓️ Schedule Manager

Turn your university advising slip into a beautiful class routine — in seconds.

![Python](https://img.shields.io/badge/Python-3.11-blue)
![Framework](https://img.shields.io/badge/Framework-Flask-000000)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Deploy-Render-46e3b7)

[**▶ Live Demo**](https://schedule-manager-0fsd.onrender.com) &nbsp;·&nbsp; [Deploy on Render](https://render.com/deploy?repo=https://github.com/nirjon001/schedule_manager)

</div>

---

A Flask web app that parses university **advising slip Excel files** (`.xlsx` / `.xls`) and generates an **interactive class routine** — with one-click **PDF** and **PNG** export. No accounts, no databases, no setup: upload, and go.

## ✨ Features

- 📤 Upload university advising slips (Excel format)
- 🧠 Auto-detects **courses, sections, time slots, rooms, and days** (column positions are detected dynamically — no hardcoded layouts)
- 🗓️ Interactive schedule grid with a **live clock** and **current/next class** highlighting
- 👨‍🏫 Faculty initials input — shown inline on the grid and embedded into exports
- 📄 Download your routine as **PDF** (ReportLab)
- 🖼️ Download your routine as **PNG** (Pillow)
- 🧹 Auto-cleanup of uploaded files (30-minute TTL)

## 🚀 Quick Start

```bash
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5000` in your browser, upload your advising slip, and watch the routine generate. Debug mode:

```bash
python app.py --debug
```

## ☁️ Deploy to Render

This repo ships with a `render.yaml` **Blueprint** — deploy in one click:

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/nirjon001/schedule_manager)

Or create a service manually:

| Setting | Value |
|---------|-------|
| Build command | `pip install -r requirements.txt` |
| Start command | `gunicorn app:app` |
| Runtime | Python |

## 📋 Expected Excel Format

The parser reads a standard university advising slip with columns for:

| Column | Description |
|--------|-------------|
| **Course(s)** | Course code (e.g. `CSE101`) |
| **Sec** | Section number |
| **Time-WeekDay** | Format: `STWRF 8:00AM-10:00AM` (day codes: `S`=Sun, `M`=Mon, `T`=Tue, `W`=Wed, `R`=Thu, `F`=Fri, `A`=Sat) |
| **Room** | Room number |

Student **name** and **ID** are auto-extracted from cells labeled `Name:` and `ID#`.

## 📁 Project Structure

```
schedule_manager/
├── app.py              # Flask app, routes & PDF/PNG generation
├── scheduler.py        # Excel parsing & grid building logic
├── render.yaml         # Render Blueprint (one-click deploy)
├── requirements.txt    # Python dependencies
├── static/
│   └── style.css       # Stylesheet
├── templates/
│   ├── index.html      # Upload page
│   └── schedule.html   # Interactive schedule view
└── uploads/            # Temporary uploaded files (auto-cleaned)
```

## 🛠️ Tech Stack

- **Backend:** Flask · openpyxl · ReportLab · Pillow
- **Frontend:** Vanilla JavaScript · CSS
- **Deployment:** Render (Blueprint / gunicorn)

## 📜 License

Released under the [MIT License](LICENSE).

---

<div align="center">

Developed with 💙 by **Ratul Hasan Nirjon**

</div>