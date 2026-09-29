# 🌾 Kentucky Wheat Yield Contest — Digital Entry Form

**University of Kentucky Cooperative Extension**  
Built with Streamlit · Google Sheets · Cloudinary · FormSubmit

🔗 **Live App:** https://wheatcontest.streamlit.app

---

## Overview

A web-based entry form that allows county agents to submit official Kentucky Wheat Yield Contest entries directly from the field. Entries are saved to Google Sheets, scale ticket photos are stored on Cloudinary, and an email notification is automatically sent to the state office on every submission.

---

## Features

- **8-section structured form** — Producer Info, Agronomic Data, Fertilizer, Pest Management, Harvest Area, Grain Characteristics, Yield Calculation, Notes & Photo
- **Live yield calculation** — official yield (bu/acre) computed in real time from grain weight, moisture, and harvest area
- **Up to 3 moisture readings** with automatic average
- **Scale ticket photo upload** — stored permanently on Cloudinary with a clickable URL saved in the sheet
- **Google Sheets sync** — every entry appended automatically with all 53 fields
- **Email notification** sent to state office (+ CC) on every submission via FormSubmit
- **Per-entry Excel receipt** — agent downloads only their own entry, not the master database
- **Agent certification checkbox** required before submission unlocks
- **Field validation** with clear error messages
- **Session log** — shows all entries submitted in the current browser session

---

## Tech Stack

| Component | Service | Notes |
|---|---|---|
| App hosting | Streamlit Community Cloud | Free tier |
| Data storage | Google Sheets | Via service account |
| Photo storage | Cloudinary | Free tier, 25 GB |
| Email | FormSubmit.co | Free, no backend needed |
| Language | Python 3.14 | |
| Keep-alive | GitHub Actions | Pings every 6 hours |

---

## Setup

### 1. Clone & Deploy

```bash
git clone https://github.com/your-username/wheat-yield-contest
```

Deploy to [Streamlit Community Cloud](https://streamlit.io/cloud) by connecting your GitHub repo.

### 2. requirements.txt

```
streamlit
pandas
openpyxl
requests
cryptography
Pillow
```

### 3. Streamlit Secrets

Go to your app dashboard → **Settings → Secrets** and add:

```toml
[gdrive]
sheet_id                 = "your-google-sheet-id"
client_email             = "your-service-account@project.iam.gserviceaccount.com"
private_key              = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
cloudinary_cloud_name    = "dxxxxxxxxx"
cloudinary_upload_preset = "YieldContest"
```

### 4. Google Sheets Service Account

1. Go to [console.cloud.google.com](https://console.cloud.google.com) → create a project
2. Enable **Google Sheets API** and **Google Drive API**
3. Create a **Service Account** → generate a JSON key
4. Share your Google Sheet with the service account email (Editor access)
5. Paste `client_email` and `private_key` into Streamlit secrets

### 5. Cloudinary Photo Upload

1. Sign up free at [cloudinary.com](https://cloudinary.com) (no credit card)
2. Copy your **Cloud Name** from the dashboard
3. Go to **Settings → Upload → Upload Presets → Add Upload Preset**
4. Set **Signing Mode = Unsigned**, Asset Folder = `ScalePhotos` → Save
5. Add `cloudinary_cloud_name` and `cloudinary_upload_preset` to secrets

### 6. Keep App Awake (GitHub Actions)

Streamlit Community Cloud sleeps apps after **12 hours** of inactivity. UptimeRobot HTTP pings do **not** work — Streamlit requires a real WebSocket connection. Use GitHub Actions instead.

Create `.github/workflows/keep_alive.yml` in your repo:

```yaml
name: Keep App Awake
on:
  schedule:
    - cron: '0 */6 * * *'   # every 6 hours
  workflow_dispatch:

jobs:
  wake:
    runs-on: ubuntu-latest
    steps:
      - name: Ping Wheat Contest App
        run: curl -L --max-time 30 https://wheat-yield-contest.streamlit.app/ || true
```

> Uses ~180 GitHub Action minutes/month. Free tier provides 2,000 min/month per account.

---

## Form Sections

| Section | Fields |
|---|---|
| 1 — Producer & Supervisor | County, Name, Email, Phone, Mobile, Address, Profession, Supervisor Name/Phone/Date |
| 2 — Agronomic Data | Division, Previous Crop, Variety, Planting/Harvest Dates, Row Width, Seeding Rate |
| 3 — Fertilizer | Fall N/P/K, Spring N applications (2), Manure type/rate/date |
| 4 — Pest Management | Growth regulators, Fall/Spring/Heading pest products, Biologicals, Tillage |
| 5 — Harvest Area | Length × Width (ft) → auto-calculates ft² and acres (min 1.50 ac required) |
| 6 — Grain Characteristics | Up to 3 moisture readings → auto-average, Test Weight (lb/bu) |
| 7 — Yield Calculation | Grain Weight (lbs) → official yield auto-calculated |
| 8 — Notes & Photo | Agent notes, Scale ticket photo (JPG/PNG/HEIC/WebP) |

### Official Yield Formula

```
Yield (bu/ac) = Grain Weight × [(100 − %Moisture) ÷ 86.5] ÷ 60 ÷ Acres
```

---

## Data Flow on Submission

1. Entry saved to master Excel database on server
2. Photo uploaded to Cloudinary → permanent URL returned
3. All 53 fields + photo URL appended to Google Sheet
4. Email notification sent to state office via FormSubmit
5. Success banner shown with yield summary
6. Single-entry Excel receipt generated for agent download (private — no other entries)
7. Entry added to session log at bottom of page

---

## Known Limitations

- Session data is lost if the app goes to sleep mid-form — mitigated by the GitHub Actions keep-alive
- Master Excel file (`/tmp/`) is wiped on every Streamlit server restart — Google Sheets is the permanent record
- Free Cloudinary tier: 25 GB storage, more than sufficient for contest photos
- FormSubmit requires a one-time email confirmation on first use — check inbox after first submission
- HEIC photos (iPhone) show a fallback preview in the app but upload to Cloudinary correctly

---

## Contact

| Role | Contact |
|---|---|
| App Development | Mohammad Jan Shamim — mshamim11@uky.edu |
| Contest Administration | Dr. Chad Lee — chad.lee@uky.edu |
| UK Extension | University of Kentucky Cooperative Extension |

---

*Kentucky Wheat Yield Contest · University of Kentucky Cooperative Extension · 2026*
