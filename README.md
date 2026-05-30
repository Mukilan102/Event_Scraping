# HARTS Event Scout

Automated event intelligence system for HARTS Consulting.
Runs every **Monday at 9:00 AM IST** on GitHub's servers — no local machine required.

---

## What it does

1. Searches the web using 50 targeted queries across HARTS' 6 service areas
2. Extracts event details (name, dates, location, format, organizer)
3. Scores each event 1–10 for relevance
4. Keeps only events scored **6 or above**
5. Writes results to `HARTS_Events.xlsx` with colour-coded rows
6. Commits the updated file back to this repository automatically

---

## One-time setup (do this once)

### Step 1 — Create a GitHub account
Go to [github.com](https://github.com) and sign up (free).

### Step 2 — Create a new repository
- Click **+** → **New repository**
- Name it: `harts-event-scout`
- Set to **Private**
- Click **Create repository**

### Step 3 — Push this code to GitHub

Install Git if you don't have it: https://git-scm.com/downloads

Open a terminal in this folder and run:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/harts-event-scout.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username.

### Step 4 — Verify the workflow is active

- Go to your repo on GitHub
- Click the **Actions** tab
- You should see **HARTS Event Scout** listed
- To test immediately: click the workflow → **Run workflow**

---

## Viewing results

After each Monday run:
- Go to your repo on GitHub
- Download `HARTS_Events.xlsx`
- Open in Excel — events are colour-coded by score

| Colour | Score | Meaning |
|--------|-------|---------|
| Dark green | 9–10 | Core HARTS focus — always include |
| Light green | 8 | Strong match |
| Yellow | 7 | Good match |
| Pale yellow | 6 | Borderline — review manually |

---

## Enabling email reminders (later)

When you are ready to connect Microsoft 365:

1. Register an app in Azure Active Directory
2. Grant it `Mail.Send` permission
3. Add these 3 secrets to GitHub (Settings → Secrets → Actions):
   - `M365_CLIENT_ID`
   - `M365_CLIENT_SECRET`
   - `M365_TENANT_ID`
4. Uncomment Phase 4 in `event_scout.py`

---

## Files

| File | Purpose |
|------|---------|
| `event_scout.py` | Main script — all 5 phases |
| `config.py` | Settings, search queries, file paths |
| `requirements.txt` | Python dependencies |
| `.github/workflows/schedule.yml` | GitHub Actions — runs every Monday |
| `HARTS_Events.xlsx` | Output — created/updated on each run |
| `event_scout.log` | Run log — appended each week |

---

## Troubleshooting

**Workflow not running?**
- Check the Actions tab for error messages
- Make sure the repo has at least one commit after the workflow file was pushed

**No events found?**
- Check `event_scout.log` in the repo for details
- DuckDuckGo Search may have rate-limited — re-run manually after a few hours

**Duplicate events appearing?**
- The system deduplicates by URL — safe to re-run anytime
