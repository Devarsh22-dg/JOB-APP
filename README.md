# JobMatch — Deployment Steps

## 1. Push to GitHub
Create a new repo (e.g. `jobmatch`) and add these two files to it:
- `jobmatch_app.py`
- `requirements.txt`

## 2. Deploy on Streamlit Community Cloud
1. Go to https://share.streamlit.io and sign in
2. "New app" → pick the repo → set main file path to `jobmatch_app.py`
3. Deploy

## 3. Add your Anthropic API key
1. In the deployed app's dashboard: **⋮ → Settings → Secrets**
2. Add:
   ```
   ANTHROPIC_API_KEY = "sk-ant-your-actual-key-here"
   ```
3. Save — the app auto-reboots

If you don't have a key yet: console.anthropic.com → Settings → Billing (add payment method,
this is separate/pay-per-use from any claude.ai subscription) → API Keys → Create Key.

## 4. Add Adzuna API credentials (for the Job Feed tab)
1. Sign up free at https://developer.adzuna.com/
2. Confirm your email, then go to your dashboard to get your `App ID` and `App Key`
3. Add both to the same Secrets panel as above:
   ```
   ADZUNA_APP_ID = "your-app-id"
   ADZUNA_APP_KEY = "your-app-key"
   ```
4. Save — the app auto-reboots

Free tier covers plenty of searches for personal job hunting use.

## 5. Add Jooble API key (second job source)
1. Sign up free at https://jooble.org/api/about
2. Get your API key
3. Add to Secrets:
   ```
   JOOBLE_API_KEY = "your-jooble-key"
   ```

## About the Company Watchlist and Quick Links
- **Watchlist** (Greenhouse/Lever): only works for companies that actually host their careers
  page on Greenhouse or Lever — mostly tech startups/scaleups. To add one, find their board
  token from their careers URL (e.g. `boards.greenhouse.io/stripe` → token is `stripe`) and
  add it in the app.
- **Quick Links**: LinkedIn, Indeed, Job Bank (Canada's official government/general job engine),
  and the specific companies requested (FAANG, Big 4, Robert Half, CGI, TD, CIBC, TCS,
  government) don't offer public search APIs — most large enterprises (TD, CIBC, etc.) run on
  Workday, which doesn't have a documented public jobs API either. These links jump straight to
  each site's own search, prefilled with your keywords/location where the platform's URL
  supports it, so there's nothing to retype, but results aren't pulled into the app itself.

## What this app does
- **Profile**: stores Kahini's resume as editable text (pre-loaded from her current resume)
- **Job Feed**: pulls live listings from Adzuna for a keyword + location search. Per listing:
  check fit score, generate a tailored resume + cover letter, open the real posting to apply,
  then click "Mark as Applied" to log it (with the exact tailored resume/cover letter attached)
- **Analyze Fit**: one-off — paste any job posting (text or URL) → fit score, matched skills, gaps
- **Tailor Resume**: generates a JD-matched resume (summary + bullet rewrites), exports to .docx
- **Cover Letter**: generates a matching cover letter, exports to .docx
- **Dashboard**: every logged application — status, fit score, editable status/notes, summary
  metrics and a chart, plus a "view tailored resume/cover letter used" expander per application
  so you can always see exactly what was submitted for a given job
- **Email Check**: placeholder tab — backlog item, not yet built (see below)

## Data persistence
Applications and the job feed are stored in a local SQLite file (`jobmatch.db`) inside the
app's own storage. This survives normal use and app restarts, but **will be wiped if you
redeploy the app** (new git push) or if Streamlit Cloud rebuilds the container from scratch
after a long period of inactivity. Use the "Download full tracker (.xlsx)" button on the
Dashboard regularly as a backup.

## What it deliberately does NOT do
- It does not submit applications automatically. Every tailored resume and cover letter is
  generated for review — Kahini opens the real posting and applies herself. This avoids
  bot-detection blocks on job boards and keeps her in control of what goes out under her name.
  "Mark as Applied" only logs the application here; it doesn't submit anything.
- URL fetching (Analyze Fit tab) works for plain-HTML job postings but many boards (LinkedIn,
  Indeed) render content via JavaScript that a simple fetch won't see — pasting the JD text
  directly is more reliable in those cases. The Job Feed tab avoids this problem entirely by
  pulling structured data straight from Adzuna's API.

## Backlog (not yet built)
- **Automated email check**: read-only scan of Gmail via an app password to flag likely
  responses (interviews, rejections) against logged applications, for Kahini to confirm and
  update status. Deferred for now — status updates are manual via the Dashboard.
