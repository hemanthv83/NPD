# NPD Sourcing Tracker (Streamlit)

Web application for E&E BOP alternate development (NPD de-risking) at Sun Mobility CPMO:
project stage-gate tracking, pending/completed activities, parts MPN and price coverage,
supplier-wise view, DC fan landed-cost comparisons and Teknic domestic-source savings.

It replaces the earlier HTML/JavaScript version (`index.html`, `app.js`, `xlsx.full.min.js`).
All its features are included, plus supplier-wise tracking, charts and an Excel download that
keeps your original formatting.

## Your data is never stored in this repository

The app starts empty. Each user loads their own data:

1. **Data Management → Load Excel**: upload `NPD_Master_Tracker_Updated.xlsx`.
   The app reads these sheets: `MQIS AND SINGLE PHASE PART`, `6DS PROJECT PARTS`,
   `ALTERNATE DEVELOPNMENT`, `DC Fan Comparison` and `Teknic Parts (Domestic)`.
   If a name doesn't match, it uses the sheet in the same position (1st to 5th).
2. You can also **Restore backup** from a JSON file. Backups from the old HTML tracker
   (`NPD_Tracker_Data.json` or `NPD_Tracker_Backup_*.json`) work too.
3. Work in the app.
4. **Before closing, download your data** from Data Management:
   - **Updated master Excel (keeps your formatting)**: writes your changes into the workbook you uploaded.
   - **Clean data export (.xlsx)**: plain tables.
   - **JSON backup**: everything, for restoring later.

> **Streamlit Community Cloud does not keep data.** The app restarts after inactivity and on every
> redeploy, and its memory is cleared. Whatever you have not downloaded is lost.

`.gitignore` blocks `.xlsx` and `.json` files so tracker data is not committed by mistake.
**Uploading through the GitHub website ignores `.gitignore`, so never drag data files into the repository.**

## Pages

| Page | What it does |
|---|---|
| Dashboard | KPIs, a card per project with status and progress, progress chart, activity-status chart, delayed activities, parts missing data, annual savings |
| Project Tracking | Add a project (optionally with the standard stage-gate list); edit activities in a table (add/delete rows, plan/actual dates, remarks); plan-vs-actual timeline; rename, change status or delete a project |
| Pending & Completed | All activities across projects with project, status and text filters; pending/completed tabs; chart by project |
| MQIS & Single Phase Parts / 6DS Project Parts | Search, filter by project and supplier, show only parts missing MPN or price (highlighted red), sort by any column; **Edit mode** to add, edit or delete parts, with duplicate Part No checking |
| Supplier-wise | Parts per supplier, MPN and price coverage, BOM value, chart, parts list per supplier |
| DC Fan Comparison | The original sheet, plus added comparisons with automatic landed cost (price × FX + freight + insurance → CIF + BCD + SWS + CHA), saving per unit, saving %, annual saving; add, edit or delete |
| Teknic Parts (Domestic) | Editable parts table; calculated saving per unit, saving %, saving per station, totals; stations per year; annual saving |
| Data Management | Upload Excel, restore JSON, download formatted Excel, clean Excel or JSON, clear session |

Status rules are the same as before:
- **Delayed**: the plan date has passed and there is no actual date.
- **Due in 7 days**: no actual date, and the plan date is within the next 7 days.
- **Done**: the activity has an actual date.
- Trigger rows and rows with neither a plan nor an actual date are left out of the progress count.

## Run on your PC (optional)

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (macOS/Linux: source .venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py
```

The app opens at http://localhost:8501.
For single-user local use, you can turn on automatic saving to a file:
`set NPD_AUTOSAVE_FILE=C:\NPD\npd_data.json` (Windows) before `streamlit run app.py`.
**Do not set this on Streamlit Cloud.** That file would be shared by every visitor.

## Upload to GitHub (repository `hemanthv83/NPD`)

1. Open https://github.com/hemanthv83/NPD.
2. **Delete the old files**: `index.html`, `app.js` and `xlsx.full.min.js`.
   For each one, open the file, click **⋯ → Delete file**, then **Commit changes**.
3. **Add file → Upload files** and drag in:
   - `app.py`
   - `requirements.txt`
   - `README.md` (replaces the old one)
   - `.gitignore`
4. **Upload the `.streamlit/config.toml` file** (it sets the theme colours and upload size):
   **Add file → Create new file**, type the name `.streamlit/config.toml`, paste its contents,
   then click **Commit changes**.
   Windows hides files that start with a dot. In Explorer, turn on **View → Show → Hidden items**
   to find `.gitignore` and the `.streamlit` folder.
5. Check that the repository root shows `app.py`, `requirements.txt`, `README.md`, `.gitignore`
   and the `.streamlit` folder.

## Deploy on Streamlit Community Cloud

1. Go to https://share.streamlit.io and **sign in with GitHub** (account `hemanthv83`).
   Allow Streamlit to access your repositories.
2. Click **Create app → Deploy a public app from GitHub** (wording may vary slightly).
3. **Repository**: `hemanthv83/NPD` · **Branch**: `main` · **Main file path**: `app.py`.
4. Optional: under **Advanced settings → Python version**, choose **3.12** or **3.13**.
5. Click **Deploy**. The first build takes 2–5 minutes while it installs `requirements.txt`.
6. Your app URL looks like `https://<name>.streamlit.app`.
7. **Restrict access**: in the app's **⋮ → Settings → Sharing**, set who can view it, for example
   only invited email addresses. Do this before you load real supplier data.

GitHub Pages is no longer needed. You can turn it off under **Settings → Pages → Source: None**.

## Updating the app

Edit `app.py` on GitHub and commit. Streamlit Cloud redeploys automatically within about a minute.
Redeploying clears the app's memory, so download your data first.
