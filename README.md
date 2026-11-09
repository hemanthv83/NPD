# NPD Sourcing Tracker

Browser-based tracker for E&E BOP alternate development, parts MPN/price coverage and cost-saving comparisons.

**This repository contains no tracker data.** Data stays on each user's PC / OneDrive:

- **Load Excel** – reads `NPD_Master_Tracker_Updated.xlsx` (sheets: MQIS AND SINGLE PHASE PART, 6DS PROJECT PARTS, ALTERNATE DEVELOPNMENT, DC Fan Comparison, Teknic Parts (Domestic)).
- **Link data file** – saves every change to `NPD_Tracker_Data.json` on your PC/OneDrive (Chrome or Edge).
- **Backup / Restore** – JSON backup of all data.
- **Export Excel** – downloads current data as .xlsx.

Do **not** commit `.xlsx` or `.json` data files to this repository (blocked by `.gitignore`).

## Files
| File | Purpose |
|---|---|
| `index.html` | Page layout and styles |
| `app.js` | All logic (add/edit, save, import/export) |
| `xlsx.full.min.js` | SheetJS 0.18.5 (Apache-2.0) for reading/writing Excel |
| `.nojekyll`, `robots.txt` | GitHub Pages settings; keep out of search engines |
