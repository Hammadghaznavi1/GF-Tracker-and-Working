# Karandaaz Treasury Dashboard (live)

- data/GF_Balance.xlsx   your treasury workbook (upload a new one to update the dashboard)
- web/                   the dashboard page (index.html) and its code (app.js)
- scripts/build_data.py  reads the Excel file and encrypts the data with the password
- .github/workflows/     automation that rebuilds and publishes the site on every upload

Setup: create a secret named DASHBOARD_PASSWORD (Settings > Secrets and variables > Actions),
then Settings > Pages > Source: GitHub Actions.
