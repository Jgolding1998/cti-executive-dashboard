# CTI Executive Dashboard

Auto-updating sales and AR dashboard for CTI Gas.

## Features
- **Sales Tracking**: Yesterday, MTD, YTD with Product/Service breakdown
- **AR Aging**: Current, 1-30, 31-60, 61-90, 90+ days
- **Cash Flow Prediction**: 6-week forecast based on AR pipeline and collection patterns
- **Top Customers & Products**: MTD rankings

## Auto-Update
Dashboard updates automatically at 6am Central via GitHub Actions.

## View Dashboard
Visit: https://[username].github.io/cti-dashboard/

## Manual Update
```bash
npm install
node generate-dashboard.js
```

## Setup (for new deployment)
1. Fork this repo
2. Add repository secrets:
   - `SYTELINE_USERNAME`: SyteLine API username
   - `SYTELINE_PASSWORD`: SyteLine API password
3. Enable GitHub Pages (Settings → Pages → Source: main branch)
4. Dashboard will auto-update daily at 6am Central
