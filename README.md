# Executive Apartment Portfolio Dashboard

A premium, static executive-analytics website presenting the KPIs and insights
from the **Executive Apartment Portfolio Reporting System** (Excel + Power Query
+ Power Pivot). Built with HTML, CSS, JavaScript and Chart.js. No backend.

## Structure
```
portfolio-dashboard/
├── index.html
├── css/
│   └── styles.css
└── js/
    └── script.js
```

## Run locally
The page loads Chart.js from a CDN and embeds all data in `script.js`,
so you can open it directly:

1. **Quick way** — double-click `index.html` (opens in your browser).

2. **Recommended (local server)** — from inside the `portfolio-dashboard` folder:
   ```bash
   python3 -m http.server 8000
   ```
   then visit **http://localhost:8000**.
   (Or use the VS Code "Live Server" extension.)

## Deploy to GitHub Pages
1. Create a new GitHub repo and push the contents of `portfolio-dashboard/`
   to the repository root (so `index.html` is at the top level):
   ```bash
   git init
   git add .
   git commit -m "Executive portfolio dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```
2. In the repo: **Settings → Pages**.
3. Under **Build and deployment → Source**, choose **Deploy from a branch**.
4. Select branch **main** and folder **/ (root)**, then **Save**.
5. Your site goes live at `https://<you>.github.io/<repo>/` within a minute.

## Editing data
All numbers live in the `DATA` object at the top of `js/script.js`
(`months`, `occTrend`, `noiTrend`, `revActual`, `revBudget`, `portfolio`,
and the `properties` array). Update those values and the KPI cards, charts,
and exception table re-render automatically.
