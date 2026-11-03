# squats

Squats Challenge — a live leaderboard for friends and family.

- `index.html` is the whole page. It is hosted with GitHub Pages.
- Scores are stored in a Google Sheet; the page talks to its Google Apps Script web app.
- Power score = total squats ÷ days since you started (start day counts as day 1).

## Pages
- **Board** — log squats, six leaderboards (Power, Total, Today, This week, Streak, Most improved), recent updates.
- **Progress** — squats together per day, and everyone's running total or power over time (compare up to 4 people).
- **Players** — each person's stats, daily chart, running total and history.

`apps-script-Code.gs` is the Google Apps Script code (paste into Code.gs in the sheet's Apps Script project).
