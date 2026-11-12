# PM Surya Ghar Report Maker (web page)

The browser version of the Report Maker. Choose the SBPDCL and NBPDCL portal exports, press
**Create report**. The page makes the same Excel report as the desktop app, saves it to the
Downloads folder, sends it to the dashboard and opens the dashboard.

- The CSV files are read on your computer. They are never uploaded; only the summary report
  (counts) goes to the dashboard.
- The upload password is asked once and kept only while the page is open.
- Works in Chrome or Edge on a desktop or laptop. The first run fetches the report engine
  (about 10 MB), which takes about a minute; later runs start at once.

## Files

| File | What it is |
|---|---|
| `index.html` | The page |
| `engine.js` | Runs the report engine in the background |
| `pmsg_report_app.py` | The report engine: the same file as the desktop app |
| `config.js` | Settings: the dashboard link |

## Put it on GitHub (one time, about 10 minutes)

1. Sign in at github.com. Press **+** (top right) > **New repository**.
   Name it, for example, `pmsg-report`. Choose **Public**. Press **Create repository**.
2. On the new repository page press **uploading an existing file**. Drag in the four files
   (`index.html`, `engine.js`, `pmsg_report_app.py`, `config.js`) and this `README.md`.
   Press **Commit changes**.
3. Open `config.js` in the repository, press the pencil icon, paste the dashboard's web app
   link between the quotes of `dashboardLink: ""` (it starts with `https://script.google.com/`
   and ends with `/exec`), then **Commit changes**.
4. **Settings** > **Pages**. Under "Build and deployment" set Source: **Deploy from a branch**,
   Branch: **main**, folder **/ (root)**. Press **Save**.
5. After a minute or two the page is live at `https://<your-github-name>.github.io/pmsg-report/`
   (Settings > Pages shows the exact link). Bookmark it.
6. Open the page, press **Test connection**, type the upload password. It should say
   "Connected".

Nothing changes on the dashboard side: it uses the same `Upload.gs` and password as the
desktop app.

## Later

- **New version of the engine:** upload the new `pmsg_report_app.py` (Add file > Upload files),
  then in `config.js` change `engineVersion` (for example to `"3.1"`) so browsers fetch the new one.
- **New dashboard link:** edit `config.js` again.
- **Who can see it:** anyone who has the page link can open it, because the repository is public.
  They cannot change the dashboard without the upload password. No data is kept on GitHub.
