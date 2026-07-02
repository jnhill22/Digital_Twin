# FacilityTwin

Browser-based facility digital twin. Import DXF floor plans, manage equipment with BOMs, route utility connections (HVAC, chilled water, power, compressed air), and version snapshots to GitHub.

## Run locally (fastest)

```bash
# 1. Make sure you have Node.js 18+ installed
node --version

# 2. Install dependencies
npm install

# 3. Start the dev server
npm run dev
```

Opens at **http://localhost:3000**. GitHub push/pull works immediately from here.

## Deploy to Vercel (free, public URL)

1. Push this folder to a GitHub repo
2. Go to [vercel.com](https://vercel.com) → **Add New Project**
3. Import your GitHub repo
4. Vercel auto-detects Vite — click **Deploy**
5. Done. You get a URL like `facility-twin.vercel.app`

## Deploy to Netlify (free, public URL)

1. Push this folder to a GitHub repo
2. Go to [app.netlify.com](https://app.netlify.com) → **Add new site → Import from Git**
3. Pick your repo
4. Build command: `npm run build`
5. Publish directory: `dist`
6. Click **Deploy site**

## Project structure

```
facility-twin/
├── index.html          # Entry HTML
├── package.json        # Dependencies
├── vite.config.js      # Vite config
├── .gitignore
└── src/
    ├── main.jsx        # React mount
    └── App.jsx         # Full application (1260 lines)
```

## GitHub integration

The app has built-in GitHub sync. In the app, click **⎇ GITHUB** in the toolbar and enter:

- **Owner**: your GitHub username
- **Repo**: the repo name
- **Branch**: `main` (default)
- **Path**: where the project JSON lives in the repo (default: `digital-twin/project.json`)
- **Token**: a fine-grained PAT with **Contents: Read and write** on the repo

The token stays in your browser's localStorage. It's stripped from committed files and exports.
