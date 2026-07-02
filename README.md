# SAP Profile Analyzer

## Run locally
```
npm install
npm run dev
```
Opens at http://localhost:5173

## Build for production
```
npm install
npm run build
```
This produces a `dist/` folder containing static files (HTML/JS/CSS) — that folder is the entire app.

## Deploy
Any static host works, since this has no backend. Easiest options:

- **Netlify / Vercel**: drag-and-drop the `dist/` folder after `npm run build` onto their dashboard, or connect this folder as a GitHub repo for auto-deploys on every push. Build command: `npm run build`, publish directory: `dist`.
- **Cloudflare Pages**: same idea — connect the repo, build command `npm run build`, output directory `dist`.
- **GitHub Pages**: push `dist/` contents to a `gh-pages` branch (or use the `gh-pages` npm package to automate it).

No environment variables, API keys, or backend services are required — everything (parsing, comparison, Excel/PDF export, session save/load) runs entirely in the browser.
