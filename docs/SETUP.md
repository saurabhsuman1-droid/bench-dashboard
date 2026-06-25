# Setup Guide — Bench Dashboard

This guide covers a complete fresh setup from zero: Azure AD registration, local development, and GitHub Pages deployment.

---

## Prerequisites

- **Node.js 20+** — https://nodejs.org (LTS version)
- **Git** — https://git-scm.com
- **EPAM Microsoft account** — needed for Azure AD App Registration and OneDrive data storage
- **GitHub account** — free account at github.com is sufficient

---

## Step 1 — Clone and Install

```bash
git clone https://github.com/saurabhsuman1-droid/bench-dashboard
cd bench-dashboard
npm install
```

---

## Step 2 — Azure AD App Registration

This registers the app with EPAM's Microsoft identity platform so users can sign in with their EPAM accounts.

1. Go to https://portal.azure.com
2. Search for **App registrations** → click **New registration**
3. Fill in:
   - **Name:** `BenchDashboard` (or any name)
   - **Supported account types:** Accounts in this organizational directory only (single tenant)
   - **Redirect URI:** Select **Single-page application (SPA)** → enter `http://localhost:5173`
4. Click **Register**
5. **Copy these two values** (you need them in Step 3):
   - **Application (client) ID** — shown on the Overview page
   - **Directory (tenant) ID** — shown on the Overview page

### Grant API Permission

6. Left sidebar → **API permissions** → **Add a permission**
7. **Microsoft Graph** → **Delegated permissions** → search **Files.ReadWrite** → check it → **Add permissions**
8. Click **Grant admin consent** (if you have admin rights; otherwise the user consents on first login)

---

## Step 3 — Local Environment File

```bash
cp .env.example .env
```

Edit `.env`:

```
VITE_AAD_CLIENT_ID=paste-your-application-client-id-here
VITE_AAD_TENANT_ID=paste-your-directory-tenant-id-here
```

`.env` is listed in `.gitignore` — it will never be committed.

---

## Step 4 — Run Locally

```bash
npm run dev
```

Open http://localhost:5173 in your browser.

Click **Sign in with Microsoft** → sign in with your EPAM account.

On first use, `BenchDashboard/bench-data.json` will be created automatically in your OneDrive the first time you add data.

---

## Step 5 — Deploy to GitHub Pages

### 5a — Push code to GitHub

On GitHub.com, create a new repository named `bench-dashboard` (can be private — GitHub Pages still serves the built files publicly).

```bash
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/bench-dashboard.git
git push -u origin main
```

### 5b — Add GitHub Secrets

In your GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**

Add these two secrets:

| Name | Value |
|---|---|
| `VITE_AAD_CLIENT_ID` | Your Application (client) ID from Step 2 |
| `VITE_AAD_TENANT_ID` | Your Directory (tenant) ID from Step 2 |

### 5c — Enable GitHub Pages

Repo → **Settings → Pages**
- Source: **GitHub Actions**

The `.github/workflows/deploy.yml` file (already in the repo) handles the rest automatically.

### 5d — Trigger first deployment

Push any change or go to **Actions tab → Deploy to GitHub Pages → Run workflow**.

After ~2 minutes, your app is live at:
`https://YOUR_USERNAME.github.io/bench-dashboard/`

---

## Step 6 — Register GitHub Pages URL in Azure AD

The redirect URI must be registered or Microsoft will reject the login.

1. `portal.azure.com` → **App registrations → BenchDashboard → Authentication**
2. Under **Single-page application** → **Add URI**
3. Add: `https://YOUR_USERNAME.github.io/bench-dashboard/`
4. Click **Save**

---

## Step 7 (Optional) — Pin as Teams Tab

1. Open Microsoft Teams → go to the bench associates channel
2. Click **+** (Add tab) at the top
3. Search for **Website** → paste the GitHub Pages URL
4. Name it **Bench Dashboard** → Save

The app is now accessible from Teams without opening a separate browser tab.

---

## Ongoing Deployment

Every push to the `main` branch triggers an automatic redeploy via GitHub Actions. No manual steps needed after initial setup.

To update:

```bash
# make changes locally, test with npm run dev, then:
git add .
git commit -m "description of change"
git push
```

GitHub Actions deploys automatically within ~2 minutes.

---

## Data Ownership & Handover

All bench data is stored in the **manager's OneDrive** at `BenchDashboard/bench-data.json`.

If the manager changes, before their EPAM account is deactivated:
1. Download `BenchDashboard/bench-data.json` from OneDrive
2. Upload it to the new manager's OneDrive under the same path `BenchDashboard/bench-data.json`
3. Update the Azure AD App Registration (or create a new one under the new manager's account)

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `AADSTS50011: redirect URI mismatch` | GitHub Pages URL not registered in Azure AD | Add it in App Registration → Authentication |
| Blank page on GitHub Pages | `base` not set to `'./'` in vite.config.js | Ensure `base: './'` is in vite.config.js |
| `sync failed` in header | Graph API returned error | Check browser DevTools Network tab; re-login if token expired |
| Login popup opens and stays open | Don't use `loginPopup` | Use `loginRedirect` (already configured) |
| Data not showing after login | `skipSave` or `cloudReady` ref issue | Refresh browser; if persistent, check App.jsx load effect |
