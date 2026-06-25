# Bench Dashboard

An internal EPAM tool for managers to track bench associates — course completion, client proposal pipeline, and placement outcomes — all secured behind EPAM Microsoft SSO with data stored in OneDrive.

**Live URL:** https://saurabhsuman1-droid.github.io/bench-dashboard/
**Access:** EPAM Microsoft account required

## Documentation

| Document | Purpose |
|---|---|
| [User Guide](docs/USER_GUIDE.md) | How to use the app as a manager |
| [Architecture](docs/ARCHITECTURE.md) | Security model, data flow, component design |
| [Setup Guide](docs/SETUP.md) | Fresh environment setup for developers |

## Tech Stack

| Layer | Technology |
|---|---|
| Frontend | React 19, Vite 8, Tailwind CSS 3 |
| Auth | `@azure/msal-react` — Azure AD OAuth 2.0 + PKCE |
| Storage | Microsoft Graph API → OneDrive (`bench-data.json`) |
| Charts | Chart.js 4 (in-app analytics + PDF export) |
| PDF | jsPDF + jspdf-autotable |
| Hosting | GitHub Pages (static files only, zero server) |
| CI/CD | GitHub Actions (auto-deploy on push to `main`) |

## Quick Start

```bash
git clone https://github.com/saurabhsuman1-droid/bench-dashboard
cd bench-dashboard
npm install
cp .env.example .env        # fill in your Azure AD credentials
npm run dev                  # http://localhost:5173
```

See [Setup Guide](docs/SETUP.md) for Azure AD registration steps and GitHub Pages deployment.
