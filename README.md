# NN-ON Dashboard — Deployment Guide

## What you have
- `server.py` — Python backend (fetches your SharePoint Excel every 60s)
- `requirements.txt` — Python dependencies
- `render.yaml` — Render.com config
- `index.html` — Dashboard (auto-updates from server)

---

## STEP 1 — Push to GitHub (2 minutes)

1. Go to github.com → New repository → name it `nn-dashboard` → Create
2. Upload all 4 files (drag & drop into GitHub)

---

## STEP 2 — Deploy server on Render.com (3 minutes)

1. Go to https://render.com → Sign up free (use Google login)
2. Click **New** → **Web Service**
3. Connect your GitHub → select `nn-dashboard` repo
4. Render auto-detects `render.yaml` — just click **Deploy**
5. Wait ~2 minutes for build to finish
6. Copy your Render URL — looks like: `https://nn-dashboard-api.onrender.com`

### ⚠ Important: Update the SharePoint URL in render.yaml
In render.yaml, the SHAREPOINT_URL is already set.
If your SharePoint link changes, update it in Render dashboard:
  Settings → Environment → SHAREPOINT_URL → paste new link

---

## STEP 3 — Update dashboard with your Render URL (1 minute)

Open `index.html` and find this line near the top of the `<script>` section:

```javascript
const RENDER_URL = window.RENDER_URL || 'YOUR_RENDER_URL_HERE';
```

Replace `YOUR_RENDER_URL_HERE` with your actual Render URL:

```javascript
const RENDER_URL = window.RENDER_URL || 'https://nn-dashboard-api.onrender.com';
```

Save the file.

---

## STEP 4 — Deploy dashboard on Netlify (1 minute)

1. Go to https://netlify.com → Sign up free
2. Click **Add new site** → **Deploy manually**
3. Drag & drop just the `index.html` file
4. Netlify gives you a URL like: `https://nn-deals.netlify.app`
5. Share this URL with everyone — that's your live dashboard link ✅

---

## How it works after setup

```
Your SharePoint Excel (editors update it anytime)
         ↓  every 60 seconds
Render server fetches latest data
         ↓  every 60 seconds
Dashboard auto-refreshes for all viewers
```

- Green dot = live and connected
- Orange dot = connecting / refreshing  
- Red dot = server error (check Render logs)

---

## Updating the SharePoint file link

If the SharePoint sharing link changes:
1. Go to Render dashboard → your service → Environment
2. Update SHAREPOINT_URL value
3. Click Save → server redeploys automatically

---

## Free tier limits
- **Render free tier**: Server sleeps after 15 min of inactivity
  → First load after sleep takes ~30 seconds (one-time delay)
  → Upgrade to Render Starter ($7/month) to keep it always awake
- **Netlify free tier**: 100GB bandwidth/month — more than enough
