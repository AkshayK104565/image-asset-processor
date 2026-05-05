# 🚀 Deploy Guide — Image Asset Processor

Follow these steps once. After that, users just visit your URL — no setup needed.

---

## What you need
- A **GitHub account** (free) → github.com
- A **Google account** (to set up login)
- A **Render account** (free) → render.com

Total time: ~20 minutes.

---

## STEP 1 — Push the code to GitHub

1. Go to **github.com** → click **"New repository"**
2. Name it `image-asset-processor` → click **"Create repository"**
3. On your computer, open a terminal / command prompt in this folder and run:

```
git init
git add .
git commit -m "initial deploy"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/image-asset-processor.git
git push -u origin main
```

*(Replace `YOUR_USERNAME` with your GitHub username)*

---

## STEP 2 — Create Google OAuth credentials

This lets users sign in with their Google account.

1. Go to **console.cloud.google.com**
2. Click **"Select a project"** → **"New Project"** → name it anything → **Create**
3. In the left menu: **APIs & Services → OAuth consent screen**
   - Choose **External** → **Create**
   - App name: `Image Asset Processor`
   - User support email: your email
   - Scroll down → **Save and Continue** (skip the rest, click through)
4. In the left menu: **APIs & Services → Credentials**
   - Click **"+ Create Credentials"** → **"OAuth client ID"**
   - Application type: **Web application**
   - Name: `Image Asset Processor`
   - Under **Authorised redirect URIs** → click **"+ Add URI"**
   - Enter: `https://YOUR-APP-NAME.onrender.com/auth/callback`
     *(you'll fill in the actual Render URL in Step 3 — come back and update this)*
   - Click **Create**
5. Copy the **Client ID** and **Client Secret** — you'll need them in Step 3.

---

## STEP 3 — Deploy to Render

1. Go to **render.com** → sign up with GitHub
2. Click **"New +"** → **"Web Service"**
3. Connect your `image-asset-processor` GitHub repo
4. Render auto-detects the `render.yaml` — click **"Apply"**
5. In the **Environment** tab, add these variables:

| Key | Value |
|-----|-------|
| `GOOGLE_CLIENT_ID` | (paste from Step 2) |
| `GOOGLE_CLIENT_SECRET` | (paste from Step 2) |
| `ALLOWED_EMAILS` | `you@gmail.com,colleague@gmail.com` (comma-separated, or leave **blank** to allow any Google account) |
| `SECRET_KEY` | click "Generate" |

6. Click **"Deploy"** — wait ~3 minutes for the build to finish.
7. Your app URL will be shown at the top, e.g. `https://image-asset-processor.onrender.com`

---

## STEP 4 — Update Google with your real URL

1. Go back to **console.cloud.google.com → APIs & Services → Credentials**
2. Click your OAuth client → edit the **Authorised redirect URIs**
3. Replace the placeholder with your real URL:
   `https://image-asset-processor.onrender.com/auth/callback`
4. Click **Save**

---

## ✅ Done!

Share the URL with your team. Users:
1. Visit the URL
2. Sign in with Google (one click)
3. Download the template (if needed)
4. Upload their filled Excel file
5. Set dimensions, click process
6. Download the ZIP

---

## Notes

**Free tier sleep:** Render free services pause after 15 minutes of inactivity.
The first visit after a pause takes ~30 seconds to wake up.
To avoid this, upgrade to Render's **Starter plan ($7/month)**.

**Restricting access:** Set `ALLOWED_EMAILS` to a comma-separated list of
allowed email addresses. Leave it blank to allow any Google account.

**Updating the app:** Push changes to GitHub → Render auto-redeploys.
