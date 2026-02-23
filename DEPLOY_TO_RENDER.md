# Deploy BMS Point Tool to Render.com

Follow these steps to get a live URL (e.g. `https://bms-point-tool-xxx.onrender.com`) that works from any PC.

---

## Step 1: Push the project to GitHub

If the project is **not** on GitHub yet:

1. Go to **https://github.com** and sign in (or create an account).
2. Click **+** (top right) → **New repository**.
3. Name it (e.g. `bms-point-tool`). Leave "Add a README" **unchecked**. Click **Create repository**.
4. On your PC, open **Command Prompt** in the project folder:
   ```bat
   cd /d "c:\Users\Admin\OneDrive\Desktop\bms_point_tool_FINAL"
   ```
5. If Git is installed, run:
   ```bat
   git init
   git add .
   git commit -m "BMS Point Tool"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/bms-point-tool.git
   git push -u origin main
   ```
   Replace `YOUR_USERNAME` and `bms-point-tool` with your GitHub username and repo name.

   **If you don't have Git:** On the new GitHub repo page, click **"uploading an existing file"**, then drag and drop all project files and folders **except** the `.venv` folder. Click **Commit changes**.

---

## Step 2: Create a Web Service on Render

1. Go to **https://render.com** and sign up or log in (you can use your GitHub account).
2. In the dashboard, click **New +** → **Web Service**.
3. Under **Connect a repository**, find and select your **bms-point-tool** repo (or the repo you used). Click **Connect** if asked.
4. Use these settings:

   | Field | Value |
   |-------|--------|
   | **Name** | `bms-point-tool` (or any name) |
   | **Region** | Choose nearest |
   | **Branch** | `main` |
   | **Runtime** | `Python 3` |
   | **Build Command** | `pip install -r requirements.txt` |
   | **Start Command** | `gunicorn app:app --bind 0.0.0.0:$PORT` |

5. Under **Advanced**, add an environment variable (optional but recommended):
   - **Key:** `SECRET_KEY`  
   - **Value:** any long random string (e.g. paste from https://randomkeygen.com/)

6. Click **Create Web Service**.

---

## Step 3: Wait for deploy

Render will build and start your app. This can take a few minutes. When the build finishes, the **Logs** tab should show no errors and the service status will be **Live**.

---

## Step 4: Open your app

At the top of the service page you’ll see a URL like:

**`https://bms-point-tool-xxxx.onrender.com`**

Click it or copy it and open it in any browser (on any PC). That is your deployed app. Share this link with anyone who should use the tool.

---

## Troubleshooting

- **Build fails:** Check the **Logs** tab. Often it’s a missing dependency – ensure `requirements.txt` lists all packages the app needs.
- **"Application failed to respond":** Make sure the **Start Command** is exactly:  
  `gunicorn app:app --bind 0.0.0.0:$PORT`
- **Not found (404):** Open `https://your-app-url/health` – if you see "ok", the app is running. Then open the root URL: `https://your-app-url/`

---

## Free tier note

On the free tier, the service may spin down after a short time of no use. The first visit after that may take 30–60 seconds to wake up. The app will then work normally.
