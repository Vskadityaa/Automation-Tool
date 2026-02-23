# Deploy BMS Point Tool – open on another PC

## Open on another PC – two ways

| Goal | What to do |
|------|------------|
| **Hosted online (anyone, any network)** | Deploy to **Render.com** (below). You get a URL like `https://bms-point-tool-xxx.onrender.com`. Share it – works from any PC. |
| **Run on your PC, share a link** | Double‑click **Host for other PCs.bat**. One-time: get a free token at [ngrok.com](https://ngrok.com), run `set NGROK_AUTHTOKEN=your_token`, then run the bat again. Share the link it prints (on other PC, click "Visit Site" if ngrok asks). |

---

## Host online (recommended – no local PC, no firewall)

Deploy the app to a cloud so **anyone can open it with a link** – no local server, no ngrok, no same network.

### Option A: Render.com (free tier)

1. Push this project to a **GitHub** (or GitLab) repository.
2. Go to [render.com](https://render.com) → Sign up → **New** → **Web Service**.
3. Connect your repo and select this project.
4. **Build command:** `pip install -r requirements.txt`  
   **Start command:** `gunicorn app:app --bind 0.0.0.0:$PORT`
5. Click **Create Web Service**. Render will build and run the app.
6. When it’s live, you get a URL like `https://bms-point-tool-xxxx.onrender.com`. Share that link – it works from any device.

**Or use the blueprint:** If your repo has `render.yaml`, use **New** → **Blueprint** and select the repo; Render will read the config.

### Option B: Railway

1. Push the project to GitHub and go to [railway.app](https://railway.app).
2. **New Project** → **Deploy from GitHub** → select the repo.
3. Railway detects the app. Set **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT`.
4. Deploy. Use the generated public URL.

### Option C: Other platforms

- **PythonAnywhere:** Upload the project, set a Web app with WSGI pointing to `app:app`, and use a production server (e.g. gunicorn) in the start command.
- **Heroku:** Use the included `Procfile`: `web: gunicorn app:app --bind 0.0.0.0:$PORT`.
- **VPS (Linux):** Install Python, run `pip install -r requirements.txt` and `gunicorn app:app --bind 0.0.0.0:8000`. Put Nginx/Caddy in front if you want HTTPS.

### If you see "Not found" (404)

- Open **`https://your-app-url/health`** – if you see "ok", the app is running; then try **`https://your-app-url/`** (root). The main page is at `/`, not `/step1`.
- Ensure the **Start command** is exactly: `gunicorn app:app --bind 0.0.0.0:$PORT` (Render sets `$PORT` automatically).
- On Render, check the **Logs** tab for startup errors (e.g. missing dependency or import error).

### When the app is hosted online

- The **/link** page shows only the **App URL** (the public link). No localhost, no “other PC”, no firewall/ngrok instructions.
- Set **SECRET_KEY** in the platform’s environment variables (e.g. a long random string) for session security.
- Uploaded files (Excel, SVG) are stored on the server’s disk; on free tiers the filesystem may reset on deploy. For permanent storage you’d add a database or object storage later.

---

## Run on your own PC (local / same network)

If you prefer to run the app on your machine:

1. **Install Python 3.8+** and run: `pip install -r requirements.txt`
2. **Start:** double‑click `start_server.bat` or run `python app.py`
3. **This PC:** open `http://localhost:6001` (or the port in `PORT.txt`)
4. **Other PCs (same WiFi/LAN):** open `http://<THIS_PC_IP>:6001`. If it doesn’t open, run `allow_firewall.bat` as Administrator once.
5. **Share from any network (open on any PC):** Get a free token at [ngrok.com](https://ngrok.com), then run `set NGROK_AUTHTOKEN=your_token` in CMD, and double‑click **Host for other PCs.bat**. Share the link it prints – it works from any PC (if you see an ngrok page, click "Visit Site").
