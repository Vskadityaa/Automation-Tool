# Push this project to GitHub

## Option 1: Using Git (command line)

### 1. Install Git
- Download: **https://git-scm.com/download/win**
- Run the installer (use default options).
- Close and reopen your terminal/Command Prompt after installing.

### 2. Open Command Prompt in this folder
- Press `Win + R`, type `cmd`, press Enter.
- Run:
  ```bat
  cd /d "c:\Users\Admin\OneDrive\Desktop\bms_point_tool_FINAL"
  ```

### 3. Initialize and commit
```bat
git init
git add .
git commit -m "BMS Point Tool - initial commit"
```

### 4. Create a new repo on GitHub
- Go to **https://github.com** and sign in.
- Click **+** (top right) → **New repository**.
- **Repository name:** e.g. `bms-point-tool`
- Leave "Add a README" **unchecked** (you already have files).
- Click **Create repository**.

### 5. Push from your PC
On GitHub, after creating the repo, you’ll see commands. Use these (replace `YOUR_USERNAME` and `bms-point-tool` with your GitHub username and repo name):

```bat
git remote add origin https://github.com/YOUR_USERNAME/bms-point-tool.git
git branch -M main
git push -u origin main
```

When asked for password, use a **Personal Access Token** (GitHub no longer accepts account password):
- GitHub → **Settings** → **Developer settings** → **Personal access tokens** → **Generate new token** → give it repo scope → copy the token and paste it when Git asks for password.

---

## Option 2: Using GitHub Desktop

1. Install **GitHub Desktop**: https://desktop.github.com
2. Open it → **File** → **Add local repository** → choose the folder `bms_point_tool_FINAL`.
3. If it says "not a Git repository", click **create a repository** and choose this folder.
4. Write a summary (e.g. "Initial commit") and click **Commit to main**.
5. **Publish repository** (choose your GitHub account and a name like `bms-point-tool`).

---

## Option 3: Upload folder on GitHub (no Git)

1. Go to **https://github.com** → Sign in.
2. Click **+** → **New repository**.
3. Name it (e.g. `bms-point-tool`) → **Create repository**.
4. Click **uploading an existing file**.
5. Drag and drop the **contents** of `bms_point_tool_FINAL` (all files and folders **except** `.venv`).
6. Do **not** upload the `.venv` folder (it’s large and not needed for deployment).
7. Add commit message → **Commit changes**.

---

After the code is on GitHub, go to **https://render.com** → **New** → **Web Service** → connect the repo and deploy (see **DEPLOY.md**).
