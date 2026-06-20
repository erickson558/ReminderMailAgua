---
description: Commit and push changes to GitHub using the erickson558 account. Stages project files, creates a descriptive commit, and pushes to origin. Use when asked to "subir a GitHub", "commit y push", or at the end of a coding session.
---

## GitHub Push Skill

Commit and push current changes to GitHub (account: erickson558, authenticated via keyring).

### Step 1 — Check status and recent history
```bash
git status
git log --oneline -5
```

### Step 2 — Initialize repo if needed (first time only)
```bash
# Only run if .git directory does not exist
git init
git remote add origin https://github.com/erickson558/ReminderMailAgua.git
git branch -M main
```

### Step 3 — Stage project files (NEVER commit binaries or build artifacts)
```bash
git add main.py
git add src/
git add locales/
git add docs/
git add .claude/
git add config.json
git add requirements.txt
git add reminderagua.spec
git add CLAUDE.md
git add .gitignore
git add tests/
```

If the task included a rebuild, verify locally that both `dist/reminderagua.exe` and `./reminderagua.exe` were refreshed before pushing source changes. Do not stage those binaries.

### Step 4 — Commit with descriptive message
```bash
git commit -m "$(cat <<'EOF'
<summary of what changed and why>

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>
EOF
)"
```

### Step 5 — Push
```bash
git push -u origin main
```

### Files to NEVER commit
- `reminderagua.exe` (binary, 16MB+, generated)
- `dist_tmp/` directory (temporary build output)
- `build/` directory (PyInstaller intermediates)
- `dist/` directory (built executables)
- `*.log` (runtime logs)
- `.env` or any credentials
- `__pycache__/` or `*.pyc`

### Account details
- Username: erickson558
- Protocol: HTTPS
- Auth: keyring (already authenticated — no password needed)
- Token scopes: repo, workflow, read:org, gist

### After pushing
Report the commit hash and confirm the push succeeded:
```bash
git log --oneline -1
```
