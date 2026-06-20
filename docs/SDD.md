# Software Design Document — ReminderMailAgua

**Version:** 2.0
**Date:** 2026-06-19
**Author:** erickson558

---

## 1. Project Overview

**Name:** ReminderMailAgua
**Purpose:** Send water bill payment reminder emails to a configurable list of recipients using the user's Outlook 365 account on Windows.
**Platform:** Windows 10/11 (requires Microsoft Outlook 365 installed and configured)
**Language:** Python 3.10+

---

## 2. Functional Requirements (Spec)

| ID | Requirement | Status |
|----|-------------|--------|
| F-01 | Send email to a configurable list of recipients via Outlook 365 | ✓ |
| F-02 | Allow adding and removing recipients from the GUI | ✓ |
| F-03 | Support selecting which Outlook account to send from | ✓ |
| F-04 | Replace `[Mes anterior en letras]` with previous month name in Spanish | ✓ |
| F-05 | Replace `[año en numero]` with the year of the previous month | ✓ |
| F-06 | Persist configuration to config.json | ✓ |
| F-07 | Auto-close application after send with configurable countdown | ✓ |
| F-08 | Support Hotmail/Outlook.com personal accounts | ✓ (v2 fix) |
| F-09 | GUI must not freeze during email sending | ✓ (v2 threading) |
| F-10 | Support Spanish and English UI languages | ✓ (v2 i18n) |
| F-11 | Beer donation button linking to PayPal | ✓ (v2) |
| F-12 | Compile to standalone Windows .exe (no console) | ✓ |

---

## 3. Architecture

### 3.1 Separation of Concerns

| Layer | Module | Responsibility |
|-------|--------|----------------|
| Entry Point | `main.py` | Bootstrap: BASE_PATH, logging, launch |
| Frontend | `src/frontend/app.py` | All tkinter widget code, event handlers |
| i18n | `src/frontend/i18n.py` | Load and serve translated strings |
| Config | `src/backend/config_manager.py` | Read/write config.json, defaults |
| Email | `src/backend/email_service.py` | Outlook COM integration, Hotmail fix |
| Dates | `src/backend/date_utils.py` | Placeholder token resolution |

**Rule:** `src/backend/` modules have **zero tkinter dependencies**. They are pure Python and testable without a GUI.

### 3.2 Data Flow

```
User clicks "Send"
  → app.py validates recipients
  → date_utils.py resolves placeholders in subject/body
  → threading.Thread starts → email_service.py sends via Outlook COM
  → root.after(0, ...) updates status label (thread-safe)
  → [if auto_close] app.py starts countdown → root.destroy()
```

### 3.3 Configuration Flow

```
On startup:
  main.py sets BASE_PATH
  → ConfigManager.__init__(BASE_PATH) reads config.json
  → I18n.__init__(BASE_PATH, language) loads locales/{lang}.json
  → ReminderApp._load_config_into_ui() populates widgets

On save:
  ReminderApp._save_config() collects widget values
  → ConfigManager.save(data) writes config.json
```

---

## 4. Email Sending Strategy

### 4.1 Outlook COM

The application uses `win32com.client.Dispatch('Outlook.Application')` to interface with the locally installed Outlook 365.

Sending uses `mail.CreateItem(0)` (olMailItem = type 0).

### 4.2 Hotmail/Outlook.com Fix (v2)

Personal Microsoft accounts (hotmail.com, outlook.com, live.com, msn.com) have `AccountType = 3` (HTTP) in Outlook's account model.

`SendUsingAccount` may fail with a COM error for these accounts.

**Fix applied in v2:**
```python
try:
    mail.SendUsingAccount = matched_account   # Works for Exchange / Office 365
except Exception:
    mail.SentOnBehalfOfName = sender_account  # Fallback for Hotmail / Outlook.com
```

### 4.3 Recipient Filtering

The sender's own SMTP address is automatically removed from the recipient list to prevent self-sends. Comparison is case-insensitive after `.strip().lower()`.

### 4.4 Threading

All Outlook COM calls run in a `daemon=True` background thread.
Status bar updates from the thread use `root.after(0, callback)` — the only thread-safe way to modify tkinter widgets.

---

## 5. Internationalization

### 5.1 Structure

```
locales/
  es.json    # Spanish (default)
  en.json    # English
```

### 5.2 Adding a New Language

1. Copy `locales/es.json` → `locales/{lang_code}.json`
2. Translate all **values** (never the keys)
3. Add `lang_code` to `SUPPORTED_LANGUAGES` in `src/frontend/i18n.py`
4. Add menu item in `app.py._build_menu()`

### 5.3 Language Persistence

Language preference is stored as `"language"` key in config.json.
On language change, the app saves config and recreates the ReminderApp with the new I18n instance (simple restart approach vs. complex StringVar binding).

---

## 6. Configuration

### 6.1 File Location

`config.json` is located **next to the executable** (or next to `main.py` in script mode).
For local development builds, if the executable is launched from `dist/` and no
`dist/config.json` exists yet, the app reuses the repository-root `config.json`
instead of creating a second divergent config file.

Frozen detection:
```python
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))
```

### 6.2 Backward Compatibility

v2 config.json adds two new keys: `cuenta_seleccionada` and `language`.
Existing v1 config.json files are fully compatible — missing keys get defaults via `config.setdefault()` pattern.

---

## 7. Compilation

### 7.1 Tool
PyInstaller 6.x, configured in `reminderagua.spec`.

### 7.2 Bundled Resources
| Resource | Why bundled |
|----------|------------|
| `locales/` | Runtime i18n strings — needed at `sys._MEIPASS/locales` |
| `config.json` | Default config for first run |
| `reminderagua.ico` | App icon |

### 7.3 PyInstaller Path Handling
When frozen (one-file exe), PyInstaller extracts bundled data to `sys._MEIPASS`.
- **Bundled read-only resources** (locales): read from `sys._MEIPASS`
- **User-writable data** (config.json, logs): read/written next to `sys.executable`
- **Development fallback:** if `sys.executable` lives in `dist/` inside the repo and
  `dist/config.json` is missing, config reads/writes use the repo-root `config.json`
  to keep manual edits and test runs aligned

### 7.4 Hidden Imports
win32com is not auto-detected by PyInstaller; added explicitly:
`win32com.client`, `win32api`, `win32con`, `pywintypes`, `pythoncom`

---

## 8. Bugs Fixed (v1 → v2)

| ID | Bug | Root Cause | Fix |
|----|-----|-----------|-----|
| B-01 | Email sent automatically 1s after launch | `root.after(1000, btn_enviar.invoke)` line 234 in v1 | Removed entirely |
| B-02 | GUI freezes while sending | COM call on tkinter main thread | Background thread |
| B-03 | Hotmail account fails silently | `SendUsingAccount` incompatible with AccountType=3 | `SentOnBehalfOfName` fallback |
| B-04 | Unhelpful error messages | Bare `except Exception as e: update_status(str(e))` | Specific error messages per COM error code |

---

## 9. Deployment

1. Run `pyinstaller reminderagua.spec --clean`
2. Test `dist/reminderagua.exe` from the dist folder
3. Distribute the `dist/` folder
4. Users need: Windows 10/11 + Outlook 365 installed and configured
5. Optionally schedule via Windows Task Scheduler using `matarreminder.xml` as reference template

---

## 10. Decisions Log

| Date | Decision | Reason |
|------|----------|--------|
| 2026-06-19 | Class-based ReminderApp instead of module-level functions | Allows proper lifecycle management, easier testing, no global state |
| 2026-06-19 | Restart approach for language change | Simpler than binding every label to a StringVar; avoids complex widget rebuild |
| 2026-06-19 | daemon=True on send thread | Ensures process exits cleanly if user closes window while sending |
| 2026-06-19 | SentOnBehalfOfName as Hotmail fallback | Only documented working alternative to SendUsingAccount for HTTP accounts |
| 2026-06-20 | Reuse repo-root config from dist/ during local builds | Prevents testing the exe against a stale or newly created second config.json |
