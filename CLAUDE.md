# ReminderMailAgua — Claude Code Guide

## What this project does
Sends water bill payment reminder emails to a configurable list of recipients via Outlook 365.
Windows-only (uses Outlook COM interface). GUI built with tkinter.

## Quick Start
```bash
pip install -r requirements.txt
python main.py
```

## Compile to EXE
```bash
pyinstaller reminderagua.spec --clean
# Output: dist/reminderagua.exe
```

---

## Architecture

```
main.py                         # Entry point: sets BASE_PATH, logging, launches app
src/
  backend/
    config_manager.py           # Reads/writes config.json (ConfigManager class)
    email_service.py            # Outlook COM send — handles Hotmail fix
    date_utils.py               # Resolves [Mes anterior en letras] / [año en numero]
  frontend/
    app.py                      # ReminderApp class — all tkinter GUI code
    i18n.py                     # I18n class — loads locales/{lang}.json
locales/
  es.json                       # Spanish UI strings (default)
  en.json                       # English UI strings
config.json                     # User configuration (persisted at runtime)
reminderagua.ico                # App icon
reminderagua.spec               # PyInstaller build spec
docs/
  SDD.md                        # Software Design Document
.claude/
  agents/
    email-debugger.md           # Agent: diagnose Outlook/Hotmail issues
    code-improver.md            # Agent: safe refactoring process
  skills/
    github-push.md              # Skill: commit & push to erickson558/ReminderMailAgua
    comment-code.md             # Skill: add meaningful comments to code
    improve-code.md             # Skill: senior engineer improvement process
    compile-exe.md              # Skill: build Windows executable
```

---

## Key Behaviors
| Behavior | Details |
|----------|---------|
| **Placeholders** | `[Mes anterior en letras]` → previous month in Spanish; `[año en numero]` → year |
| **Account filtering** | Sender's email is auto-excluded from the recipient list |
| **Auto-close** | Optional countdown after successful send (configurable seconds) |
| **Multi-language** | ES/EN, persisted as `"language"` key in config.json |
| **Threading** | Email send runs in daemon thread — GUI never freezes |
| **Beer button** | Opens PayPal donation link in default browser |

---

## config.json Keys
| Key | Type | Description |
|-----|------|-------------|
| `destinatarios` | `list[str]` | Recipient email addresses |
| `asunto` | `str` | Email subject template (supports placeholders) |
| `cuerpo` | `str` | Email body template (supports placeholders) |
| `auto_close` | `bool` | Enable auto-close after successful send |
| `auto_close_delay` | `int` | Seconds before auto-close (default: 60) |
| `cuenta_seleccionada` | `str` | Last selected sender account (SMTP) |
| `language` | `str` | UI language: `"es"` or `"en"` |

---

## Skills Available
| Skill | Trigger phrase |
|-------|---------------|
| `/github-push` | "subir a github", "commit y push" |
| `/comment-code` | "comentar el código", "agregar comentarios" |
| `/improve-code` | "mejorar el código", "refactorizar" |
| `/compile-exe` | "compilar", "generar el exe" |

## Agents Available
| Agent | Use case |
|-------|---------|
| `email-debugger` | App fails to send from Hotmail/Outlook.com |
| `code-improver` | Needs careful refactoring with analysis-first process |

---

## Known Bugs Fixed (v1 → v2)
| Bug | Root Cause | Fix Applied |
|-----|-----------|-------------|
| Auto-send on launch | `root.after(1000, btn_enviar.invoke)` in v1 | Line removed in v2 |
| GUI freezes on send | Synchronous COM call on main thread | Background `threading.Thread` |
| Hotmail send fails | `SendUsingAccount` incompatible with HTTP account type | Added `SentOnBehalfOfName` fallback |
| Generic error messages | Bare `except` with unhelpful text | Specific error messages per error type |

---

## GitHub Account
- Username: **erickson558**
- Protocol: HTTPS (authenticated via keyring — no password needed)
- Repo: `https://github.com/erickson558/ReminderMailAgua`
- Token scopes: `repo`, `workflow`, `read:org`, `gist`

---

## SDD Update Workflow (Spec Driven Development)
After any significant change:
1. Update `docs/SDD.md` with the design decision
2. Update this `CLAUDE.md` if behavior/files changed
3. Run `/github-push` to commit everything
4. Run `/compile-exe` to recompile the executable
