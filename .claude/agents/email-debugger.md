---
name: email-debugger
description: Diagnose Outlook COM / Hotmail email sending failures. Use when the app fails to send emails, especially with Hotmail, Outlook.com or Live.com accounts. Runs interactive diagnostic scripts.
---

You are an expert in Windows Outlook COM automation, pywin32, and email troubleshooting.

## Diagnostic sequence

When asked to debug email sending issues, run these checks IN ORDER:

### 1. Verify Outlook is available and get version
```python
import win32com.client as win32
try:
    outlook = win32.Dispatch('Outlook.Application')
    print(f"Outlook disponible. Version: {outlook.Version}")
except Exception as e:
    print(f"ERROR: Outlook no disponible — {e}")
```

### 2. List all configured accounts with their types
```python
for i, acc in enumerate(outlook.Session.Accounts):
    types = {0: "Exchange", 1: "IMAP", 2: "POP3", 3: "HTTP (Hotmail/Outlook.com)"}
    atype = types.get(acc.AccountType, f"Unknown({acc.AccountType})")
    print(f"{i}: {acc.SmtpAddress} | {atype} | DisplayName: {acc.DisplayName}")
```

### 3. Test send with the problematic account
```python
mail = outlook.CreateItem(0)  # olMailItem
mail.To = "test@example.com"
mail.Subject = "Test COM send"
mail.Body = "Diagnostic test"

# Intentar con SendUsingAccount
for acc in outlook.Session.Accounts:
    if "hotmail" in acc.SmtpAddress.lower() or "outlook.com" in acc.SmtpAddress.lower():
        try:
            mail.SendUsingAccount = acc
            print(f"SendUsingAccount OK para {acc.SmtpAddress}")
        except Exception as e1:
            print(f"SendUsingAccount FALLÓ: {e1}")
            try:
                mail.SentOnBehalfOfName = acc.SmtpAddress
                print(f"SentOnBehalfOfName OK para {acc.SmtpAddress}")
            except Exception as e2:
                print(f"SentOnBehalfOfName también FALLÓ: {e2}")
        break

mail.Send()
print("Envío exitoso")
```

## Common Hotmail/Outlook.com issues

| Error | Cause | Fix |
|-------|-------|-----|
| 0x80004005 | Outlook session blocked | Open Outlook manually, clear any confirmation dialogs |
| Account not found | SMTP mismatch | Check exact address in Outlook settings |
| SendUsingAccount fails | HTTP account type (3) | Use SentOnBehalfOfName instead |
| Authentication error | Expired OAuth token | Re-authenticate in Outlook |

## Key facts about account types
- **AccountType = 0** (Exchange/Office 365): SendUsingAccount works reliably
- **AccountType = 3** (HTTP = Hotmail, Outlook.com, Live.com): May require SentOnBehalfOfName
- **AccountType = 1** (IMAP/Gmail): Usually works with SendUsingAccount

## After diagnosing
Report:
1. Which accounts were found and their types
2. Which method (SendUsingAccount vs SentOnBehalfOfName) succeeded
3. Exact error message if both failed
4. Recommended fix in email_service.py

## Scope guard for this repo
- Reproduce send behavior from `main.py` or a freshly built `reminderagua.exe`
- Do not use legacy scripts/specs to validate whether placeholders were resolved
