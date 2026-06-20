---
description: Compile ReminderMailAgua to a Windows .exe using PyInstaller. Uses reminderagua.ico, no console window, bundles locales/ and config.json. Use when asked to "compilar", "generar el exe", or "build".
---

## Compile to EXE Skill

Build `dist/reminderagua.exe` — self-contained Windows executable with no console window.

### Prerequisites
```bash
pip install pyinstaller pywin32
```

### Build command
```bash
pyinstaller reminderagua.spec --clean
```
- `--clean` ensures no stale cached artifacts cause issues
- Output: `dist/reminderagua.exe`

### What the spec bundles
| Source | Destination in exe | Purpose |
|--------|--------------------|---------|
| `locales/` | `locales/` | UI translations |
| `config.json` | `.` | Default configuration |
| `reminderagua.ico` | `.` | App icon |

### After compilation
1. Test the exe **from the dist/ folder** (not from build/)
2. Verify it opens without console window
3. Verify it connects to Outlook and loads config
4. The `config.json` in dist/ is the default — user data is written next to the exe

### Common PyInstaller errors for this project

| Error | Fix |
|-------|-----|
| Missing win32com | Add to hiddenimports in spec |
| Missing pywintypes | Add `pywintypes` to hiddenimports |
| Locales not found | Verify datas tuple: `('locales', 'locales')` |
| `sys._MEIPASS` error | Check i18n.py frozen path handling |

### Distributing
- Share only the `dist/` folder contents
- Users need **Outlook 365 installed and configured**
- No Python installation required (runtime is bundled)
- Schedule via Windows Task Scheduler (see matarreminder.xml as reference)

### Spec file location
`reminderagua.spec` — at project root. Edit this file if you:
- Change the entry point
- Add new data files to bundle
- Need to add new hiddenimports
