---
description: Compile ReminderMailAgua to a Windows .exe using PyInstaller. Uses reminderagua.ico, no console window, bundles locales/ and config.json, and refreshes both dist/reminderagua.exe and the root reminderagua.exe beside main.py. Use when asked to "compilar", "generar el exe", or "build".
---

## Compile to EXE Skill

Build a fresh executable from `main.py` and keep two local copies:
- `dist/reminderagua.exe`
- `./reminderagua.exe` beside `main.py`

### Prerequisites
```bash
pip install pyinstaller pywin32
```

### Build command
```bash
pyinstaller reminderagua.spec --noconfirm --workpath build_tmp --distpath dist_tmp
New-Item -ItemType Directory -Force dist | Out-Null
Copy-Item dist_tmp\reminderagua.exe dist\reminderagua.exe -Force
Copy-Item dist_tmp\reminderagua.exe .\reminderagua.exe -Force
```
- Uses alternate work/dist folders to avoid OneDrive lock issues during rebuilds
- Outputs: `dist/reminderagua.exe` and `./reminderagua.exe`

### What the spec bundles
| Source | Destination in exe | Purpose |
|--------|--------------------|---------|
| `locales/` | `locales/` | UI translations |
| `config.json` | `.` | Default configuration |
| `reminderagua.ico` | `.` | App icon |

### After compilation
1. Test the exe from `dist/reminderagua.exe` or `./reminderagua.exe`
2. Verify it opens without console window
3. Verify subject/body placeholders resolve from the current PC date at send time
4. Verify it connects to Outlook and loads config
5. The `config.json` in dist/ is the default — user data is written next to the exe

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

### Important validation rule
- Do not validate placeholder behavior with legacy scripts such as `reminderagua.py`
- The authoritative build path is `main.py` through `reminderagua.spec`

### Spec file location
`reminderagua.spec` — at project root. Edit this file if you:
- Change the entry point
- Add new data files to bundle
- Need to add new hiddenimports
