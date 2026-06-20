---
description: Senior engineer improvement process — analyze architecture, identify problems, propose and implement improvements without breaking existing behavior. Use when asked to "mejorar el código", "refactorizar", or "aplicar mejoras de ingeniero senior".
---

## Senior Engineer Improvement Skill

Act as a senior Python engineer. Follow the MANDATORY analysis-first process from the code-improver agent.

## Improvement areas to evaluate (in order)

### 1. Critical bugs
- [ ] Auto-send on launch (`root.after` with invoke) — **FIXED in v2**
- [ ] GUI freeze during COM call — **FIXED in v2 (threading)**
- [ ] Hotmail send failure — **FIXED in v2 (dual method)**

### 2. Architecture
- [ ] Frontend/backend separation — **DONE in v2**
- [ ] Class-based GUI instead of globals — **DONE in v2**
- [ ] Config isolated in ConfigManager — **DONE in v2**

### 3. Code quality
- [ ] Error messages are specific and actionable
- [ ] All exceptions are caught at the right boundary
- [ ] No print() statements — use logger
- [ ] Constants defined at module level
- [ ] No magic strings or numbers inline

### 4. User experience
- [ ] Status bar shows clear feedback
- [ ] Send button disabled while sending (prevents double-send)
- [ ] Language preference persisted
- [ ] Beer donation button visible

### 5. Maintainability
- [ ] All UI strings in locales/ (not hardcoded)
- [ ] New language added by creating locales/{lang}.json only
- [ ] Config keys documented in CLAUDE.md
- [ ] Architecture documented in docs/SDD.md

## Rules (always apply)
- **Never break existing config.json compatibility**
- Keep placeholder tokens: `[Mes anterior en letras]`, `[año en numero]`
- Never add tkinter imports to src/backend/
- Comment non-obvious code with WHY, not WHAT
- One function = one responsibility

## After improving
1. Update docs/SDD.md
2. Update CLAUDE.md
3. Run `python main.py` to verify
4. Run `/compile-exe` to recompile
5. Run `/github-push` to commit

## Adding a new language
1. Copy `locales/es.json` → `locales/{lang}.json`
2. Translate all string values (keep the keys in English)
3. Add `lang` to `SUPPORTED_LANGUAGES` in `src/frontend/i18n.py`
4. Add menu item in `app.py` `_build_menu()`
