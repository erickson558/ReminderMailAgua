---
description: Add meaningful comments and docstrings to Python source files explaining WHY each part exists, not just what it does. Follows project commenting conventions. Use when asked to "comentar el código", "agregar comentarios", or "explicar qué hace cada parte".
---

## Comment Code Skill

Add useful comments and docstrings to the project's Python files.

## What to comment

### Module docstring (top of file)
```python
"""
Nombre del módulo — qué responsabilidad tiene en la aplicación.

Contexto importante: por qué existe este módulo separado,
qué problema resuelve, qué dependencias externas usa.
"""
```

### Class docstring
```python
class MyClass:
    """
    Qué representa o gestiona esta clase.
    Una o dos oraciones son suficientes para clases simples.
    """
```

### Method/function docstring
```python
def my_method(self, arg: str) -> bool:
    """
    Descripción breve de lo que hace.

    Args:
        arg: Qué representa este argumento y valores válidos.

    Returns:
        True si ..., False si ...

    Raises:
        ValueError: Cuándo y por qué se lanza.
    """
```

### Inline comments — solo para lógica no obvia
```python
# Comparación case-insensitive porque Outlook puede retornar la dirección
# con capitalización diferente a la guardada en config.json
if acc.SmtpAddress.strip().lower() == sender_lower:
    ...

# root.after(0, ...) es la única forma thread-safe de actualizar tkinter
# desde un hilo de fondo — llamar directamente causaría race conditions
self.root.after(0, lambda: self._status_label.config(text=message))
```

## What NOT to comment

```python
# MAL: describe lo que ya es obvio por el nombre
i += 1  # incrementar contador

# MAL: repite lo que dice el nombre de la función
def save_config():
    # Guardar la configuración    ← innecesario

# MAL: describe el "qué", no el "por qué"
mail.To = "; ".join(recipients)  # juntar con punto y coma
```

## Rules for this project

1. All docstrings in **Spanish** (matches existing code language)
2. Args/Returns/Raises sections only when non-trivial
3. Mark win32com quirks with a comment explaining the COM behavior
4. Mark Hotmail-specific workarounds with `# FIX: Hotmail — <reason>`
5. Mark tkinter thread-safety points with `# thread-safe: root.after`
6. Constants: comment origin/reason if not obvious from name

## Files to comment in this project
Priority order:
1. `src/backend/email_service.py` — most complex, win32com behavior
2. `src/frontend/app.py` — threading and tkinter patterns
3. `src/backend/config_manager.py` — path resolution logic
4. `src/frontend/i18n.py` — PyInstaller path handling
5. `src/backend/date_utils.py` — date calculation logic
6. `main.py` — entry point bootstrapping
