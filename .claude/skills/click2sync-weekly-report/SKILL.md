---
name: click2sync-weekly-report
description: Genera el xlsx semanal consolidado de Click2Sync con todas las personas. Use when user says "genera el xlsx", "consolida el c2c", "weekly report", "click2sync report", "genera reporte semanal consolidado".
---

# Click2Sync Weekly Report Builder

## IMPORTANT
ALWAYS use `/usr/bin/python3` to run scripts. NEVER use `python` or `python3` without the full path.

## Instructions

### Step 1: Build Report
```bash
/usr/bin/python3 scripts/build-report.py
```

If the script asks for the first requestNo (no previous xlsx found), ask the user:
"¿Cuál es el primer Request No para esta semana? (ej: 29582-3-10050)"

Then re-run with the argument:
```bash
/usr/bin/python3 scripts/build-report.py --first-request-no=29582-3-10050
```

### Step 2: Present Result

Parse the JSON output:
- If status is "ok": Say "✅ Reporte generado: {filename} ({row_count} rows, {person_count} personas)"
- If status is "error": Show the error message

## Presentation Rules

### NEVER show or mention:
- Script names or file paths
- JSON output raw
- Technical details
- Exit codes

### ALWAYS:
- Present in Spanish
- Be concise

## Error: Auth expired

If any script fails with "OAuth" or "401" or "auth" errors, tell the user:

> Tu sesión expiró. Para renovarla:
> 1. Abre https://www.internalfb.com/intern/jf/authenticate/
> 2. Da click en "Generate new token"
> 3. Copia URL
> 4. Pegala en tu terminal y da enter
> 5. Vuelve a correr /click2sync-weekly-report
