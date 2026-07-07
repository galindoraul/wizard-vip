---
name: click2sync-status-checker
description: Muestra quién ya hizo su C2C semanal y quién falta. Use when user says "c2c status", "quién falta", "status del c2c", "click2sync status", "quien no ha hecho su c2c".
---

# Click2Sync Status

## IMPORTANT
ALWAYS use `/usr/bin/python3` to run scripts. NEVER use `python` or `python3` without the full path.

## Instructions

### Step 1: Check Status
```bash
/usr/bin/python3 scripts/read-submissions.py
```

### Step 2: Present Results

Parse the JSON output and present like this:

```
📊 C2C Status — Semana Jul 7-13

✅ Ya hicieron su C2C (3):
  - Johan Jaramillo
  - Maria Lopez
  - Raul Galindo

❌ Faltan por hacer su C2C (2):
  - Carlos Perez
  - Ana Rodriguez

⏸️ En PTO toda la semana (1):
  - Pedro Sanchez
```

### Step 3: Ask about notification

After showing the status, IF there are people in the "missing" list, ask:
"¿Quieres que mande el recordatorio al grupo?"

If user says yes, run:
```bash
/usr/bin/python3 scripts/notify.py
```

After the script runs, say: "✅ Recordatorio enviado al grupo."

If there is no one missing, do NOT ask about notifications. Just say:
"✅ Todos hicieron su C2C esta semana."

## Presentation Rules

### NEVER show or mention:
- Script names or file paths
- JSON output raw
- Technical details
- Exit codes
- Webhook URLs

### ALWAYS:
- Present in Spanish
- Use the emoji format shown above
- Sort names alphabetically in each group
- Be concise

## Error: Auth expired

If any script fails with "OAuth" or "401" or "auth" errors, tell the user:

> Tu sesión expiró. Para renovarla:
> 1. Abre https://www.internalfb.com/intern/jf/authenticate/
> 2. Busca "Legacy Options"
> 3. Copia UID y NONCE
> 4. Corre: `jf auth --skip-legacy-auth-upgrade <UID> <NONCE>`
> 5. Vuelve a correr /click2sync-status
