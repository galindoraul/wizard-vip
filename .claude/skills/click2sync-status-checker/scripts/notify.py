#!/usr/bin/env python3
"""Send C2C reminder to Google Chat space via webhook.
Uses @all + individual @mentions for missing people.
Resolves gchat user IDs from the space member list (cached 1hr)."""

import json
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta

WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/AAQA6X4LVDY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=9y7jxKd60u_-m6D2OvDktHb3vBnwOnzFsN0y4snwe5w"
SPACE_NAME = "spaces/AAQA6X4LVDY"

SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = "/tmp"
MEMBERS_CACHE_TTL = 3600  # 1 hour


def get_status():
    """Run read-submissions.py and parse its output."""
    result = subprocess.run(
        ["/usr/bin/python3", os.path.join(SCRIPTS_DIR, "read-submissions.py")],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"Error running read-submissions: {result.stderr}", file=sys.stderr)
        sys.exit(1)

    lines = result.stdout.strip().split("\n")
    json_start = next(
        (i for i, l in enumerate(lines) if l.strip().startswith("{")), None
    )
    if json_start is None:
        print("Error: No JSON from read-status", file=sys.stderr)
        sys.exit(1)

    return json.loads("\n".join(lines[json_start:]))


def get_space_members():
    """Get all space members with their gchat IDs and emails. Cached 1hr."""
    cache_path = os.path.join(CACHE_DIR, "c2c_status_space_members.json")

    if os.path.exists(cache_path):
        age = time.time() - os.path.getmtime(cache_path)
        if age < MEMBERS_CACHE_TTL:
            with open(cache_path) as f:
                return json.load(f)

    result = subprocess.run(
        [
            "meta",
            "google.chat.member",
            "list",
            f"--space-name={SPACE_NAME}",
            "-o",
            "json",
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"Warning: Could not get space members: {result.stderr}", file=sys.stderr)
        return {}

    lines = result.stdout.strip().split("\n")
    json_start = next(
        (i for i, l in enumerate(lines) if l.strip().startswith("[")), None
    )
    if json_start is None:
        return {}

    members = json.loads("\n".join(lines[json_start:]))

    # Build email -> gchat_id map
    email_to_id = {}
    for member in members:
        email = member.get("email", "")
        name_parts = member.get("name", "").split("/")
        gchat_id = name_parts[-1] if name_parts else ""
        if email and gchat_id:
            email_to_id[email.lower()] = gchat_id

    # Save to cache
    if email_to_id:
        with open(cache_path, "w") as f:
            json.dump(email_to_id, f)

    return email_to_id


def build_message(status, email_to_id):
    """Build webhook message with @all and individual @mentions."""
    missing = status.get("missing", [])
    week = status.get("week", "")

    if not missing:
        return None

    # Build @mentions for each missing person
    mentions = []
    for person in sorted(missing, key=lambda x: x["name"]):
        name = person["name"]
        meta_unixname = person.get("meta_unixname", "")
        email = f"{meta_unixname}@meta.com".lower() if meta_unixname else ""
        gchat_id = email_to_id.get(email, "")

        if gchat_id:
            mentions.append(f"• <users/{gchat_id}>")
        else:
            mentions.append(f"• {name}")

    mentions_text = "\n".join(mentions)

    payload = {
        "text": f"<users/all> ⚠️ C2C Reminder — Semana {week}\n\nFaltan por hacer su C2C ({len(missing)}):\n{mentions_text}\n\nRealicen sus tasks y corran /tasks-to-click2sync antes del viernes."
    }

    return payload


def send_webhook(payload):
    """Send message via webhook."""
    result = subprocess.run(
        [
            "curl",
            "-s",
            "-X",
            "POST",
            WEBHOOK_URL,
            "-H",
            "Content-Type: application/json",
            "-d",
            json.dumps(payload),
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print(f"Error sending webhook: {result.stderr}", file=sys.stderr)
        sys.exit(1)

    try:
        response = json.loads(result.stdout)
        if "error" in response:
            print(f"Webhook error: {response['error']}", file=sys.stderr)
            sys.exit(1)
    except json.JSONDecodeError:
        pass

    print("ok")


def main():
    status = get_status()
    missing = status.get("missing", [])

    if not missing:
        print("No one missing — no notification needed.")
        return

    # Get space members for @mentions
    email_to_id = get_space_members()

    # Build and send message
    payload = build_message(status, email_to_id)
    if payload:
        send_webhook(payload)


if __name__ == "__main__":
    main()
