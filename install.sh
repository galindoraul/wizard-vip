#!/bin/bash
# install.sh — Wizard VIP Skills Installer
# Downloads and installs ALL skills (standard + VIP) for your Claude environment.
# Run this script once to install, or again anytime to update.

QA_REPO_URL="https://github.com/galindoraul/wizard.git"
VIP_REPO_URL="https://github.com/galindoraul/wizard-vip.git"
QA_REPO_DIR="$HOME/.wizard/wizard"
VIP_REPO_DIR="$HOME/.wizard/wizard-vip"
SKILLS_DIR="$HOME/.claude/skills"

echo ""
echo "🧙‍♂️ Wizard VIP"
echo "───────────────────────────────────"
echo ""

# Clone or update standard repo
if [ -d "$QA_REPO_DIR/.git" ]; then
    echo "📥 Updating standard skills..."
    cd "$QA_REPO_DIR" && git pull
else
    echo "📦 Installing standard skills..."
    mkdir -p "$(dirname "$QA_REPO_DIR")"
    git clone "$QA_REPO_URL" "$QA_REPO_DIR"
fi

# Clone or update VIP repo
if [ -d "$VIP_REPO_DIR/.git" ]; then
    echo "📥 Updating VIP skills..."
    cd "$VIP_REPO_DIR" && git pull
else
    echo "📦 Installing VIP skills..."
    mkdir -p "$(dirname "$VIP_REPO_DIR")"
    git clone "$VIP_REPO_URL" "$VIP_REPO_DIR"
fi

# Auto-detect and symlink all skills from both repos
mkdir -p "$SKILLS_DIR"
count=0
echo ""
echo "🔗 Installed skills:"

QA_SKILLS_SRC="$QA_REPO_DIR/.claude/skills"
if [ -d "$QA_SKILLS_SRC" ]; then
    for skill_dir in "$QA_SKILLS_SRC"/*/; do
        if [ -f "$skill_dir/SKILL.md" ]; then
            skill_name=$(basename "$skill_dir")
            ln -sf "$skill_dir" "$SKILLS_DIR/$skill_name"
            echo "   ✅ /$skill_name"
            count=$((count + 1))
        fi
    done
fi

VIP_SKILLS_SRC="$VIP_REPO_DIR/.claude/skills"
if [ -d "$VIP_SKILLS_SRC" ]; then
    for skill_dir in "$VIP_SKILLS_SRC"/*/; do
        if [ -f "$skill_dir/SKILL.md" ]; then
            skill_name=$(basename "$skill_dir")
            ln -sf "$skill_dir" "$SKILLS_DIR/$skill_name"
            echo "   ✅ /$skill_name ⭐"
            count=$((count + 1))
        fi
    done
fi

echo ""
echo "───────────────────────────────────"
echo "✅ Done! $count skill(s) ready to use in Claude."
echo ""
