#!/bin/bash
# install.sh — Wizard VIP
# Downloads and installs ALL skills (standard + VIP) for your Claude environment.
# Run this script once to install, or again anytime to update.

WIZARD_DIR="$HOME/.wizard"
QA_REPO_DIR="$WIZARD_DIR/wizard"
VIP_REPO_DIR="$WIZARD_DIR/wizard-vip"
CLAUDE_DIR="$WIZARD_DIR/.claude"

echo ""
echo "🧙‍♂️ Wizard VIP"
echo "───────────────────────────────────"
echo ""

# Clone or update repos
mkdir -p "$WIZARD_DIR"
if [ -d "$QA_REPO_DIR/.git" ]; then
    echo "📥 Updating wizard..."
    cd "$QA_REPO_DIR" && git pull
else
    echo "📦 Cloning wizard..."
    git clone https://github.com/galindoraul/wizard.git "$QA_REPO_DIR"
fi

if [ -d "$VIP_REPO_DIR/.git" ]; then
    echo "📥 Updating wizard-vip..."
    cd "$VIP_REPO_DIR" && git pull
else
    echo "📦 Cloning wizard-vip..."
    git clone https://github.com/galindoraul/wizard-vip.git "$VIP_REPO_DIR"
fi

# Clean old symlinks
rm -rf "$CLAUDE_DIR"
mkdir -p "$CLAUDE_DIR"

# Process all .claude/ contents from a repo
process_repo() {
    local repo_dir="$1"
    local label="$2"
    local src="$repo_dir/.claude"

    [ ! -d "$src" ] && return

    for subdir in "$src"/*/; do
        [ ! -d "$subdir" ] && continue
        subdir_name=$(basename "$subdir")
        target="$CLAUDE_DIR/$subdir_name"
        mkdir -p "$target"

        for item in "$subdir"*/; do
            [ ! -d "$item" ] && continue
            item_name=$(basename "$item")
            ln -sf "$item" "$target/$item_name"
            echo "   ✅ $subdir_name/$item_name ($label)"
        done
    done

    for file in "$src"/*; do
        [ -d "$file" ] && continue
        [ ! -f "$file" ] && continue
        file_name=$(basename "$file")
        ln -sf "$file" "$CLAUDE_DIR/$file_name"
        echo "   ✅ $file_name ($label)"
    done
}

echo "🔗 Installed:"
process_repo "$QA_REPO_DIR" "wizard"
process_repo "$VIP_REPO_DIR" "wizard-vip"

# Add wizard alias with auto-update
SHELL_RC="$HOME/.zshrc"
ALIAS_LINE='alias wizard="cd ~/.wizard && (cd wizard && git pull -q &) 2>/dev/null; (cd wizard-vip && git pull -q &) 2>/dev/null; claude"'
if grep -q 'alias wizard=' "$SHELL_RC" 2>/dev/null; then
    sed -i '' 's|alias wizard=.*|'"$ALIAS_LINE"'|' "$SHELL_RC"
else
    echo '' >> "$SHELL_RC"
    echo "$ALIAS_LINE" >> "$SHELL_RC"
fi
echo ""
echo "   ⚡ Alias 'wizard' ready (auto-updates on launch)"

echo ""
echo "───────────────────────────────────"
echo "✅ Done! Restart terminal, then type: wizard"
echo ""
