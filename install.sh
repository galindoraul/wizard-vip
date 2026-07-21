#!/bin/bash
# install.sh — Wizard VIP
# Downloads and installs ALL skills (standard + VIP) for your Claude environment.
# Run this script once to install, or again anytime to update.

WIZARD_DIR="$HOME/.wizard"
QA_REPO_DIR="$WIZARD_DIR/wizard"
VIP_REPO_DIR="$WIZARD_DIR/wizard-vip"
CLAUDE_DIR="$WIZARD_DIR/.claude"
RAW_BASE_QA="https://raw.githubusercontent.com/galindoraul/wizard/main"
RAW_BASE_VIP="https://raw.githubusercontent.com/galindoraul/wizard-vip/main"

echo ""
echo "🧙‍♂️ Wizard VIP "
echo "───────────────────────────────────"
echo ""

mkdir -p "$WIZARD_DIR"

# --- Git clone/pull with fallback to curl ---

clone_or_update() {
    local repo_url="$1"
    local repo_dir="$2"
    local raw_base="$3"
    local label="$4"

    # Try git first
    if [ -d "$repo_dir/.git" ]; then
        if git -c http.https://github.com.sslVerify=false -c credential.helper= -C "$repo_dir" pull 2>/dev/null; then
            echo "📥 Updated $label via git"
            return 0
        else
            # Pull failed (conflict, network, etc.) — nuke and re-clone
            rm -rf "$repo_dir"
            if git -c http.https://github.com.sslVerify=false -c credential.helper= clone "$repo_url" "$repo_dir" 2>/dev/null; then
                echo "📥 Re-cloned $label via git"
                return 0
            fi
        fi
    else
        if git -c http.https://github.com.sslVerify=false -c credential.helper= clone "$repo_url" "$repo_dir" 2>/dev/null; then
            echo "📦 Cloned $label via git"
            return 0
        fi
    fi

    # Fallback: download via curl using manifest
    echo "📥 Downloading $label files..."
    rm -rf "$repo_dir"
    mkdir -p "$repo_dir"

    local manifest
    manifest=$(curl -sL "$raw_base/manifest.txt")
    if [ -z "$manifest" ]; then
        echo "   ❌ Could not download $label manifest"
        return 1
    fi

    while IFS= read -r file; do
        [ -z "$file" ] && continue
        local dir=$(dirname "$repo_dir/$file")
        mkdir -p "$dir"
        curl -sL "$raw_base/$file" -o "$repo_dir/$file"
    done <<< "$manifest"
}

clone_or_update "https://github.com/galindoraul/wizard.git" "$QA_REPO_DIR" "$RAW_BASE_QA" "wizard"
clone_or_update "https://github.com/galindoraul/wizard-vip.git" "$VIP_REPO_DIR" "$RAW_BASE_VIP" "wizard-vip"

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

# Add wizard alias (auto-updates + ensures claude is installed)
SHELL_RC="$HOME/.zshrc"
ALIAS_LINE='alias wizard="(curl -sL https://raw.githubusercontent.com/galindoraul/wizard-vip/main/install.sh | bash > /dev/null 2>&1 &); command -v claude >/dev/null 2>&1 || devfeature install claude_code; cd ~/.wizard && claude"'
grep -v 'alias wizard=' "$SHELL_RC" > "$SHELL_RC.tmp" 2>/dev/null && mv "$SHELL_RC.tmp" "$SHELL_RC"
echo '' >> "$SHELL_RC"
echo "$ALIAS_LINE" >> "$SHELL_RC"

echo ""
echo "   ⚡ Alias 'wizard' ready (auto-updates on launch)"
echo ""
echo "───────────────────────────────────"
echo "✅ Done! Restart terminal, then type: wizard"
echo ""
