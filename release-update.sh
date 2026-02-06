#!/bin/bash

# Konfiguracja
REPO_DIR="/c/Projekty/github/cmdt"
TAG="latest"
REPO="wesmar/cmdt"  # ✅ Poprawiona nazwa repo

cd "$REPO_DIR" || exit 1

echo "======================================"
echo "🔧 KROK 1: Pakowanie plików"
echo "======================================"
./pack-data.sh
if [ $? -ne 0 ]; then
    echo "❌ Błąd pakowania!"
    exit 1
fi

echo ""
echo "======================================"
echo "🗑️  KROK 2: Usuwanie starych assetów"
echo "======================================"

# Usuń stare cmdt.7z
gh release delete-asset "$TAG" cmdt.7z --yes 2>/dev/null && echo "✅ Usunięto cmdt.7z" || echo "⚠️  cmdt.7z nie istniało"

echo ""
echo "======================================"
echo "📤 KROK 3: Upload nowych plików"
echo "======================================"

gh release upload "$TAG" \
    "cmdt.7z#cmdt.7z" \
    --clobber

if [ $? -eq 0 ]; then
    echo ""
    echo "======================================"
    echo "✅ SUKCES!"
    echo "======================================"
    echo "Release zaktualizowany: https://github.com/$REPO/releases/tag/$TAG"
    echo ""
    echo "📦 Zawartość archiwum:"
    echo "   - cmdt_x64.exe (~20KB)"
    echo "   - cmdt_x86.exe (~16KB)"
else
    echo "❌ Błąd uploadu!"
    exit 1
fi
