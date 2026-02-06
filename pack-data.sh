#!/bin/bash

REPO_DIR="/c/Projekty/github/cmdt"
cd "$REPO_DIR" || exit 1

ARCHIVE="cmdt.7z"
PASSWORD="github.com"

# Sprawdź czy pliki istnieją
if [ ! -f "./data/cmdt_x64.exe" ] || [ ! -f "./data/cmdt_x86.exe" ]; then
    echo "❌ Błąd: Nie znaleziono plików w katalogu data/"
    echo "   Oczekiwane:"
    echo "   - ./data/cmdt_x64.exe"
    echo "   - ./data/cmdt_x86.exe"
    exit 1
fi

# Usuń stare archiwum
rm -f "$ARCHIVE"

echo "======================================"
echo "📦 Pakuję pliki do $ARCHIVE"
echo "🔒 Hasło: $PASSWORD"
echo "======================================"

# Pakuj wszystkie pliki z katalogu data/
"/c/Program Files/7-Zip/7z.exe" a -t7z -mx=9 -p"$PASSWORD" "$ARCHIVE" \
    ./data/cmdt_x64.exe \
    ./data/cmdt_x86.exe

if [ $? -eq 0 ]; then
    echo ""
    echo "======================================"
    echo "✅ Sukces!"
    echo "======================================"
    SIZE=$(du -h "$ARCHIVE" | cut -f1)
    echo "   Rozmiar: $SIZE"
    ls -lh "$ARCHIVE"
else
    echo "❌ Błąd pakowania!"
    exit 1
fi
