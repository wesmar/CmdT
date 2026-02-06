#!/bin/bash

# Konfiguracja
REPO_DIR="/c/Projekty/github/cmdt"
BRANCH="main"

# Funkcja do wypychania zmian
push_changes() {
    echo "Przechodzę do katalogu: $REPO_DIR"
    cd "$REPO_DIR" || { echo "Błąd: Nie można przejść do katalogu!"; exit 1; }
    
    # Fetch remote info
    git fetch origin
    
    # Sprawdź czy są commity do wypchnięcia
    LOCAL=$(git rev-parse @)
    REMOTE=$(git rev-parse @{u} 2>/dev/null || echo "no-upstream")
    
    if [ "$LOCAL" != "$REMOTE" ] && [ "$REMOTE" != "no-upstream" ]; then
        echo "⚠️  Są commity do wypchnięcia!"
        git log --oneline @{u}..@
    fi
    
    # Sprawdź czy są zmiany w working dir
    echo "Sprawdzam zmiany..."
    HAS_CHANGES=false
    
    if ! git diff --quiet || ! git diff --staged --quiet; then
        HAS_CHANGES=true
        echo "Zmiany do commitowania:"
        git status --short
        
        # Dodaj i commituj
        git add .
        if [ -n "$1" ]; then
            git commit -m "$1"
        else
            git commit -m "Update: $(date '+%Y-%m-%d %H:%M:%S')"
        fi
    fi
    
    # Sprawdź znowu czy coś do pushowania
    if [ "$LOCAL" == "$REMOTE" ] && [ "$HAS_CHANGES" == "false" ]; then
        echo "✅ Wszystko zsynchronizowane."
        return 0
    fi
    
    # Push
    echo "Wypycham zmiany na GitHub..."
    if git push origin "$BRANCH"; then
        echo ""
        echo "✅ Zmiany wypchnięte pomyślnie!"
        echo "   Branch: $BRANCH"
        echo "   Repo: https://github.com/wesmar/cmdt"
    else
        echo "❌ Push failed! Spróbuj:"
        echo "   git pull --rebase origin $BRANCH"
        echo "   git push origin $BRANCH"
        exit 1
    fi
}

# Obsługa parametrów
if [ "$#" -gt 1 ]; then
    echo "Użycie: $0 [wiadomość_commit]"
    exit 1
fi

# Uruchom funkcję
push_changes "$1"
