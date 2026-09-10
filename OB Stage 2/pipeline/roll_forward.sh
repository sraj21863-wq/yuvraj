#!/bin/sh
# Freeze the filled live block into history and lay a fresh one underneath.
cd "$(dirname "$0")" || exit 1
python3 roll_forward.py --dry-run || exit 1
printf '\nGo ahead? (y/N) '
read -r go
case "$go" in y|Y) exec python3 roll_forward.py ;; *) echo 'nothing changed.' ;; esac
