#!/usr/bin/env bash
# Reconstruit un classeur Excel.
# Usage : ./make.sh <nom_sans_extension> [export.csv]
#   ./make.sh empreinte-carbone-copilot-modele                       -> modèle vierge
#   ./make.sh empreinte-carbone-copilot-exemple ../exemple-export-copilot.csv
# Prérequis : Python 3 + openpyxl, LibreOffice (commande soffice).
set -e
cd "$(dirname "$0")"
python3 build.py "$1.xlsx" $2
mkdir -p _recalc
soffice --headless --calc --convert-to xlsx --outdir _recalc "$1.xlsx" >/dev/null
python3 inject.py "$1.xlsx" "_recalc/$1.xlsx" "$1_final.xlsx"
mv "$1_final.xlsx" "$1.xlsx"
rm -rf _recalc
echo "Classeur prêt : $(pwd)/$1.xlsx"
