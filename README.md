# Empreinte carbone de GitHub Copilot

Outils pour estimer l'empreinte carbone (kgCO2e) de l'usage de GitHub Copilot à partir de l'export de facturation GitHub, jour par jour, par modèle et par utilisateur.

Deux versions sont proposées, avec la même méthode de calcul :

| Version | Fichier | Pour qui |
|---|---|---|
| Tableau de bord HTML | [`tableau-carbone/empreinte-carbone-copilot.html`](tableau-carbone/empreinte-carbone-copilot.html) | Usage courant : filtres, graphiques interactifs, projection, contrôles, FR/EN |
| Classeur Excel | [`tableau-carbone/empreinte-carbone-copilot-modele.xlsx`](tableau-carbone/empreinte-carbone-copilot-modele.xlsx) | Quand il faut rester dans Excel (formules visibles, pas de navigateur) |

Les autres fichiers :

- `tableau-carbone/empreinte-carbone-copilot-exemple.xlsx` : le classeur déjà rempli avec les données d'exemple.
- `tableau-carbone/exemple-export-copilot.csv` : un export **fictif** (utilisateurs, organisations et dépôts inventés) pour tester les deux outils.
- `tableau-carbone/empreinte-carbone-copilot.txt` : copie texte du code du tableau HTML, identique au `.html`, pour le lire ou le transférer là où les fichiers HTML sont bloqués.
- `tableau-carbone/source-excel/` : les scripts qui génèrent le classeur Excel.

## Méthode en bref

Pour chaque ligne de l'export (un jour, un utilisateur, un modèle) :

1. **Tokens d'entrée frais** : `F = I − δ·CR − ε·CW`, où `I` est l'input, `CR` le cache lu et `CW` le cache écrit. Les coefficients δ et ε dépendent de la convention de comptage du fournisseur (cache inclus ou non dans l'input).
2. **Tokens équivalent-sortie** : `T_O = F·s_I + CR·s_CR + CW·s_CW + O`. Chaque type de token est ramené à un token de sortie selon le rapport de prix de la grille GitHub (`s_I = prix input / prix output`, etc.).
3. **Énergie** : `E (kWh) = T_O × e_O × PUE / 1000`.
4. **Carbone** : `CO2e (kg) = T_O × facteur d'émission par token` (table de facteurs par modèle). Une variante calcule `E × mix électrique + fabrication`, à partir des scénarios EcoLogits v2 du « Calculateur de Boris » (borisruf/carbon-footprint-modeling-tool). Les modèles absents de ce calculateur sont rapprochés d'un modèle proche.

Des **contrôles** automatiques (C1 à C6) vérifient la cohérence de l'export : structure des tokens, convention de cache, rapprochement des coûts recalculés avec `gross_amount` (rapport attendu entre 0,90 et 1,00), et sensibilité aux hypothèses. Une **projection** jusqu'au 31 décembre donne l'atterrissage de l'année avec une fourchette.

Tous les paramètres (formules, facteurs, rapports de prix, PUE, mix électrique) sont affichés et modifiables dans l'onglet **Calcul**, et la méthodologie complète est dans l'onglet **Méthodologie**.

## Utiliser le tableau de bord HTML

1. Téléchargez `empreinte-carbone-copilot.html` (bouton *Download raw file* sur GitHub) et ouvrez-le dans Edge, Chrome ou Firefox (Edge 111 minimum). Aucune installation et aucune connexion ne sont nécessaires : tout tourne dans le navigateur et les données ne quittent pas le poste.
2. Cliquez sur **Importer l'extraction GitHub** et choisissez (ou glissez-déposez) le CSV de facturation GitHub Copilot (colonnes `date, username, product, sku, model, quantity, …, input, output, cache_read, cache_write`). Pour essayer, prenez `exemple-export-copilot.csv`.
3. Naviguez dans les onglets : **Synthèse** (chiffres clés, tendance, atterrissage), **Empreinte** (détail par modèle, cumul, intensité, coûts), **Calcul** (paramètres et projection), **Contrôles**, **Méthodologie**.
4. Les réglages modifiés sont gardés dans le navigateur. Les exports CSV de résultats s'ouvrent directement dans Excel en français.

## Utiliser le classeur Excel

1. Ouvrez `empreinte-carbone-copilot-modele.xlsx`.
2. Dans l'onglet **Données**, collez l'export en cellule `A2` (sans la ligne d'en-tête). Les colonnes bleues se calculent seules.
3. Choisissez l'année dans la cellule jaune de **Synthèse**. Les cellules jaunes de **Calcul** sont modifiables, et tout le classeur se recalcule.

Le classeur reprend la logique du tableau HTML, avec quelques simplifications : projection au rythme moyen des 28 derniers jours (sans fourchette P10–P90), contrôle des coûts agrégé par famille de modèles, français uniquement, pas de filtre par utilisateur. Il n'utilise que des fonctions compatibles Excel 2007.

## Reconstruire le classeur Excel

Prérequis : Python 3 avec `openpyxl`, et LibreOffice (commande `soffice`).

```bash
cd tableau-carbone/source-excel
./make.sh empreinte-carbone-copilot-modele                                    # modèle vierge
./make.sh empreinte-carbone-copilot-exemple ../exemple-export-copilot.csv      # classeur rempli
```

- `build.py` écrit le classeur (onglets, formules, graphiques) à partir de `src.json` (grille de prix GitHub et familles de modèles) et `eco.json` (scénarios EcoLogits v2 du Calculateur de Boris).
- LibreOffice recalcule le classeur, puis `inject.py` recopie les valeurs calculées dans le fichier d'origine, pour qu'il s'affiche rempli dès l'ouverture sans perdre les colonnes calculées du tableau Excel.

Le tableau HTML n'a pas d'étape de construction : le fichier `.html` est lui-même la source. Après chaque modification, la copie `.txt` est mise à jour à l'identique.

## Limites

Les résultats sont des **estimations**. Les facteurs d'émission par token des modèles récents ne sont pas publiés par les fournisseurs : ils viennent d'une table de facteurs fournie et, à défaut, de modèles proches dans EcoLogits. La fourchette d'incertitude et l'onglet Contrôles indiquent la sensibilité des résultats à ces hypothèses.
