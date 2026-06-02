# Excel Synthèse Python — Balance des Échanges Internationaux

Script Python qui génère automatiquement un fichier Excel consolidé à partir des fichiers de suivi des étudiants entrants et sortants d'une école de commerce.

## Fonctionnement

```
Fichiers Incoming*.xlsx + Outgoing*.xlsx
            ↓
    Script Python (assistant_mobility.py)
            ↓
    Balance_Echanges.xlsx
    ├── Résumé (par année et université)
    ├── Par_Programme (par année, programme, semestre)
    └── Infos (logs d'exécution)
```

## Fonctionnalités

- Lecture automatique de tous les fichiers `Incoming*.xlsx` et `Outgoing*.xlsx` du dossier
- Mapping dynamique des colonnes — robuste aux variations de nommage entre fichiers sources
- Normalisation automatique des semestres (`Fall Semester`, `Spring Semester`, `Whole Year`, etc.)
- Détection et ignorance des lignes barrées (désistements) et des lignes `free mover / study abroad`
- Enrichissement via table de référence des universités partenaires (codes Erasmus → nom / pays / zone)
- Auto-enrichissement de la table de référence : toute nouvelle université détectée est ajoutée automatiquement
- Dérivation du pays depuis le préfixe du code Erasmus si le pays est absent
- Mise en forme Excel automatique : tableaux, couleurs, totaux, colonnes auto-ajustées
- Logs d'exécution complets dans la feuille `Infos`

## Stack technique

- **Python** — pandas, numpy, openpyxl, pathlib, unicodedata
- **openpyxl** — génération et mise en forme du fichier Excel final
- **pandas** — consolidation et agrégation des données

## Structure du projet

```
├── assistant_mobility.py              # Script principal
├── requirements.txt                   # Dépendances Python
├── Générer Balance.bat                # Lanceur Windows (double-clic)
└── Template fichiers suivi IN-OUT.xlsx  # Table de référence des universités
```

## Utilisation

### Nommage des fichiers source

```
Incoming [Programme] [Année].xlsx
Outgoing [Programme] [Année].xlsx
```

Exemples :
```
Incoming BBA 2025-2026.xlsx
Outgoing PGE DD 2024-2025.xlsx
Incoming Summer School 2024-2025.xlsx
```

**Programmes reconnus :** BBA · BBA DD · PGE · PGE DD · Summer School · Winter School

### Lancement

```bash
# Via terminal
python assistant_mobility.py

# Via Windows
Double-cliquer sur "Générer Balance.bat"
```

### Installation des dépendances

```bash
pip install -r requirements.txt
```

## Fichier résultat — Balance_Echanges.xlsx

### Feuille Résumé
Vue consolidée par année et université (tous programmes confondus).

`AcademicYear · Partner · InstCode · Country · Zone · Incoming · Outgoing · Balance`

### Feuille Par_Programme
Détail par année, programme, université et semestre.

`AcademicYear · Exchange_program · Partner · InstCode · Country · Zone · Semester · Incoming · Outgoing · Balance`

### Code couleur — colonne Balance

| Couleur | Valeur |
|---------|--------|
| 🟢 Vert | Balance > 0 |
| 🟡 Jaune | -3 à -1 |
| 🟠 Orange | -6 à -4 |
| 🔴 Rouge | < -6 |

## Table de référence

L'onglet `Codes institutionnels` du template contient les universités partenaires avec leur code Erasmus, pays et zone géographique.

- Le script enrichit automatiquement cette table si une nouvelle université est trouvée
- Toutes les universités de la table apparaissent dans le résultat final, même avec 0 étudiants

## Contexte

Projet développé durant un stage à **emlyon Business School** (2026).  
Développé par Danee Ayasamy — étudiant ingénieur CESI Lyon.
