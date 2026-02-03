# Guide d'installation - Automatisation SUIVI Kiabi

## Prerequis

- Windows 10/11
- Microsoft Excel installe (avec Power Query)
- Les fichiers SUIVI synchronises via OneDrive

---

## Installation

1. Dezippez le dossier `Automatisation_SUIVI.zip`
2. Placez le dossier ou vous voulez (ex: Bureau)
3. Double-cliquez sur `Automatisation_SUIVI.exe`
4. Au premier lancement, entrez le chemin du dossier OneDrive contenant les sous-dossiers SUIVI

   Exemple : `C:\Users\VotreNom\OneDrive - Kiabi\MonDossier`

   Ce chemin doit contenir les dossiers :
   - `SUIVI_KPIS\`
   - `SUIVI_MDR\`
   - `SUIVI_PMA\`
   - `SUIVI_PRODUIT\`
   - `SUIVI_CRM\`

5. Le chemin est enregistre et ne sera plus demande aux prochains lancements
6. Pour changer le chemin : relancez avec `Automatisation_SUIVI.exe --config`

---

## Configuration prealable des fichiers (A FAIRE UNE SEULE FOIS)

Avant le premier lancement du script, il faut configurer certains parametres dans les fichiers Excel source (les derniers fichiers en date, ex: S04).

### 1. Niveaux de confidentialite (TOUS les fichiers)

Pour chaque fichier SUIVI (KPIS, MDR, PMA, PRODUIT, CRM) :

1. Ouvrir le fichier dans Excel
2. Aller dans **Donnees** > **Obtenir des donnees** > **Parametres de la source de donnees**
3. Pour chaque source listee :
   - Cliquer sur la source
   - Cliquer sur **"Modifier les autorisations..."**
   - Dans "Niveau de confidentialite", choisir **"Public"**
   - Cliquer **OK**
4. Fermer la fenetre des parametres
5. **Enregistrer** le fichier

> Ce parametre sera conserve dans les fichiers dupliques par le script.

### 2. Requetes selligent - Navigation (fichier SUIVI_CRM uniquement)

Les requetes `selligent_all` et `selligent_all_histo` utilisent un identifiant (UUID) qui change a chaque semaine. Il faut modifier la formule pour ne plus dependre de cet identifiant.

Pour chaque requete (`selligent_all` et `selligent_all_histo`) :

1. Ouvrir le fichier SUIVI_CRM dans Excel
2. Aller dans **Donnees** > **Requetes et connexions**
3. Double-cliquer sur la requete (ex: `selligent_all`)
4. Dans le panneau **Etapes appliquees** (a droite), cliquer sur l'etape **Navigation**
5. Dans la barre de formule en haut, vous verrez :
   ```
   = Source{[Item="3770e733-5c4e-...", Kind="Sheet"]}[Data]
   ```
6. **Remplacer** cette formule par :
   ```
   = Source{0}[Data]
   ```
7. Appuyer sur Entree pour valider
8. Faire la meme chose pour `selligent_all_histo`
9. Cliquer sur **Fermer et charger**
10. **Enregistrer** le fichier

> Cette modification permet au script de fonctionner sans connaitre l'identifiant du fichier source.

---

## Ce que fait le script automatiquement

A chaque lancement, le script effectue les operations suivantes pour chaque fichier :

| Etape | Description |
|-------|-------------|
| 1 | Trouve le dernier fichier (ex: SUIVI_KPIS_S04.xlsx) |
| 2 | Le duplique avec le numero de semaine suivant (S04 -> S05) |
| 3 | Met a jour la date dans la cellule A1 (+7 jours) |
| 4 | Met a jour les requetes Power Query (chemins, dates) |
| 5 | Actualise toutes les connexions de donnees |
| 6 | Recalcule toutes les formules |
| 7 | Sauvegarde et ferme |

### Ordre de traitement

1. **SUIVI_KPIS** - KPIs hebdomadaires
2. **SUIVI_MDR** - Rapport MDR
3. **SUIVI_PMA** - Rapport PMA (format .xlsm)
4. **SUIVI_PRODUIT** - Rapport Produit (format .xlsm)
5. **SUIVI_CRM** - Rapport CRM (requetes selligent, push et piano)

### Fichier CRM - Details des requetes mises a jour

| Requete | Type de mise a jour |
|---------|-------------------|
| selligent_all | Chemin fichier : 2025_S04 -> 2025_S05 |
| selligent_all_histo | Chemin fichier : 2025_S04 -> 2025_S05 |
| push_all | Chemin fichier : 2025_S04 -> 2025_S05 |
| push_all_histo | Chemin fichier : 2025_S04 -> 2025_S05 |
| piano_all | Dates start/end : +7 jours |
| piano_all_histo | Dates start/end : +7 jours |

---

## Utilisation hebdomadaire

1. **Le lundi**, double-cliquez sur `Automatisation_SUIVI.exe`
2. Le script s'execute automatiquement (environ 5-10 min)
3. Excel s'ouvre et se ferme pour chaque fichier (ne pas toucher)
4. A la fin, un resume s'affiche avec le statut de chaque fichier
5. Appuyez sur Entree pour fermer

### En cas d'erreur

- Verifiez que tous les fichiers Excel sont fermes avant de lancer le script
- Verifiez la connexion au lecteur U: (pour les requetes selligent/push du CRM)
- Si le script plante, ouvrez le Gestionnaire des taches et fermez tous les processus "Excel" avant de relancer
- Pour relancer un fichier specifique, utilisez les scripts individuels :
  ```
  python scripts/update_kpis.py
  python scripts/update_mdr.py
  python scripts/update_pma.py
  python scripts/update_produit.py
  python scripts/update_crm.py
  ```

---

## Structure des dossiers attendue

```
[Dossier OneDrive]\
    SUIVI_KPIS\
        SUIVI_KPIS_S04.xlsx
    SUIVI_MDR\
        SUIVI_MDR_S04.xlsx
    SUIVI_PMA\
        SUIVI_PMA_S04.xlsm
    SUIVI_PRODUIT\
        SUIVI_PRODUIT_S04.xlsm
    SUIVI_CRM\
        SUIVI_CRM_S04.xlsx
```
