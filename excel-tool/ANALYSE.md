# Analyse technique — Suite « Piste » (piste - V3.xlsm / Visu.xlsm / DispoVierge.xlsx)

Analyse statique du code VBA (macros extraites via `oletools`/`olevba`) et de la structure OOXML
des trois classeurs. Aucune exécution réelle n'a été faite (pas d'Excel disponible dans cet
environnement) : les points listés ci-dessous sont issus de la lecture du code, pas d'un test
runtime. À valider en conditions réelles avant correctif.

**Correctifs déjà appliqués** : les 3 liens externes morts de `Visu.xlsm` (§2) ont été retirés
par édition directe de l'OOXML (validée par re-parsing XML + réouverture openpyxl + comparaison
cellule à cellule avec l'original) — aucune macro touchée. Le VBA de déduplication/robustesse
(§3, points 1-2) est fourni prêt à coller dans `VBA_CORRECTIONS.md`, à appliquer et tester
toi-même dans l'éditeur VBA.

**Correctif du 19/09 sur `Visu.xlsm`** : le nettoyage des liens morts avait renuméroté la
référence externe restante (de `[4]` à `[1]`) dans les formules des cellules, mais pas dans les
`calculatedColumnFormula` du tableau structuré `Tableau1` (`xl/tables/table1.xml`), qui
contiennent leur propre copie de ces formules et référençaient encore `[4]` — un index qui
n'existait plus après suppression des 3 liens morts. Résultat : Excel jugeait le fichier corrompu
à l'ouverture standalone de `Visu.xlsm` (message « Enregistrements supprimés : Tableau ») et
supprimait le tableau. Corrigé (32 occurrences), revalidé par un balayage de tout le paquet OOXML
(feuilles + tables) confirmant qu'aucune formule ne référence plus un index externe inexistant —
ce que ma validation initiale n'avait pas couvert, d'où le bug passé inaperçu jusqu'à ton test.

**Évolution** : une colonne « Roues » (valeurs 1/2/3, vert/orange/rouge) a été ajoutée dans
`piste - V3.xlsm` entre DEM VAL et Masquer, sur `Situation` et `Gestion` — détails, choix de
placement et VBA à coller dans `VBA_CORRECTIONS_ROUES.md`.

**Évolution du 20/09 — Visu.xlsm remplacé par un export HTML** : `Visu.xlsm` (deuxième instance
Excel pilotée par `GetObject`) est retiré du pipeline de rafraîchissement. À la place, la macro
`GenererHTMLSituation` (Module1) lit `Situation!B7:S31` (mapping colonnes confirmé sur le fichier
réel : B=AVION, D=Pos., E=Camp., F=O2, G=AM, H=PM, I=VDN, J=PTR, K=AVQ, L=PLEIN, N=V, O=OBSERVATIONS,
S=Masquer, + date Kannad en V23), injecte les données dans le template
`situation-avions-template.html` (thème visuel façon Airbus) et écrit `situation-avions.html` à
côté du classeur, à chaque Actualiser/Valider. Le fichier généré se recharge lui-même toutes les
30 s dans le navigateur de la TV. Détails, VBA complet et vérifications faites avant livraison
dans `VBA_CORRECTIONS_HTML.md` ; retrait des appels `PousserRafraichissementVisu` devenus inutiles
dans `VBA_CORRECTIONS_RETRAIT_VISU.md`.

## 1. Fonctionnement d'ensemble

Outil de suivi de disponibilité/occupation des postes de stationnement avion sur une piste
(vocabulaire : AVION, Pos., Camp(ement), O2, AM/PM/VDN, PTR, AVQ, PLEIN, TRG (technicien),
Resa/Indispo/Dispo/Vol). Trois fichiers en interaction :

| Fichier | Rôle |
|---|---|
| **piste - V3.xlsm** | Classeur maître. 3 feuilles : `Situation` (tableau de bord, visible, protégé), `Gestion` (grille de saisie, masquée), `Paramètre` (config/techniciens, masquée). Contient tout le code métier. |
| **Visu.xlsm** | Écran de restitution plein écran (type « écran TV » pour la piste), lancé dans **un second processus Excel** depuis `piste - V3.xlsm`. Se resynchronise seul toutes les 5 s. |
| **DispoVierge.xlsx** | Gabarit vierge (feuilles `10min` + `Enregistrements`) recopié chaque jour pour créer l'archive `Archives\Dispo_AAAA-MM-JJ.xlsx`, qui journalise un instantané (Resa/Indispo/Dispo/Vol) toutes les 10 minutes sur 25 h glissantes. |

### Séquence à l'ouverture de `piste - V3.xlsm`
1. `Workbook_Open` neutralise l'UI (barre de formule, grille, plein écran).
2. Lance un **nouveau** `Excel.Application` (`CreateObject`) et y ouvre `Visu.xlsm`.
3. Vérifie/crée l'archive du jour (`DispoVierge.xlsx` → `Dispo_AAAA-MM-JJ.xlsx` si absente), puis remplit/complète les créneaux de 10 min dans `Enregistrements` avec les totaux `Paramètre!Q30:T30` (Resa/Indispo/Dispo/Vol).
4. Réapplique la mise en forme conditionnelle (`MiseEnForme`) sur `Situation`.

### Flux de saisie
`Gestion` (déprotégée, éditable) → bouton **Valider** (`valid_click`) recopie en valeurs dans
`Situation`, reverrouille les cellules, masque `Gestion`, sauvegarde. Boutons **PlanG/PlanS**
appellent `Visu.xlsm!Module1.Switchvisu` dans l'instance déjà ouverte (bascule feuille
`visu`/`PKG`) via `GetObject`.

### Rafraîchissement de `Visu.xlsm`
Boucle infinie `Application.OnTime` toutes les 5 s (`HorlogeEnA1`) qui met à jour l'horloge et
force `ThisWorkbook.UpdateLink` sur le lien externe pointant vers `piste - V3.xlsm` (lien
relatif, correct). Annulée proprement dans `Workbook_BeforeClose` via un flag `bstop`.

## 2. Verdict sécurité (macro-analyse)

Pas de charge malveillante détectée : aucun `Shell`, `WScript.Shell`, `PowerShell`, appel
réseau/URL, ni chaîne encodée servant de dropper. Les alertes « Suspicious » d'`olevba`
(`Chr`, `Base64 Strings`, `Hex Strings`, `Call`, `Open`, `Run`) sont des **faux positifs** :
`Chr(10)` sert à insérer un retour ligne dans un nom de colonne de tableau, la « chaîne
Base64 » détectée n'est que le mot « Visu » qui matche accidentellement l'heuristique, et
`.Run`/`GetObject`/`CreateObject` ne servent qu'à piloter la seconde instance Excel (usage
légitime d'automation Office). Je n'ai identifié aucune faille d'exécution de code arbitraire.

En revanche, deux points d'hygiène sécurité réels :

- **Protection de feuille sans mot de passe** (`sheet1.xml`/`sheet2.xml` : `<sheetProtection
  sheet="1" .../>` sans attribut `password`/`saltValue`). La protection de `Situation` et
  `Gestion` est purement cosmétique : n'importe qui fait *Révision → Ôter la protection* sans
  mot de passe et modifie directement les cellules « verrouillées », contournant la validation
  UCase et le circuit Gestion→Valider. À corriger avec `.Protect Password:="..."` si l'objectif
  est réellement d'empêcher les modifications accidentelles.
- **Liens externes morts** dans `Visu.xlsm` : 3 des 4 `externalLinks` pointent vers des
  chemins absolus obsolètes d'anciens postes (`/Users/Olivier/Desktop/Piste/Suivi conf
  avion/Visu.xlsm`, `/Users/dasio/Desktop/Eric/Test/piste - V2.xlsm`, etc.) — restes d'une
  V2. Seul le 4ᵉ lien (relatif, vers `piste - V3.xlsm`) est utile. Ces liens morts déclenchent
  l'invite Excel « Mettre à jour les liens » à chaque ouverture et exposent, dans le fichier,
  le nom d'utilisateur Windows/Mac et l'arborescence de postes tiers (fuite d'information
  mineure mais réelle si le fichier est partagé hors de l'équipe). **Action recommandée :**
  Données → Modifier les liens → Rompre le lien sur les 3 références obsolètes.

## 3. Bugs / fragilités identifiées

1. **Aucune gestion d'erreur autour des I/O disque** (`Workbook_Open`, `Workbook_BeforeClose`,
   `Actualiser_Click`). Le `SaveAs` vers `...\Archives\Dispo_AAAA-MM-JJ.xlsx` suppose que le
   dossier `Archives` existe déjà — rien ne le crée (`MkDir` absent) et il n'y a aucun `On
   Error`. Si le dossier est absent, si le fichier journalier est verrouillé par un autre
   poste (partage réseau), ou si le disque est plein, l'erreur VBA non interceptée remonte en
   plein `Workbook_Open`/`Workbook_BeforeClose` : le classeur peut rester dans un état
   incohérent (barre de formule masquée, plein écran actif, fichier non sauvegardé), voire
   bloquer la fermeture d'Excel. C'est le point le plus critique du point de vue robustesse.

2. **Code dupliqué trois fois à l'identique** : le bloc « vérifier/créer l'archive du jour +
   remplir les créneaux de 10 min » (~30 lignes) est copié-collé dans `Workbook_Open`,
   `Workbook_BeforeClose` et `Actualiser_Click`. Un correctif appliqué dans une des trois
   copies (ce qui semble déjà être arrivé vu les commentaires `'*****-----Modif----------*****`
   et le code mort en commentaire) risque de ne pas être répercuté ailleurs. À factoriser en
   une seule `Sub` appelée aux trois endroits.

3. **`GetObject(Chemin)` sur `Visu.xlsm` dans `PlanG_Click`/`PlanS_Click`** : si l'utilisateur
   clique sur le bouton de bascule alors que `Visu.xlsm` a été fermé entre-temps (fermé
   manuellement, instance Excel plantée), le comportement VBA standard de `GetObject` sur un
   fichier non ouvert est de **rouvrir le fichier** (potentiellement dans une nouvelle instance
   cachée) plutôt que de renvoyer une erreur propre — risque de doublons d'instances Excel
   fantômes/invisibles et de conflits de verrou fichier. Aucun test `Nothing`/`On Error` ne
   protège cet appel malgré le commentaire `If Not Wb Is Nothing`.

4. **Deux instances Excel indépendantes** (`piste - V3.xlsm` + un second `Excel.Application`
   pour `Visu.xlsm`) au lieu d'une seconde fenêtre de la même instance. Coût mémoire/CPU
   inutile (~150-250 Mo et plusieurs secondes de lancement pour un simple miroir en lecture),
   et un `Excel.exe` orphelin peut rester actif en tâche de fond si `Visu.xlsm` n'est pas fermé
   proprement par le code (crash, fermeture forcée, fin de session Windows) puisque le `OnTime`
   auto-réarmé toutes les 5 s n'est annulé que dans `Workbook_BeforeClose`.

5. **Fraîcheur des données du tableau de bord `Visu.xlsm`** : `UpdateLink` ne relit que la
   **dernière version sauvegardée sur disque** de `piste - V3.xlsm`. Or ce dernier n'enregistre
   qu'à des points précis (`Actualiser`, `Valider`, fermeture) — pas à chaque frappe. Un agent
   qui modifie `Gestion` sans cliquer « Valider »/« Actualiser » ne verra donc **aucune mise à
   jour** sur l'écran `Visu` pendant plusieurs minutes malgré le rafraîchissement de 5 s, ce qui
   peut être trompeur si l'écran est présenté comme « temps réel ».

6. **Chemins de fichiers codés en dur** (`"piste - V3.xlsm"`, `"\Visu.xlsm"`,
   `"\DispoVierge.xlsx"`) plutôt que dérivés de `ThisWorkbook.Name`/relations dynamiques. Le
   nom même du fichier (« V3 ») trahit un historique de renommage (V2 → V3, visible dans les
   liens externes morts de `Visu.xlsm`) : un futur renommage vers une V4 cassera silencieusement
   tous les appels `Workbooks.Open`/`GetObject`/`UpdateLink` qui référencent le nom en clair.

7. **Boucle de recopie des créneaux de 10 min** (`For i = 1 To 150`, avec boucle interne `For j`)
   effectue des accès cellule-par-cellule via COM (`Cells(...)`) plutôt que par tableau VBA en
   mémoire — jusqu'à 150 × 4 lectures/écritures COM à chaque ouverture/fermeture/clic
   « Actualiser ». La logique de remplissage avant/arrière est correcte (propage le total du
   jour sur les créneaux futurs, comble les trous des créneaux passés), mais son implémentation
   COM cellule-à-cellule est nettement plus lente qu'un passage par `Variant` array, et
   `ScreenUpdating` n'est désactivé que dans `Actualiser_Click` (pas dans `Workbook_Open`/
   `Workbook_BeforeClose`), ce qui peut produire un scintillement visible à l'ouverture/fermeture.

8. **Calcul itératif activé** (`<calcPr calcId="191029" iterate="1"/>`) sur les deux classeurs
   `.xlsm`. Cela autorise silencieusement les références circulaires au lieu de les signaler par
   une erreur Excel — utile si une formule circulaire est volontaire quelque part, mais masque
   aussi toute référence circulaire introduite par erreur lors d'une future modification (aucune
   alerte ne remontera).

## 4. Pistes d'optimisation

- Factoriser le bloc « archive + créneaux 10 min » dans une unique procédure partagée.
- Créer le dossier `Archives`/`Archives\Backup` par code (`If Dir(...) = "" Then MkDir ...`) au
  lieu de supposer son existence.
- Remplacer les boucles `Cells(...)` par une lecture/écriture en une fois via un tableau
  `Variant` (`Range(...).Value`), surtout pour les 150 itérations exécutées à chaque
  ouverture/fermeture/actualisation.
- Envisager de piloter `Visu.xlsm` comme une fenêtre de la **même** instance Excel (`Windows`
  API / `NewWindow`) plutôt qu'un second `Excel.Application`, pour réduire la consommation
  mémoire et éviter les processus orphelins.
- Nettoyer les 3 liens externes morts de `Visu.xlsm` (Données → Modifier les liens).
- Ajouter un vrai mot de passe sur les protections de feuille si l'objectif est d'empêcher les
  modifications hors circuit (actuellement contournable en un clic).
- Ajouter systématiquement `On Error Resume Next`/`On Error GoTo` avec message utilisateur
  explicite autour des opérations fichier (ouverture/sauvegarde d'archive), pour éviter qu'une
  erreur silencieuse ne laisse le classeur dans un état incohérent (plein écran actif, barre de
  formule masquée) sans message clair pour l'utilisateur.

## 5. Ce qui n'a pas pu être vérifié ici

Analyse purement statique (pas d'Excel/Windows disponible dans cet environnement) : les
scénarios d'erreur ci-dessus (dossier `Archives` absent, fichier verrouillé par un autre poste,
`Visu.xlsm` fermé puis bouton de bascule cliqué) sont déduits de la lecture du code, pas
reproduits en conditions réelles. À confirmer par un test manuel avant toute correction en
production.
