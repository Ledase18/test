# Colonne « Roues » — ce qui a été fait et ce qu'il reste à coller

## Ce qui a déjà été modifié dans `piste - V3.xlsm` (fait, poussé sur le repo)

J'ai directement édité le classeur (XML interne, zip) — validé par re-parsing XML, réouverture
openpyxl, et comparaison **cellule par cellule** avec l'original (0 écart en dehors de ce qui
est décrit ci-dessous). Le `vbaProject.bin` (macros) est **resté strictement identique**
(comparaison d'octets) : je ne l'ai pas touché, je ne peux pas l'éditer sans Excel pour le
valider — d'où le fichier VBA à coller toi-même ci-dessous.

- **Emplacement retenu : entre DEM VAL et Masquer**, pas entre GILETS et DEM VAL. Sur
  `Situation`, l'en-tête affiché à la colonne juste après GILETS montre « V » mais c'est en
  réalité la colonne « DEM VAL » (le libellé affiché est trompeur — bug préexistant de nommage,
  pas introduit par moi, je n'y ai pas touché). Le fichier `Visu.xlsm` récupère aujourd'hui les
  colonnes de `Situation` jusqu'à DEM VAL inclus, mais jamais au-delà (jamais Masquer). En
  insérant Roues juste après DEM VAL, **aucune formule de `Visu.xlsm` n'a besoin d'être
  modifiée** — Visu continue de lire exactement les mêmes colonnes qu'avant, Roues lui est
  invisible par construction. C'est le seul emplacement qui garantit ça sans toucher à Visu.
- Nouvelle colonne **Roues** insérée dans les tableaux structurés `Tableau1` (Situation) et
  `Tableau2` (Gestion), largeur alignée sur les colonnes de statut courtes du tableau (comme
  « 21 »/« 22 »).
- **Mise en forme conditionnelle** ajoutée sur `Roues` (mêmes techniques que le reste du
  tableau — dégradé blanc→couleur→blanc, comme les colonnes « 21 »/« 22 »/O2) :
  - 1 → vert
  - 2 → orange (même orange que la colonne O2)
  - 3 → rouge (même rouge que les colonnes « 21 »/« 22 »)
- La formule de la colonne « Num av. » (calcul de préfixe « x » selon Masquer) a été corrigée :
  elle testait `R7="x"` (Masquer était en colonne R) ; Masquer étant décalé en colonne S, la
  formule teste maintenant `S7="x"`. Sans ça, cette formule aurait continué à lire l'ancienne
  position de Masquer — désormais occupée par Roues — et se serait cassée silencieusement.
  Vérifié colonne par colonne sur les 25 lignes (7 à 31), dans les deux feuilles.
  - Comme demandé, cette colonne **reste éditable directement sur `Situation`** (comme GILETS)
    après validation, plutôt que verrouillée comme DEM VAL/Masquer.
- Boutons (Actualiser, PlanS, exit1, FullScr1 sur Situation ; PlanG, Tri, valid, Annul, eff,
  Exit2 sur Gestion) : leur position (ancrage colonne) a été décalée pour rester alignés
  visuellement, sans toucher à leur macro associée.
- `calcChain.xml` supprimé du fichier (liste interne de recalcul) — Excel le régénère
  automatiquement à l'ouverture ; c'est une pratique standard pour éviter d'avoir à corriger à
  la main une troisième structure alors qu'aucune formule affectée n'y était listée par
  cellule propre.

## Ce qu'il reste à coller dans l'éditeur VBA (Alt+F11)

Ces 6 procédures contiennent des plages codées en dur (`"...R31"`, `"Q7:R31"`, `"U23"`) qui
supposaient Masquer en colonne R. Comme les plages Excel ne se corrigent pas toutes seules quand
une colonne est insérée par édition directe du fichier (contrairement à un vrai clic-droit
« Insérer une colonne » dans Excel), il faut les mettre à jour à la main. Fais une copie de
sauvegarde du fichier avant de coller.

### 1. Feuil2 (feuille `Situation`) — `Worksheet_Change`

Une seule ligne à changer (`P7:R31` → `P7:S31`, pour que la mise en majuscules continue de
couvrir GILETS/DEM VAL/Masquer — Roues est numérique, la majuscule n'a pas d'effet dessus, sans
danger qu'elle soit incluse) :

**Avant :**
```vba
If b = False And Target.Count = 1 And Not Intersect(Target, Range("D7:F31,J7:L31,N7:N31,P7:R31")) Is Nothing Then
```
**Après :**
```vba
If b = False And Target.Count = 1 And Not Intersect(Target, Range("D7:F31,J7:L31,N7:N31,P7:S31")) Is Nothing Then
```

### 2. Feuil3 (feuille `Gestion`) — `Worksheet_Change`

Même correctif que ci-dessus :

**Avant :**
```vba
If b = False And Target.Count = 1 And Not Intersect(Target, Range("E7:F31,J7:L31,N7:N31,P7:R31")) Is Nothing Then
```
**Après :**
```vba
If b = False And Target.Count = 1 And Not Intersect(Target, Range("E7:F31,J7:L31,N7:N31,P7:S31")) Is Nothing Then
```

### 3. Feuil3 (feuille `Gestion`) — `valid_click`

Deux changements : la copie de la date « MàJ FDS/KANNAD » (`U23`→`V23`, puisqu'elle a été
décalée par l'insertion de colonne), et la plage verrouillée après validation (Roues exclue,
comme GILETS, donc éditable directement sur Situation) :

**Avant :**
```vba
Private Sub valid_click()
Application.ScreenUpdating = False
    Sheets("Situation").Unprotect
    Range("Tableau2").Select
    Selection.Copy
    Sheets("Situation").Select
    Sheets("Situation").Range("B7").Select
    ActiveSheet.Paste
    
    ''''''''''''''''''''''''''''''''
    Sheets("Gestion").Select
    Sheets("Gestion").Range("U23").Select
    Selection.Copy
    Sheets("Situation").Select
    Sheets("Situation").Range("U23").Select
    ActiveSheet.Paste
    ''''''''''''''''''''''''''''''''
    
    Sheets("Situation").Range("B7:B31,G7:I31,L7:L31,N7:O31,Q7:R31").Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    

    Sheets("Situation").Range("D7").Select
    Sheets("Gestion").Visible = False
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
```

**Après :**
```vba
Private Sub valid_click()
Application.ScreenUpdating = False
    Sheets("Situation").Unprotect
    Range("Tableau2").Select
    Selection.Copy
    Sheets("Situation").Select
    Sheets("Situation").Range("B7").Select
    ActiveSheet.Paste
    
    ''''''''''''''''''''''''''''''''
    Sheets("Gestion").Select
    Sheets("Gestion").Range("V23").Select
    Selection.Copy
    Sheets("Situation").Select
    Sheets("Situation").Range("V23").Select
    ActiveSheet.Paste
    ''''''''''''''''''''''''''''''''
    
    Sheets("Situation").Range("B7:B31,G7:I31,L7:L31,N7:O31,Q7:Q31,S7:S31").Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    

    Sheets("Situation").Range("D7").Select
    Sheets("Gestion").Visible = False
    
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
```

### 4. Feuil3 (feuille `Gestion`) — `Annul_Click`

Une ligne (`D7:R31` → `D7:S31`, pour que l'annulation déverrouille bien toute la ligne y
compris Roues) :

**Avant :**
```vba
    Range("B7:B31,D7:R31").Select
    Sheets("Gestion").Unprotect
    Selection.Locked = False
```
**Après :**
```vba
    Range("B7:B31,D7:S31").Select
    Sheets("Gestion").Unprotect
    Selection.Locked = False
```

### 5. Feuil3 (feuille `Gestion`) — `eff_Click`

Une ligne (`D7:R31` → `D7:S31`, pour que l'effacement vide bien toute la ligne y compris
Roues) :

**Avant :**
```vba
    Range("B7:B31,D7:R31").Select
    Application.CutCopyMode = False
    Selection.ClearContents
```
**Après :**
```vba
    Range("B7:B31,D7:S31").Select
    Application.CutCopyMode = False
    Selection.ClearContents
```

### 6. Module2 — `switchp`

Deux occurrences (`D7:R31` → `D7:S31`) :

**Avant :**
```vba
    Sheets("Gestion").Visible = True
    Sheets("Gestion").Unprotect
    Sheets("Gestion").Select
    Range("B7:B31,D7:R31").Select
    Selection.ClearContents
    Sheets("Gestion").Unprotect
    Range("C7:C31").Select
    Selection.Locked = False
    Sheets("Situation").Select
    Range("Tableau1").Select
    Selection.Copy
    Range("D7").Select
    Sheets("Gestion").Select
    Range("B7").Select
    ActiveSheet.Paste
    Sheets("Gestion").Unprotect
    Range("B7:B31,D7:R31").Select
    Sheets("Gestion").Unprotect
    Selection.Locked = False
```
**Après :**
```vba
    Sheets("Gestion").Visible = True
    Sheets("Gestion").Unprotect
    Sheets("Gestion").Select
    Range("B7:B31,D7:S31").Select
    Selection.ClearContents
    Sheets("Gestion").Unprotect
    Range("C7:C31").Select
    Selection.Locked = False
    Sheets("Situation").Select
    Range("Tableau1").Select
    Selection.Copy
    Range("D7").Select
    Sheets("Gestion").Select
    Range("B7").Select
    ActiveSheet.Paste
    Sheets("Gestion").Unprotect
    Range("B7:B31,D7:S31").Select
    Sheets("Gestion").Unprotect
    Selection.Locked = False
```

## Ce qui n'a **pas** été touché, volontairement

- `Visu.xlsm` : aucun changement. Confirmé par relecture de ses formules — elles s'arrêtent à
  DEM VAL, jamais au-delà, donc Roues (insérée après DEM VAL) lui reste invisible sans rien
  faire de plus.
- `Paramètre` (feuille) : aucune formule n'y référence les colonnes au-delà de K sur
  `Situation` — vérifié, aucun changement nécessaire, et confirmé identique à l'octet près
  entre l'ancien et le nouveau fichier.
- Le mot de passe de protection des feuilles (déjà signalé absent dans `ANALYSE.md`) — hors
  périmètre de cette tâche.

## Comment tester avant de garder

1. Sauvegarde le fichier actuel avant de coller.
2. Colle les 6 correctifs ci-dessus.
3. Ouvre le fichier : vérifie que la colonne **Roues** apparaît bien entre DEM VAL et Masquer
   sur `Situation` et sur `Gestion` (Ctrl+G pour l'afficher), avec le dégradé vert/orange/rouge
   quand tu saisis 1, 2 ou 3.
4. Modifie Roues directement sur `Situation` (hors Gestion) : la valeur doit rester modifiable
   même après un clic sur **Valider**.
5. Ouvre `Visu.xlsm` : confirme que Roues n'apparaît nulle part dessus et que le reste de
   l'affichage (jusqu'à DEM VAL) est inchangé.
6. Clique **Valider**, **Annuler**, **Effacer**, Ctrl+G (switch Gestion/Situation) : confirme
   qu'aucune erreur VBA ne remonte et que toutes les colonnes (y compris Masquer et O2 Psi, qui
   ont juste glissé d'une colonne) se comportent comme avant.
