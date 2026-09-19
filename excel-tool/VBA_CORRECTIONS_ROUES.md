# Colonne « Roues » — ce qui a été fait et ce qu'il reste à coller

## Correctif du 19/09 (2) : format Date + mise en forme disparue + liste 1/2/3 ajoutée

Après ton retour (« liste 1/2/3, et j'ai testé 1/2/3, il n'y a pas les mef ») — deux bugs, tous
les deux de mon fait, trouvés cette fois en testant réellement le fichier avec un moteur tableur
(LibreOffice, installé et piloté ici pour l'occasion) plutôt qu'en relisant juste le XML :

1. **Les cellules Roues avaient hérité un format « Date »** (`dd/mm/yy;@`) de la colonne GILETS,
   dont j'avais cloné le style par erreur lors de l'insertion. Résultat : Excel affichait tes 1/2/3
   comme des dates (« 01/01/00 » etc.) au lieu du chiffre — la case semblait ne rien faire.
   Corrigé : format remis en « Standard » sur les 25 lignes, dans les deux feuilles, sans toucher
   au format de GILETS (qui partageait le même style — j'ai dû créer des styles dédiés pour ne
   pas y toucher).
2. **La mise en forme conditionnelle (vert/orange/rouge) que j'avais ajoutée avait disparu** du
   fichier — perdue au fil des réenregistrements Excel successifs. Cause racine trouvée : la
   macro `MiseEnForme` (qui supprime **toute** la mise en forme de Situation et la reconstruit à
   chaque ouverture du classeur) ne connaissait pas Roues, donc rien ne la recréait après le
   nettoyage. Je l'avais mise uniquement dans le fichier, pas dans le VBA — erreur de ma part,
   il fallait les deux. Recréée dans le fichier ; **il faut aussi l'ajouter dans `MiseEnForme`**
   (§ ci-dessous) sinon elle redisparaîtra à la prochaine ouverture dans Excel.
3. **Liste déroulante 1/2/3 ajoutée** sur Roues (comme demandé) — validation de saisie simple,
   sans toucher à `Paramètre`.

Revalidé cette fois avec un vrai moteur de calcul (pas juste une relecture XML) : ouverture du
fichier réel, saisie simulée de 1/2/3/4/vide dans Roues, lecture de la règle de mise en forme qui
matche effectivement et de sa couleur résultante — vert/orange/rouge confirmés, 4 et vide ne
déclenchent rien. `vbaProject.bin` et la feuille Paramètre confirmés identiques à l'octet près.

### VBA à ajouter dans `MiseEnForme` (Module2) — obligatoire pour que la couleur survive au prochain Workbook_Open

Colle ce bloc juste **avant** les deux dernières lignes de la macro (`Application.ScreenUpdating
= True` / `Sheets("Situation").Protect`) :

```vba
    Range("R7:R31").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$R7=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0.5)
        .Color = 5287936
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Range("R7:R31").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$R7=2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0.5)
        .Color = 39423
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Range("R7:R31").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$R7=3"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 90
        .Gradient.ColorStops.Clear
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(0.5)
        .Color = 255
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
```

(Même technique — dégradé blanc→couleur→blanc — que toutes les autres règles de cette macro ;
`.Color = 5287936` = vert, `39423` = orange (même orange que la colonne O2), `255` = rouge (même
rouge que « 21 »/« 22 »). Pas besoin de toucher à Gestion : sa mise en forme est statique, pas
reconstruite par macro, donc la version déjà dans le fichier suffit et va tenir.)

Pour tester : ferme et rouvre le fichier dans Excel après avoir collé ce bloc, remets une valeur
dans Roues, vérifie que la couleur apparaît toujours après un Ctrl+S puis réouverture.

## Correctif du 19/09 : liste déroulante « Masquer » restait collée sur Roues

Après ton retour (« dans Roues je ne peux mettre que x ») : la colonne Masquer avait, en plus de
sa mise en forme, une **liste déroulante de validation** (blanc ou « X », source
`Paramètre!$B$46:$B$47`) enregistrée dans une zone du fichier que je n'avais pas vérifiée lors de
l'insertion (validations "étendues" `x14:dataValidations`, distinctes du bloc `dataValidations`
classique que j'avais bien contrôlé). Cette validation ciblait la position R par coordonnées
(`R7:R31` / `R6:R31`), donc quand Roues a pris la place de la colonne R, elle en a hérité — d'où
le blocage sur « x » uniquement. Corrigé : la validation est remise sur Masquer, à sa position
actuelle (colonne S). Roues n'a plus aucune contrainte de saisie (texte libre, comme prévu).
Revalidé (XML, diff binaire, VBA intact) et repoussé sur le repo — retélécharge le fichier avant
de retester.

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
