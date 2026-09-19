# Rafraîchissement Visu : push au lieu de poll — VBA à coller

Ça touche les deux fichiers (`piste - V3.xlsm` et `Visu.xlsm`), en VBA uniquement — aucune
structure de feuille/tableau/mise en forme n'est modifiée cette fois, donc rien à corriger côté
fichier de mon côté, tout se fait en collant ces blocs dans l'éditeur VBA.

## Ce que ça change

Aujourd'hui, `Visu.xlsm` fait **une seule chose toutes les 5 secondes** : mettre à jour l'heure
affichée ET relire `piste - V3.xlsm` sur le disque (`UpdateLink`). La relecture disque est plus
coûteuse que l'affichage de l'heure, et inutile la plupart du temps puisque les données ne
changent pas toutes les 5 secondes — d'où l'impression de décalage entre une action sur piste-V3
et sa répercussion sur Visu.

Nouveau fonctionnement :
- **Horloge** : toujours mise à jour toutes les 5 s, léger, ne touche pas au disque.
- **Données (le tableau)** : ne sont relues en fond que toutes les **30 s** (filet de sécurité),
  MAIS piste-V3 pousse maintenant un rafraîchissement immédiat vers Visu juste après chaque
  sauvegarde réelle (bouton **Actualiser**, bouton **Valider**, ouverture, fermeture) — donc dans
  l'usage normal, l'écran se met à jour à l'instant où tu cliques, pas au prochain sondage.

## 1. `Visu.xlsm` — remplace tout le code de `ThisWorkbook`

**Avant** (le code actuel) :
```vba
Dim bstop As Boolean
Dim HeureProchainAppel
Private Sub Workbook_BeforeClose(Cancel As Boolean)
 bstop = True
 HorlogeEnA1
 Application.DisplayAlerts = False
 Application.Quit
End Sub
Private Sub Workbook_Open()

    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.DisplayFullScreen = True
    
    Range("A1").Select
    
    HorlogeEnA1
 End Sub
Public Sub HorlogeEnA1()
Dim hrs As Date
Dim cor As String
Application.DisplayAlerts = False
Application.ScreenUpdating = False
 If bstop = True Then
 'Annuler le paramétrage du OnTime programmé précédemment.
 Application.OnTime EarliestTime:=HeureProchainAppel, _
  Procedure:="ThisWorkbook.HorlogeEnA1", Schedule:=False
  Exit Sub
 End If
 
hrs = Format(Now, "HH:MM:SS")
cor = Format("16/11/2012 00:02", "HH:MM:SS")
 
Sheets("visu").Range("O4").Value = Format(hrs + cor, "HH:MM:SS" & " Z")

On Error GoTo Error_MayCauseAnError
ThisWorkbook.UpdateLink Name:=ThisWorkbook.Path & "\piste - V3.xlsm", Type:=xlExcelLinks
Error_MayCauseAnError:

 HeureProchainAppel = Now + TimeValue("00:00:05")
 Application.OnTime HeureProchainAppel, "ThisWorkbook.HorlogeEnA1", False

 Range("A1:Z1").Select

 End Sub
```

**Après** :
```vba
Dim bstop As Boolean
Dim HeureProchainTic
Dim HeureProchainRafraichissement

Private Sub Workbook_BeforeClose(Cancel As Boolean)
 bstop = True
 TicHorloge
 RafraichirDonnees
 Application.DisplayAlerts = False
 Application.Quit
End Sub

Private Sub Workbook_Open()

    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.DisplayFullScreen = True
    
    Range("A1").Select
    
    TicHorloge
    RafraichirDonnees
 End Sub

' Met à jour uniquement l'horloge affichée. Léger, cadence 5 s, ne touche pas au disque.
Public Sub TicHorloge()
Dim hrs As Date
Dim cor As String
Application.ScreenUpdating = False
 If bstop = True Then
 Application.OnTime EarliestTime:=HeureProchainTic, _
  Procedure:="ThisWorkbook.TicHorloge", Schedule:=False
  Exit Sub
 End If

hrs = Format(Now, "HH:MM:SS")
cor = Format("16/11/2012 00:02", "HH:MM:SS")

Sheets("visu").Range("O4").Value = Format(hrs + cor, "HH:MM:SS" & " Z")

 HeureProchainTic = Now + TimeValue("00:00:05")
 Application.OnTime HeureProchainTic, "ThisWorkbook.TicHorloge", False
End Sub

' Relit piste - V3.xlsm sur le disque (UpdateLink). Plus coûteux que TicHorloge, donc
' cadence espacée (30 s, en filet de sécurité) -- et rappelable immédiatement depuis
' piste - V3.xlsm juste après une sauvegarde, via :
'   GetObject(...\Visu.xlsm").Parent.Run "Visu.xlsm!ThisWorkbook.RafraichirDonnees"
Public Sub RafraichirDonnees()
' Annule le rafraîchissement déjà programmé (s'il y en a un), pour ne pas empiler des
' appels en double quand on est déclenché en push depuis piste-V3 entre deux tics.
On Error Resume Next
Application.OnTime EarliestTime:=HeureProchainRafraichissement, _
    Procedure:="ThisWorkbook.RafraichirDonnees", Schedule:=False
On Error GoTo 0

Application.DisplayAlerts = False
Application.ScreenUpdating = False

 If bstop = True Then Exit Sub

On Error GoTo Error_MayCauseAnError
ThisWorkbook.UpdateLink Name:=ThisWorkbook.Path & "\piste - V3.xlsm", Type:=xlExcelLinks
Error_MayCauseAnError:

 HeureProchainRafraichissement = Now + TimeValue("00:00:30")
 Application.OnTime HeureProchainRafraichissement, "ThisWorkbook.RafraichirDonnees", False

 Range("A1:Z1").Select
 Application.ScreenUpdating = True
 Application.DisplayAlerts = True
End Sub
```

(Pour changer la cadence du filet de sécurité, une seule ligne à toucher :
`HeureProchainRafraichissement = Now + TimeValue("00:00:30")`.)

## 2. `piste - V3.xlsm` — Module1 : nouvelle procédure

Ajoute, à la suite de `ArchiverDispoDuJour` (ou de `ExisteFichier`/`WbOpen`) :

```vba
' Pousse un rafraîchissement immédiat vers l'écran Visu (s'il est ouvert), plutôt que
' d'attendre son prochain sondage de fond (30 s). Sans effet si Visu n'est pas ouvert.
Sub PousserRafraichissementVisu()
    On Error Resume Next
    Dim WbVisu As Object
    Set WbVisu = GetObject(ThisWorkbook.Path & "\Visu.xlsm")
    If Not WbVisu Is Nothing Then WbVisu.Parent.Run "Visu.xlsm!ThisWorkbook.RafraichirDonnees"
    On Error GoTo 0
End Sub
```

## 3. `piste - V3.xlsm` — Module1 : appel dans `ArchiverDispoDuJour`

**Avant** (fin de la procédure) :
```vba
    ActiveWorkbook.Save
    Workbooks(myName).Activate
    ActiveWorkbook.Save

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
```

**Après** :
```vba
    ActiveWorkbook.Save
    Workbooks(myName).Activate
    ActiveWorkbook.Save

    Call PousserRafraichissementVisu

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
```

## 4. `piste - V3.xlsm` — Feuil3 (`Gestion`) : appel dans `valid_click`

C'est le bouton **Valider** — le vrai point d'entrée des changements de statut avion, donc le
plus important des deux déclencheurs.

**Avant** (fin de la procédure) :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
```

**Après** :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    Call PousserRafraichissementVisu

Application.ScreenUpdating = True
End Sub
```

## Limite connue (pas nouvelle, déjà présente avant ce correctif)

`PousserRafraichissementVisu` réutilise exactement la même technique que les boutons
PlanG/PlanS existants (`GetObject` sur le chemin de `Visu.xlsm`). Comportement VBA standard : si
Visu.xlsm a été fermé entre-temps, `GetObject` peut le rouvrir silencieusement en arrière-plan
plutôt que d'échouer proprement — ce n'est pas quelque chose que ce correctif introduit, c'est
déjà le comportement des boutons de bascule de vue.

## Comment tester

1. Sauvegarde tes fichiers actuels avant de coller.
2. Colle les 4 blocs ci-dessus (2 dans `Visu.xlsm`, 2 dans `piste - V3.xlsm`).
3. Ouvre les deux fichiers, modifie une valeur dans `Gestion`, clique **Valider** : l'écran Visu
   doit se mettre à jour quasi instantanément, sans attendre.
4. Laisse tourner sans rien toucher : l'heure sur Visu doit continuer à avancer normalement
   toutes les 5 s.
5. Ferme `piste - V3.xlsm` : `Visu.xlsm` doit se fermer proprement (comme avant), sans laisser de
   processus Excel fantôme.
