# Lancer la TV au démarrage de piste-V3 — VBA à coller

Contexte : un seul PC, deuxième écran = la TV. Donc "afficher la TV" = ouvrir
`situation-avions.html` dans un navigateur, positionné sur le 2ᵉ écran, en plein écran, au
lancement de `piste - V3.xlsm` — exactement ce que faisait l'ancien bloc qui ouvrait `Visu.xlsm`
dans une deuxième instance Excel, mais avec un navigateur à la place.

## Ce que fait `LancerTV`

- Repère la largeur de l'écran principal (`GetSystemMetrics`) et l'utilise comme position X du
  navigateur — hypothèse standard "2ᵉ écran étendu à droite, même hauteur". **Si ta TV est
  positionnée autrement** (au-dessus, à gauche, décalage vertical...), ajuste `TV_OFFSET_X` /
  `TV_OFFSET_Y` en dur dans le code ci-dessous — les vraies valeurs sont visibles dans Windows,
  Paramètres → Système → Affichage → schéma des écrans (survole chaque écran pour voir sa
  position en pixels).
- Cherche Chrome ou Edge aux emplacements d'installation standards.
- Lance le navigateur trouvé en `--kiosk` (plein écran total, sans barre d'adresse ni onglets —
  l'équivalent du plein écran que Visu imposait) positionné sur le 2ᵉ écran.
- Si ni Chrome ni Edge n'est trouvé à un emplacement standard : ouvre avec le navigateur par
  défaut (`FollowHyperlink`) sans position/plein écran automatique — à faire à la main la
  première fois dans ce cas.

## `piste - V3.xlsm` — Module1 : nouvelle procédure

Ajoute ce bloc à la suite de `GenererHTMLSituation` :

```vba
' ============================================================
' Lancement TV -- ouvre situation-avions.html en plein écran sur le 2e
' écran au démarrage de piste-V3, à la place de l'ancien lancement de
' Visu.xlsm dans une 2e instance Excel.
' ============================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const SM_CXSCREEN As Long = 0

Sub LancerTV()
    On Error GoTo ErrHandler

    Dim htmlPath As String
    htmlPath = ThisWorkbook.Path & "\situation-avions.html"

    If Dir(htmlPath) = "" Then Call GenererHTMLSituation
    If Dir(htmlPath) = "" Then Exit Sub ' toujours rien -- rien à afficher

    ' Position du 2e écran -- hypothèse "étendu à droite, y=0".
    ' À corriger ici si ta disposition d'écrans est différente.
    Dim tvOffsetX As Long, tvOffsetY As Long
    tvOffsetX = GetSystemMetrics(SM_CXSCREEN)
    tvOffsetY = 0

    Dim candidates() As String
    candidates = Split( _
        Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe|" & _
        Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe|" & _
        Environ$("LocalAppData") & "\Google\Chrome\Application\chrome.exe|" & _
        Environ$("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedge.exe|" & _
        Environ$("ProgramFiles") & "\Microsoft\Edge\Application\msedge.exe", "|")

    Dim browserPath As String
    Dim i As Long
    For i = LBound(candidates) To UBound(candidates)
        If Dir(candidates(i)) <> "" Then
            browserPath = candidates(i)
            Exit For
        End If
    Next i

    Dim fileUrl As String
    fileUrl = "file:///" & Replace(htmlPath, "\", "/")

    If browserPath <> "" Then
        Dim cmd As String
        cmd = """" & browserPath & """ --new-window --window-position=" & tvOffsetX & "," & tvOffsetY & _
              " --kiosk """ & fileUrl & """"
        Shell cmd, vbNormalFocus
    Else
        ThisWorkbook.FollowHyperlink htmlPath
    End If

    Exit Sub

ErrHandler:
    ' Ne bloque jamais l'ouverture de piste-V3 pour un problème d'affichage TV.
    Debug.Print "LancerTV : " & Err.Description
End Sub
```

## `piste - V3.xlsm` — ThisWorkbook : appel dans `Workbook_Open`

Remplace le bloc qui ouvrait `Visu.xlsm` dans une deuxième instance Excel.

**Avant** :
```vba
    Application.DisplayAlerts = False

    Set appExcel = CreateObject("Excel.Application")
    appExcel.Visible = True
    Application.CutCopyMode = False

    Application.DisplayAlerts = False
    Set wbExcel = appExcel.Workbooks.Open(ThisWorkbook.Path & "\Visu.xlsm", UpdateLinks:=0)
    Application.DisplayAlerts = True

    Call ArchiverDispoDuJour

    Sheets("Situation").Select
    Call MiseEnForme
    Sheets("Situation").Range("B7").Select

End Sub
```

**Après** :
```vba
    Application.DisplayAlerts = False

    Call ArchiverDispoDuJour
    Call LancerTV

    Sheets("Situation").Select
    Call MiseEnForme
    Sheets("Situation").Range("B7").Select

End Sub
```

(Les déclarations `Dim appExcel As Excel.Application`, `Dim wbExcel As Excel.Workbook`,
`Dim wsExcel As Excel.Worksheet` en haut de `Workbook_Open` deviennent inutiles — tu peux les
laisser ou les retirer, sans effet.)

## Comportement inchangé par rapport à l'ancien système

Comme l'ancien lancement de `Visu.xlsm`, `LancerTV` s'exécute à **chaque** ouverture de
`piste - V3.xlsm`, sans vérifier si un navigateur/onglet TV est déjà ouvert (l'ancien code ne le
faisait pas non plus pour Visu). Si tu rouvres `piste - V3.xlsm` en cours de journée, ça ouvrira
une nouvelle fenêtre TV en plus de la précédente — même limite qu'avant, je n'en ajoute pas de
nouvelle.

## Point non vérifiable d'ici

Je n'ai pas de Windows/Chrome/Edge disponible dans cet environnement pour tester réellement le
`Shell` et le positionnement multi-écran — c'est une technique standard (arguments de ligne de
commande Chrome/Edge documentés), mais à valider chez toi :
- Si `--window-position`/`--kiosk` sont ignorés (ça arrive quand une fenêtre Chrome utilisant le
  même profil est déjà ouverte ailleurs : Chrome route parfois le nouveau lancement vers le
  processus existant plutôt que d'en créer un nouveau avec les mêmes options) : dis-le-moi, la
  solution est d'ajouter `--user-data-dir="C:\...\ChromeKiosk"` à la commande pour forcer un
  profil dédié, garantissant une fenêtre neuve qui respecte les options.
- Si ton PC utilise un layout d'écrans qui n'est pas "TV à droite, même hauteur" : ajuste
  `tvOffsetX`/`tvOffsetY` avec les vraies coordonnées (Windows Paramètres d'affichage).
