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
- Lance le navigateur trouvé en `--start-fullscreen` (plein écran, comme le faisait Visu)
  positionné sur le 2ᵉ écran. **Pas `--kiosk`** : ce mode-là est conçu pour des bornes publiques
  et bloque volontairement la sortie (F11/Esc ne fonctionnent pas, seul Alt+F4 ferme la fenêtre)
  — `--start-fullscreen` donne le même rendu visuel mais F11 bascule normalement en fenêtré, donc
  tu peux ressortir et déplacer la fenêtre à la main si besoin.
- Ajoute `--allow-file-access-from-files` : c'est ce qui permet à `situation-avions.html` de se
  relire lui-même sur disque toutes les 5s (`fetch`) sans recharger toute la page -- Chrome/Edge
  bloquent `fetch()` entre fichiers locaux par défaut, ce qui forçait le rechargement complet (le
  flash visible toutes les 5s que tu as signalé) tant que ce drapeau n'était pas là.
- Lance le navigateur avec un profil dédié (`--user-data-dir=...\ChromeTV`) : **c'est le correctif
  de ce tour**. Si Chrome/Edge tournait déjà (même profil par défaut) quand piste-V3 s'ouvre, un
  nouveau lancement se contente d'ouvrir une fenêtre dans le processus existant et **ignore tous
  les drapeaux ci-dessus** -- position, plein écran, et surtout l'accès fichier local, ce qui
  explique que le flash persistait malgré le drapeau ajouté au tour précédent. Un profil séparé
  force une fenêtre réellement neuve qui les respecte.
- Si ni Chrome ni Edge n'est trouvé à un emplacement standard : ouvre avec le navigateur par
  défaut (`FollowHyperlink`) sans position/plein écran automatique — à faire à la main la
  première fois dans ce cas.

## `piste - V3.xlsm` — Module1 : nouvelle procédure

`PtrSafe` sans variante conditionnelle `#If VBA7` : c'est requis depuis Office 2010 (32 et 64
bits confondus), donc une branche `#Else` sans `PtrSafe` ne sert plus à rien en pratique — et
l'éditeur VBA a tendance à surligner cette branche en rouge à tort (faux positif connu du
vérificateur de syntaxe en temps réel sur les blocs `#Else`), donc autant l'omettre.

Ajoute ce bloc à la suite de `GenererHTMLSituation` :

```vba
' ============================================================
' Lancement TV -- ouvre situation-avions.html en plein écran (--start-fullscreen,
' PAS --kiosk : ce dernier bloque la sortie) sur le 2e écran au démarrage de
' piste-V3, à la place de l'ancien lancement de Visu.xlsm dans une 2e instance
' Excel. F11 bascule normalement en fenêtré depuis --start-fullscreen.
' ============================================================

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXSCREEN As Long = 0

' PID du navigateur lancé par LancerTV, utilisé par FermerTV pour le fermer.
' Public (pas juste module) car les boutons Situation/Gestion doivent tous les
' deux pouvoir déclencher LancerTV/FermerTV et lire/écrire la même valeur.
' Reste à 0 tant qu'aucun lancement via Shell n'a réussi dans cette session
' Excel -- si tu fermes puis rouvres piste-V3, la valeur repart à 0 (normal,
' une variable Public ne survit pas à la fermeture du classeur), donc FermerTV
' ne pourra fermer que la fenêtre lancée par la session en cours.
Public tvProcessID As Double

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

    ' Profil Chrome/Edge dédié : si le navigateur tourne déjà (même profil), un
    ' nouveau lancement route vers le processus existant et IGNORE tous les
    ' drapeaux ci-dessous (position, plein écran, accès fichier local). Un
    ' --user-data-dir séparé force une fenêtre neuve qui les respecte vraiment.
    Dim profileDir As String
    profileDir = Environ$("LocalAppData") & "\ChromeTV"

    If browserPath <> "" Then
        Dim cmd As String
        cmd = """" & browserPath & """ --user-data-dir=""" & profileDir & """ --new-window" & _
              " --window-position=" & tvOffsetX & "," & tvOffsetY & _
              " --start-fullscreen --allow-file-access-from-files """ & fileUrl & """"
        ' Shell() en fonction (pas juste "Shell cmd") pour récupérer le PID --
        ' nécessaire pour que FermerTV puisse fermer précisément cette fenêtre.
        tvProcessID = Shell(cmd, vbNormalFocus)
    Else
        ThisWorkbook.FollowHyperlink htmlPath
        ' Pas de PID récupérable via FollowHyperlink (ouverture par le
        ' navigateur par défaut, hors de notre contrôle) -- FermerTV ne pourra
        ' pas fermer cette fenêtre-là automatiquement.
    End If

    Exit Sub

ErrHandler:
    ' Ne bloque jamais l'ouverture de piste-V3 pour un problème d'affichage TV.
    Debug.Print "LancerTV : " & Err.Description
End Sub

' ============================================================
' Fermeture TV -- termine le processus navigateur lancé par LancerTV
' (identifié par tvProcessID). Utilise taskkill /F car Shell() ne donne pas de
' handle de fenêtre exploitable pour un "clic sur la croix" propre -- /F tue le
' process direct, sans confirmation ni sauvegarde à gérer côté navigateur
' (aucune perte : situation-avions.html n'a pas d'état à sauvegarder).
' ============================================================
Sub FermerTV()
    On Error GoTo ErrHandler

    If tvProcessID > 0 Then
        Shell "taskkill /PID " & tvProcessID & " /F", vbHide
        tvProcessID = 0
    Else
        MsgBox "Aucune fenêtre TV n'a été lancée par ce classeur depuis son ouverture " & _
               "(ou elle a été ouverte via le navigateur par défaut, non fermable " & _
               "automatiquement). Ferme-la manuellement si besoin.", vbInformation, "Fermer TV"
    End If

    Exit Sub

ErrHandler:
    Debug.Print "FermerTV : " & Err.Description
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

## Ajout de 2 boutons "Lancer TV" / "Fermer TV" sur Situation et Gestion

Objectif : pouvoir déclencher `LancerTV` / `FermerTV` à la main depuis chaque feuille, sans
attendre l'ouverture/fermeture du classeur (utile si tu dois réafficher la TV en cours de
journée, ou la fermer sans fermer piste-V3).

**Je ne les insère pas par édition XML directe du fichier**, contrairement aux repositionnements
de boutons déjà faits pour la colonne Roues. Différence importante : là, je déplaçais des formes
*déjà existantes* (juste leurs coordonnées `<xdr:from>`/`<xdr:to>`) sur un fichier que je pouvais
rouvrir et vérifier avec `openpyxl`. Créer une *nouvelle* forme Bouton de formulaire liée à une
macro touche plusieurs fichiers XML liés entre eux (`drawing.xml`, `vmlDrawing*.vml`,
`legacyDrawing`, la table des macros) que je ne peux pas valider en le réouvrant dans Excel ici
(pas de Windows/Excel disponible dans cet environnement). Sur un classeur de production que tu
utilises tous les jours, le risque de livrer un fichier corrompu ou un bouton non cliquable
l'emporte largement sur les 60 secondes que ça prend à la main. Donc : instructions ci-dessous,
identiques pour Situation et Gestion.

Par feuille (Situation, puis Gestion) :

1. Onglet **Développeur** (si absent : Fichier → Options → Personnaliser le ruban → cocher
   "Développeur") → groupe **Contrôles** → **Insérer** → sous "Contrôles de formulaire", choisis
   **Bouton**.
2. Dessine-le à l'endroit voulu (une zone libre proche des boutons existants, ex. à côté de
   `exit1`/`FullScr1` sur Situation, ou `Exit2` sur Gestion, pour rester cohérent visuellement) --
   une boîte de dialogue "Affecter une macro" s'ouvre automatiquement.
3. Choisis **`LancerTV`** dans la liste (macro du Module1, visible depuis n'importe quelle
   feuille) → OK.
4. Clic droit sur le bouton → **Modifier le texte** → renomme en `Lancer TV`.
5. Répète les étapes 1 à 4 avec un second bouton, macro **`FermerTV`**, texte `Fermer TV`.
6. Redimensionne/aligne les deux boutons à la taille des boutons existants (clic droit → Taille
   et propriétés, ou glisser les poignées) pour rester "petits" comme demandé.
7. Recommence sur l'autre feuille.

Pas besoin de VBA supplémentaire pour les boutons eux-mêmes : un bouton de formulaire affecté à
`LancerTV` ou `FermerTV` appelle directement ces Subs du Module1, déjà ajoutées ci-dessus -- aucun
`*_Click()` séparé n'est nécessaire (ce mécanisme est différent des contrôles ActiveX, qui eux
exigeraient un `Private Sub NomDuBouton_Click()` dans le module de la feuille).

## Point non vérifiable d'ici

Je n'ai pas de Windows/Chrome/Edge disponible dans cet environnement pour tester réellement le
`Shell`, le profil dédié et le positionnement multi-écran — ce sont des techniques standards
(arguments de ligne de commande Chrome/Edge documentés), mais à valider chez toi :
- Si ton PC utilise un layout d'écrans qui n'est pas "TV à droite, même hauteur" : ajuste
  `tvOffsetX`/`tvOffsetY` avec les vraies coordonnées (Windows Paramètres d'affichage).
- **Si le flash persiste malgré le profil dédié** : c'est le signe que ce n'est pas (ou plus) un
  problème de drapeaux ignorés. Ouvre les outils de développement du navigateur sur la fenêtre TV
  (F12) → onglet Console, laisse tourner un cycle de 5s, et regarde s'il y a une erreur rouge au
  moment du saut. Dis-moi ce qui s'affiche (ou une capture) — ça me dira si `fetch` échoue pour
  une autre raison (chemin de fichier incorrect, fichier verrouillé pendant l'écriture VBA...) et
  je pourrai corriger précisément plutôt que deviner un correctif de plus.
