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
- Ajoute `--disable-session-crashed-bubble` : censé supprimer le bandeau "Chrome ne s'est pas
  arrêté correctement, restaurer les pages ?" qu'affiche Chrome/Edge après un arrêt qu'il
  considère anormal. **Insuffisant seul en pratique** sur les versions récentes de Chrome/Edge --
  le vrai correctif est dans `FermerTV`, qui fait maintenant un arrêt "propre" (`taskkill` sans
  `/F`, équivalent à un clic sur la croix, laisse Chrome écrire lui-même qu'il s'est arrêté
  normalement) avant de forcer en filet de sécurité. Le drapeau reste en place en défense
  complémentaire (coupure de courant, plantage réel, redémarrage du PC...), mais ne suffit plus à
  lui seul contre un `taskkill /F` systématique.
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
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' PID du navigateur lancé par LancerTV, utilisé par FermerTV pour le fermer.
' Public (pas juste module) car les boutons Situation/Gestion doivent tous les
' deux pouvoir déclencher LancerTV/FermerTV et lire/écrire la même valeur.
' Reste à 0 tant qu'aucun lancement via Shell n'a réussi dans cette session
' Excel -- si tu fermes puis rouvres piste-V3, la valeur repart à 0 (normal,
' une variable Public ne survit pas à la fermeture du classeur), donc FermerTV
' ne pourra fermer que la fenêtre lancée par la session en cours.
Public tvProcessID As Double
' /!\ CETTE LIGNE EST INDISPENSABLE -- sans elle, LancerTV et FermerTV
' compilent quand même (pas d'Option Explicit dans ce projet) mais chacune se
' crée sa propre variable tvProcessID locale et invisible de l'autre : le PID
' capturé par LancerTV disparaît à la fin de la Sub, et FermerTV voit toujours
' 0 -> message "Aucune fenêtre TV n'a été lancée...". C'est exactement ce qui
' se produit si tu as collé LancerTV/FermerTV sans cette déclaration.

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
              " --start-fullscreen --allow-file-access-from-files --disable-session-crashed-bubble """ & fileUrl & """"
        ' Shell() en fonction (pas juste "Shell cmd") pour récupérer le PID --
        ' nécessaire pour que FermerTV puisse fermer précisément cette fenêtre.
        tvProcessID = Shell(cmd, vbNormalFocus)
    Else
        ThisWorkbook.FollowHyperlink htmlPath
        ' Pas de PID récupérable via FollowHyperlink (ouverture par le
        ' navigateur par défaut, hors de notre contrôle) -- FermerTV ne pourra
        ' pas fermer cette fenêtre-là automatiquement.
    End If

    ' Rend la main à piste-V3 : le navigateur (vbNormalFocus, ou le navigateur
    ' par défaut via FollowHyperlink) passe au premier plan une fois sa fenêtre
    ' ouverte, ce qui laisse Excel en arrière-plan. Sleep laisse le temps à
    ' cette fenêtre de s'afficher et de prendre le focus avant qu'on le
    ' reprenne -- sans ce délai, AppActivate s'exécuterait avant que le
    ' navigateur n'ait fini de s'activer, qui reprendrait la main juste après.
    ' Corrige aussi le double-clic nécessaire ensuite sur ImageOffSit/
    ' ImageOffGest : tant qu'Excel n'a pas le focus, le 1er clic sur son
    ' contrôle ActiveX ne fait que réactiver la fenêtre Excel (comportement
    ' Windows standard), le 2e seulement atteint réellement le contrôle.
    Sleep 800
    AppActivate Application.Caption

    Exit Sub

ErrHandler:
    ' Ne bloque jamais l'ouverture de piste-V3 pour un problème d'affichage TV.
    Debug.Print "LancerTV : " & Err.Description
End Sub

' ============================================================
' Fermeture TV -- termine le processus navigateur lancé par LancerTV
' (identifié par tvProcessID). taskkill /F seul déclenche systématiquement le
' bandeau "Chrome ne s'est pas arrêté correctement" au lancement suivant, et
' --disable-session-crashed-bubble ne suffit plus à le supprimer sur les
' versions récentes de Chrome/Edge -- d'où l'arrêt en 2 temps ci-dessous.
' ============================================================
Sub FermerTV()
    On Error GoTo ErrHandler

    If tvProcessID > 0 Then
        ' 1) Fermeture "propre" -- taskkill SANS /F envoie WM_CLOSE aux
        ' fenêtres du process, exactement comme un clic sur sa croix : Chrome
        ' exécute alors son arrêt normal et écrit lui-même dans son profil
        ' qu'il s'est fermé correctement (aucune confirmation à attendre côté
        ' page, situation-avions.html n'a pas de gestionnaire beforeunload).
        Shell "taskkill /PID " & tvProcessID, vbHide
        Sleep 1000

        ' 2) Filet de sécurité -- si la fenêtre n'a pas réagi dans le délai
        ' ci-dessus (navigateur bloqué), on force. Si le process s'est déjà
        ' fermé à l'étape 1, ce taskkill échoue juste silencieusement (rien à
        ' tuer) : aucun effet, aucun message (vbHide).
        Shell "taskkill /PID " & tvProcessID & " /F", vbHide

        tvProcessID = 0
    End If
    ' tvProcessID = 0 : aucune TV lancée depuis ce classeur dans cette session
    ' (jamais lancée, déjà fermée, ou ouverte via le navigateur par défaut hors
    ' de notre contrôle) -- ne rien faire, sans message, comme demandé : ce
    ' n'est pas une erreur, c'est l'état normal la plupart du temps (ex. clic
    ' sur OFF par réflexe sans avoir rallumé la TV).

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

**En pratique** : plutôt que des boutons de formulaire, les boutons ont été ajoutés en contrôles
ActiveX Image (`ImageOnSit`/`ImageOffSit` sur Situation -- module `Feuil2` --, `ImageOnGest`/
`ImageOffGest` sur Gestion -- module `Feuil3`), cohérent avec le style déjà en place sur ce
classeur (`Actualiser`, `exit1`, `Exit2`... sont aussi des Image ActiveX). Ça marche tout aussi
bien, avec la nuance ci-dessous sur l'événement à utiliser.

## Corrections suite aux tests (double-clic, focus, fermeture par la croix)

Trois points remontés après un premier test réel :

**1. Double-clic nécessaire sur `ImageOffSit` pour couper la TV.** Cause : le code collé
utilisait l'événement `MouseDown` au lieu de `Click` --

```vba
Private Sub ImageOffSit_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call FermerTV
End Sub
```

-- alors que `ImageOnGest`/`ImageOffGest` (Gestion) utilisent déjà `Click`, comme tous les autres
boutons du classeur. Au-delà du double-clic, `MouseDown` sans filtrer `Button` se déclenche aussi
sur un **clic droit**, ce qui fermerait la TV par accident. À remplacer dans le module `Feuil2`
par :

```vba
Private Sub ImageOffSit_Click()
Call FermerTV
End Sub
```

(Supprime l'ancienne `ImageOffSit_MouseDown` en entier -- garde juste ce `_Click`.)

**2. Focus resté sur la fenêtre TV après le lancement, et cause probable du double-clic.**
`Shell(cmd, vbNormalFocus)` donne le focus au navigateur dès que sa fenêtre s'affiche -- Excel
repasse en arrière-plan. C'est très probablement la vraie cause du point 1 aussi : tant qu'une
fenêtre Windows n'a pas le focus, cliquer sur un de ses contrôles ActiveX ne fait, au premier
clic, que réactiver la fenêtre (comportement standard Windows, pas un bug Excel) -- le clic
suivant seulement atteint le contrôle. `ImageOnSit` n'en souffre pas parce qu'au moment où tu
cliques dessus, Excel a déjà le focus (rien ne l'en a fait sortir) ; `ImageOffSit` en souffre
parce que tu cliques juste après que la TV a pris le focus. Le correctif (`Sleep 800` +
`AppActivate Application.Caption`, ajouté à la fin de `LancerTV` ci-dessus) redonne la main à
piste-V3 automatiquement dès que la fenêtre TV est ouverte -- ça couvre le point 2 directement, et
devrait faire disparaître le point 1 par la même occasion (à confirmer chez toi ; si le double-clic
persiste malgré tout après ce correctif, le remplacement `MouseDown` → `Click` du point 1 reste
nécessaire de toute façon).

**3. Fermer piste-V3 (croix de la fenêtre) doit aussi couper la TV.** `Workbook_BeforeClose`
n'appelle actuellement pas `FermerTV` -- fermer piste-V3 par la croix laisse le navigateur TV
tourner tout seul. Ajoute `Call FermerTV` dans `ThisWorkbook`, `Workbook_BeforeClose` :

**Avant** :
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim FichierSource As Workbook
Dim MyPath, myName As String
Dim i, j As Integer
   Call ArchiverDispoDuJour
If WbOpen("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx") = True Then Workbooks("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx").Close True
ActiveWorkbook.Save
End Sub
```

**Après** :
```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Dim FichierSource As Workbook
Dim MyPath, myName As String
Dim i, j As Integer
   Call FermerTV
   Call ArchiverDispoDuJour
If WbOpen("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx") = True Then Workbooks("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx").Close True
ActiveWorkbook.Save
End Sub
```

`FermerTV` place `Call FermerTV` avant `ArchiverDispoDuJour` volontairement : `FermerTV` a sa
propre gestion d'erreur interne (elle ne peut pas faire échouer la fermeture de piste-V3), donc
autant garantir que la TV est coupée même si l'archivage du jour lève un problème ensuite. Le
bouton OFF (`ImageOffSit_Click`/`ImageOffGest_Click`) coupe déjà la TV via ce même `FermerTV`,
donc ce point ne concerne que la fermeture par la croix.

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
- **`Sleep 800` dans `LancerTV`** : délai choisi au jugé (le temps qu'une fenêtre Chrome/Edge
  s'affiche et prenne le focus sur un PC "normal") -- si le focus revient à Excel *avant* que la
  TV n'ait fini de s'afficher, ou au contraire si la TV garde encore le focus après, augmente ou
  diminue cette valeur (en millisecondes) selon ce que tu observes chez toi. Pas de moyen fiable
  de le déterminer sans tester sur ta machine réelle.
- **`Sleep 1000` dans `FermerTV`** entre l'arrêt propre et le filet de sécurité forcé : même
  remarque, délai au jugé. S'il est trop court, le forçage `/F` interviendra alors que Chrome
  était en train de se fermer proprement -- ce qui ramène le bandeau de restauration dans ce
  cas précis. S'il te semble que le bandeau réapparaît occasionnellement, augmente cette valeur.
