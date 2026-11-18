# Export de la colonne Gilet et des avions hors ligne — VBA à coller dans `GenererHTMLSituation`

Contexte : le template HTML affiche maintenant une colonne **Gilet**, construite à partir de
deux colonnes déjà présentes dans `Situation!P7:Q31` (confirmées en inspectant directement le
fichier `piste - V4.xlsm` que tu as envoyé), et une liste **Avions hors ligne** sur le plan,
construite à partir de `Paramètre!B59:C78` (`Tableau11`, confirmé dans le même fichier) :

- **P = GILETS** : taille affectée (liste déroulante `Paramètre!H46:H115`, ex. `"+2 M-W"`,
  `"-3 A"`...). Texte libre selon la liste, affiché tel quel dans le tableau HTML.
- **Q = V** : validation (liste déroulante `Paramètre!F46:F48`, valeurs `"X"` ou `"V"`).
- **Paramètre!B59:C78** : colonne B = position (`"B6"`, `"C1"`...), colonne C = numéro d'avion
  (ex. `96`, correspond à `Num` dans `Tableau3`/`Num_Avion`). Seules les lignes où C est rempli
  sont exportées.

Règle de couleur Gilet reprise **telle quelle de la mise en forme conditionnelle du fichier
Excel** (`Situation!P7:Q31`, dxfId 8/9/10) :
- `Q = "X"` → rouge (dxfId 10, `FFFF0000`)
- `Q = "V"` → vert (dxfId 8, `FF33CC33`)
- `P` renseigné mais `Q` ni `"X"` ni `"V"` → orange (dxfId 9, `FFFF9900`)
- `P` et `Q` tous les deux vides → pas de couleur (rien dans le fichier Excel n'indique un rouge
  par défaut ici, contrairement à Plein qui a cette règle explicite -- dis-moi si Gilet doit se
  comporter pareil)

## `GenererHTMLSituation` (Module1) — à coller

D'après le code que tu as envoyé, voici le `Sub` complet avec Gilet et les avions hors ligne
ajoutés :

```vba
Sub GenererHTMLSituation()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Situation")

    Dim tail As String, pos As String, o2 As String, am As String, pm As String, vdn As String
    Dim ptr As String, avq As String, plein As String, trg As String, obs As String
    Dim gilet As String, giletv As String
    Dim camp As Boolean, pleinValide As Boolean, masque As Boolean
    Dim item As String
    Dim itemsArr(0 To 24) As String
    Dim n As Long
    Dim r As Long

    n = 0
    For r = 7 To 31
        tail = CellStr(ws.Cells(r, "B"))
        If tail <> "" Then
            pos = CellStr(ws.Cells(r, "D"))
            camp = (UCase(CellStr(ws.Cells(r, "E"))) = "X")
            o2 = CellStr(ws.Cells(r, "F"))
            am = CellStr(ws.Cells(r, "G"))
            pm = CellStr(ws.Cells(r, "H"))
            vdn = CellStr(ws.Cells(r, "I"))
            ptr = CellStr(ws.Cells(r, "J"))
            avq = CellStr(ws.Cells(r, "K"))
            plein = CellStr(ws.Cells(r, "L"))
            trg = CellStr(ws.Cells(r, "M"))
            pleinValide = (UCase(CellStr(ws.Cells(r, "N"))) = "V")
            obs = CellStr(ws.Cells(r, "O"))
            gilet = CellStr(ws.Cells(r, "P"))
            giletv = CellStr(ws.Cells(r, "Q"))
            masque = (UCase(CellStr(ws.Cells(r, "S"))) = "X")

            item = "    {tail:" & JsStr(tail) & ", pos:" & JsStr(pos) & _
                   ", camp:" & JsBool(camp) & ", o2:" & JsStr(o2) & _
                   ", am:" & JsStr(am) & ", pm:" & JsStr(pm) & ", vdn:" & JsStr(vdn) & _
                   ", ptr:" & JsStr(ptr) & ", avq:" & JsStr(avq) & _
                   ", gilet:" & JsStr(gilet) & ", giletv:" & JsStr(giletv) & _
                   ", plein:" & JsStr(plein) & ", trg:" & JsStr(trg) & ", v:" & JsBool(pleinValide)
            If masque Then item = item & ", masque:true"
            item = item & ", obs:" & JsStr(obs) & "}"

            itemsArr(n) = item
            n = n + 1
        End If
    Next r

    Dim aircraftJs As String
    Dim i As Long
    If n = 0 Then
        aircraftJs = "[]"
    Else
        aircraftJs = "[" & vbCrLf
        For i = 0 To n - 1
            aircraftJs = aircraftJs & itemsArr(i)
            If i < n - 1 Then aircraftJs = aircraftJs & ","
            aircraftJs = aircraftJs & vbCrLf
        Next i
        aircraftJs = aircraftJs & "  ]"
    End If

    ' Avions hors ligne (Paramètre!B59:C78, Tableau11) -- affichés dans une
    ' liste à part sur le plan, pas sur la voie/les hangars.
    Dim wsParam As Worksheet
    Set wsParam = ThisWorkbook.Sheets("Paramètre")
    Dim hlPos As String, hlNum As String
    Dim hlArr(0 To 19) As String
    Dim hn As Long
    hn = 0
    For r = 59 To 78
        hlNum = CellStr(wsParam.Cells(r, "C"))
        If hlNum <> "" Then
            hlPos = CellStr(wsParam.Cells(r, "B"))
            hlArr(hn) = "    {pos:" & JsStr(hlPos) & ", num:" & JsStr(hlNum) & "}"
            hn = hn + 1
        End If
    Next r

    Dim horsLigneJs As String
    Dim hi As Long
    If hn = 0 Then
        horsLigneJs = "[]"
    Else
        horsLigneJs = "[" & vbCrLf
        For hi = 0 To hn - 1
            horsLigneJs = horsLigneJs & hlArr(hi)
            If hi < hn - 1 Then horsLigneJs = horsLigneJs & ","
            horsLigneJs = horsLigneJs & vbCrLf
        Next hi
        horsLigneJs = horsLigneJs & "  ]"
    End If

    Dim kannadStr As String
    If IsDate(ws.Range("V23").Value) Then
        kannadStr = Format(ws.Range("V23").Value, "dd/mm/yyyy")
    Else
        kannadStr = ""
    End If

    Dim templatePath As String, outPath As String
    templatePath = ThisWorkbook.Path & "\situation-avions-template.html"
    outPath = ThisWorkbook.Path & "\situation-avions.html"

    If Dir(templatePath) = "" Then
        MsgBox "Template introuvable : " & templatePath & vbCrLf & _
               "Dépose situation-avions-template.html à côté de ce classeur.", vbExclamation
        Exit Sub
    End If

    ' Lecture/écriture en UTF-8 (accents, guillemets) via ADODB.Stream --
    ' Open/Print natif VBA est en ANSI et casserait les caractères accentués.
    Dim streamIn As Object
    Set streamIn = CreateObject("ADODB.Stream")
    streamIn.Type = 2 ' adTypeText
    streamIn.Charset = "utf-8"
    streamIn.Open
    streamIn.LoadFromFile templatePath
    Dim html As String
    html = streamIn.ReadText
    streamIn.Close

    ' Thème clair/sombre (bouton "Thème", voir BasculerTheme) -- stocké dans
    ' Paramètre!A1 ("light" ou vide/autre -> sombre par défaut).
    Dim themeStr As String
    themeStr = LCase(Trim(CellStr(wsParam.Range("A1"))))
    If themeStr <> "light" Then themeStr = "dark"

    html = Replace(html, "__AIRCRAFT_DATA__", aircraftJs)
    html = Replace(html, "__KANNAD_DATE__", kannadStr)
    html = Replace(html, "__HORS_LIGNE_DATA__", horsLigneJs)
    html = Replace(html, "__THEME__", themeStr)

    Dim streamOut As Object
    Set streamOut = CreateObject("ADODB.Stream")
    streamOut.Type = 2
    streamOut.Charset = "utf-8"
    streamOut.Open
    streamOut.WriteText html
    streamOut.SaveToFile outPath, 2 ' adSaveCreateOverWrite
    streamOut.Close

    Exit Sub

ErrHandler:
    MsgBox "GenererHTMLSituation : " & Err.Description, vbExclamation
End Sub
```

Changements par rapport à ta version :
- `Dim gilet As String, giletv As String` + les deux lectures `gilet = CellStr(ws.Cells(r, "P"))`
  / `giletv = CellStr(ws.Cells(r, "Q"))` + leur injection dans `item` juste après `avq`.
- Le bloc "Avions hors ligne" entier (nouveau), juste après le calcul de `aircraftJs`.
- `html = Replace(html, "__HORS_LIGNE_DATA__", horsLigneJs)`, ajouté à la suite des deux
  `Replace` déjà présents.
- Lecture du thème (`Paramètre!A1`) et `html = Replace(html, "__THEME__", themeStr)` -- voir
  section suivante.

Tout le reste est identique à ce que tu as envoyé.

## Bascule thème clair/sombre

Le template gère maintenant deux thèmes (sombre par défaut, clair en option -- moins de
reflets au soleil). Le choix est stocké dans une cellule libre confirmée vide en inspectant
`piste - V4.xlsm` : **`Paramètre!A1`**, contenant soit `"light"`, soit vide/autre chose (traité
comme sombre). `GenererHTMLSituation` la lit à chaque export (voir ci-dessus) et l'injecte dans
`__THEME__`.

Il faut un deuxième `Sub`, qui bascule la valeur de `Paramètre!A1` puis relance tout de suite
`GenererHTMLSituation` pour que le changement soit visible dans les 5 secondes (le navigateur
relit le fichier exporté automatiquement, pas besoin de rouvrir la page) :

```vba
Sub BasculerTheme()
    On Error GoTo ErrHandler

    Dim wsParam As Worksheet
    Set wsParam = ThisWorkbook.Sheets("Paramètre")

    Dim current As String
    current = LCase(Trim(CStr(wsParam.Range("A1").Value)))

    If current = "light" Then
        wsParam.Range("A1").Value = "dark"
    Else
        wsParam.Range("A1").Value = "light"
    End If

    GenererHTMLSituation
    Exit Sub

ErrHandler:
    MsgBox "BasculerTheme : " & Err.Description, vbExclamation
End Sub
```

### Bouton Excel

Même principe que les boutons "Lancer TV" / "Fermer TV" déjà en place : un bouton Image
ActiveX (ou un bouton de formulaire, selon ce que tu utilises déjà) sur la feuille **Situation**
et sur la feuille **Gestion**, assigné à la macro `BasculerTheme`. Pas de nouvelle icône
fournie ici -- reprends le même style/gabarit que les boutons existants (ou une icône simple
"soleil/lune" si tu en as une sous la main), l'important est juste l'assignation de macro.

Rien d'autre à changer : `Paramètre!A1` vide au premier lancement se lit comme `"dark"`
(comportement actuel inchangé tant que personne n'a cliqué sur le bouton).

