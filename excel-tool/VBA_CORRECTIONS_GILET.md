# Export de la colonne Gilet — VBA à coller dans `GenererHTMLSituation`

Contexte : le template HTML affiche maintenant une colonne **Gilet**, construite à partir de
deux colonnes déjà présentes dans `Situation!P7:Q31` (confirmées en inspectant directement le
fichier `piste - V4.xlsm` que tu as envoyé) :

- **P = GILETS** : taille affectée (liste déroulante `Paramètre!H46:H115`, ex. `"+2 M-W"`,
  `"-3 A"`...). Texte libre selon la liste, affiché tel quel dans le tableau HTML.
- **Q = V** : validation (liste déroulante `Paramètre!F46:F48`, valeurs `"X"` ou `"V"`).

Règle de couleur reprise **telle quelle de la mise en forme conditionnelle du fichier Excel**
(`Situation!P7:Q31`, dxfId 8/9/10) :
- `Q = "X"` → rouge (dxfId 10, `FFFF0000`)
- `Q = "V"` → vert (dxfId 8, `FF33CC33`)
- `P` renseigné mais `Q` ni `"X"` ni `"V"` → orange (dxfId 9, `FFFF9900`)
- `P` et `Q` tous les deux vides → pas de couleur (rien dans le fichier Excel n'indique un rouge
  par défaut ici, contrairement à Plein qui a cette règle explicite -- dis-moi si Gilet doit se
  comporter pareil)

## `GenererHTMLSituation` (Module1) — à coller

D'après le code que tu as envoyé, voici le `Sub` complet avec les deux colonnes ajoutées
(déclarations, lecture, et injection dans `item`) :

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

    html = Replace(html, "__AIRCRAFT_DATA__", aircraftJs)
    html = Replace(html, "__KANNAD_DATE__", kannadStr)

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

Seuls changements par rapport à ta version : la ligne `Dim gilet As String, giletv As String`, les deux
lectures `gilet = CellStr(ws.Cells(r, "P"))` / `giletv = CellStr(ws.Cells(r, "Q"))`, et leur
injection dans `item` juste après `avq`. Tout le reste est identique à ce que tu as envoyé.
