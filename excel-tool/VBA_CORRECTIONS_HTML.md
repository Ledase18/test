# Export HTML "Situation avions" — VBA à coller

Ça câble le thème HTML (jusqu'ici une maquette avec des données d'exemple) sur les vraies
données de `piste - V3.xlsm`. Rien ne change dans la structure du classeur — tout se fait en
VBA, plus un fichier HTML à déposer à côté.

## Fichiers concernés

- `situation-avions-template.html` — le **template** (structure, CSS, JS). C'est le fichier à
  déposer **dans le même dossier que `piste - V3.xlsm`**, une fois. Il n'est jamais réécrit par
  la macro — il contient deux marqueurs (`__AIRCRAFT_DATA__` et `__KANNAD_DATE__`) que la macro
  remplace par les vraies valeurs.
- `situation-avions.html` — le fichier **généré**, écrit dans le même dossier par la macro à
  chaque rafraîchissement. C'est lui que le navigateur de la TV doit ouvrir/garder ouvert (il se
  recharge tout seul toutes les 30 s pour prendre en compte la dernière version écrite).

Ne modifie jamais `situation-avions.html` à la main — il est écrasé à chaque export. Les
retouches de mise en forme se font dans `situation-avions-template.html`.

## D'où viennent les données

Confirmé en lisant directement la feuille `Situation` (ligne d'en-têtes 6, données lignes 7 à 31) :

| Colonne | En-tête          | Champ JS   | Traitement |
|---------|------------------|------------|------------|
| B       | AVION            | `tail`     | texte tel quel |
| D       | Pos.             | `pos`      | texte (les positions numériques type `11` sont reconverties en texte sans décimale) |
| E       | Camp.            | `camp`     | `true` si = "X" |
| F       | O2               | `o2`       | `true` si = "X" |
| G       | AM               | `am`       | texte tel quel |
| H       | PM               | `pm`       | texte tel quel |
| I       | VDN              | `vdn`      | texte tel quel |
| J       | PTR              | `ptr`      | texte tel quel |
| K       | AVQ              | `avq`      | texte tel quel |
| L       | PLEIN            | `plein`    | texte tel quel |
| N       | V                | `v`        | `true` si = "V" (plein validé) |
| O       | OBSERVATIONS     | `obs`      | texte tel quel |
| S       | Masquer          | `masque`   | champ ajouté seulement si = "X" (sinon absent) |
| V23     | (MAJ FDS/KANNAD) | date pill  | formatée `jj/mm/aaaa` |

Une ligne est ignorée si la colonne B (AVION) est vide.

Ce mapping a été vérifié directement sur le fichier (en-têtes ligne 6 + données réelles), donc
fiable — pas une supposition cette fois.

## 1. `piste - V3.xlsm` — Module1 : nouvelles procédures

Ajoute ce bloc entier (functions utilitaires + macro d'export) à la suite de
`PousserRafraichissementVisu` :

```vba
' ============================================================
' Export HTML "Situation avions" -- génère situation-avions.html
' à côté du classeur à partir de situation-avions-template.html
' et des données réelles de Situation!B7:S31.
' ============================================================

Private Function JsStr(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    JsStr = """" & s & """"
End Function

Private Function JsBool(ByVal b As Boolean) As String
    JsBool = IIf(b, "true", "false")
End Function

Private Function CellStr(ByVal c As Range) As String
    Dim v As Variant
    v = c.Value
    If IsEmpty(v) Then
        CellStr = ""
    ElseIf IsNumeric(v) Then
        CellStr = Trim(CStr(CLng(v)))
    Else
        CellStr = Trim(CStr(v))
    End If
End Function

Sub GenererHTMLSituation()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Situation")

    Dim tail As String, pos As String, am As String, pm As String, vdn As String
    Dim ptr As String, avq As String, plein As String, obs As String
    Dim camp As Boolean, o2 As Boolean, pleinValide As Boolean, masque As Boolean
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
            o2 = (UCase(CellStr(ws.Cells(r, "F"))) = "X")
            am = CellStr(ws.Cells(r, "G"))
            pm = CellStr(ws.Cells(r, "H"))
            vdn = CellStr(ws.Cells(r, "I"))
            ptr = CellStr(ws.Cells(r, "J"))
            avq = CellStr(ws.Cells(r, "K"))
            plein = CellStr(ws.Cells(r, "L"))
            pleinValide = (UCase(CellStr(ws.Cells(r, "N"))) = "V")
            obs = CellStr(ws.Cells(r, "O"))
            masque = (UCase(CellStr(ws.Cells(r, "S"))) = "X")

            item = "    {tail:" & JsStr(tail) & ", pos:" & JsStr(pos) & _
                   ", camp:" & JsBool(camp) & ", o2:" & JsBool(o2) & _
                   ", am:" & JsStr(am) & ", pm:" & JsStr(pm) & ", vdn:" & JsStr(vdn) & _
                   ", ptr:" & JsStr(ptr) & ", avq:" & JsStr(avq) & ", plein:" & JsStr(plein) & _
                   ", v:" & JsBool(pleinValide)
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

## 2. `piste - V3.xlsm` — Module1 : appel dans `ArchiverDispoDuJour`

Même point d'ancrage que `PousserRafraichissementVisu` (voir
`VBA_CORRECTIONS_RAFRAICHISSEMENT.md`) — ça couvre `Workbook_Open`, `Workbook_BeforeClose` et
`Actualiser_Click` d'un coup, puisque les trois appellent déjà `ArchiverDispoDuJour`.

**Avant** (fin de la procédure) :
```vba
    ActiveWorkbook.Save
    Workbooks(myName).Activate
    ActiveWorkbook.Save

    Call PousserRafraichissementVisu

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
    Call GenererHTMLSituation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
```

## 3. `piste - V3.xlsm` — Feuil3 (`Gestion`) : appel dans `valid_click`

Le bouton **Valider** — point d'entrée principal des changements de statut avion.

**Avant** (fin de la procédure) :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    Call PousserRafraichissementVisu

Application.ScreenUpdating = True
End Sub
```

**Après** :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    Call PousserRafraichissementVisu
    Call GenererHTMLSituation

Application.ScreenUpdating = True
End Sub
```

## Mise en place

1. Dépose `situation-avions-template.html` dans le même dossier que `piste - V3.xlsm`
   (et `Visu.xlsm`).
2. Colle les 3 blocs VBA ci-dessus dans l'éditeur VBA de `piste - V3.xlsm`.
3. Sauvegarde, ouvre le classeur, clique **Valider** (ou **Actualiser**) une fois : un fichier
   `situation-avions.html` doit apparaître dans le même dossier.
4. Ouvre ce fichier dans le navigateur destiné à la TV, laisse-le affiché : il se recharge tout
   seul toutes les 30 s, donc les rafraîchissements suivants (nouveaux Valider/Actualiser)
   apparaissent sans rien rouvrir à la main.

## Vérifié avant livraison

J'ai testé la transformation (lecture Situation → tableau JS → injection dans le template) en
dehors d'Excel, avec les vraies valeurs de mon dernier instantané local de `piste - V3.xlsm` (19
avions en ligne 7-31, dont 8 masqués), et vérifié :
- le JS généré est syntaxiquement valide (`node --check`),
- le rendu correspond aux données (avions masqués bien exclus de la liste et du plan, couleurs
  PTR/AVQ/Plein correctes, badge Kannad à la bonne date).

Point que je n'ai **pas** pu vérifier : le comportement réel de `ADODB.Stream` dans ton Excel
(version, paramètres de sécurité macro) — c'est une technique standard mais si `CreateObject
("ADODB.Stream")` échoue chez toi (rare, mais arrive sur des postes très verrouillés), dis-le
moi, il y a une alternative en `Scripting.FileSystemObject` mais qui gère moins bien l'UTF-8.

## Connu, pas encore fait

- `Camp.`, `O2`, etc. sont lus mais aucune règle particulière n'existe pour eux au-delà de
  l'affichage en chip — si un de ces champs doit influer sur la couleur globale de la ligne,
  dis-le moi.
- La colonne M (`TRG + Qté`) et la deuxième colonne "V" (Q, liée à GILETS/DEM VAL) ne sont pas
  exportées — le thème actuel ne les affiche pas. Je peux les ajouter si besoin.
