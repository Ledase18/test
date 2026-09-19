# Correctifs VBA à coller — `piste - V3.xlsm`

## Ajout du 19/09 : supprimer l'avertissement « informations personnelles » à l'enregistrement, pour tout le monde

Ce message vient d'un réglage Excel local à chaque poste (Centre de gestion de la
confidentialité), donc le désactiver là ne vaudrait que pour ton poste. Pour que ça s'applique
**automatiquement à tout le monde qui ouvre ce fichier**, sans toucher au réglage de chacun, la
bonne approche est en VBA — via la propriété `Workbook.RemovePersonalInformation`, qui est
conservée avec le fichier et se comporte comme une commande à réappliquer avant chaque
enregistrement (`Workbook_BeforeSave` se déclenche à chaque Ctrl+S, bouton, ou `.Save` en VBA) :

Dans **ThisWorkbook**, ajoute cette nouvelle procédure (elle n'existe pas encore, donc pas de
« avant/après » ici, juste à coller) :

```vba
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Me.RemovePersonalInformation = False
End Sub
```

Je n'ai pas pu tester ça dans un vrai Excel (juste dans LibreOffice, qui n'a pas ce système
d'avertissement de confidentialité pour comparer) — c'est une propriété VBA standard et
documentée, mais je ne peux pas garantir à 100 % qu'elle supprime exactement le message que tu
vois dans ta version d'Excel. Si le message persiste malgré ce correctif, la solution de repli
reste le réglage Centre de gestion de la confidentialité (Fichier → Options → Centre de gestion
de la confidentialité → Paramètres → Options de confidentialité) — mais cette fois-là il faudrait
le déployer sur chaque poste (ou via une stratégie de groupe si vous êtes en domaine Windows,
ce que je ne peux pas configurer d'ici).

## Correctif du 19/09 : Actualiser ne sauvegardait plus piste-V3 (bug de mon refactor)

Si tu as déjà collé la version précédente de ce fichier : **une ligne manque**, ajoute-la. En
factorisant le bloc dupliqué dans `ArchiverDispoDuJour`, j'ai fusionné par erreur deux
`ActiveWorkbook.Save` qui visaient deux classeurs différents à deux moments différents en un
seul. Résultat : le bouton **Actualiser** sauvegardait bien l'archive du jour, mais plus
`piste - V3.xlsm` lui-même — donc `Visu.xlsm` (qui ne lit que ce qui est enregistré sur le
disque) ne voyait jamais la mise à jour. Dans la section « 1. Module1 » ci-dessous, juste après
```vba
    ActiveWorkbook.Save
    Workbooks(myName).Activate
```
ajoute une seconde ligne `ActiveWorkbook.Save` (elle sauvegarde maintenant `piste - V3.xlsm`,
puisqu'on vient de l'activer juste avant). Le code complet plus bas est déjà à jour avec ce
correctif.

---

Ce fichier n'est **pas exécuté automatiquement** : je n'ai pas d'Excel/Windows dans cet
environnement pour éditer/tester `vbaProject.bin` (format binaire compilé — une modification à
l'aveugle serait invérifiable et pourrait corrompre le classeur). Tu copies-colles toi-même ces
blocs dans l'éditeur VBA (Alt+F11), tu enregistres, tu testes. Ça ne touche à rien d'autre que
ces 3 procédures + l'ajout d'une nouvelle.

Ce que ça corrige, sans changer le comportement dans le cas normal :
- le bloc « vérifier/créer l'archive du jour + remplir les créneaux de 10 min », dupliqué à
  l'identique 3 fois, devient une seule procédure `ArchiverDispoDuJour` appelée aux 3 endroits ;
- ajout de la création du dossier `Archives` s'il est absent (`MkDir`) — avant, un `SaveAs` vers
  un dossier manquant plantait sans message ;
- ajout d'une gestion d'erreur (`On Error GoTo`) autour de l'ouverture/sauvegarde de l'archive,
  avec un message clair plutôt qu'un plantage silencieux en plein `Workbook_Open`/`Close` ;
- `Application.ScreenUpdating` désactivé pendant la boucle des 150 créneaux dans les 3 cas (avant,
  seul `Actualiser_Click` le faisait) — évite le scintillement à l'ouverture/fermeture.

Rien d'autre n'est changé : mêmes cellules lues/écrites, même logique de remplissage
avant/arrière, même ordre d'exécution.

---

## 1. Module1 — ajouter la nouvelle procédure

Dans **VBAProject → Modules → Module1**, colle ce bloc à la suite de `WbOpen` (ne touche pas à
`ExisteFichier`/`WbOpen`, qui restent utilisées telles quelles) :

```vba
Sub ArchiverDispoDuJour()
' Factorisation du bloc "archive du jour + créneaux 10 min", auparavant
' dupliqué à l'identique dans Workbook_Open, Workbook_BeforeClose et
' Actualiser_Click. Ajout : création du dossier Archives s'il est absent,
' gestion d'erreur autour des accès fichier, ScreenUpdating désactivé
' pendant la boucle. Logique de remplissage inchangée.
On Error GoTo GestionErreur

Dim FichierSource As Workbook
Dim MyPath As String, myName As String
Dim i As Integer, j As Integer
Dim NomArchive As String

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set FichierSource = ThisWorkbook
    myName = ThisWorkbook.Name
    MyPath = FichierSource.Path & "\Archives\"
    NomArchive = "Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx"

    If Dir(MyPath, vbDirectory) = "" Then MkDir MyPath

    If ExisteFichier(MyPath & NomArchive) = True Then
        If WbOpen(MyPath & NomArchive) = False Then
            Workbooks.Open Filename:=MyPath & NomArchive
        End If
    Else
        Workbooks.Open Filename:=FichierSource.Path & "\DispoVierge.xlsx"
        ActiveWorkbook.SaveAs MyPath & NomArchive
    End If

    '10 minutes
    For i = 1 To 150
        If ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + i) = "FIN" Then Exit For
        If Format(Now(), "HH:MM") < Format(ActiveWorkbook.Sheets("Enregistrements").Cells(10, 2 + i), "HH:MM") Then
            ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + i) = FichierSource.Sheets("Paramètre").Range("Q30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(12, 2 + i) = FichierSource.Sheets("Paramètre").Range("R30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(13, 2 + i) = FichierSource.Sheets("Paramètre").Range("S30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(14, 2 + i) = FichierSource.Sheets("Paramètre").Range("T30")
        End If
        If Format(Now(), "HH:MM") >= Format(ActiveWorkbook.Sheets("Enregistrements").Cells(10, 2 + i), "HH:MM") Then
            ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + i) = FichierSource.Sheets("Paramètre").Range("Q30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(12, 2 + i) = FichierSource.Sheets("Paramètre").Range("R30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(13, 2 + i) = FichierSource.Sheets("Paramètre").Range("S30")
            ActiveWorkbook.Sheets("Enregistrements").Cells(14, 2 + i) = FichierSource.Sheets("Paramètre").Range("T30")
            For j = i + 1 To 150
                If ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + j) <> "" Then Exit For
                If ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + j) = "FIN" Then Exit For
                If ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + j) = "" Then
                    ActiveWorkbook.Sheets("Enregistrements").Cells(11, 2 + j) = FichierSource.Sheets("Paramètre").Range("Q30")
                    ActiveWorkbook.Sheets("Enregistrements").Cells(12, 2 + j) = FichierSource.Sheets("Paramètre").Range("R30")
                    ActiveWorkbook.Sheets("Enregistrements").Cells(13, 2 + j) = FichierSource.Sheets("Paramètre").Range("S30")
                    ActiveWorkbook.Sheets("Enregistrements").Cells(14, 2 + j) = FichierSource.Sheets("Paramètre").Range("T30")
                Else
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i

    ActiveWorkbook.Save
    Workbooks(myName).Activate
    ActiveWorkbook.Save

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

GestionErreur:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Impossible de mettre à jour l'archive du jour (" & NomArchive & ")." & vbCrLf & _
           "Erreur " & Err.Number & " : " & Err.Description & vbCrLf & _
           "Vérifie que le dossier Archives est accessible et que le fichier n'est pas déjà " & _
           "ouvert par un autre poste.", vbExclamation, "Archivage Dispo"
End Sub
```

## 2. ThisWorkbook — `Workbook_Open`

Remplace tout le bloc entre `Application.DisplayAlerts = False` (celui du milieu, après
l'ouverture de Visu.xlsm) et `Workbooks(myName).Activate` par un seul appel. **Avant** :

```vba
    Application.DisplayAlerts = False
    Set FichierSource = ThisWorkbook
    myName = ThisWorkbook.Name
    MyPath = FichierSource.Path & "\Archives\"
    If ExisteFichier(MyPath & "Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx") = True Then
        ' ... (tout le bloc dupliqué)
    Next i
    ActiveWorkbook.Save
    Workbooks(myName).Activate
```

**Après** :

```vba
    Call ArchiverDispoDuJour
```

(Les lignes `Dim FichierSource As Workbook`, `Dim MyPath, myName As String`, `Dim i, j As
Integer` en haut de `Workbook_Open` deviennent inutiles si plus rien d'autre dans la procédure ne
les utilise — tu peux les laisser, elles ne gênent pas, ou les supprimer.)

## 3. ThisWorkbook — `Workbook_BeforeClose`

Même remplacement, mais **garde** les 2 lignes qui suivent (fermeture de l'archive) :

**Avant** (tout le bloc dupliqué, puis) :
```vba
ActiveWorkbook.Save
Workbooks(myName).Activate
   '***************************************************************
If WbOpen("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx") = True Then Workbooks("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx").Close True
ActiveWorkbook.Save
End Sub
```

**Après** :
```vba
    Call ArchiverDispoDuJour

If WbOpen("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx") = True Then Workbooks("Dispo_" & Format(Now(), "YYYY-MM-DD") & ".xlsx").Close True
ActiveWorkbook.Save
End Sub
```

## 4. Feuil2 (feuille `Situation`) — `Actualiser_Click`

**Avant** : tout le corps entre `Application.DisplayAlerts = False` et `Application.DisplayAlerts
= True` final.

**Après** :
```vba
Private Sub Actualiser_Click()
    Call ArchiverDispoDuJour
End Sub
```

---

## Comment tester avant de garder

1. Fais une copie de sauvegarde de `piste - V3.xlsm` avant de coller quoi que ce soit (Ctrl+C/V
   dans l'explorateur de fichiers).
2. Colle les 4 blocs ci-dessus.
3. Renomme temporairement le dossier `Archives` (ou son contenu du jour) pour vérifier que le
   `MkDir` recrée bien le dossier sans plantage à l'ouverture.
4. Ouvre, clique sur **Actualiser**, ferme le classeur : vérifie que `Dispo_AAAA-MM-JJ.xlsx` est
   bien créé/mis à jour dans `Archives`, avec les mêmes valeurs qu'avant le correctif.
5. Si un test échoue, restaure la sauvegarde de l'étape 1 — aucune de ces modifications n'a été
   appliquée au fichier déposé sur le repository, qui reste dans son état d'origine.
