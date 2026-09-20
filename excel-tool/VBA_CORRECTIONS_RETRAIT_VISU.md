# Retrait de Visu.xlsm — VBA à retirer

Maintenant que `situation-avions.html` (généré par `GenererHTMLSituation`, voir
`VBA_CORRECTIONS_HTML.md`) affiche les mêmes données sur la TV, `Visu.xlsm` n'a plus besoin
d'être ouvert. Ce correctif retire l'appel qui le rafraîchissait — le reste (le fichier lui-même,
son code VBA) n'a pas besoin d'être touché : il ne sera simplement plus lancé.

## `piste - V3.xlsm` — Module1 : retire l'appel dans `ArchiverDispoDuJour`

Si tu as déjà collé `VBA_CORRECTIONS_RAFRAICHISSEMENT.md`, la fin de la procédure ressemble à
ça :

**Avant** :
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

**Après** (retire uniquement la ligne `Call PousserRafraichissementVisu`) :
```vba
    ActiveWorkbook.Save
    Workbooks(myName).Activate
    ActiveWorkbook.Save

    Call GenererHTMLSituation

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
```

## `piste - V3.xlsm` — Feuil3 (`Gestion`) : retire l'appel dans `valid_click`

**Avant** :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    Call PousserRafraichissementVisu
    Call GenererHTMLSituation

Application.ScreenUpdating = True
End Sub
```

**Après** :
```vba
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True

    Call GenererHTMLSituation

Application.ScreenUpdating = True
End Sub
```

## La procédure `PousserRafraichissementVisu` elle-même

Tu peux la laisser en place dans Module1 (elle ne fait plus rien puisque plus rien ne l'appelle)
ou la supprimer entièrement — les deux sont sans risque. Je n'ai pas de préférence forte ; la
garder ne coûte rien et évite une manipulation VBA de plus si jamais tu changeais d'avis sur
Visu.

## Côté poste TV

- Le navigateur affiché sur la TV doit maintenant ouvrir `situation-avions.html` (celui généré
  par `GenererHTMLSituation`, à côté de `piste - V3.xlsm`) au lieu de lancer `Visu.xlsm`.
- `Visu.xlsm` n'a plus besoin d'être ouvert du tout. Je ne l'ai pas supprimé du dossier ni du
  dépôt — pas de raison de le faire disparaître tant que le nouveau pipeline n'a pas tourné en
  vrai sur la TV pendant quelques jours ; il reste un filet de sécurité simple à réactiver si
  jamais le HTML posait un problème imprévu.

## Point d'attention

`Visu.xlsm` reste un fichier lié à `piste - V3.xlsm` par liaison externe (`xlExcelLinks`) côté
classeur. Le retirer du pipeline de rafraîchissement ne casse rien dans `piste - V3.xlsm` --
c'est `Visu.xlsm` qui lisait piste-V3, pas l'inverse. Si un jour tu veux supprimer Visu.xlsm pour
de bon, ce sera un nettoyage séparé, pas nécessaire pour que la TV fonctionne en HTML.
