# Comptes – app mobile pour le Google Sheets « Suivi des dépenses »

Web app Apps Script liée au Google Sheets (onglet `Situation_`). Le fichier reste la source de vérité.

## Installation
1. Drive : ouvrir le .xlsx → *Fichier → Enregistrer en tant que Google Sheets*.
2. Dans le Sheets : *Extensions → Apps Script*. Coller `Code.gs` (remplace le contenu) et créer un fichier HTML nommé `Index` avec `Index.html`.
3. *Déployer → Nouveau déploiement → Application Web* ; exécuter en tant que **moi**, accès **moi uniquement**.
4. Ouvrir l'URL `/exec` sur le téléphone → menu du navigateur → *Ajouter à l'écran d'accueil*.

## Logique
`Reste = Solde + CA − Prêt conso − Dépenses en attente − Revenus mis de côté − Charges fixes à payer (col. D)`
(`B4` n'est soustrait qu'une fois ; l'ancienne formule le soustrayait deux fois.)
Les zones sont lues depuis les formules `SUM` du fichier : insérer une ligne dans une zone est sans risque.
