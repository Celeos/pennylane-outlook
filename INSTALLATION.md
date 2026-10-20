# Complément Outlook « Pennylane CELEOS » - installation

## 1. Hébergement sur GitHub Pages (dépôt privé de l'organisation)

Créer le dépôt `pennylane-outlook` avec l'organisation comme « Owner », visibilité **Private**. Y déposer le **contenu** du dossier `pennylane-outlook` (Add file > Upload files > Commit changes). Puis Settings > Pages : « Deploy from a branch », `main`, `/ (root)`, Save.

Sécurisation du compte GitHub (Settings de l'organisation) : authentification à deux facteurs obligatoire pour les membres (Authentication security > Require two-factor authentication) ; accès en écriture au dépôt limité à vous seul ; dans le dépôt, Settings > Branches, règle de protection sur `main` interdisant la suppression et la réécriture d'historique.

## 2. Inscription de l'application dans Entra ID

entra.microsoft.com > Applications > Inscriptions d'applications > Nouvelle inscription. Nom `Pennylane Outlook CELEOS`, comptes de cet annuaire uniquement, URI de redirection de type **Application monopage (SPA)** : `brk-multihub://<organisation>.github.io`.

API autorisées > Microsoft Graph > Autorisations déléguées : `Mail.ReadWrite` et `Mail.Send`, puis « Accorder un consentement d'administrateur ». Ne créer aucun secret client ni certificat : l'application n'en a pas besoin.

Restriction des utilisateurs : Applications d'entreprise > `Pennylane Outlook CELEOS` > Propriétés > « Affectation requise » sur **Oui** ; puis Utilisateurs et groupes > ajouter uniquement les personnes autorisées.

## 3. Finalisation

Transmettre à Claude le nom de l'organisation GitHub, l'ID d'application (client) et l'ID de l'annuaire (locataire) pour compléter `config.js` et `manifest.xml`, puis déposer ces deux fichiers dans le dépôt.

## 4. Installation dans le nouvel Outlook

`https://aka.ms/olksideload` > Mes compléments > Ajouter un complément personnalisé > Ajouter à partir du fichier > `manifest.xml`. Au premier usage, le panneau demande l'adresse de collecte Pennylane ; elle est enregistrée dans la boîte aux lettres et n'apparaît nulle part sur le site.

## Mesures de sécurité intégrées

Aucune donnée CELEOS sur le site : l'adresse Pennylane est stockée dans les paramètres du complément de la boîte aux lettres. Seules les adresses `@suppliers.pennylane.com` sont acceptées, au moment de la saisie puis avant l'envoi. Une politique de sécurité de contenu (CSP) n'autorise que le site lui-même, la bibliothèque Office de Microsoft et Microsoft Graph ; les polices et scripts tiers ont été supprimés. Le jeton de connexion n'est conservé que pour la session du panneau. Le lecteur PDF (pdf.js 6.4.299) est la version corrigée de la faille GHSA-hq66-cqwq-w95j, avec exécution de scripts et formulaires XFA désactivés ; les bibliothèques sont servies depuis le dépôt, sans dépendance à un CDN tiers. Aucun secret n'est stocké, ni dans le code ni dans Entra ID.

## Risque résiduel

Le complément s'exécute avec vos droits sur la messagerie (lecture, écriture, envoi). Quiconque pourrait modifier les fichiers du dépôt modifierait le code exécuté dans votre Outlook au prochain chargement. La protection porte donc sur le compte GitHub de l'organisation (2FA, un seul membre en écriture). En cas de doute : supprimer le complément dans Outlook et désactiver l'application dans Entra ID (Applications d'entreprise > Propriétés > « Activé pour la connexion des utilisateurs » sur Non), ce qui coupe immédiatement tout accès.
