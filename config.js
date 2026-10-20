// Paramètres du complément « Pennylane CELEOS ».
// Aucune donnée confidentielle ici : l'adresse de collecte Pennylane est saisie dans le panneau
// et enregistrée dans la boîte aux lettres (paramètres itinérants du complément), jamais sur le site.
window.CELEOS_CONFIG = {
  clientId: "__CLIENT_ID__",          // ID d'application (client) Entra ID
  tenantId: "__TENANT_ID__",          // ID de l'annuaire (locataire) Entra ID
  allowedRecipientDomain: "suppliers.pennylane.com", // seul domaine de destination accepté
  pennylaneName: "Pennylane",
  targetFolderName: "Transfert vers Pennylane", // sous-dossier de la Boîte de réception
  maxPdfBytesToScan: 15 * 1024 * 1024,
  sentCheckTimeoutSeconds: 90
};
