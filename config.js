// Paramètres du complément « Pennylane CELEOS ».
// Aucune donnée confidentielle ici : l'adresse de collecte Pennylane est saisie dans le panneau
// et enregistrée dans la boîte aux lettres (paramètres itinérants du complément), jamais sur le site.
window.CELEOS_CONFIG = {
  clientId: "5bcb60ed-9402-4bfa-b944-76e97f1e9c2a",          // ID d'application (client) Entra ID
  tenantId: "4f91dd1f-b2f0-4a3a-824b-711610a06e5d",          // ID de l'annuaire (locataire) Entra ID
  allowedRecipientDomain: "suppliers.pennylane.com", // seul domaine de destination accepté
  pennylaneName: "Pennylane",
  targetFolderName: "Transfert vers Pennylane", // sous-dossier de la Boîte de réception
  maxPdfBytesToScan: 15 * 1024 * 1024,
  sentCheckTimeoutSeconds: 90
};
