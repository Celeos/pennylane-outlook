/* Complément Outlook « Pennylane CELEOS »
 * Transfert d'un e-mail de facture vers l'adresse de collecte Pennylane :
 * corps vidé, seules les pièces jointes retenues sont conservées,
 * envoi vérifié dans Éléments envoyés, puis classement du message d'origine.
 */
(function () {
  "use strict";

  const CFG = window.CELEOS_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Mail.ReadWrite", "Mail.Send", "Mail.ReadWrite.Shared", "Mail.Send.Shared"];
  // Propriété MAPI nommée posée sur le brouillon pour retrouver l'envoi de façon certaine.
  const PROP_ID = "String {6f1e8a3c-2b7d-4c5e-9a10-3d4b5c6e7f80} Name CeleosPennylaneRef";

  const SETTING_KEY = "pennylaneAddress";
  const ADDR_RE = new RegExp("^[a-z0-9._%+-]+@" + CFG.allowedRecipientDomain.replace(/\./g, "\\.") + "$", "i");

  let pca = null;
  let pdfjsLib = null;
  let state = null;
  let busy = false;
  let dest = null; // adresse de collecte, lue dans les paramètres itinérants de la boîte

  const $ = (id) => document.getElementById(id);
  const enc = encodeURIComponent;
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  /* ---------- Démarrage ---------- */

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Outlook) return;
    $("dom").textContent = CFG.allowedRecipientDomain;
    $("btn-transfer").addEventListener("click", onTransfer);
    $("btn-save").addEventListener("click", saveAddress);
    $("edit-dest").addEventListener("click", (e) => { e.preventDefault(); showSettings(); });
    try {
      pdfjsLib = await import("./lib/pdf.min.js");
      pdfjsLib.GlobalWorkerOptions.workerSrc = "lib/pdf.worker.min.js";
    } catch (e) {
      pdfjsLib = null; // l'analyse se limitera aux noms de fichiers
    }
    try {
      pca = await msal.createNestablePublicClientApplication({
        auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId },
        cache: { cacheLocation: "sessionStorage" }
      });
    } catch (e) {
      return fail("Initialisation de l'authentification impossible.\n" + e.message);
    }
    if (Office.context.mailbox.addHandlerAsync) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { if (!busy && dest) loadItem(); });
    }
    dest = Office.context.roamingSettings.get(SETTING_KEY) || null;
    if (!dest || !ADDR_RE.test(dest)) return showSettings();
    showDest();
    loadItem();
  });

  /* ---------- Adresse de collecte ---------- */

  function showDest() {
    $("dest").textContent = dest;
    $("edit-dest").classList.remove("hidden");
  }

  function showSettings() {
    resetUi();
    $("settings").classList.remove("hidden");
    $("addr").value = dest || "";
    $("addr").focus();
  }

  function saveAddress() {
    const v = $("addr").value.trim().toLowerCase();
    if (!ADDR_RE.test(v)) {
      $("settings-err").textContent = "Adresse refusée : elle doit se terminer par @" + CFG.allowedRecipientDomain + ".";
      return;
    }
    Office.context.roamingSettings.set(SETTING_KEY, v);
    Office.context.roamingSettings.saveAsync((r) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) {
        $("settings-err").textContent = "Enregistrement impossible : " + (r.error && r.error.message);
        return;
      }
      dest = v;
      $("settings").classList.add("hidden");
      $("settings-err").textContent = "";
      showDest();
      loadItem();
    });
  }

  /* ---------- Authentification et Graph ---------- */

  let signedInUser = null; // compte Microsoft 365 connecté (UPN)

  async function getToken() {
    const req = { scopes: SCOPES };
    let r;
    try {
      r = await pca.acquireTokenSilent(req);
    } catch (e) {
      r = await pca.acquireTokenPopup(req);
    }
    if (r && r.account && r.account.username) signedInUser = r.account.username.toLowerCase();
    return r.accessToken;
  }

  // Boîte aux lettres du message affiché : la sienne, ou une boîte partagée (ex. compta@).
  function getSharedProps(item) {
    return new Promise((resolve) => {
      if (!item.getSharedPropertiesAsync) return resolve(null);
      try {
        item.getSharedPropertiesAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null));
      } catch (e) { resolve(null); }
    });
  }

  async function resolveMailbox(item) {
    await getToken();
    const sp = await getSharedProps(item);
    let addr = sp && (sp.targetMailbox || sp.owner);
    if (!addr) {
      const prof = Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress;
      if (prof && signedInUser && prof.toLowerCase() !== signedInUser) addr = prof;
    }
    addr = addr ? addr.toLowerCase() : null;
    if (!addr || addr === signedInUser) return { base: "/me", address: signedInUser, shared: false };
    return { base: "/users/" + enc(addr), address: addr, shared: true };
  }

  async function graph(method, path, body, opts = {}) {
    const token = await getToken();
    const headers = { Authorization: "Bearer " + token };
    if (body !== undefined) headers["Content-Type"] = "application/json";
    const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
      method,
      headers,
      body: body !== undefined ? JSON.stringify(body) : undefined
    });
    if (!res.ok) {
      const t = await res.text();
      throw new Error("Graph " + method + " " + path.split("?")[0] + " - " + res.status + " " + t.slice(0, 400));
    }
    if (opts.raw) return res;
    const txt = await res.text();
    return txt ? JSON.parse(txt) : null;
  }

  /* ---------- Chargement du message affiché ---------- */

  async function loadItem() {
    resetUi();
    const item = Office.context.mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
      return setStatus("Sélectionnez un message.", "");
    }
    const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    let mb;
    try { mb = await resolveMailbox(item); } catch (e) { return fail("Connexion impossible : " + e.message); }
    state = { restId, atts: [], mb };

    $("message").classList.remove("hidden");
    $("subject").textContent = item.subject || "(sans objet)";
    $("sender").textContent = item.from ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "";

    try {
      if (mb.shared) showWarning("Boîte partagée : " + mb.address + ". Le transfert partira de cette adresse et le message sera classé dans son dossier « " + CFG.targetFolderName + " ».");
      setStatus("Analyse des pièces jointes en cours...", "");
      const [msg, folder, list] = await Promise.all([
        graph("GET", state.mb.base + "/messages/" + enc(restId) + "?$select=id,parentFolderId,conversationId"),
        findTargetFolder(),
        graph("GET", state.mb.base + "/messages/" + enc(restId) + "/attachments?$select=id,name,contentType,size,isInline")
      ]);
      state.folder = folder;
      if (folder && msg.parentFolderId === folder.id) {
        showWarning((mb.shared ? $("warning").textContent + " " : "") + "Ce message se trouve déjà dans « " + CFG.targetFolderName + " » : il a probablement déjà été transféré.");
      }
      checkAlreadySent(msg.conversationId).catch(() => {});
      state.atts = await classify(list.value || []);
      renderAttachments();
      setStatus("", "");
    } catch (e) {
      fail(e.message);
    }
  }

  /* ---------- Détection des factures ---------- */

  const NAME_POS = /(factur|invoice|rechnung|fattura|factuur|\bfact[\s_-]?\d|\binv[\s_-]?\d|\bfa[\s_-]?\d{4,}|avoir|credit[\s_-]?note)/i;
  const NAME_NEG = /(receipt|re[çc]u|quittance|\bcgv\b|conditions|terms|devis|quote|pro[\s_-]?forma|relev[ée]|statement|bon[\s_-]de[\s_-](livraison|commande)|purchase[\s_-]order)/i;
  const TITLE_POS = /(facture|invoice|rechnung|fattura|factuur|avoir|credit note)/i;
  const TITLE_NEG = /(receipt|re[çc]u|payment confirmation|confirmation de paiement|devis|quotation|pro ?forma|bon de commande|purchase order|conditions g[ée]n[ée]rales|terms and conditions)/i;
  const TEXT_STRONG = /(n[°o]\s*(de\s*)?facture|num[ée]ro de facture|date de facture|invoice\s*(no\b|number|n[°o]|#|date))/i;
  const TEXT_TAX = /(\bTVA\b|\bVAT\b|\bHT\b|net amount|montant hors taxe)/i;

  async function classify(list) {
    const out = [];
    for (const a of list) {
      const isFile = a["@odata.type"] === "#microsoft.graph.fileAttachment";
      const isImage = /^image\//i.test(a.contentType || "") || /\.(png|jpe?g|gif|bmp|svg|webp)$/i.test(a.name || "");
      const isPdf = /pdf/i.test(a.contentType || "") || /\.pdf$/i.test(a.name || "");
      const r = { id: a.id, name: a.name || "(sans nom)", size: a.size || 0, isInline: !!a.isInline, isFile, isPdf, isImage, score: 0, tag: "", tagClass: "", checked: false };

      if (!isFile) { r.score = -10; r.tag = "Élément joint (e-mail)"; r.tagClass = "no"; out.push(r); continue; }
      if (r.isInline || (isImage && r.size < 60000)) { r.score = -10; r.tag = "Image intégrée / signature"; r.tagClass = "no"; out.push(r); continue; }

      if (NAME_POS.test(r.name)) r.score += 3;
      if (NAME_NEG.test(r.name)) r.score -= 4;

      if (isPdf && r.size <= CFG.maxPdfBytesToScan && pdfjsLib) {
        try {
          const txt = await pdfText(state.restId, a.id);
          if (!txt.trim()) {
            r.note = "PDF sans texte (scan)";
          } else {
            const title = txt.slice(0, 400);
            if (TITLE_POS.test(title)) r.score += 3;
            if (TITLE_NEG.test(title)) r.score -= 4;
            if (TEXT_STRONG.test(txt)) r.score += 2;
            if (TEXT_TAX.test(txt)) r.score += 1;
          }
        } catch (e) {
          r.note = "Lecture du PDF impossible";
        }
      }
      out.push(r);
    }

    const candidates = out.filter((r) => r.score > -10);
    if (candidates.length === 1) {
      candidates[0].checked = true;
      if (!candidates[0].tag) setTag(candidates[0]);
    } else {
      candidates.forEach((r) => { r.checked = r.score > 0; setTag(r); });
      if (!candidates.some((r) => r.checked) && candidates.length) {
        const best = candidates.filter((r) => r.isPdf).sort((x, y) => y.score - x.score)[0] || candidates[0];
        best.checked = true;
        best.tag = "Incertain, à vérifier"; best.tagClass = "doubt";
      }
    }
    return out;
  }

  function setTag(r) {
    if (r.score >= 3) { r.tag = "Facture détectée"; r.tagClass = "ok"; }
    else if (r.score > 0) { r.tag = "Facture probable"; r.tagClass = "doubt"; }
    else if (r.score < 0) { r.tag = "Pas une facture"; r.tagClass = "no"; }
    else { r.tag = "Indéterminé"; r.tagClass = "doubt"; }
  }

  async function pdfText(msgId, attId) {
    const res = await graph("GET", state.mb.base + "/messages/" + enc(msgId) + "/attachments/" + enc(attId) + "/$value", undefined, { raw: true });
    const buf = await res.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, enableXfa: false, disableFontFace: true }).promise;
    let txt = "";
    for (let i = 1; i <= Math.min(pdf.numPages, 2); i++) {
      const page = await pdf.getPage(i);
      const c = await page.getTextContent();
      txt += c.items.map((it) => it.str).join(" ") + "\n";
    }
    return txt;
  }

  /* ---------- Affichage ---------- */

  function renderAttachments() {
    const box = $("att-list");
    box.innerHTML = "";
    $("attachments").classList.remove("hidden");
    $("actions").classList.remove("hidden");

    if (!state.atts.length) {
      box.innerHTML = '<div class="muted">Aucune pièce jointe.</div>';
      $("att-hint").textContent = "Pennylane a besoin d'une pièce jointe pour créer la facture.";
      updateButton();
      return;
    }
    state.atts.forEach((r, i) => {
      const row = document.createElement("label");
      row.className = "att";
      row.innerHTML =
        '<input type="checkbox" data-i="' + i + '"' + (r.checked ? " checked" : "") + ">" +
        '<div><div class="name">' + esc(r.name) + (r.tag ? '<span class="tag ' + r.tagClass + '">' + esc(r.tag) + "</span>" : "") + "</div>" +
        '<div class="meta">' + formatSize(r.size) + (r.note ? " - " + esc(r.note) : "") + "</div></div>";
      box.appendChild(row);
    });
    box.querySelectorAll("input").forEach((cb) => cb.addEventListener("change", (ev) => {
      state.atts[+ev.target.dataset.i].checked = ev.target.checked;
      updateButton();
    }));
    $("att-hint").textContent = "Les éléments décochés seront retirés du transfert. Le texte du message n'est pas transmis.";
    updateButton();
  }

  function updateButton() {
    const n = state ? state.atts.filter((a) => a.checked).length : 0;
    $("btn-transfer").disabled = busy || n === 0;
    $("btn-transfer").textContent = n > 1 ? "Transférer " + n + " pièces jointes et classer" : "Transférer et classer";
  }

  function formatSize(b) {
    if (b < 1024) return b + " o";
    if (b < 1048576) return (b / 1024).toFixed(0) + " Ko";
    return (b / 1048576).toFixed(1).replace(".", ",") + " Mo";
  }

  function resetUi() {
    ["settings", "message", "attachments", "actions", "progress"].forEach((id) => $(id).classList.add("hidden"));
    $("warning").classList.add("hidden");
    $("steps").innerHTML = "";
    setStatus("", "");
  }

  function showWarning(t) { $("warning").textContent = t; $("warning").classList.remove("hidden"); }
  function setStatus(t, cls) { const s = $("status"); s.textContent = t; s.className = "status " + (cls || ""); }
  function fail(t) { setStatus(t, "err"); }

  let currentStep = null;
  function step(t) {
    if (currentStep) currentStep.classList.add("done");
    $("progress").classList.remove("hidden");
    currentStep = document.createElement("li");
    currentStep.textContent = t;
    $("steps").appendChild(currentStep);
  }

  /* ---------- Transfert ---------- */

  async function onTransfer() {
    if (busy) return;
    const selected = state.atts.filter((a) => a.checked);
    if (!selected.length) return;
    if (!dest || !ADDR_RE.test(dest)) return showSettings();
    busy = true; updateButton();
    $("steps").innerHTML = ""; currentStep = null; setStatus("", "");
    const msgId = state.restId;
    let draftId = null;

    try {
      step("Création du brouillon de transfert");
      const draft = await graph("POST", state.mb.base + "/messages/" + enc(msgId) + "/createForward", {});
      draftId = draft.id;

      step("Destinataire Pennylane, corps vidé");
      const ref = (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random());
      await graph("PATCH", state.mb.base + "/messages/" + enc(draftId), {
        toRecipients: [{ emailAddress: { address: dest, name: CFG.pennylaneName } }],
        ccRecipients: [],
        bccRecipients: [],
        body: { contentType: "HTML", content: "<html><body></body></html>" },
        singleValueExtendedProperties: [{ id: PROP_ID, value: ref }]
      });

      step("Retrait des pièces jointes non retenues");
      const dAtts = (await graph("GET", state.mb.base + "/messages/" + enc(draftId) + "/attachments?$select=id,name,size,isInline")).value || [];
      const keep = matchKeep(dAtts, selected);
      await Promise.all(dAtts.filter((d) => !keep.has(d.id))
        .map((d) => graph("DELETE", state.mb.base + "/messages/" + enc(draftId) + "/attachments/" + enc(d.id))));

      step("Contrôle du brouillon avant envoi");
      const check = await graph("GET", state.mb.base + "/messages/" + enc(draftId) + "?$select=toRecipients,ccRecipients,bccRecipients,body,internetMessageId,conversationId&$expand=attachments($select=id,name)");
      verifyMessage(check, selected, true);

      step("Envoi");
      await graph("POST", state.mb.base + "/messages/" + enc(draftId) + "/send");
      draftId = null;

      step("Vérification dans Éléments envoyés");
      await waitForSent(ref, selected, { imid: check.internetMessageId, cid: check.conversationId, t0: Date.now() - 60000 });

      step("Classement dans « " + CFG.targetFolderName + " »");
      const folder = state.folder || (await findTargetFolder()) || (await createTargetFolder());
      await graph("POST", state.mb.base + "/messages/" + enc(msgId) + "/move", { destinationId: folder.id });

      if (currentStep) currentStep.classList.add("done");
      setStatus("Transfert envoyé, vérifié et message classé.", "ok");
      $("actions").classList.add("hidden");
    } catch (e) {
      if (currentStep) currentStep.classList.add("err");
      let msg = e.message;
      if (draftId) {
        try { await graph("DELETE", state.mb.base + "/messages/" + enc(draftId)); msg += "\nLe brouillon a été supprimé, rien n'a été envoyé."; }
        catch (_) { msg += "\nLe brouillon n'a pas pu être supprimé : vérifiez le dossier Brouillons."; }
      }
      fail(msg);
    } finally {
      busy = false; updateButton();
    }
  }

  // Associe les pièces jointes du brouillon (nouveaux identifiants) aux pièces retenues, par nom.
  function matchKeep(dAtts, selected) {
    const want = {};
    selected.forEach((s) => { want[s.name] = (want[s.name] || 0) + 1; });
    const keep = new Set();
    const ordered = dAtts.slice().sort((a, b) => (a.isInline === b.isInline ? 0 : a.isInline ? 1 : -1));
    for (const d of ordered) {
      if (want[d.name] > 0) { keep.add(d.id); want[d.name]--; }
    }
    const missing = Object.keys(want).filter((k) => want[k] > 0);
    if (missing.length) throw new Error("Pièce(s) jointe(s) introuvable(s) dans le brouillon : " + missing.join(", "));
    return keep;
  }

  function verifyMessage(m, selected, isDraft) {
    const errs = [];
    const to = (m.toRecipients || []).map((r) => (r.emailAddress.address || "").toLowerCase());
    if (to.length !== 1 || to[0] !== dest || !ADDR_RE.test(to[0])) errs.push("destinataire inattendu (" + to.join(", ") + ")");
    if (isDraft && ((m.ccRecipients || []).length || (m.bccRecipients || []).length)) errs.push("copie(s) présente(s)");
    if (isDraft && m.body) {
      const text = (m.body.content || "").replace(/<style[\s\S]*?<\/style>/gi, "").replace(/<[^>]+>/g, "").replace(/&nbsp;|\s/g, "");
      if (text.length) errs.push("le corps du message n'est pas vide");
    }
    const got = (m.attachments || []).map((a) => a.name).sort();
    const exp = selected.map((s) => s.name).sort();
    if (JSON.stringify(got) !== JSON.stringify(exp)) errs.push("pièces jointes attendues : " + exp.join(", ") + " / présentes : " + (got.join(", ") || "aucune"));
    if (errs.length) throw new Error("Contrôle " + (isDraft ? "du brouillon" : "de l'envoi") + " en échec : " + errs.join(" ; "));
  }

  // Recherche du message envoyé par trois critères successifs : identifiant Internet du message,
  // propriété posée sur le brouillon, puis conversation + destinataire + heure d'envoi.
  async function waitForSent(ref, selected, meta) {
    const SEL = "&$select=id,toRecipients,sentDateTime&$expand=attachments($select=id,name)";
    const q = (base, filter) => base + "/mailFolders/sentitems/messages?$filter=" + enc(filter) + SEL;
    const bases = state.mb.shared ? [state.mb.base, "/me"] : [state.mb.base];
    const filters = [];
    if (meta && meta.imid) filters.push({ f: "internetMessageId eq '" + meta.imid.replace(/'/g, "''") + "'" });
    filters.push({ f: "singleValueExtendedProperties/any(ep:ep/id eq '" + PROP_ID + "' and ep/value eq '" + ref + "')" });
    if (meta && meta.cid) filters.push({ f: "conversationId eq '" + meta.cid.replace(/'/g, "''") + "'", conv: true });
    const deadline = Date.now() + CFG.sentCheckTimeoutSeconds * 1000;
    const delays = [400, 600, 800, 1000, 1500, 2000];
    let i = 0;
    while (Date.now() < deadline) {
      await sleep(delays[Math.min(i++, delays.length - 1)]);
      for (const b of bases) {
        for (const c of filters) {
          let r;
          try { r = await graph("GET", q(b, c.f)); } catch (e) { continue; } // critère non pris en charge : on passe au suivant
          let hits = (r && r.value) || [];
          if (c.conv) {
            hits = hits.filter((m) => new Date(m.sentDateTime).getTime() >= meta.t0 &&
              (m.toRecipients || []).some((t) => (t.emailAddress.address || "").toLowerCase() === dest));
          }
          if (hits.length) {
            verifyMessage(hits[0], selected, false);
            return hits[0];
          }
        }
      }
    }
    throw new Error("Le message envoyé n'apparaît pas dans Éléments envoyés après " + CFG.sentCheckTimeoutSeconds + " s. Le mail d'origine n'a pas été déplacé : vérifiez Éléments envoyés avant de relancer.");
  }

  // Signale un transfert déjà envoyé à Pennylane dans la même conversation.
  async function checkAlreadySent(cid) {
    if (!cid || !dest) return;
    const bases = state.mb.shared ? [state.mb.base, "/me"] : [state.mb.base];
    for (const b of bases) {
      const r = await graph("GET", b + "/mailFolders/sentitems/messages?$filter=" + enc("conversationId eq '" + cid.replace(/'/g, "''") + "'") + "&$select=toRecipients,sentDateTime");
      const prev = ((r && r.value) || []).filter((m) => (m.toRecipients || []).some((t) => (t.emailAddress.address || "").toLowerCase() === dest));
      if (prev.length) {
        const d = prev.map((m) => new Date(m.sentDateTime)).sort((a, b2) => b2 - a)[0];
        const txt = "Déjà transféré à Pennylane le " + d.toLocaleDateString("fr-FR") + " à " + d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" }) +
          (prev.length > 1 ? " (" + prev.length + " envois)" : "") + ".";
        showWarning(($("warning").classList.contains("hidden") ? "" : $("warning").textContent + " ") + txt);
        return;
      }
    }
  }

  async function findTargetFolder() {
    const name = CFG.targetFolderName.replace(/'/g, "''");
    const r = await graph("GET", state.mb.base + "/mailFolders/inbox/childFolders?$filter=" + enc("displayName eq '" + name + "'") + "&$select=id,displayName");
    return (r.value && r.value[0]) || null;
  }

  async function createTargetFolder() {
    return graph("POST", state.mb.base + "/mailFolders/inbox/childFolders", { displayName: CFG.targetFolderName });
  }
})();
