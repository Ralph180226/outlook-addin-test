// ===== Kleine logger =====
function log(msg) {
  console.log(msg);
  const box = document.getElementById("status");
  if (box) box.textContent += msg + "\n";
}

// ===== Helpers =====
function isCompose() {
  const item = Office.context.mailbox?.item;
  return !!(item && typeof item.addItemAttachmentAsync === "function");
}

// ===== READ: Meld phishing → nieuw concept openen =====
function forwardPhishing() {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox?.item;

    if (!item || !item.itemId) {
      log("Fout: geen itemId in READ mode.");
      return;
    }

    const id = item.itemId;

    // Id opslaan voor compose-handler
    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", id);

    settings.saveAsync(() => {
      log("ItemId opgeslagen. Open compose...");

      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
        // Bijlagen via displayNewMessageForm laten we achterwege;
        // we voegen de .eml bijlage in compose zelf toe via de knop.
      });
    });
  } catch (e) {
    log("Fout in forwardPhishing: " + e);
  }
}

// ===== COMPOSE: originele mail toevoegen als .eml =====
function attachOriginalMail(done) {
  if (!isCompose()) {
    log("Niet in compose, bijlage toevoegen overslaan.");
    if (typeof done === "function") done();
    return;
  }

  const settings = Office.context.roamingSettings;
  const id = settings.get("phishOriginalId");

  if (!id) {
    log("Geen opgeslagen itemId gevonden.");
    if (typeof done === "function") done();
    return;
  }

  log("Bijlage toevoegen...");

  const composeItem = Office.context.mailbox.item;
  composeItem.addItemAttachmentAsync(
    id,                   // EWS ItemId van de originele mail
    "Originele e-mail",   // weergavenaam
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        log("Bijlage toegevoegd ✓");
        settings.remove("phishOriginalId");
        settings.saveAsync();
      } else {
        log("Bijlage fout: " + JSON.stringify(res.error));
      }
      if (typeof done === "function") done();
    }
  );
}

// ===== Wacht tot compose-API beschikbaar is (soms async geladen) =====
function waitForComposeReady(cb) {
  const item = Office.context.mailbox?.item;
  if (item && typeof item.addItemAttachmentAsync === "function") {
    cb();
    return;
  }
  setTimeout(() => waitForComposeReady(cb), 300);
}

// ===== Click-handler voor de COMPOSE-knop =====
function onComposeAttachClick() {
  // geen auto-attach bij load; pas bij klik
  waitForComposeReady(() => attachOriginalMail());
}

// ===== Auto-start: alleen logging (geen auto-attach) =====
Office.onReady(() => {
  try {
    if (isCompose()) {
      log("Mode: COMPOSE");
      // GEEN automatische bijlage meer hier.
    } else {
      log("Mode: READ");
    }
  } catch (e) {
    console.error("Startup error:", e);
  }
});
