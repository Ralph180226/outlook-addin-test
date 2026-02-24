// Kleine helper om te loggen naar UI én console
function log(msg) {
  console.log(msg);
  const box = document.getElementById("status");
  if (box) box.textContent += msg + "\n";
}

// Controle of we in COMPOSE zitten
function isCompose() {
  const item = Office.context.mailbox?.item;
  return item && typeof item.addItemAttachmentAsync === "function";
}

// ---------- READ MODE ----------
function forwardPhishing() {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox?.item;

    if (!item || !item.itemId) {
      log("Fout: geen itemId in READ mode.");
      return;
    }

    const id = item.itemId;

    // opslaan in roaming settings
    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", id);

    settings.saveAsync(() => {
      log("ItemId opgeslagen. Open compose...");

      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
      });

      // READY: Compose zal nu ons script opnieuw laden.
    });
  } catch (e) {
    log("Fout in forwardPhishing: " + e);
  }
}

// ---------- COMPOSE MODE ----------
function attachOriginalMail() {
  if (!isCompose()) {
    log("Niet in compose, bijlage toevoegen overslaan.");
    return;
  }

  const settings = Office.context.roamingSettings;
  const id = settings.get("phishOriginalId");

  if (!id) {
    log("Geen opgeslagen itemId gevonden.");
    return;
  }

  log("Bijlage toevoegen...");

  const composeItem = Office.context.mailbox.item;
  composeItem.addItemAttachmentAsync(
    id,
    "Originele e-mail",
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        log("Bijlage toegevoegd ✓");
        settings.remove("phishOriginalId");
        settings.saveAsync();
      } else {
        log("Bijlage fout: " + JSON.stringify(res.error));
      }
    }
  );
}

// ---------- AUTO: detecteer compose en voeg bijlage toe ----------
Office.onReady(() => {
  const item = Office.context.mailbox?.item;

  if (isCompose()) {
    log("Mode: COMPOSE (itemCompose gedetecteerd)");
    // kleine delay zodat compose API klaar is
    setTimeout(attachOriginalMail, 400);
  } else {
    log("Mode: READ");
  }
});
