// Kleine helper om te loggen naar UI én console
function log(msg) {
  console.log(msg);
  const box = document.getElementById("status");
  if (box) box.textContent += msg + "\n";
}

// Controle: zijn we in compose mode?
function isCompose() {
  const item = Office.context.mailbox?.item;
  return item && typeof item.addItemAttachmentAsync === "function";
}

// ------------------------------------------------------------
// READ MODE – Start phishing forwarding
// ------------------------------------------------------------
function forwardPhishing() {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox?.item;

    if (!item || !item.itemId) {
      log("Fout: geen itemId in READ mode.");
      return;
    }

    const id = item.itemId;

    // itemId opslaan zodat compose het later kan ophalen
    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", id);

    settings.saveAsync(() => {
      log("ItemId opgeslagen. Open compose...");

      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
      });

      // ⭐ BELANGRIJK:
      // GEEN displayTaskPaneAsync → dit werkt NIET in Outlook Classic.
      // Compose-mode wordt automatisch geactiveerd door FormSettings.
    });
  } catch (e) {
    log("Fout in forwardPhishing: " + e);
  }
}

// ------------------------------------------------------------
// COMPOSE MODE – Id ophalen en bijlage toevoegen
// ------------------------------------------------------------
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

// ------------------------------------------------------------
// WACHTRUTINE – Compose API verschijnt soms vertraagd
// ------------------------------------------------------------
function waitForComposeReady() {
  const item = Office.context.mailbox?.item;

  if (item && typeof item.addItemAttachmentAsync === "function") {
    log("Compose API klaar → bijlage toevoegen...");
    attachOriginalMail();
    return;
  }

  log("Compose nog niet klaar, opnieuw proberen...");
  setTimeout(waitForComposeReady, 300);
}

// ------------------------------------------------------------
// AUTO-START – Detecteer compose en start bijlage-proces
// ------------------------------------------------------------
Office.onReady(() => {
  try {
    if (isCompose()) {
      log("Mode: COMPOSE");
      waitForComposeReady();
    } else {
      log("Mode: READ");
    }
  } catch (e) {
    console.error("Startup error:", e);
  }
});
