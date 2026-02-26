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

    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", id);

    settings.saveAsync(() => {
      log("ItemId opgeslagen. Open compose...");

      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
      });

      // ⭐ BELANGRIJK: taskpane opnieuw openen zodat compose-JS draait
      Office.context.ui.displayTaskPaneAsync(
        "https://ralph180226.github.io/outlook-addin-test/function-file.html"
      );
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

// ---------- NIEUW: Betrouwbare wachtroutine voor compose ----------
function waitForComposeReady() {
  const item = Office.context.mailbox?.item;

  // Compose API beschikbaar?
  if (item && typeof item.addItemAttachmentAsync === "function") {
    log("Compose API klaar → bijlage toevoegen...");
    attachOriginalMail();
    return;
  }

  // Nog niet klaar → opnieuw proberen
  log("Compose nog niet klaar, opnieuw proberen...");
  setTimeout(waitForComposeReady, 300);
}

// ---------- AUTO: detecteer compose en wacht tot API klaar is ----------
Office.onReady(() => {
  const item = Office.context.mailbox?.item;

  if (isCompose()) {
    log("Mode: COMPOSE (direct gedetecteerd)");
    waitForComposeReady();
  } else {
    log("Mode: READ");
  }
});

