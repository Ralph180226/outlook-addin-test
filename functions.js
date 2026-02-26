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
        // Let op: attachments via displayNewMessageForm zijn beperkt/fragiel,
        // we voegen de .eml bijlage straks betrouwbaar toe in compose zelf.
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
    log("Compose API klaar → bijlage toevoegen...");
    cb();
    return;
  }
  log("Compose nog niet klaar, opnieuw proberen...");
  setTimeout(() => waitForComposeReady(cb), 300);
}

// ===== Auto-start wanneer pagina/JS wordt geladen =====
Office.onReady(() => {
  try {
    if (isCompose()) {
      log("Mode: COMPOSE");
      // In UI-context (taskpane/function-file) ook automatisch proberen
      waitForComposeReady(() => attachOriginalMail());
    } else {
      log("Mode: READ");
    }
  } catch (e) {
    console.error("Startup error:", e);
  }
});

// ====== ExecuteFunction handlers (lint-knoppen) ======
function forwardPhishingCommand(event) {
  try {
    forwardPhishing();
  } finally {
    // Commands MUST call event.completed()
    // (zie docs event-based/commands)
    event.completed();
  }
}

function attachOriginalMailCommand(event) {
  try {
    waitForComposeReady(() => attachOriginalMail(() => event.completed()));
    return; // event.completed() wordt in callback aangeroepen
  } catch (e) {
    console.error(e);
  }
  event.completed();
}

// ====== LaunchEvent handler: auto-run bij nieuw compose venster ======
function onNewMessageCompose(event) {
  // Wordt getriggerd via ExtensionPoint LaunchEvent (OnNewMessageCompose)
  // Classic Outlook laadt dit via JS runtime; Web/Nieuw Outlook via HTML runtime.
  try {
    waitForComposeReady(() => attachOriginalMail(() => event.completed()));
    return; // completed in callback
  } catch (e) {
    console.error("onNewMessageCompose error:", e);
  }
  event.completed();
}

// ====== Koppel namen zodat Outlook ze kan aantreffen ======
Office.actions.associate("forwardPhishingCommand", forwardPhishingCommand);
Office.actions.associate("attachOriginalMailCommand", attachOriginalMailCommand);
Office.actions.associate("onNewMessageCompose", onNewMessageCompose);


