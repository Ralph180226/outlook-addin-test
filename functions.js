// ===== Logger =====
function log(msg) {
  console.log(msg);
  const box = document.getElementById("status");
  if (box) box.textContent += msg + "\n";
}

// ===== Helpers =====
function isCompose() {
  const it = Office.context.mailbox?.item;
  return !!(it && typeof it.addItemAttachmentAsync === "function");
}

// ===== READ: stuur melding & open compose =====
function sendPhishingReport() {
  const item = Office.context.mailbox?.item;
  if (!item || !item.itemId) {
    log("Geen itemId gevonden.");
    return;
  }

  const id = item.itemId;
  const extra = document.getElementById("extra")?.value || "";

  const settings = Office.context.roamingSettings;
  settings.set("origMailId", id);
  settings.set("extraText", extra);

  settings.saveAsync(() => {
    log("Gegevens opgeslagen. Open compose...");
    
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["ondersteuning@itssunday.nl"],
      subject: "Phishingmelding ITS Sunday",
      htmlBody:
        "<p>Er is een phishingmelding verstuurd.</p>" +
        "<p><b>Extra info:</b><br>" +
        (extra.trim() ? extra : "(geen)") + "</p>"
    });
  });
}

// ===== COMPOSE: voeg originele mail toe =====
function attachOriginalMail(done) {
  if (!isCompose()) {
    done?.();
    return;
  }

  const settings = Office.context.roamingSettings;
  const id = settings.get("origMailId");

  if (!id) {
    log("Geen originele mail-id gevonden.");
    done?.();
    return;
  }

  log("Originele e-mail wordt toegevoegd...");

  Office.context.mailbox.item.addItemAttachmentAsync(
    id,
    "Originele-email.eml",
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        log("Originele e-mail toegevoegd ✓");
        settings.remove("origMailId");
        settings.remove("extraText");
        settings.saveAsync();
      } else {
        log("Bijlage fout: " + JSON.stringify(res.error));
      }
      done?.();
    }
  );
}

// ===== Wacht tot compose API klaar is =====
function waitForComposeReady(cb) {
  const it = Office.context.mailbox?.item;
  if (it && typeof it.addItemAttachmentAsync === "function") {
    cb();
    return;
  }
  setTimeout(() => waitForComposeReady(cb), 300);
}

// ===== Auto-run in compose pane =====
Office.onReady(() => {
  if (isCompose()) {
    waitForComposeReady(() => attachOriginalMail());
  }
});

// ===== Ribbon command handler =====
function attachOriginalMailCommand(event) {
  waitForComposeReady(() => attachOriginalMail(() => event.completed()));
}
