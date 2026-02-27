// ===== Kleine logger =====
function log(msg) {
  console.log(msg);
  const box = document.getElementById("status");
  if (box) box.textContent += msg + "\n";
}

// Check compose mode
function isCompose() {
  const item = Office.context.mailbox?.item;
  return !!(item && typeof item.addItemAttachmentAsync === "function");
}

// ===== READ: sla originele mail-ID op + open compose =====
function sendPhishingReport() {
  try {
    const item = Office.context.mailbox?.item;
    if (!item || !item.itemId) {
      log("Fout: geen itemId beschikbaar.");
      return;
    }

    const id = item.itemId;
    const extra = document.getElementById("extra")?.value || "";

    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", id);
    settings.set("phishExtra", extra);
    settings.saveAsync(() => {
      log("Gegevens opgeslagen. Open compose...");
      
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding ITS Sunday",
        htmlBody: "<p>Er is een phishingmelding verstuurd.</p><p><b>Extra info:</b><br>" +
                  (extra.trim() ? extra : "(geen)") +
                  "</p>"
      });
    });
  } catch (e) {
    log("Fout sendPhishingReport(): " + e);
  }
}

// ===== COMPOSE: voeg originele mail toe =====
function attachOriginalMail(done) {
  if (!isCompose()) {
    log("Niet in compose mode.");
    done && done();
    return;
  }

  const settings = Office.context.roamingSettings;
  const id = settings.get("phishOriginalId");

  if (!id) {
    log("Geen opgeslagen itemId gevonden.");
    done && done();
    return;
  }

  log("Bijlage toevoegen...");

  Office.context.mailbox.item.addItemAttachmentAsync(
    id,
    "Originele e-mail.eml",
    (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        log("Originele e-mail toegevoegd ✓");
        settings.remove("phishOriginalId");
        settings.remove("phishExtra");
        settings.saveAsync();
      } else {
        log("Bijlage fout: " + JSON.stringify(res.error));
      }
      done && done();
    }
  );
}

// ===== Compose ready helper =====
function waitForComposeReady(cb) {
  const item = Office.context.mailbox?.item;
  if (item && typeof item.addItemAttachmentAsync === "function") {
    cb();
    return;
  }
  setTimeout(() => waitForComposeReady(cb), 300);
}

// ===== Auto-run in compose =====
Office.onReady(() => {
  try {
    if (isCompose()) {
      log("Compose modus gedetecteerd.");
      waitForComposeReady(() => attachOriginalMail());
    } else {
      log("Read modus gedetecteerd.");
    }
  } catch (e) {
    console.error(e);
  }
});

