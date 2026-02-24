Office.onReady(() => {
  const item = Office.context.mailbox?.item;

  // Als we in COMPOSE zitten → bijlage toevoegen
  if (item && item.addItemAttachmentAsync) {
    setTimeout(attachOriginalMail, 300);
  }
});

/**
 * Stap 1 — Vanuit READ: slaan we ItemId op via RoamingSettings 
 * en openen we een nieuw compose venster
 */
function forwardPhishing(event) {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (!item || !item.itemId) {
      console.error("Geen item of itemId in READ-mode.");
      event.completed();
      return;
    }

    // Sla op in roaming settings (gedeelde opslag!)
    const settings = Office.context.roamingSettings;
    settings.set("phishOriginalId", item.itemId);

    settings.saveAsync(() => {
      // Open compose venster
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
      });

      event.completed();
    });

  } catch (e) {
    console.error("Fout in forwardPhishing:", e);
    event.completed();
  }
}

/**
 * Stap 2 — In COMPOSE: originele mail als bijlage toevoegen
 */
function attachOriginalMail() {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    const settings = Office.context.roamingSettings;
    const originalId = settings.get("phishOriginalId");

    if (!originalId) {
      console.warn("Geen originele mail gevonden in roamingSettings.");
      return;
    }

    item.addItemAttachmentAsync(
      originalId,
      "Originele e-mail",
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Bijlage toegevoegd:", res.value);

          // Opruimen
          settings.remove("phishOriginalId");
          settings.saveAsync();
        } else {
          console.error("Bijlage fout:", res.error);
        }
      }
    );

  } catch (e) {
    console.error("Fout in attachOriginalMail:", e);
  }
}
