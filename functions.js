Office.onReady(() => {
  const item = Office.context.mailbox?.item;

  // In compose mode: probeer automatisch bijlage toe te voegen
  if (item && item.addItemAttachmentAsync) {
    setTimeout(attachOriginalMail, 300);
  }
});

/**
 * READ-modus:
 * 1. ItemId opslaan (RoamingSettings = gedeelde opslag)
 * 2. Nieuw bericht openen
 */
function forwardPhishing() {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (!item || !item.itemId) {
      console.error("Geen itemId in READ-modus.");
      return;
    }

    const originalId = item.itemId;
    const settings = Office.context.roamingSettings;

    // Sla originele mail op
    settings.set("phishOriginalId", originalId);

    settings.saveAsync(() => {
      // Open nieuw compose bericht
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "Deze e-mail is gemeld als phishing."
      });
    });

  } catch (e) {
    console.error("Fout in forwardPhishing:", e);
  }
}

/**
 * COMPOSE-modus:
 * Originele mail automatisch als bijlage toevoegen
 */
function attachOriginalMail() {
  try {
    const settings = Office.context.roamingSettings;
    const originalId = settings.get("phishOriginalId");

    if (!originalId) {
      console.warn("Geen originele mail gevonden in roamingSettings.");
      return;
    }

    const composeItem = Office.context.mailbox.item;

    composeItem.addItemAttachmentAsync(
      originalId,
      "Originele e-mail",
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Bijlage toegevoegd:", res.value);

          // Opschonen
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
