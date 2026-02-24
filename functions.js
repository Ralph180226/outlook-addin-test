Office.onReady(() => {});

/**
 * Hoofdactie vanuit de leesweergave (ItemRead)
 */
function forwardPhishing(event) {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    if (!mailbox || !item || !item.itemId) {
      console.error("Geen mailbox of itemId.");
      if (event && typeof event.completed === "function") event.completed();
      return;
    }

    const originalId = item.itemId;

    // Bewaar het ItemId zodat compose het later kan ophalen
    localStorage.setItem("phishOriginalId", originalId);

    // Open een nieuwe mail (COMPOSE)
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["ondersteuning@itssunday.nl"],
      subject: "Phishingmelding",
      htmlBody: "Deze e-mail is gemeld als phishing."
    });

    if (event && typeof event.completed === "function") event.completed();
  } catch (e) {
    console.error(e);
    if (event && typeof event.completed === "function") event.completed();
  }
}

/**
 * Functie die automatisch draait als we in COMPOSE zitten.
 * Voegt de originele mail toe als bijlage.
 */
function attachOriginalEmail() {
  try {
    const originalId = localStorage.getItem("phishOriginalId");
    if (!originalId) {
      console.log("Geen originele mail gevonden in storage.");
      return;
    }

    Office.context.mailbox.item.addItemAttachmentAsync(
      originalId,
      "Originele email",
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Originele email toegevoegd als bijlage:", result.value);
          localStorage.removeItem("phishOriginalId"); // opschonen
        } else {
          console.error("Bijlage mislukt:", result.error);
        }
      }
    );
  } catch (e) {
    console.error("Fout bij het toevoegen van de bijlage:", e);
  }
}

/**
 * Automatische COMPOSE-detectie
 */
Office.onReady(() => {
  const item = Office.context.mailbox?.item;

  // Check COMPOSE-modus
  if (
    item &&
    (item.displayReplyForm || item.addItemAttachmentAsync) // kenmerken van compose
  ) {
    // Kleine delay zodat compose UI stabiel is
    setTimeout(attachOriginalEmail, 300);
  }
});

if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}
