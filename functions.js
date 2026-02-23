/* globals Office */

Office.onReady(() => {});

function forwardPhishing(event) {
  try {
    // 1) Moderne route: echte forward als de API beschikbaar is
    if (Office.context.mailbox.item &&
        typeof Office.context.mailbox.item.forwardAsync === "function") {

      Office.context.mailbox.item.forwardAsync(
        { toRecipients: ["ondersteuning@itssunday.nl"] },
        function () {
          if (event && typeof event.completed === "function") event.completed();
        }
      );
      return;
    }

    // 2) Fallback voor Outlook Classic (attachments niet meegeven, want dat breekt hier)
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ["ondersteuning@itssunday.nl"],
      subject: "Phishingmelding",
      // Tip: zet korte instructie in de body; gebruiker kan desgewenst zelf 'Doorsturen' op de originele mail klikken
      htmlBody:
        "<p>Deze e-mail is gemeld als phishing.</p>" +
        "<p>Tip: gebruik de knop <b>Doorsturen</b> op de originele mail als je de volledige headers wilt meesturen.</p>"
    });

    if (event && typeof event.completed === "function") event.completed();

  } catch (e) {
    
    if (event && typeof event.completed === "function") event.completed();
    // (optioneel) console.error(e);
  }
}

// (optioneel voor bundlers/tests)
if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}

