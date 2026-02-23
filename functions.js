Office.onReady(() => {});

function forwardPhishing(event) {
  try {
    // Als forwardAsync bestaat → echte forward gebruiken
    if (Office.context.mailbox.item && typeof Office.context.mailbox.item.forwardAsync === "function") {

      Office.context.mailbox.item.forwardAsync(
        { toRecipients: ["ondersteuning@itssunday.nl"] },
        function () {
          if (event && typeof event.completed === "function") event.completed();
        }
      );

    } else {
      // FALLBACK voor Outlook Classic & clients zonder forwardAsync
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Phishingmelding",
        htmlBody: "<p>Deze e-mail is gemeld als phishing.</p><p>De originele e-mail is als bijlage toegevoegd.</p>",
        attachments: [
          {
            type: Office.MailboxEnums.AttachmentType.Item,
            itemId: Office.context.mailbox.item.itemId
          }
        ]
      });

      if (event && typeof event.completed === "function") event.completed();
    }

  } catch (e) {
    console.error(e);
    if (event && typeof event.completed === "function") event.completed();
  }
}

// Voor debug/bundlers
if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}
