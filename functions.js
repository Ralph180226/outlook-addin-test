Office.onReady(() => {});

function forwardPhishing(event) {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    // Als forwardAsync bestaat (nieuwe Outlook/web), gebruik die – opent compose maar voegt geen .eml toe.
    if (item && typeof item.forwardAsync === "function") {
      item.forwardAsync({ toRecipients: ["ondersteuning@itssunday.nl"] }, () => {
        if (event && typeof event.completed === "function") event.completed();
      });
      return;
    }

    // Classic Outlook EWS-pad
    if (!mailbox || !item || !item.itemId) {
      console.error("Geen mailbox of itemId.");
      if (event && typeof event.completed === "function") event.completed();
      return;
    }

    // In Classic is item.itemId al EWS-compatibel → direct gebruiken
    const originalId = item.itemId;

    const ews =
      <?xml version="1.0" encoding="utf-8"?>
       <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                      xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                      xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
         <soap:Body>
           <m:CreateItem MessageDisposition="SaveOnly">
             <m:Items>
               <t:Message>
                 <t:Subject>Phishingmelding</t:Subject>
                 <t:Body BodyType="HTML">Deze e-mail is gemeld als phishing.</t:Body>
                 <t:ToRecipients>
                   <t:Mailbox>
                     <t:EmailAddress>ondersteuning@itssunday.nl</t:EmailAddress>
                   </t:Mailbox>
                 </t:ToRecipients>
                 <t:Attachments>
                   <t:ItemAttachment>
                     <t:Name>Originele email.eml</t:Name>
                     <t:ItemId>${escapeXml(originalId)}</t:ItemId>
                   </t:ItemAttachment>
                 </t:Attachments>
               </t:Message>
             </m:Items>
           </m:CreateItem>
         </soap:Body>
       </soap:Envelope>;

    mailbox.makeEwsRequestAsync(ews, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const xml = asyncResult.value;
        // Haal de nieuwe ItemId uit de EWS-response
        const match = xml.match(/<t:ItemId Id="([^"]+)"/);
        if (match) {
          const newId = match[1];
          mailbox.displayMessageForm(newId); // Open concept met .eml bijlage
        } else {
          console.error("Kon nieuw ItemId niet vinden in EWS-response.");
          notifyUser("Concept is aangemaakt, maar kon niet automatisch geopend worden.");
        }
      } else {
        console.error("EWS-fout:", asyncResult.error);
        notifyUser("Kon geen concept maken via EWS.");
      }

      if (event && typeof event.completed === "function") event.completed();
    });

  } catch (e) {
    console.error(e);
    if (event && typeof event.completed === "function") event.completed();
  }
}

// Optionele notificatie in de leesweergave
function notifyUser(message) {
  try {
    const nm = Office.context?.mailbox?.item?.notificationMessages;
    if (nm && typeof nm.addAsync === "function") {
      nm.addAsync("phishingInfo", {
        type: "informationalMessage",
        message,
        icon: "icon16",
        persistent: false
      });
    }
  } catch (_) { /* ignore */ }
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

// Voor bundlers/debug
if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}

