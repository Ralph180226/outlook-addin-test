Office.onReady(() => {});

function forwardPhishing(event) {
  try {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;

    // Moderne route indien beschikbaar
    if (item && typeof item.forwardAsync === "function") {
      item.forwardAsync(
        { toRecipients: ["ondersteuning@itssunday.nl"] },
        () => {
          // Je zou hier eventueel ook meteen item.notificationMessages.addAsync kunnen doen
          if (event && typeof event.completed === "function") event.completed();
        }
      );
      return;
    }

    if (!mailbox || !item || !item.itemId) {
      console.error("Mailbox of itemId niet beschikbaar.");
      completeWithInfo(event, "Kon dit bericht niet melden (geen itemId).");
      return;
    }

    const trySend = (ewsId) => sendDirectWithAttachment(mailbox, ewsId, event);

    const convertFn = mailbox.convertToEwsIdAsync;
    if (typeof convertFn === "function") {
      convertFn.call(mailbox, item.itemId, (conv) => {
        if (conv.status === Office.AsyncResultStatus.Succeeded) {
          trySend(conv.value);
        } else {
          trySend(item.itemId); // fallback
        }
      });
    } else {
      trySend(item.itemId); // Classic
    }
  } catch (e) {
    console.error(e);
    completeWithInfo(event, "Er ging iets mis bij het melden van dit bericht.");
  }
}

function sendDirectWithAttachment(mailbox, ewsId, event) {
  if (typeof mailbox.makeEwsRequestAsync !== "function") {
    console.warn("EWS wordt niet ondersteund in deze client.");
    completeWithInfo(event, "Deze Outlook-versie ondersteunt dit niet.");
    return;
  }

  const ews =
    `<?xml version="1.0" encoding="utf-8"?>
     <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                    xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"
                    xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
       <soap:Body>
         <m:CreateItem MessageDisposition="SendAndSaveCopy">
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
                   <t:ItemId>${escapeXml(ewsId)}</t:ItemId>
                 </t:ItemAttachment>
               </t:Attachments>
             </t:Message>
           </m:Items>
         </m:CreateItem>
       </soap:Body>
     </soap:Envelope>`;

  mailbox.makeEwsRequestAsync(ews, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // Succes: toon nette bevestiging
      try {
        mailbox.item.notificationMessages.addAsync("phishingSent", {
          type: "informationalMessage",
          message: "Dit bericht is gemeld aan IT's Sunday. Bedankt!",
          icon: "icon16",
          persistent: false
        });
      } catch (_) { /* no-op */ }
    } else {
      console.error("EWS-fout:", asyncResult.error);
      notifyUser("Kon het bericht niet automatisch verzenden.");
    }

    if (event && typeof event.completed === "function") event.completed();
  });
}

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
  } catch (e) { /* no-op */ }
}

function completeWithInfo(event, message) {
  notifyUser(message);
  if (event && typeof event.completed === "function") event.completed();
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}
