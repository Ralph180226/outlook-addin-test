
Office.onReady(() => { });

function forwardPhishing(event) {
  try {

    // 1) Moderne route als forwardAsync beschikbaar is
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

    // 2) Fallback EWS voor Outlook Classic
    const mailbox = Office.context.mailbox;
    const itemId = mailbox.item.itemId;

    // Eerst REST-ID naar EWS-ID converteren
    mailbox.convertToEwsIdAsync(itemId, (idResult) => {

      if (idResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("EWS conversie mislukt:", idResult.error);
        if (event) event.completed();
        return;
      }

      const ewsId = idResult.value;

      // Correcte SOAP envelope
      const ews =
        `<?xml version="1.0" encoding="utf-8"?>
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
                       <t:ItemId>${ewsId}</t:ItemId>
                     </t:ItemAttachment>
                   </t:Attachments>
                 </t:Message>
               </m:Items>
             </m:CreateItem>
           </soap:Body>
         </soap:Envelope>`;

      mailbox.makeEwsRequestAsync(ews, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

          // ID van nieuw item uit response halen
          const response = asyncResult.value;
          const match = response.match(/<t:ItemId Id="([^"]+)"/);

          if (match) {
            const newId = match[1];
            mailbox.displayMessageForm(newId);
          } else {
            console.error("Kon nieuw ItemId niet vinden in EWS response.");
          }
        }

        if (event && typeof event.completed === "function") event.completed();
      });
    });

  } catch (e) {
    console.error(e);
    if (event && typeof event.completed === "function") event.completed();
  }
}

// Voor debugging / bundlers
if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}

