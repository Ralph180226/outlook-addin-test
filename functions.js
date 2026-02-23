Office.onReady(() => {});

function forwardPhishing(event) {
  try {

    // 1) Moderne route
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

    // 2) EWS route (Classic Outlook)
    const mailbox = Office.context.mailbox;
    const originalId = mailbox.item.itemId; // GEWOON DIRECT gebruiken

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
                     <t:ItemId>${originalId}</t:ItemId>
                   </t:ItemAttachment>
                 </t:Attachments>
               </t:Message>
             </m:Items>
           </m:CreateItem>
         </soap:Body>
       </soap:Envelope>`;

    mailbox.makeEwsRequestAsync(ews, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

        // ItemId uit EWS-response halen:
        const xml = asyncResult.value;
        const match = xml.match(/<t:ItemId Id="([^"]+)"/);

        if (match) {
          mailbox.displayMessageForm(match[1]); // Toon concept met .eml bijlage
        } else {
          console.error("Kon nieuw ItemId niet vinden.");
        }
      }

      if (event && typeof event.completed === "function") event.completed();
    });

  } catch (e) {
    console.error(e);
    if (event && typeof event.completed === "function") event.completed();
  }
}
