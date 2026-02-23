Office.onReady(() => {});

function forwardPhishing(event) {
  try {

    // 1) Moderne route als forwardAsync beschikbaar is
    if (Office.context.mailbox.item &&
        typeof Office.context.mailbox.item.forwardAsync === "function") {

      Office.context.mailbox.item.forwardAsync(
        { toRecipients: ["ondersteuning@itssunday.nl"] },
        function() {
          if (event && typeof event.completed === "function") event.completed();
        }
      );
      return;
    }

    // 2) Fallback EWS voor Outlook Classic
    const itemId = Office.context.mailbox.item.itemId;

    const ews =
      `<CreateItem MessageDisposition="SaveOnly" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
        <Items>
          <Message xmlns="http://schemas.microsoft.com/exchange/services/2006/types">
            <Subject>Phishingmelding</Subject>
            <Body BodyType="HTML">Deze e-mail is gemeld als phishing.</Body>
            <ToRecipients>
              <Mailbox>
                <EmailAddress>ondersteuning@itssunday.nl</EmailAddress>
              </Mailbox>
            </ToRecipients>
            <Attachments>
              <ItemAttachment>
                <Name>Originele email.eml</Name>
                <ItemId>${itemId}</ItemId>
              </ItemAttachment>
            </Attachments>
          </Message>
        </Items>
      </CreateItem>`;

    Office.context.mailbox.makeEwsRequestAsync(ews, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const response = asyncResult.value;
        Office.context.mailbox.displayMessageForm(response);
      }

      if (event && typeof event.completed === "function") event.completed();
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
