Office.onReady(() => {});

function forwardPhishing(event) {
  try {
    // 1) Moderne forward als forwardAsync beschikbaar is
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

    // 2) Fallback voor Outlook Classic: EWS gebruiken om mail als bijlage te versturen
    let itemId = Office.context.mailbox.item.itemId;

    let ews = `
      <CreateItem MessageDisposition="SaveOnly" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
        <Items>
          <Message>
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
                <ItemId Id="${itemId}" />
              </ItemAttachment>
            </Attachments>
          </Message>
        </Items>
      </CreateItem>`;

    Office.context.mailbox.makeEwsRequestAsync(ews, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        // Draft gemaakt — open hem in een nieuw venster
        Office.context.mailbox.displayMessageForm(asyncResult.value);
      }

      if (event && typeof event.completed === "function") event.completed();
    });
