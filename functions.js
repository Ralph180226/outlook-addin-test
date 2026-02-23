Office.onReady(() => {});

function forwardPhishing() {
  Office.context.mailbox.item.forwardAsync(
    {
      toRecipients: ["ondersteuning@itssunday.nl"]
    },
    function () {
      Office.context.ui.closeContainer();
    }
  );
}
