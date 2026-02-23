Office.onReady(() => {});

function forwardPhishing() {
  Office.context.mailbox.item.forwardAsync(
    { toRecipients: ["ondersteuning@itssunday.nl"] },
    function () {
      // Sluit het paneel (optioneel in klassieke varianten)
      try { Office.context.ui.closeContainer(); } catch (e) {}
    }
  );
}
