Office.onReady(() => {
  // Optional: console.log("ITS Sunday add-in ready");
});

function forwardPhishing(event) {
  Office.context.mailbox.item.forwardAsync(
    { toRecipients: ["ondersteuning@itssunday.nl"] },

    function (asyncResult) {
      // Event moet worden afgesloten zodat Outlook weet dat de actie klaar is
      if (event && typeof event.completed === "function") {
        event.completed();
      }
    }
  );
}

// Nodig voor debug of bundlers
if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}
