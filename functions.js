Office.onReady(() => {});

function forwardPhishing(event) {
  Office.context.mailbox.item.forwardAsync(
    {
      toRecipients: ["ondersteuning@itssunday.nl"]
    },
    function () {
      event.completed();
    }
  );
}

if (typeof module !== "undefined") {
  module.exports = { forwardPhishing };
}
