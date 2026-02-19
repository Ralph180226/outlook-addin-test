Office.onReady(() => {});

function reportPhishing(event) {

    const item = Office.context.mailbox.item;

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["ondersteuning@itssunday.nl"],
        subject: "Verdachte mail gemeld: " + item.subject,
        htmlBody: "<p>Deze e-mail is gemeld als verdacht.</p>"
    });

    event.completed();
}
