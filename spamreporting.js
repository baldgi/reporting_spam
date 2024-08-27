Office.onReady(() => {
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onSpamReport", onSpamReport);
  }
});

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(
    Office.CoercionType.EML,
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
        return;
      }

      const file = asyncResult.value;
      const base64Content = file.content;

      // Create an email item to send the reported message to the support team.
      Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["xigabik938@avashost.com"],
        subject: "Reported Suspicious Email",
        attachments: [
          {
            type: Office.MailboxEnums.AttachmentType.File,
            name: "reported_email.eml",
            url: `data:message/rfc822;base64,${base64Content}`
          }
        ],
        body: {
          contentType: Office.MailboxEnums.BodyType.Text,
          content: "Please find the attached suspicious email reported by the user."
        }
      });

      // Complete the event and show a confirmation dialog.
      event.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Contoso Spam Reporting",
          description: "Thank you for reporting this message.",
        },
      });
    }
  );
}
