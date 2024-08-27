Office.onReady(() => {
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onSpamReport", onSpamReport);
  }
});

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(Office.CoercionType.EML, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
      return;
    }

    // Get the EML file as a Base64 string
    const emlFile = asyncResult.value;

    // Create a new email item to send the reported email as an attachment
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: ['xigabik938@avashost.com'],
      subject: 'Reported Spam Email',
      body: `
        <p>Hello IT Team,</p>
        <p>A suspicious email has been reported. Please find the attached email for your review.</p>
        <p>Best regards,<br>Your Outlook Add-in</p>`,
      attachments: [
        {
          type: Office.MailboxEnums.AttachmentType.File,
          name: 'reportedEmail.eml',
          url: URL.createObjectURL(new Blob([emlFile], { type: 'message/rfc822' }))
        }
      ]
    });

    // Complete the event
    event.completed({
      onErrorDeleteItem: true,
      moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
      showPostProcessingDialog: {
        title: "Contoso Spam Reporting",
        description: "Thank you for reporting this message.",
      },
    });
  });
}
