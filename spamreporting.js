Office.onReady(() => {
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onSpamReport", onSpamReport); }
});

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(
          Error encountered during message processing: ${asyncResult.error.message}
        );
        return;
      }

      // Get the user's responses to the options and text box in the preprocessing dialog.
      const spamReportingEvent = asyncResult.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;
      const event = asyncResult.asyncContext;
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
