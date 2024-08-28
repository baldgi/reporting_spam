// https://aka.ms/olksideload

/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
 */

// Ensures the Office.js library is loaded.
Office.onReady(() => {
  /**
   * IMPORTANT: To ensure your add-in is supported in the classic Outlook client on Windows,
   * remember to map the event handler name specified in the manifest to its JavaScript counterpart.
   */
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onSpamReport", onSpamReport); }
});

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the user's email address
  const userEmailAddress = Office.context.mailbox.userProfile.emailAddress;
  // Log the user's email address to the console
  console.log(`User's email address: ${userEmailAddress}`);

  // Get the Base64-encoded EML format of a reported message.
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(
          `Error encountered during message processing: ${asyncResult.error.message}`
        );
        return;
      }

      // Get the user's responses to the options and text box in the preprocessing dialog.
      const spamReportingEvent = asyncResult.asyncContext;
      console.log(`spamReportingEvent: ${spamReportingEvent}`);
      const reportedOptions = spamReportingEvent.options;
      console.log(`reportedOptions: ${reportedOptions}`);
      const additionalInfo = spamReportingEvent.freeText;
      console.log(`additionalInfo: ${additionalInfo}`);

      // Send new email with attachement
      // https://learn.microsoft.com/fr-fr/javascript/api/overview/azure/communication-email-readme?view=azure-node-latest
      const message = {
        senderAddress: ${userEmailAddress} //"sender@contoso.com",
        content: {
          subject: "This is the subject",
          plainText: "This is the body",
        },
        recipients: {
          to: [
            {
              address: "xigabik938@avashost.com",
              displayName: "Customer Name",
            },
          ],
        },
        attachments: [
          {
            name: spamReportingEvent//path.basename(filePath),
            contentType: "text/plain",
            //contentInBase64: readFileSync(filePath, "base64"),
          },
        ],
      };
      
      const poller = await emailClient.beginSend(message);
      const response = await poller.pollUntilDone();
      
      /**
       * Signals that the spam-reporting event has completed processing.
       * It then moves the reported message to the Junk Email folder of the mailbox,
       * then shows a post-processing dialog to the user.
       * If an error occurs while the message is being processed,
       * the `onErrorDeleteItem` property determines whether the message will be deleted.
       */
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
