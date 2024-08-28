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
async function onSpamReport(event) {
  try {
    // Get the user's email address
    const userEmailAddress = Office.context.mailbox.userProfile.emailAddress;
    console.log(`User's email address: ${userEmailAddress}`);

    // Get the Base64-encoded EML format of the reported message.
    Office.context.mailbox.item.getAsFileAsync(Office.CoercionType.EML, async (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
        return;
      }

      const emlContent = asyncResult.value;
      console.log("EML content retrieved successfully");

      // Example email message setup
      const message = {
        senderAddress: userEmailAddress,
        content: {
          subject: "Reported Spam Email",
          plainText: "Please see the attached email for details.",
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
            name: "reportedEmail.eml",
            contentType: "message/rfc822",
            contentInBase64: emlContent,
          },
        ],
      };

      // Log the email details for debugging
      console.log("Email Message:", JSON.stringify(message, null, 2));

      // Send the email using the Azure Communication Service (replace with actual client setup)
      try {
        const emailClient = "xigabik938@avashost.com";
        const poller = await emailClient.beginSend(message);
        const response = await poller.pollUntilDone();
        console.log("Email sent successfully:", response);
      } catch (sendError) {
        console.log("Error sending email:", sendError);
      }

      // Complete the event and move the reported email to Junk
      event.completed({
        onErrorDeleteItem: true,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.JunkFolder,
        showPostProcessingDialog: {
          title: "Contoso Spam Reporting",
          description: `Thank you for reporting this message. Email sent from ${userEmailAddress} to ${message.recipients.to[0].address}.`,
        },
      });
    });
  } catch (error) {
    console.error("An unexpected error occurred:", error);
  }
}
