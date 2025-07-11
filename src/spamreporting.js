Office.onReady();

// Handles the SpamReporting event to process a reported message.
function onSpamReport(event) {
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
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      // Run additional processing operations here.

      /**
       * Signals that the spam-reporting event has completed processing and shows a post-processing dialog to the user.
       * If an error occurs while the message is being processed, the `onErrorDeleteItem` property determines whether the message will be deleted.
       */
      const event = asyncResult.asyncContext;
      event.completed({
        onErrorDeleteItem: false,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.None,
        showPostProcessingDialog: {
          title: "Bedankt voor je melding",
          description: "We hebben de verdachte mail ontvangen en bekijken het zo snel mogelijk.",
        },
      });
    }
  );
}

/**
 * IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name
 * specified in the manifest to its JavaScript counterpart.
 */
Office.actions.associate("onSpamReport", onSpamReport);
