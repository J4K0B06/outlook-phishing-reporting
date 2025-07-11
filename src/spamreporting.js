Office.onReady();

function onSpamReport(event) {
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(`Error encountered during message processing: ${asyncResult.error.message}`);
        return;
      }

      const spamReportingEvent = asyncResult.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      const base64Eml = asyncResult.value; // Already base64

      // Send to backend
      /*
      fetch("https://webhook.site/1c18c494-96fb-4bd2-a480-2f0f34bebd6c", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          eml: base64Eml,
          options: reportedOptions,
          comment: additionalInfo,
        }),
      })
        .then((res) => {
          console.log("Verzonden naar backend", res.status);
        })
        .catch((err) => {
          console.error("Fout bij verzenden naar backend:", err);
        });
      */
      
      fetch("https://webhook.site/1c18c494-96fb-4bd2-a480-2f0f34bebd6c", {
        method: "POST",
        body: JSON.stringify({ test: "hello" }),
      });

      // Show post-processing dialog
      event.completed({
        onErrorDeleteItem: false,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
        showPostProcessingDialog: {
          title: "Bedankt voor je melding",
          description: "We hebben de verdachte mail ontvangen en bekijken het zo snel mogelijk.",
        },
      });
    }
  );
}

Office.actions.associate("onSpamReport", onSpamReport);
