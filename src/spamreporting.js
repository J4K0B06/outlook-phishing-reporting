Office.onReady();

function onSpamReport(event) {
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(
          `Error encountered during message processing: ${asyncResult.error.message}`
        );
        return;
      }

      const spamReportingEvent = asyncResult.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      const file = asyncResult.value;
      file.getSliceAsync(0, (sliceResult) => {
        if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
          const base64Eml = sliceResult.value.data;

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
              spamReportingEvent.completed({
                onErrorDeleteItem: false,
                moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
                showPostProcessingDialog: {
                  title: "Bedankt voor je melding",
                  description:
                    "We hebben de verdachte mail ontvangen en bekijken het zo snel mogelijk.",
                },
              });
            })
            .catch((err) => {
              console.error("Fout bij verzenden naar backend:", err);
              spamReportingEvent.completed({
                onErrorDeleteItem: false,
                moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
                showPostProcessingDialog: {
                  title: "Fout bij verzending",
                  description:
                    "Er ging iets mis bij het verzenden. Probeer het later opnieuw.",
                },
              });
            });
        } else {
          console.log("Error in getSliceAsync:", sliceResult.error.message);
          spamReportingEvent.completed({
            onErrorDeleteItem: false,
            moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
            showPostProcessingDialog: {
              title: "Fout bij ophalen mail",
              description: "De mail kon niet correct opgehaald worden.",
            },
          });
        }
      });
    }
  );
}

Office.actions.associate("onSpamReport", onSpamReport);
