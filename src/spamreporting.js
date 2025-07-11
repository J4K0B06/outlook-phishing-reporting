Office.onReady();

function onSpamReport(event) {
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Fout bij ophalen bestand:", result.error.message);
        return;
      }

      const file = result.value; // Office.File
      const spamReportingEvent = result.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      file.getSliceAsync(0, (sliceResult) => {
        if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
          const slice = sliceResult.value;
          const base64Eml = slice.data; // Dit is de base64-encoded inhoud

          // Versturen naar backend
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
            file.close(); // Sluit het bestand
            event.completed({
              onErrorDeleteItem: false,
              moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
              showPostProcessingDialog: {
                title: "Bedankt voor je melding",
                description: "We hebben de verdachte mail ontvangen en bekijken het zo snel mogelijk.",
              },
            });
          })
          .catch((err) => {
            console.error("Fout bij versturen naar backend:", err);
            file.close();
            event.completed({
              onErrorDeleteItem: false,
              moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
              showPostProcessingDialog: {
                title: "Er ging iets mis",
                description: "De melding kon niet worden verstuurd. Probeer het later opnieuw.",
              },
            });
          });
        } else {
          console.error("Fout bij lezen van bestandsslice:", sliceResult.error.message);
          file.close();
        }
      });
    }
  );
}

Office.actions.associate("onSpamReport", onSpamReport);
