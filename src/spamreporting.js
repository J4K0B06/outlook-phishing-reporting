Office.onReady();

function onSpamReport(event) {
  Office.context.mailbox.item.getAsFileAsync(
    { asyncContext: event },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Fout bij ophalen bestand:", result.error.message);
        return;
      }

      const file = result.value;
      console.log("file object:", file);
      
      const spamReportingEvent = result.asyncContext;
      const reportedOptions = spamReportingEvent.options;
      const additionalInfo = spamReportingEvent.freeText;

      // Sommige omgevingen geven meteen slice data mee
      if (file && file.slice && typeof file.slice === "function") {
        file.slice(0, (sliceResult) => {
          if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
            const base64Eml = sliceResult.value.data;

            sendToBackend(base64Eml, reportedOptions, additionalInfo, event);
            file.close();
          } else {
            console.error("Fout bij slice ophalen:", sliceResult.error.message);
            file.close();
          }
        });
      } else {
        console.error("File object ondersteunt geen slicing.");
        event.completed({
          onErrorDeleteItem: false,
          moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
          showPostProcessingDialog: {
            title: "Er ging iets mis",
            description: "Deze versie van Outlook ondersteunt deze functie niet correct.",
          },
        });
      }
    }
  );
}

function sendToBackend(eml, options, comment, event) {
  fetch("https://webhook.site/1c18c494-96fb-4bd2-a480-2f0f34bebd6c", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ eml, options, comment }),
  })
    .then((res) => {
      console.log("Verzonden naar backend", res.status);
    })
    .catch((err) => {
      console.error("Fout bij verzenden naar backend:", err);
    })
    .finally(() => {
      event.completed({
        onErrorDeleteItem: false,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
        showPostProcessingDialog: {
          title: "Bedankt voor je melding",
          description: "We hebben de verdachte mail ontvangen en bekijken het zo snel mogelijk.",
        },
      });
    });
}

Office.actions.associate("onSpamReport", onSpamReport);
