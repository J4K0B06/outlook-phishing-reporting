reader.onload = function () {
  const base64Eml = reader.result.split(',')[1];

  fetch("https://webhook.site/1c18c494-96fb-4bd2-a480-2f0f34bebd6c", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      eml: base64Eml,
      options: reportedOptions,
      comment: additionalInfo,
    }),
  })
    .then((res) => {
      console.log("Verzonden naar backend", res.status);
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
      console.error("Fout bij verzenden naar backend:", err);
      // Event alsnog afronden, maar met een foutmelding misschien
      event.completed({
        onErrorDeleteItem: false,
        moveItemTo: Office.MailboxEnums.MoveSpamItemTo.NoMove,
        showPostProcessingDialog: {
          title: "Er ging iets mis",
          description: "De melding kon niet worden verstuurd. Probeer het later opnieuw.",
        },
      });
    });
};
