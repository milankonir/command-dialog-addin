function openDialog(event: Office.AddinCommands.Event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog.html",
    { displayInIframe: false, promptBeforeOpen: false, width: 30, height: 30 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        result.value.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
          if ("error" in args) {
            if (args.error === 12006) {
              event.completed();
            }
          }
        });
      } else {
        console.log(`Failed to open dialog ${result}`);
        event.completed();
      }
    }
  );
}

Object.assign(globalThis, {
  openDialog,
});

Office.onReady();
