/* commands.js — OnMessageSend event handler */

let globalSendEvent;

Office.onReady(() => {
  // Office is ready — nothing to initialize here
});

function onMessageSendHandler(event) {
  // Store the event so we can call .completed() after dialog closes
  globalSendEvent = event;

  const dialogUrl =
    "https://sadikemreikiz.github.io/outlook-smart-alerts/dialog.html";

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 45, width: 35, displayInIframe: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Could not open dialog — allow send so user is not blocked
        console.error("Dialog failed to open:", asyncResult.error.message);
        globalSendEvent.completed({ allowEvent: true });
        return;
      }

      const dialog = asyncResult.value;

      // Listen for message from dialog.js → messageParent()
      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        function (args) {
          dialog.close();

          if (args.message === "confirm") {
            // User clicked Send → allow the email to go
            globalSendEvent.completed({ allowEvent: true });
          } else {
            // User clicked Cancel → block the email
            globalSendEvent.completed({
              allowEvent: false,
              errorMessage: "Email sending was cancelled. Please review before sending.",
            });
          }
        }
      );

      // User closed the dialog with the X button — treat as cancel
      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        function () {
          globalSendEvent.completed({
            allowEvent: false,
            errorMessage: "Email sending was cancelled.",
          });
        }
      );
    }
  );
}

// Associate the function name used in manifest LaunchEvent
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
