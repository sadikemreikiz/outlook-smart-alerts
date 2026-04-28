/* dialog.js — form logic, sends result back to commands.js */

Office.onReady(() => {
  document.getElementById("sendBtn").addEventListener("click", function () {
    Office.context.ui.messageParent("confirm");
  });

  document.getElementById("cancelBtn").addEventListener("click", function () {
    Office.context.ui.messageParent("cancel");
  });
});
