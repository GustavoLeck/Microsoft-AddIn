/* global Office */

Office.onReady(() => {
  // Office.js is ready to be called, if needed.
});

/**
 * "Perform an action" ribbon button: shows a notification on the email being written.
 * @param {Office.AddinCommands.Event} event
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message,
  );

  // Always tell Office the command has finished.
  event.completed();
}

// Register the function with Office (name must match <FunctionName> in manifest.xml).
Office.actions.associate("action", action);
