/* global Office */

const bookMeetingAction = (event: Office.AddinCommands.Event) => {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Saying cheese, update version :D",
    icon: "Icon.80x80",
    persistent: true,
  };
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // Be sure to indicate when the add-in command function is complete
  event.completed();
};

export default bookMeetingAction;
