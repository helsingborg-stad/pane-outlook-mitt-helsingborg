/* global Office */

const setAvailableAction = (event: Office.AddinCommands.Event) => {
  Office.context.mailbox.item.subject.setAsync("Bokningsbar tid Mitt Helsingborg");
  // Be sure to indicate when the add-in command function is complete
  event.completed();
};

export default setAvailableAction;
