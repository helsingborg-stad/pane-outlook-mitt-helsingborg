// import dayjs from "dayjs";

/* global Office, console */

// import getItemDatePromise from "../helpers/getItemDatePromise";

const setAvailableAction = async (event: Office.AddinCommands.Event) => {
  try {
    Office.context.mailbox.item.subject.setAsync("MH bokningsbar tid");

    // let startDate = await getItemDatePromise("start");
    // const formattedStartDate = dayjs(startDate).format();

    // const endDate = await getItemDatePromise("end");
    // const formattedSEndDate = dayjs(endDate).format();

    // const message: Office.NotificationMessageDetails = {
    //   type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    //   message: `${formattedStartDate} - ${formattedSEndDate}`,
    //   icon: "Icon.80x80",
    //   persistent: true,
    // };

    Office.context.mailbox.item.close();

    // Show a notification message
    // Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  } catch (error) {
    console.error("Error creating availability slot: ", error);
    Office.context.mailbox.item.subject.setAsync(error);
  }

  event.completed();
};

export default setAvailableAction;
