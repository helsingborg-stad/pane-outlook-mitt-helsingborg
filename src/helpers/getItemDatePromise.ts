/* global Office, console */

const getItemDatePromise = async (date: "start" | "end"): Promise<Date> => {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item[date].getAsync((startResult) => {
      if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${startResult.error.message}`);
        reject();
      }

      const startDate = startResult.value;
      resolve(startDate);
    });
  });
};

export default getItemDatePromise;
