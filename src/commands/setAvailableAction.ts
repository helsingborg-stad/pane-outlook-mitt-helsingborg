/* global Office */

const BOOKING = {
  SUBJECT: "MH bokningsbar tid",
};

const setAvailableAction = (event: Office.AddinCommands.Event) => {
  Office.context.mailbox.item.subject.setAsync(BOOKING.SUBJECT);

  const htmlContent = `
    <div style={{ display: 'flex', flexDirection: 'column'}}>
      <h2>Mitt Helsingborg</h2>
      <p>Valfria inställningar för användare:</p>
      <ol>
        <li>Visa mötesbokningen som "Free"</li>
        <li>Stäng av notifieringar</li>
      </ol>
      <p>Du kan nu spara den här mötesbokningen!</p>
    </div>
  `;

  Office.context.mailbox.item.body.setAsync(htmlContent, {
    coercionType: Office.CoercionType.Html,
  });

  event.completed();
};

export default setAvailableAction;
