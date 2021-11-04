/* global Office */

import { getAccessToken } from "../helpers/ssoauthhelper";

const bookMeetingAction = (event: Office.AddinCommands.Event) => {
  getAccessToken().then((token) => {
    Office.context.mailbox.item.body.setAsync(JSON.stringify(token));
    event.completed();
  });
};

export default bookMeetingAction;
