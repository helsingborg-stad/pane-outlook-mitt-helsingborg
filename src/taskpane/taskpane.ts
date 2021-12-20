import { BACKEND_BASE_URL } from "./../helpers/constants";
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import axios from "axios";
import { getAccessToken } from "./../helpers/ssoauthhelper";

/* global $, document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    $(document).ready(function () {
      $("#submitButton").click(() => queryStuff($("#query").val().toString().trim()));
    });
  }
});

function queryStuff(query): void {
  getAccessToken().then((token) => {
    axios
      .post(
        BACKEND_BASE_URL + "/users/fetchReferenceCode",
        { query },
        {
          headers: {
            Authorization: `Bearer ${token}`,
          },
        }
      )
      .then((res) => {
        Office.context.mailbox.item.subject.getAsync((titleResult) => {
          Office.context.mailbox.item.subject.setAsync(`${titleResult.value} ##${res.data.data.referenceCode}`, () => {
            $("#errorText").text("");
            $("#successText").text("Referenskoden för invånaren har hämtats.");
          });
        });
      })
      .catch(() => {
        $("#errorText").text("Den angivna personnummret/e-mailadressen finns inte i systemet.");
        $("#successText").text("");
      });
  });
}

export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];
  let userProfileInfo: string[] = [];
  userProfileInfo.push(result["displayName"]);
  userProfileInfo.push(result["mail"]);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
