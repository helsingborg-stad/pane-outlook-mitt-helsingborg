/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import axios from "axios";
import { getAccessToken } from "./../helpers/ssoauthhelper";

declare const API_URL: string;

interface BookingInformation {
  subject: string;
  startTime: Date;
  endTime: Date;
  organizationRequiredAttendees: string[];
  externalRequiredAttendees: string[];
  remoteMeeting: boolean;
  id: string;
  formData: {
    firstname: {
      name: string;
      value: string;
    };
    lastname: {
      name: string;
      value: string;
    };
    phone: {
      name: string;
      value: string;
    };
    email: {
      name: string;
      value: string;
    };
    remoteMeeting: {
      name: string;
      value: boolean;
    };
    comment: {
      name: string;
      value: string;
    };
  };
}

/* global $, document, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    $(".booking-information").hide();
    const id = await getBookingId();
    $(".status-text").text("⌛️ Laddar mötesinformation...");
    getBookingInformation(id)
      .then((response) => {
        $(".status-text").text("✅ Mötesinformationen har hämtats");
        renderBookingInformation(response);
        $(".booking-information").show();
      })
      .catch(() => {
        $(".status-text").text("🛑 Kunde inte hämta mötesinformation");
      });
  }

  $("#cancel").click(() => {
    cancelMeeting()
      .then(() => {
        $(".status-text").text("✅ Mötet avbokat");
        $(".booking-information").hide();
      })
      .catch(() => {
        $(".status-text").text("🛑 Kunde inte avboka mötet");
      });
  });
});

function renderBookingInformation(response: BookingInformation): void {
  $("#firstname").text(response.formData.firstname.value);
  $("#lastname").text(response.formData.lastname.value);
  $("#email").text(response.formData.email.value);
  $("#phone").text(response.formData.phone.value);
  $("#info").text(response.formData.comment.value);
}

function cancelMeeting(): Promise<void> {
  return Promise.resolve();
}

function getBookingId(): Promise<string> {
  return new Promise((res) => {
    res(Office.context.mailbox.item.subject.split("##").pop());
  });
}

function getBookingInformation(id: string): Promise<BookingInformation> {
  return getAccessToken().then((token) => {
    return axios
      .get(`${API_URL}/booking/${id}`, {
        headers: {
          Authorization: `Bearer ${token}`,
        },
      })
      .then((response) => response.data.data);
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
