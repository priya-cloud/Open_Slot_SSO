/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { getGraphData } from "./../helpers/ssoauthhelper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
  }
});

export async function run() {
  getGraphData(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  // eslint-disable-next-line no-undef
  console.log("inside writeDataToOfficeDocument");
  let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  userInfo =
    "Start: " +
    result.value[0].scheduleItems[0].start.dateTime +
    " End: " +
    result.value[0].scheduleItems[0].end.dateTime;
  /*for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }*/

  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
