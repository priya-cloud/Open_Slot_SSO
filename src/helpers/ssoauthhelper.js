/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { dialogFallback } from "./fallbackauthhelper";
import * as sso from "office-addin-sso";

/* global OfficeRuntime */

let retryGetAccessToken = 0;

export async function getGraphData(callback) {
  try {
    let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
    let response = await sso.makeGraphApiCall(bootstrapToken);
    /* const endpoint = "me/messages";
    const urlparams = "?top=10";

    // let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: response.claims });
    // eslint-disable-next-line prettier/prettier
    let response = sso.getGraphData("eyJ0eXAiOiJKV1QiLCJub25jZSI6IlJnSThhNWwzdHItTWgzWGEwY2gzVFZLTXg3Z1Y5VDhPRE5fN0YzNU5QWnMiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8xYWE4NGQ0ZC1kYTQ4LTQ0Y2QtYmJlMy00YTgxZTkwMDYxNmYvIiwiaWF0IjoxNjY1MjI1MTg5LCJuYmYiOjE2NjUyMjUxODksImV4cCI6MTY2NTIyOTA4OSwiYWlvIjoiRTJaZ1lIQ2M1MXJ4bnJ2ZmlMczZnYVByWGtBUEFBPT0iLCJhcHBfZGlzcGxheW5hbWUiOiJPcGVuX1Nsb3RfU1NPIiwiYXBwaWQiOiIzMTk4Yjg1Zi1jYWVkLTQwZDItYWYwZC0zMWE3NWNlMmM2MDgiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8xYWE4NGQ0ZC1kYTQ4LTQ0Y2QtYmJlMy00YTgxZTkwMDYxNmYvIiwiaWR0eXAiOiJhcHAiLCJvaWQiOiJhY2VjNzBmYS04YjEwLTQ1NDgtYTcyYS1lMDQ5Y2VmNWM3NzQiLCJyaCI6IjAuQVhBQVRVMm9Ha2phelVTNzQwcUI2UUJoYndNQUFBQUFBQUFBd0FBQUFBQUFBQUJ3QUFBLiIsInN1YiI6ImFjZWM3MGZhLThiMTAtNDU0OC1hNzJhLWUwNDljZWY1Yzc3NCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJBUyIsInRpZCI6IjFhYTg0ZDRkLWRhNDgtNDRjZC1iYmUzLTRhODFlOTAwNjE2ZiIsInV0aSI6IjRaTlZGSE5VUTAydlRUUkRIUW1tQUEiLCJ2ZXIiOiIxLjAiLCJ3aWRzIjpbIjA5OTdhMWQwLTBkMWQtNGFjYi1iNDA4LWQ1Y2E3MzEyMWU5MCJdLCJ4bXNfdGNkdCI6MTYyMjE5NDEzM30.hnnaVwAACNRsYf5z0x5vlZkGv2hQ0QG-2vKilarBS1cE1gWQXMxQyPrOAr3vwe4-TCasc1T-ThAJs0wF268RRHIVmTjpKHws2TsmEdqgAjyO_ySHtSjkG5BCav4GiN1lW2PPVzPRzEg2fKKJS_PMLvnKPW7C8lHuAdwn4gSDEPnZ6MDcByJRdCLcKfDAyWcwn6zTEHLR3NpqEKd_pUk_A_uA08xASb0qCXxiuFFuVfCpAJhs-U-R8XEXpLexWWh9j_KSzaq6IhfgmTdAou_ERPzZLFCCcxw-824hWmuuCJP7pWk5dAyWpQ1Mv5KbZBdz5-kPWrEq5kjjend2DpaJ5Q", endpoint, urlparams);
    */
    if (response.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.s
      let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: response.claims });
      response = sso.makeGraphApiCall(mfaBootstrapToken);
    }

    if (response.error) {
      // AAD errors are returned to the client with HTTP code 200, so they do not trigger
      // the catch block below.
      handleAADErrors(response);
    } else {
      // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
      // in the .fail callback of that call
      // eslint-disable-next-line no-undef
      //console.log(response.value[0].subject);
      // eslint-disable-next-line no-undef
      console.log(response);
      callback(response);
      Promise.resolve();
    }
  } catch (exception) {
    if (exception.code) {
      if (sso.handleClientSideErrors(exception)) {
        dialogFallback(callback);
      }
    } else {
      sso.showMessage("EXCEPTION: " + JSON.stringify(exception));
      Promise.reject();
    }
  }
}

function handleAADErrors(response, callback) {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.
  if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetAccessToken <= 0) {
    retryGetAccessToken++;
    getGraphData(callback);
  } else {
    dialogFallback(callback);
  }
}
