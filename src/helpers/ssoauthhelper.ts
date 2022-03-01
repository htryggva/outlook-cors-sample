/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global OfficeRuntime */
import * as sso from "office-addin-sso";
import { dialogFallback } from "./fallbackauthhelper";
import { callApi } from "./xhr";
let retryGetAccessToken = 0;

export async function getGraphAccessToken(): Promise<string> {
  try {
    let bootstrapToken: string = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });
    let exchangeResponse: any = await getGraphToken(bootstrapToken);
    if (exchangeResponse.claims) {
      // Microsoft Graph requires an additional form of authentication. Have the Office host
      // get a new token using the Claims string, which tells AAD to prompt the user for all
      // required forms of authentication.
      let mfaBootstrapToken: string = await OfficeRuntime.auth.getAccessToken({
        authChallenge: exchangeResponse.claims,
      });
      exchangeResponse = await getGraphToken(mfaBootstrapToken);
    }

    if (exchangeResponse.error) {
      // AAD errors are returned to the client with HTTP code 200, so they do not trigger
      // the catch block below.
      handleAADErrors(exchangeResponse);
    } else {
      // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
      // in the .fail callback of that call
      return exchangeResponse.access_token;
    }
  } catch (exception) {
    // if handleClientSideErrors returns true then we will try to authenticate via the fallback
    // dialog rather than simply throw and error
    if (exception.code) {
      if (sso.handleClientSideErrors(exception)) {
        dialogFallback();
      }
    } else {
      throw exception;
    }
  }
}

async function getGraphToken(bootstrapToken: string) {
  const jsonResponse = await callApi(`https://INSERT_TOKEN_EXCHANGE_SERVER/api/auth?token=${bootstrapToken}`, {
    method: "GET",
  });

  return JSON.parse(jsonResponse);
}

function handleAADErrors(exchangeResponse: any): void {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.

  if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1 && retryGetAccessToken <= 0) {
    retryGetAccessToken++;
    getGraphAccessToken();
  } else {
    dialogFallback();
  }
}
