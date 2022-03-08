/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global OfficeRuntime */
import * as sso from "office-addin-sso";
import { log } from "./debug";
import { dialogFallback } from "./fallbackauthhelper";
import { callApi } from "./xhr";
let retryGetAccessToken = 0;

export async function getGraphAccessToken(): Promise<string> {
  try {
    /*
    [Step 2] - Get SSO token
    ========================
    
    How to enable SSO in an event-activated add-in
    ----------------------------------------------
    
    - Option 1: Serve a JSON with a list of allowed files at the same domain as the add-in
      - https://some.host.com/commands.js
      - https://some.host.com/.well-known/microsoft-officeaddins-allowed.json

    - Option 2: Serve the origin as an HTTP Header
      - MS-OfficeAddins-Allowed-Origin : https://some.host.com
  
    [Docu](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/use-sso-in-event-based-activation)
    */
    log("getAccessToken");

    let bootstrapToken: string = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true });

    /*
    [Step 3] - Get MS Graph token
    =============================
    
    - Token exchange is based on OAuth2 on-behalf-of flow
    - Requires a web server that knows the client secret for the add-in client id (which is created during SSO setup)
    - This example uses the office-addin-sso NPM package to host a **local** node.js server that handles token exchange

    [office-addin-sso authentication backend implementation](https://github.com/OfficeDev/Office-Addin-Scripts/blob/master/packages/office-addin-sso/src/authRoute.ts)
    */
    log("getGraphToken");
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
  const jsonResponse = await callApi(`https://localhost:3000/auth`, {
    method: "GET",
    headers: { Authorization: `Bearer ${bootstrapToken}` },
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
