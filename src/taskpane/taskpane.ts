/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global $, document, Office */
import * as sso from "office-addin-sso";
import { getGraphAccessToken } from "../helpers/ssoauthhelper";
import { fetchDataAndInsertSignature } from "../shared/signature";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    $(document).ready(function () {
      $("#getGraphDataButton").click(getGraphDataHandler);
    });
  }
});

async function getGraphDataHandler() {
  try {
    const accessToken = await getGraphAccessToken();
    await fetchDataAndInsertSignature(accessToken);
    sso.showMessage("Your data has been added to the document.");
  } catch (exception) {
    sso.showMessage("EXCEPTION: " + JSON.stringify(exception));
  }
}
