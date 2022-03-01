/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { log, logObject } from "../helpers/debug";
import { getGraphAccessToken } from "../helpers/ssoauthhelper";
import { fetchDataAndInsertSignature } from "../shared/signature";

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function onMessageComposeHandler(event) {
  try {
    const accessToken = await getGraphAccessToken();
    await fetchDataAndInsertSignature(accessToken);
  } catch (error) {
    log("error");
    logObject(error);
  }
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action;
g.onMessageComposeHandler;

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
