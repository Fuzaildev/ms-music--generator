/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  // Register the function with Office for all supported hosts
  if (info.host === Office.HostType.Excel || 
      info.host === Office.HostType.Word || 
      info.host === Office.HostType.PowerPoint) {
    // Register the function with Office.
    Office.actions.associate("showTaskpane", showTaskpane);
  }
});

/**
 * Shows the taskpane when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function showTaskpane(event) {
  // Show the taskpane
  Office.context.document.settings.set("taskpaneVisible", true);
  Office.context.document.settings.saveAsync();
  
  // Be sure to indicate when the function is complete
  event.completed();
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
