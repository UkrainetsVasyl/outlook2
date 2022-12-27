/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

let webSocket = null;

const Messages = {
  newMessage: 'new_message',
  reply: 'reply_to',
  replyAll: 'reply_to_all',
  forward: 'forward_to',
  tag: 'tag_message',
  untag: 'untag_message',
  createFoulder: 'create_foulder',
  downloadAttachments: 'download_attachments'
}

window.addEventListener('DOMContentLoaded', () => {
  webSocket = new WebSocket('ws://localhost:9000');
})

Office.initialize = function (reason) {
  Office.onReady(function () {
    webSocket = new WebSocket('ws://localhost:9000');

    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      loadNewItem,
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          // Handle error.
        }
      });
  });
};

Office.onReady((info) => {
  webSocket = new WebSocket('ws://localhost:9000');
  webSocket.onmessage = function (message) {
    switch (message.data) {
      case Messages.newMessage: openNewMessage(); break;
      case Messages.reply: replyMessage(); break;
      case Messages.replyAll: replyMessagesAll(); break;
      case Messages.forward: forwardMessage(); break;
      case Messages.tag: tagMessage(); break;
      case Messages.untag: untagMessage(); break;
      case Messages.createFoulder: createFoulder(); break;
      case Messages.downloadAttachments: downloadAttachments(); break;
      default: logMessage(message.data);
    }
  };
});

function loadNewItem(eventArgs) {
  const item = Office.context.mailbox.item;
  console.log(item)
  // Check that item is not null.
  if (item !== null) {
    // Work with item, e.g., define and call function that
    // loads the properties of the newly selected item.
    console.log(item);
  }
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  webSocket = new WebSocket('ws://localhost:9000');
  const message = {
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

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;



function logMessage(message) {
  console.log('Message: %s', message.data);
}

function openNewMessage() {
  Office.context.mailbox.displayNewMessageForm({});
}

function replyMessage() {
  Office.context.mailbox.item.displayReplyForm('hello there');
}

function replyMessagesAll() {
  Office.context.mailbox.item.displayReplyAllForm('hello there');
}

async function forwardMessage() {
  console.log(Office.context.mailbox.item)
  const body = await new Promise((resolve) => {
    Office.context.mailbox.item.body.getAsync("html", { asyncContext: event }, (data) => {
      resolve(data.value);
    })
  });
  console.log(body)
  Office.context.mailbox.item.displayReplyForm({
    htmlBody: '',
    'attachments' :
    [
        {
            'type' : 'item',
            'name' : 'rand',
            'itemId' : Office.context.mailbox.item.itemId
        }
    ]
  });
}

async function tagMessage() {
  const categories = await new Promise((resolve) => {
    Office.context.mailbox.masterCategories.getAsync(null, (data) => {
      resolve(data.value)
    });
  });
  Office.context.mailbox.item.categories.addAsync([categories[0].displayName]);
}

async function untagMessage() {
  const categories = await new Promise((resolve) => {
    Office.context.mailbox.item.categories.getAsync((data) => {
      resolve(data.value);
    });
  });
  console.log(categories);

  Office.context.mailbox.item.categories.removeAsync([categories[0].displayName]);
}

async function createFoulder() {
  // Office.context.mailbox.makeEwsRequestAsync(getSubjectRequest(Office.context.mailbox.item.itemId), (data) => {
  //   console.log(new DOMParser().parseFromString(data.value, 'text/xml'));
  //   // console.log(JSON.parse(data.value));
  // });
  // console.log(foulderxml);
  Office.context.mailbox.makeEwsRequestAsync(getRootFoulder(), (data) => {
    console.log(data.value)
    // console.log(new DOMParser().parseFromString(data.value, 'text/xml'));
  });
}

// async function createFoulder() {
//   var item = Office.context.mailbox.item;
//   //@ts-ignore
//   easyEws.getMailItem(item.id, (ids) => {
//     console.log(ids); 
//   });
// }

async function downloadAttachments() {
  const item = Office.context.mailbox.item;
  // console.log(item.getSelectedEntities())
  // console.log(item.attachments[0]);
  console.log(Office.context.mailbox.item.body)
  Office.context.mailbox.item.body.getTypeAsync(data => {
    console.log(data)
  });
  console.log(Office.context.document.getSelectedResourceAsync((result) => {
    console.log(result)
  }));
  // Office.context.mailbox.item.close()
  // item.getAttachmentContentAsync(item.attachments[0].id, {}, res => {
  //   console.log(res);
  // });
}

function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}

function UpdateTaskPaneUI(item)
{
  console.log(item)
  // Assuming that item is always a read item (instead of a compose item).
  if (item != null) console.log(item.subject);
}