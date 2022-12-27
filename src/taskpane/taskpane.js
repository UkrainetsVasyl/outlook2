/* eslint-disable prettier/prettier */
/* eslint-disable no-undef */
/* eslint-disable office-addins/no-office-initialize */
import getSubject from "../exchangeXML/getSubject.xml";
import getRoot from "../exchangeXML/getRoot.xml";
import createFolder from "../exchangeXML/generateFolder.xml";

let webSocket = null;

function getSubjectRequest() {
  return getSubject;
}

function getRootFoulder() {
  return getRoot;
}

function generateFolder() {
  return createFolder;
}

const Messages = {
  newMessage: "new_message",
  reply: "reply_to",
  replyAll: "reply_to_all",
  forward: "forward_to",
  tag: "tag_message",
  untag: "untag_message",
  createFoulder: "create_foulder",
  downloadAttachments: "download_attachments",
};

Office.initialize = () => {
  Office.onReady(function () {
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

function loadNewItem() {
  const item = Office.context.mailbox.item;
  if (item !== null) {
      console.log(item);
  }
}

Office.onReady((info) => {
  webSocket = new WebSocket("ws://localhost:9000");

  run();

  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = displayData;
  }
});

export async function run() {
  webSocket.onopen = function () {
    console.log("connected");
  };

  webSocket.onmessage = function (message) {
    switch (message.data) {
      case Messages.newMessage:
        openNewMessage();
        break;
      case Messages.reply:
        replyMessage();
        break;
      case Messages.replyAll:
        replyMessagesAll();
        break;
      case Messages.forward:
        forwardMessage();
        break;
      case Messages.tag:
        tagMessage();
        break;
      case Messages.untag:
        untagMessage();
        break;
      case Messages.createFoulder:
        makeEWS();
        break;
      case Messages.downloadAttachments:
        downloadAttachments();
        break;
      default:
        logMessage(message.data);
    }
  };
}

function logMessage(message) {
  console.log("Message: %s", message.data);
}

function openNewMessage() {
  Office.context.mailbox.displayNewMessageForm({});
}

function replyMessage() {
  Office.context.mailbox.item.displayReplyForm("hello there");
}

function replyMessagesAll() {
  Office.context.mailbox.item.displayReplyAllForm("hello there");
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
    htmlBody: "",
    "attachments" :
    [
        {
            "type" : "item",
            "name" : "rand",
            "itemId" : Office.context.mailbox.item.itemId
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

async function makeEWS() {
  Office.context.mailbox.makeEwsRequestAsync(getSubjectRequest(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
  Office.context.mailbox.makeEwsRequestAsync(getRootFoulder(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
  Office.context.mailbox.makeEwsRequestAsync(generateFolder(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
}

async function downloadAttachments() {
  item.getAttachmentContentAsync(item.attachments[0].id, {}, res => {
    console.log(res);
  });
}