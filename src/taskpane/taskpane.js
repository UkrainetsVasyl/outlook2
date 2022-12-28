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
  downloadImages: "download_images_attachments",
  downloadFiles: "download_file_attachments",
  move: "move",
  setBCC: "set_bcc",
  setTo: "set_to",
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
  console.log(Office.context)

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
      case Messages.downloadImages:
        downloadImages();
        break;
      case Messages.downloadFiles:
        downloadFiles();
        break;
      case Messages.move:
        move();
        break;
      case Messages.setBCC:
        setBCC();
        break;
      case Messages.setTo:
        setTo();
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
  Office.context.mailbox.item.displayReplyForm({});
}

function replyMessagesAll() {
  Office.context.mailbox.item.displayReplyAllForm({});
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
    "attachments":
      [
        {
          "type": "item",
          "name": "rand",
          "itemId": Office.context.mailbox.item.itemId
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
  Office.context.mailbox.makeEwsRequestAsync(getRootFoulder(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
  Office.context.mailbox.makeEwsRequestAsync(getSubjectRequest(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
  Office.context.mailbox.makeEwsRequestAsync(generateFolder(), (data) => {
    console.log(new DOMParser().parseFromString(data.value, "text/xml"));
  });
}

async function downloadAttachments() {
  const item = Office.context.mailbox.item;
  item.attachments?.forEach(attachment => {
    item.getAttachmentContentAsync(attachment.id, {}, (res) => handleAttachmentsCallback(attachment, res));
  });
}

async function downloadImages() {
  const item = Office.context.mailbox.item;
  item.attachments?.filter(({contentType}) => contentType.includes('image'))?.forEach(attachment => {
    item.getAttachmentContentAsync(attachment.id, {}, (res) => handleAttachmentsCallback(attachment, res));
  });
}

async function downloadFiles() {
  const item = Office.context.mailbox.item;
  item.attachments?.filter(({contentType}) => contentType.includes('pdf'))?.forEach(attachment => {
    item.getAttachmentContentAsync(attachment.id, {}, (res) => handleAttachmentsCallback(attachment, res));
  });
}

function move() {
  const item = Office.context.mailbox.item;
  const archiveId = "AAMkADk1M2E0OThlLWZjZmYtNGExNy1hOTM1LTk3MmZmNTc4NzAxMgAuAAAAAACv7rTjE3bnTKFd7SYkXW3fAQBbzAuKJnTqQKuv5GK7bVH1AAACNERmAAA=";
  const folderName = "Архів";
  item.move(archiveId);
}

function setBCC() {
  Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
}

function setTo() {
  Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
}

function handleFileAttachment(data, type, name) {
  // console.log(data);
  downloadBase64File(data, type, name);
}

function downloadBase64File(base64Data, contentType, fileName) {
  const linkSource = `data:${contentType};base64,${base64Data}`;
  const downloadLink = document.createElement("a");
  downloadLink.href = linkSource;
  downloadLink.download = fileName;
  downloadLink.click();
}

function handleAttachmentsCallback(fileDescription, result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      handleFileAttachment(result.value.content, fileDescription.contentType, fileDescription.name);
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
    // Handle attachment formats that are not supported.
  }
}