let webSocket = null;


function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item. 
  const result =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"' +
    '      xmlns:xsd="https://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '  <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '      <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return result;
}

function getRootFoulder() {
  const result = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
   xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <GetFolder xmlns="https://schemas.microsoft.com/exchange/services/2006/messages"
               xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </FolderShape>
      <FolderIds>
        <t:DistinguishedFolderId Id="inbox"/>
      </FolderIds>
    </GetFolder>
  </soap:Body>
</soap:Envelope>`;

  return result;
}

function generateFoulder(id) {
  return `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
          xmlns:t="https://schemas.microsoft.com/exchange/services/2006/types">
      <soap:Body>
        <CreateFolder xmlns="https://schemas.microsoft.com/exchange/services/2006/messages">
          <ParentFolderId>
            <t:DistinguishedFolderId Id="inbox"/>
          </ParentFolderId>
          <Folders>
            <t:Folder>
              <t:DisplayName>Folder1</t:DisplayName>
            </t:Folder>
            <t:Folder>
              <t:DisplayName>Folder2</t:DisplayName>
            </t:Folder>
          </Folders>
        </CreateFolder>
      </soap:Body>
    </soap:Envelope>
  `;
}

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

Office.initialize = function (reason) {
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

Office.onReady((info) => {
  webSocket = new WebSocket('ws://localhost:9000');

  run();

  if (info.host === Office.HostType.Outlook) {
    document.getElementById('sideload-msg').style.display = 'none';
    document.getElementById('app-body').style.display = 'flex';
    // document.getElementById('run').onclick = displayData;
  }
});

export async function run() {
  webSocket.onopen = function () {
    console.log('connected');
  };

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
}

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