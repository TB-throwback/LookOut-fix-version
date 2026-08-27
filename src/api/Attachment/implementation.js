/*
 * Author: John Bieling (john@thunderbird.net)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

var { ExtensionCommon } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionCommon.sys.mjs"
);
var { ExtensionUtils } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionUtils.sys.mjs"
);
var { ExtensionError } = ExtensionUtils;

ChromeUtils.defineESModuleGetters(this, {
  AttachmentInfo: "resource:///modules/AttachmentInfo.sys.mjs",

  cal: "resource:///modules/calendar/calUtils.sys.mjs",

  invitation:
    "resource:///modules/calendar/utils/calInvitationUtils.sys.mjs",
});

Cu.importGlobalProperties([
  "File",
  "IOUtils",
  "PathUtils",
  "TextDecoder",
]);

async function getRealFileForFile(file) {
  let pathTempFile = await IOUtils.createUniqueFile(
    PathUtils.tempDir,
    file.name.replaceAll(/[/:*?\"<>|]/g, "_"),
    0o600
  );

  let tempFile = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
  tempFile.initWithPath(pathTempFile);
  let extAppLauncher = Cc[
    "@mozilla.org/uriloader/external-helper-app-service;1"
  ].getService(Ci.nsPIExternalAppLauncher);
  extAppLauncher.deleteTemporaryFileOnExit(tempFile);

  let buffer = await file.arrayBuffer();
  await IOUtils.write(pathTempFile, new Uint8Array(buffer));
  return tempFile;
}

function ClearAttachmentList(window) {
  // clear selection
  var list = window.document.getElementById("attachmentList");
  list.clearSelection();

  while (list.hasChildNodes()) {
    list.lastChild.remove();
  }
}

var Attachment = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {

    function getMessageWindow(tabId) {
      // Get about:message from the tabId.
      let { nativeTab } = context.extension.tabManager.get(tabId);
      if (nativeTab instanceof Ci.nsIDOMWindow) {
        return nativeTab.messageBrowser.contentWindow
      } else if (nativeTab.mode && nativeTab.mode.name == "mail3PaneTab") {
        return nativeTab.chromeBrowser.contentWindow.messageBrowser.contentWindow
      } else if (nativeTab.mode && nativeTab.mode.name == "mailMessageTab") {
        return nativeTab.chromeBrowser.contentWindow;
      }
      return null;
    }

    async function getAttachmentFromUrl(url) {
      let channel = Services.io.newChannelFromURI(
        Services.io.newURI(url),
        null,
        Services.scriptSecurityManager.getSystemPrincipal(),
        null,
        Ci.nsILoadInfo.SEC_ALLOW_CROSS_ORIGIN_SEC_CONTEXT_IS_NULL,
        Ci.nsIContentPolicy.TYPE_OTHER
      );

      let raw = await new Promise((resolve, reject) => {
        let listener = Cc["@mozilla.org/network/stream-loader;1"].createInstance(
          Ci.nsIStreamLoader
        );
        listener.init({
          onStreamComplete(loader, context, status, resultLength, result) {
            if (Components.isSuccessCode(status)) {
              resolve(Uint8Array.from(result));
            } else {
              reject(
                new ExtensionError(
                  `Failed to read attachment content: ${status}`
                )
              );
            }
          },
        });
        channel.asyncOpen(listener, null);
      });

      return raw;
    }

    /**
     * Returns the currently displayed message in the given tab.
     *
     * @param {integer} tabId
     * @returns {nsIMsgHdr} nsIMsgHdr
     */
    function getDisplayedMessage(tabId) {
      let { nativeTab } = context.extension.tabManager.get(tabId);
      if (nativeTab instanceof Ci.nsIDOMWindow) {
        if (nativeTab.messageBrowser) {
          return nativeTab.messageBrowser.contentWindow.gMessage;
        }
      } else if (nativeTab.mode.name == "mail3PaneTab") {
        let msgHdrs = nativeTab.chromeBrowser.contentWindow.gDBView.getSelectedMsgHdrs();
        if (msgHdrs.length == 1) {
          return msgHdrs[0];
        }
      } else if (nativeTab.mode.name == "mailMessageTab") {
        return nativeTab.chromeBrowser.contentWindow.gMessage;
      }
      return null;
    }

    /*
     * Decode the calendar attachment.
     */
    async function getCalendarData(file) {
      let buffer = await file.arrayBuffer();

      return new TextDecoder("utf-8", {
        fatal: false,
      }).decode(buffer);
    }

    /*
     * Extract the iTIP METHOD from the ICS data.
     */
    function getCalendarMethod(data) {
      let unfolded = data.replace(/\r?\n[ \t]/g, "");

      let match = unfolded.match(
        /(?:^|\r?\n)METHOD\s*:\s*([A-Za-z-]+)\s*(?:\r?\n|$)/i
      );

      return match ? match[1].trim().toUpperCase() : "";
    }

    /*
     * Create Thunderbird's native calIItipItem.
     */
    function createItipItem(data) {
      let itipItem = Cc[
        "@mozilla.org/calendar/itip-item;1"
      ].createInstance(Ci.calIItipItem);

      itipItem.init(data);

      return itipItem;
    }

    /*
     * Initialise the iTIP item with the same message information
     * Thunderbird normally supplies when processing a MIME invitation.
     */
    function initialiseItipItem(itipItem, method, msgHdr) {
      if (msgHdr) {
        try {
          if (msgHdr.author) {
            itipItem.sender = msgHdr.author;
          }
        } catch (ex) {
          console.debug(
            "LookOut: unable to set iTIP sender",
            ex
          );
        }
      }

      if (
        typeof cal?.itip?.initItemFromMsgData ==
        "function"
      ) {
        cal.itip.initItemFromMsgData(
          itipItem,
          method,
          msgHdr
        );
      }
    }

    /*
     * Recreate the HTML that CalMimeConverter normally generates for
     * a legacy iMIP invitation.
     *
     * calImipBar.showImipBar() expects #imipHTMLDetails to already exist.
     */
    function createInvitationOverlay(
      window,
      itipItem
    ) {
      let messagePane =
        window.document.getElementById(
          "messagepane"
        );

      if (!messagePane || !messagePane.contentDocument) {
        throw new ExtensionError(
          "Thunderbird messagepane contentDocument not available"
        );
      }

      let contentDocument =
        messagePane.contentDocument;

      let item = itipItem.getItemList()[0];

      if (!item) {
        throw new ExtensionError(
          "iTIP item contains no calendar item"
        );
      }

      let overlayDocument =
        invitation.createInvitationOverlay(
          item,
          itipItem
        );

      let details =
        overlayDocument.getElementById(
          "imipHTMLDetails"
        );

      if (!details) {
        throw new ExtensionError(
          "Thunderbird invitation overlay does not contain imipHTMLDetails"
        );
      }

      let existing =
        contentDocument.getElementById(
          "imipHTMLDetails"
        );

      if (existing) {
        existing.remove();
      }

      let importedDetails =
        contentDocument.importNode(
          details,
          true
        );

      if (!contentDocument.body) {
        throw new ExtensionError(
          "Thunderbird message body is unavailable"
        );
      }

      contentDocument.body.prepend(
        importedDetails
      );

      return contentDocument.getElementById(
        "imipHTMLDetails"
      );
    }

    /*
     * Display the invitation using Thunderbird's native legacy iMIP bar.
     */
    async function showCalendarInvitation(
      window,
      itipItem,
      method
    ) {
      let calImipBar =
        window.calImipBar;

      if (!calImipBar) {
        throw new ExtensionError(
          "Thunderbird calImipBar is not available"
        );
      }

      /*
       * This is the important part for TNEF invitations:
       *
       * CalMimeConverter normally creates this HTML before the iMIP bar
       * is displayed. Since LookOut bypasses CalMimeConverter, create it
       * ourselves using Thunderbird's own invitation generator.
       */
      createInvitationOverlay(
        window,
        itipItem
      );

      calImipBar.showImipBar(
        itipItem,
        method
      );

      /*
       * Allow Thunderbird's asynchronous calendar lookup to update
       * the native invitation controls.
       */
      await new Promise(resolve =>
        window.setTimeout(resolve, 100)
      );
    }

    /*
     * Handle a decoded TNEF calendar invitation.
     */
    async function handleCalendarInvitation(
      tabId,
      file,
      partName
    ) {
      let window =
        getMessageWindow(tabId);

      if (!window) {
        throw new ExtensionError(
          "Unable to obtain Thunderbird message window"
        );
      }

      if (!file) {
        throw new ExtensionError(
          `No calendar file supplied for ${partName}`
        );
      }

      let contentType =
        String(file.type || "")
          .toLowerCase()
          .split(";")[0]
          .trim();

      if (contentType != "text/calendar") {
        return false;
      }

      let data =
        await getCalendarData(file);

      if (!/BEGIN:VCALENDAR/i.test(data)) {
        return false;
      }

      let itipItem =
        createItipItem(data);

      let method =
        String(
          itipItem.receivedMethod ||
            getCalendarMethod(data)
        )
          .trim()
          .toUpperCase();

      if (!method) {
        return false;
      }

      let msgHdr =
        getDisplayedMessage(tabId);

      initialiseItipItem(
        itipItem,
        method,
        msgHdr
      );

      await showCalendarInvitation(
        window,
        itipItem,
        method
      );

      return true;
    }

    return {
      Attachment: {
        listAttachments: async function (tabId) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }
          let attachments = [];
          for (let attachmentInfo of window.currentAttachments) {
            let attachment = {
              contentType: attachmentInfo.contentType,
              name: attachmentInfo.name,
              partName: attachmentInfo.partID,
              size: attachmentInfo.size,
            }
            attachments.push(attachment);
          };
          return attachments;
        },

        getAttachmentFile: async function (tabId, partName) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }
          let attachmentInfo = window.currentAttachments.find(a => a.partID == partName);
          if (!attachmentInfo) {
            throw new ExtensionError(`Attachment with partName ${partName} not found`);
          }
          let bytes = await getAttachmentFromUrl(attachmentInfo.url);
          return new File([bytes], attachmentInfo.name, { type: attachmentInfo.contentType });
        },

        /*
         * NEW: expose TNEF calendar invitation handling.
         */
        handleCalendarInvitation: async function (
          tabId,
          file,
          partName
        ) {
          return await handleCalendarInvitation(
            tabId,
            file,
            partName
          );
        },

        addAttachments: async function (tabId, newAttachments) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }

          let modified = false;
          for (let attachment of newAttachments) {
            let msgHdr = getDisplayedMessage(tabId);
            if (!msgHdr) {
              continue;
            }
            let realFile = await getRealFileForFile(attachment.file);
            let url = `${Services.io.newFileURI(realFile).spec}?part=${attachment.partName}`;
            let attachmentInfo = new AttachmentInfo({
              contentType: attachment.contentType,
              url,
              name: attachment.name,
              uri: msgHdr.folder.getUriForMsg(msgHdr),
              isExternalAttachment: true,
              message: msgHdr,
              updateAttachmentsDisplayFn: window.updateAttachmentsDisplay,
            });
            window.currentAttachments.push(attachmentInfo);
            modified = true;
          }

          if (!modified) {
            return
          }

          ClearAttachmentList(window);
          window.gBuildAttachmentsForCurrentMsg = false;
          await window.displayAttachmentsForExpandedView();
          window.gBuildAttachmentsForCurrentMsg = true;
        },

        removeAttachments: async function (tabId, partNames) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }

          let modified = false;
          for (let index = window.currentAttachments.length; index > 0; index--) {
            let idx = index - 1;
            if (partNames.includes(window.currentAttachments[idx].partID)) {
              window.currentAttachments.splice(idx);
              modified = true;
            }
          }

          if (!modified) {
            return
          }

          ClearAttachmentList(window);
          window.gBuildAttachmentsForCurrentMsg = false;
          await window.displayAttachmentsForExpandedView();
          window.gBuildAttachmentsForCurrentMsg = true;
        },
      },
    };
  }
};
