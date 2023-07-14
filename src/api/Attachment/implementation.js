/*
 * Author: John Bieling (john@thunderbird.net)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

var { ExtensionCommon } = ChromeUtils.import(
  "resource://gre/modules/ExtensionCommon.jsm"
);
var { ExtensionUtils } = ChromeUtils.import(
  "resource://gre/modules/ExtensionUtils.jsm"
);
var { ExtensionError } = ExtensionUtils;

ChromeUtils.defineESModuleGetters(this, {
  AttachmentInfo: "resource:///modules/AttachmentInfo.sys.mjs",
});

Cu.importGlobalProperties(["File", "IOUtils", "PathUtils"]);

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
                  `Failed to read attachment ${attachment.url} content: ${status}`
                )
              );
            }
          },
        });
        channel.asyncOpen(listener, null);
      });
    
      return raw;
    }

    return {
      Attachment: {
        listAttachments: async function(tabId) {
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
        getAttachmentFile: async function(tabId, partName) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }
          let attachmentInfo = window.currentAttachments.find(a => a.partID == partName);
          if (!attachmentInfo) {
            throw new ExtensionError(`Attachment with partName ${partName} not found`);
          }
          let bytes = await getAttachmentFromUrl(attachmentInfo.url);
          return new File([bytes], attachmentInfo.name, {type: attachmentInfo.contentType});
        },
        removeParts: async function (tabId, messageId, removedParts) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }

          let modified = false;
          for (let index = window.currentAttachments.length; index > 0; index--) {
            let idx = index - 1;
            if (removedParts.includes(window.currentAttachments[idx].partID)) {
              window.currentAttachments.splice(idx);
              modified = true;
            }
          }

          if (!modified) {
            return
          }

          await window.ClearAttachmentList();
          window.gBuildAttachmentsForCurrentMsg = false;
          await window.displayAttachmentsForExpandedView();
          window.gBuildAttachmentsForCurrentMsg = true;
        },
        addFiles: async function (tabId, messageId, addedFiles) {
          let window = getMessageWindow(tabId);
          if (!window) {
            return
          }

          let modified = false;
          //let attachmentList = window.document.getElementById("attachmentList");
          for (let file of addedFiles) {
            let msgHdr = context.extension.messageManager.get(messageId);
            let realFile = await getRealFileForFile(file);
            let url = Services.io.newFileURI(realFile).spec;
            let attachmentInfo = new AttachmentInfo({
              contentType: file.type,
              url,
              name: file.name,
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

          await window.ClearAttachmentList();
          window.gBuildAttachmentsForCurrentMsg = false;
          await window.displayAttachmentsForExpandedView();
          window.gBuildAttachmentsForCurrentMsg = true;
        },
      },
    };
  }
};
