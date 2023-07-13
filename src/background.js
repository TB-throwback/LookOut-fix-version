import { TnefExtractor } from "/scripts/lookout.mjs"

async function handleMessage(tab, message) {
  // Skip if message is junk.
  if (message.junk) {
    return;
  }

  // Get preferences.
  let prefs = {
    "debug_level": 7,
    "disable_filename_character_set": true,
    "attach_raw_mapi": true,
    "body_part_prefix": "body_part_"
  }

/*
    var TNEF_PREF_PREFIX = "extensions.lookout.";

    function tnef_get_pref(name, get_type_func, default_val) {
      var pref_name = TNEF_PREF_PREFIX + name;
      var pref_val;
      try {
        var prefs = Cc["@mozilla.org/preferences-service;1"].getService(Ci.nsIPrefBranch);
        pref_val = prefs[get_type_func](pref_name);
      } catch (ex) {
        tnef_log_msg("TNEF: warning: could not retrieve setting '" + pref_name + "': " + ex, 5);
      }
      if (pref_val === void (0))
        pref_val = default_val;

      return pref_val;
    }
    function tnef_get_bool_pref(name, default_val) {
      return tnef_get_pref(name, "getBoolPref", default_val);
    }
    function tnef_get_string_pref(name, default_val) {
      return tnef_get_pref(name, "getCharPref", default_val);
    }
    function tnef_get_int_pref(name, default_val) {
      return tnef_get_pref(name, "getIntPref", default_val);
    }
*/

  // Read attachments of the message
  let attachments = await browser.messages.listAttachments(message.id);
  let removedParts = [];
  let addedFiles = [];
  for (let attachment of attachments) {
    if (attachment.name != "winmail.dat" && attachment.contentType != "application/ms-tnef") {
      continue;
    }
    let file = await browser.messages.getAttachmentFile(message.id, attachment.partName)
    let tnefExtractor = new TnefExtractor();
    let files = await tnefExtractor.parse(file, null, prefs);
    addedFiles.push(...files);

    removedParts.push(attachment.partName);
  }
  if (removedParts.length > 0) {
    await browser.Attachment.removeParts(tab.id, message.id, removedParts);
  }
  if (addedFiles.length > 0) {
    await browser.Attachment.addFiles(tab.id, message.id, addedFiles);
  }
}

// Handle all displayed messages
let tabs = (await browser.tabs.query({})).filter(t => ["messageDisplay", "mail"].includes(t.type));
for (let tab of tabs) {
  let message = await browser.messageDisplay.getDisplayedMessage(tab.id);
  // Do not await this but just fire all requests in parallel and let them finish
  // on its own.
  if (message) {
    handleMessage(tab, message);
  }
}
browser.messageDisplay.onMessageDisplayed.addListener(handleMessage);

/*
messenger.WindowListener.registerDefaultPrefs("defaults/preferences/lookout.js");
messenger.WindowListener.registerChromeUrl([
  ["content", "lookout", "chrome/content/"],
  ["locale", "lookout", "en-US", "chrome/locale/en-US/"],
  ["locale", "lookout", "hu", "chrome/locale/hu/"],
  ["locale", "lookout", "de", "chrome/locale/de/"],
  ["locale", "lookout", "ja", "chrome/locale/ja/"],
  ["locale", "lookout", "nl", "chrome/locale/nl/"]
]);
messenger.WindowListener.registerOptionsPage("chrome://lookout/content/options.xhtml");

messenger.WindowListener.registerWindow(
  "about:message",
  "chrome://lookout/content/scripts/messenger.js");
messenger.WindowListener.startListening();
*/
