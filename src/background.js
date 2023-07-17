import { TnefExtractor } from "/scripts/lookout.mjs"

const PREF_PREFIX = "extensions.lookout.";
const PREF_DEFAULTS = {
  "attach_raw_mapi": false,
  "direct_to_calendar": false,
  "disable_filename_character_set": false,
  "remove_winmail_dat": true,
  "strict_contenttype": true,
  "debug_enabled": false,
  "body_part_prefix": "body_part_",
}

let prefs = {};

for (let [name, value] of Object.entries(PREF_DEFAULTS)) {
  await browser.LegacyPrefs.setDefaultPref(`${PREF_PREFIX}${name}`, value);
  prefs[name] = await browser.LegacyPrefs.getPref(`${PREF_PREFIX}${name}`);
}

async function handleMessage(tab, message) {
  // Skip if message is junk.
  if (message.junk) {
    return;
  }

  // Read attachments of the message
  let attachments = await browser.Attachment.listAttachments(tab.id);

  let removedParts = [];
  let tnefAttachments = [];
  for (let attachment of attachments) {
    if (
      attachment.name != "winmail.dat" &&
      attachment.contentType != "application/ms-tnef" &&
      prefs["strict_contenttype"]
    ) {
      continue;
    }

    let file = await browser.Attachment.getAttachmentFile(tab.id, attachment.partName);

    let tnefExtractor = new TnefExtractor();
    let tnefFiles = await tnefExtractor.parse(file, null, prefs);
    for (let i = 0; i < tnefFiles.length; i++) {
      let partName = `${attachment.partName}.${i+1}`
      // Skip if we have added that attachment already.
      if (attachments.find(a => a.partName == partName)) {
        continue;
      }

      let tnefAttachment = {
        contentType: tnefFiles[i].type,
        name: tnefFiles[i].name,
        size: tnefFiles[i].size,
        partName,
        file: tnefFiles[i],
      }
      tnefAttachments.push(tnefAttachment);
    }
    if (prefs["remove_winmail_dat"]) {
      removedParts.push(attachment.partName);
    }
  }
  if (removedParts.length > 0) {
    await browser.Attachment.removeAttachments(tab.id, removedParts);
  }
  if (tnefAttachments.length > 0) {
    await browser.Attachment.addAttachments(tab.id, tnefAttachments);
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

// Update prefs chache, if they are changed.
browser.LegacyPrefs.onChanged.addListener(async (name, value) => {
  // The value is null if it is the default, so we always pull the new value.
  prefs[name] = await browser.LegacyPrefs.getPref(`${PREF_PREFIX}${name}`);
}, PREF_PREFIX);
