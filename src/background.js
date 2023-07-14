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
  let addedFiles = [];
  for (let attachment of attachments) {
    if (attachment.name != "winmail.dat" && attachment.contentType != "application/ms-tnef") {
      continue;
    }
    let tnefExtractor = new TnefExtractor();
    let file = await browser.Attachment.getAttachmentFile(tab.id, attachment.partName);
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

// Update prefs chache, if they are changed.
browser.LegacyPrefs.onChanged.addListener((name, value) => {
  prefs[name] = value
}, PREF_PREFIX);
