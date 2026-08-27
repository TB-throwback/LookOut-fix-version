import { TnefExtractor } from "/scripts/lookout.mjs";
import * as storage from "./scripts/storage.mjs";

// Migrate legacy prefs to local storage.
await storage.migratePrefs();

async function handleMessage(tab, message) {
  // Skip if message is junk.
  if (message.junk) {
    return;
  }

  // Read attachments of the message.
  let attachments = await browser.Attachment.listAttachments(tab.id);

  // Get the current prefs.
  let prefs = await storage.getPrefs();

  let removedParts = [];
  let tnefAttachments = [];
  let calendarInvitations = [];

  for (let attachment of attachments) {
    if (
      attachment.name != "winmail.dat" &&
      attachment.contentType != "application/ms-tnef" &&
      prefs["strict_contenttype"]
    ) {
      continue;
    }

    let file = await browser.Attachment.getAttachmentFile(
      tab.id,
      attachment.partName
    );

    let tnefExtractor = new TnefExtractor();
    let tnefFiles = await tnefExtractor.parse(file, null, prefs);

    for (let i = 0; i < tnefFiles.length; i++) {
      let partName = `${attachment.partName}.${i + 1}`;

      // Skip if we have added that attachment already.
      if (attachments.find(a => a.partName == partName)) {
        continue;
      }

      let tnefFile = tnefFiles[i];

      /*
       * TNEF can contain an iCalendar meeting request.  At this point
       * Thunderbird's normal MIME parser has already finished, so the
       * decoded calendar part will otherwise only become an attachment.
       *
       * Keep the decoded File object and pass it separately into the
       * Thunderbird iMIP/calendar integration.
       */
      let contentType = (
        tnefFile.type || ""
      ).toLowerCase().split(";")[0].trim();

      if (contentType == "text/calendar") {
        calendarInvitations.push({
          file: tnefFile,
          partName,
        });
      }

      /*
       * Keep adding the decoded file as an attachment. This preserves the
       * existing LookOut behavior and means the ICS remains available to
       * the user even if the calendar integration cannot process it.
       */
      let tnefAttachment = {
        contentType: tnefFile.type,
        name: tnefFile.name,
        size: tnefFile.size,
        partName,
        file: tnefFile,
      };

      tnefAttachments.push(tnefAttachment);
    }

    if (tnefFiles.length > 0 && prefs["remove_winmail_dat"]) {
      removedParts.push(attachment.partName);
    }
  }

  /*
   * Remove winmail.dat before adding the decoded attachments, as before.
   */
  if (removedParts.length > 0) {
    await browser.Attachment.removeAttachments(
      tab.id,
      removedParts
    );
  }

  /*
   * Add the decoded TNEF attachments.
   */
  if (tnefAttachments.length > 0) {
    await browser.Attachment.addAttachments(
      tab.id,
      tnefAttachments
    );
  }

  /*
   * Now expose decoded calendar data to Thunderbird's existing iMIP
   * processing pipeline.
   *
   * This must happen after TNEF decoding because Thunderbird's MIME parser
   * never sees the calendar part hidden inside winmail.dat.
   */
  for (let invitation of calendarInvitations) {
    try {
      await browser.Attachment.handleCalendarInvitation(
        tab.id,
        invitation.file,
        invitation.partName
      );
    } catch (error) {
      console.error(
        "LookOut: unable to process decoded calendar invitation",
        invitation.partName,
        error
      );
    }
  }
}

// Handle all currently displayed messages.
let tabs = (await browser.tabs.query({}))
  .filter(t => ["messageDisplay", "mail"].includes(t.type));

for (let tab of tabs) {
  let message = await browser.messageDisplay.getDisplayedMessage(tab.id);

  // Do not await this. Fire all requests in parallel and let them finish
  // on their own.
  if (message) {
    handleMessage(tab, message);
  }
}

browser.messageDisplay.onMessageDisplayed.addListener(handleMessage);
