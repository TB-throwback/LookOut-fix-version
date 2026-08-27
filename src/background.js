import { TnefExtractor } from "/scripts/lookout.mjs";
import * as storage from "./scripts/storage.mjs";

// Migrate legacy prefs to local storage.
await storage.migratePrefs();

// Load Junk message warning
await browser.scripting.messageDisplay.registerScripts([
  {
    id: "lookout-junk-tnef-warning",
    js: ["message-content-script.js"],
  },
]);


async function showJunkWarning(tab) {
  let message = browser.i18n.getMessage("junk_tnef_warning");
  for (let attempt = 0; attempt < 10; attempt++) {
    try {
      await browser.tabs.sendMessage(tab.id, {
        command: "showJunkTnefWarning",
        message,
      });
      return;
    } catch (error) {
      if (attempt == 9) {
        console.error(
          "LookOut: unable to show Junk TNEF warning",
          error
        );
        return;
      }

      await new Promise(resolve => setTimeout(resolve, 50));
    }
  }
}

async function showMessageBody(tab, html) {
  for (let attempt = 0; attempt < 10; attempt++) {
    try {
      await browser.tabs.sendMessage(tab.id, {
        command: "replaceMessageBody",
        html,
      });
      return;
    } catch (error) {
      if (attempt == 9) {
        console.error(
          "LookOut: unable to replace message body",
          error
        );
        return;
      }

      await new Promise(resolve => setTimeout(resolve, 50));
    }
  }
}

async function handleMessage(tab, message) {
  if (message.junk || message.folder?.specialUse?.includes("junk")) {
      console.log(
        "LookOut: TNEF processing cancelled for junk message"
      );

      await showJunkWarning(tab);

      return;
  }

  // Read attachments of the message.
  let attachments = await browser.Attachment.listAttachments(tab.id);

  // Get the current prefs.
  let prefs = await storage.getPrefs();

  let removedParts = [];
  let tnefAttachments = [];
  let calendarInvitations = [];
  let messageBody = null;

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

      /*
       * TNEF can contain the original HTML message body.
       *
       * Keep the decoded body separate so it can replace the body
       * Thunderbird displayed from the MIME message.
       */
      if (tnefFiles[i].name == "body_part_0.html" && prefs["replace_body"]) {
        try {
          messageBody = await tnefFiles[i].text();
        } catch (error) {
          console.error(
            "LookOut: unable to read TNEF message body",
            error
          );
        }

        continue;
      }

      /*
       * TNEF can contain an iCalendar meeting request. At this point
       * Thunderbird's normal MIME parser has already finished, so the
       * decoded calendar part will otherwise only become an attachment.
       *
       * Keep the decoded File object and pass it separately into the
       * Thunderbird iMIP/calendar integration.
       */
      let contentType = (
        tnefFiles[i].type || ""
      ).toLowerCase().split(";")[0].trim();

      if (contentType == "text/calendar") {
        calendarInvitations.push({
          file: tnefFiles[i],
          partName,
        });
      }

      /*
       * Keep adding the decoded file as an attachment. This preserves the
       * existing LookOut behavior and means the ICS remains available to
       * the user even if the calendar integration cannot process it.
       */
      let tnefAttachment = {
        contentType: tnefFiles[i].type,
        name: tnefFiles[i].name,
        size: tnefFiles[i].size,
        partName,
        file: tnefFiles[i],
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
    await browser.Attachment.removeAttachments(tab.id, removedParts);
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
   * Replace the displayed message body if TNEF contained
   * body_part0.html.
   */
   if (messageBody !== null) {
    await showMessageBody(tab, messageBody);
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

// Handle all displayed messages.
let tabs = (await browser.tabs.query({}))
  .filter(t => ["messageDisplay", "mail"].includes(t.type));

for (let tab of tabs) {
  let message = await browser.messageDisplay.getDisplayedMessage(tab.id);

  // Do not await this but just fire all requests in parallel
  // and let them finish on their own.
  if (message) {
    handleMessage(tab, message);
  }
}

browser.messageDisplay.onMessageDisplayed.addListener(handleMessage);
