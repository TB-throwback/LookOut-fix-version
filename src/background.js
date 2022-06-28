async function getThunderbirdVersion() {
  let browserInfo = await messenger.runtime.getBrowserInfo();
  let parts = browserInfo.version.split(".");
  return {
    major: parseInt(parts[0]),
    minor: parseInt(parts[1]),
    revision: parts.length > 2 ? parseInt(parts[2]) : 0,
  }
}

async function main() {
  var majorVersion = await getThunderbirdVersion();
  
  // Must delay startup for issue #79
  // Thunderbird 102 does this by default.
  let startupDelay;
  if (majorVersion.major < 92) {
    startupDelay = await new Promise(async (resolve) => {
      const restoreListener = (window, state = true) => {
        browser.SessionRestore.onStartupSessionRestore.removeListener(restoreListener);
        resolve(state);
      }
      browser.SessionRestore.onStartupSessionRestore.addListener(restoreListener);
  
      let isRestored = await browser.SessionRestore.isRestored();
      if (isRestored) {
        restoreListener(null, false);
      }
    });
  }
  
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
    "chrome://messenger/content/messenger.xhtml",
    "chrome://lookout/content/scripts/messenger.js");
  messenger.WindowListener.registerWindow(
    "chrome://messenger/content/messageWindow.xhtml",
    "chrome://lookout/content/scripts/messenger.js");
  messenger.WindowListener.startListening();
}

main();
