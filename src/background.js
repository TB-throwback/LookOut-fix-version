browser.SessionRestore.onStartupSessionRestore.addListener(main);

function main() {
  browser.SessionRestore.onStartupSessionRestore.removeListener(main);

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
