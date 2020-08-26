messenger.WindowListener.registerDefaultPrefs("defaults/preferences/lookout.js");
messenger.WindowListener.registerChromeUrl(
  ["content", "lookout",           "chrome/content/"],
  ["locale",  "lookout", "en-US",  "chrome/locale/en-US/"],
  ["locale",  "lookout", "hu",     "chrome/locale/ca/"],
  ["locale",  "lookout", "de",     "chrome/locale/de/"],
  ["locale",  "lookout", "ja",  "chrome/locale/ja/"],
  ["locale",  "lookout", "nl",  "chrome/locale/nl/"]
);
messenger.WindowListener.registerOptionsPage("chrome://lookout/content/options.xul");
messenger.WindowListener.registerWindow(
    "chrome://messenger/content/messenger.xul",
    "chrome://lookout/content/scripts/messenger.js");
messenger.WindowListener.startListening();