// Import any needed modules.
var Services = globalThis.Services || ChromeUtils.import(
  "resource://gre/modules/Services.jsm"
).Services;

// Load an additional JavaScript file.
Services.scriptloader.loadSubScript("chrome://lookout/content/mapi_props.js", window, "UTF-8");
Services.scriptloader.loadSubScript("chrome://lookout/content/tnef.js", window, "UTF-8");
Services.scriptloader.loadSubScript("chrome://lookout/content/lookout.js", window, "UTF-8");

function onLoad(isAddonActivation) {
  WL.injectCSS("resource://lookout/skin/lookout.css");
  window.LookoutLoad();
}

function onUnload(isAddonDeactivation) {
  if (isAddonDeactivation) {
    window.LookoutUnload();
  }
}
