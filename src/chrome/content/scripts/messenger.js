// Import any needed modules.
var { Services } = ChromeUtils.import("resource://gre/modules/Services.jsm");

// Load an additional JavaScript file.
Services.scriptloader.loadSubScript("chrome://lookout/content/mapi_props.js", window, "UTF-8");
Services.scriptloader.loadSubScript("chrome://lookout/content/tnef.js", window, "UTF-8");
Services.scriptloader.loadSubScript("chrome://lookout/content/lookout.js", window, "UTF-8");

function onLoad(activatedWhileWindowOpen) {

  WL.injectCSS("resource://lookout/skin/lookout.css");
  WL.injectElements(`
    <menupopup id="taskPopup">
      <menuitem id="lookout-settings" label="&lookout.label;" oncommand="lookout.openSettings();" insertbefore="prefSep" class="menu-iconic lookout-icon menuitem-iconic" />
    </menupopup>`);

  window.LookoutLoad();
}

function onUnload(deactivatedWhileWindowOpen) {
  window.LookoutUnload();
}
