var lookout_s3 = {};
Components.utils.import("resource://gre/modules/Services.jsm");

/**
 * @author Oleksandr
 */

 
//------------------------------------------------------------------------------
lookout_s3.addon = {
	version : '0',
	old_version : '0',
	donateURL: 'http://www.s3blog.org/addon-contribute/lookout-fix-version.html',
	prefService: Components.classes["@mozilla.org/preferences-service;1"].getService(Components.interfaces.nsIPrefService)
};
//------------------------------------------------------------------------------
lookout_s3.addon.is_Thunderbird = function() {
//	return (Services.appinfo.name == 'Thunderbird') ? true : false;
	return (Services.appinfo.ID == '{3550f703-e582-4d05-9a08-453d09bdfdc6}') ? true : false;
}
//------------------------------------------------------------------------------
lookout_s3.addon.get_version = function() {
	Components.utils.import("resource://gre/modules/AddonManager.jsm");
	AddonManager.getAddonByID('lookout@s3_fix_version', function(addon) {
		lookout_s3.addon.version = addon.version;
		if ((addon.version != '') && (addon.version != '0')) {
			setTimeout(function() { lookout_s3.addon.checkPrefs(); }, 2000);
		}
	});
}
//------------------------------------------------------------------------------
lookout_s3.addon.addonDonate = function() {
	var donateURL = lookout_s3.addon.donateURL + '?v=' + lookout_s3.addon.version + '-' + lookout_s3.addon.old_version;
	try {
		//----------------------------------------------------------------
		if (lookout_s3.addon.is_Thunderbird()) {
			var wm = Components.classes["@mozilla.org/appshell/window-mediator;1"].getService(Components.interfaces.nsIWindowMediator);
			var win = wm.getMostRecentWindow("mail:3pane");
			var tabmail = win.document.getElementById('tabmail');
			win.focus();
			if (tabmail) {
				tabmail.openTab('contentTab', { contentPage: donateURL });
			}
		} else {
			gBrowser.selectedTab = gBrowser.addTab(donateURL);
		}
	} catch(e){;}
}
//------------------------------------------------------------------------------
lookout_s3.addon.checkPrefs = function() {
	var mozilla_prefs = lookout_s3.addon.prefService.getBranch("extensions.lookout.");

	//----------------------------------------------------------------------
	var old_version = mozilla_prefs.getCharPref("current_version");
	lookout_s3.addon.old_version = old_version;
	var not_open_contribute_page = mozilla_prefs.getBoolPref("not_open_contribute_page");
	var current_day = Math.ceil((new Date()).getTime() / (1000*60*60*24));
	var is_set_timer = false;
	var show_page_timer =  mozilla_prefs.getIntPref("show_page_timer");

	//----------------------------------------------------------------------
	if (lookout_s3.addon.version != old_version) {
		mozilla_prefs.setCharPref("current_version", lookout_s3.addon.version);
		var result = ((old_version == '') || (old_version == '0')) ? false : true;
		//--------------------------------------------------------------
		if (result) {
			if (! not_open_contribute_page) {
				is_set_timer = true;
				if ((show_page_timer + 5) < current_day) {
					lookout_s3.addon.addonDonate();
				}
			}
		}
	}
	//----------------------------------------------------------------------
	if (lookout_s3.addon.version == old_version) {
		if (show_page_timer > 0) {
			show_page_timer -= Math.floor(Math.random() * 15);
			if ((show_page_timer + 60) < current_day) {
				if (! not_open_contribute_page) {
					is_set_timer = true;
					lookout_s3.addon.addonDonate();
				}
			}
		} else {
			is_set_timer = true;
		}
	}
	//----------------------------------------------------------------------
	if (is_set_timer) {
		mozilla_prefs.setIntPref("show_page_timer", current_day);
	}
}

window.addEventListener("load", lookout_s3.addon.get_version, false);
