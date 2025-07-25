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

export async function migratePrefs() {
  for (let name of Object.keys(PREF_DEFAULTS)) {
    let value = await browser.LegacyPrefs.getUserPref(`${PREF_PREFIX}${name}`);
    if (value !== null) {
      console.info(`Migrating LookOut preference ${name} to local storage`);
      await browser.storage.local.set({ [name]: value });
      await browser.LegacyPrefs.clearUserPref(`${PREF_PREFIX}${name}`);
    }
  }
}

export async function getPrefs() {
  // One option to access stored values is
  //    let { option } = await browser.storage.local.get({ option: defaultValue })
  // We can therefore simply pass in our PREF_DEFAULTS object and get back the
  // desired prefs object, with the current values for each stored pref, or the
  // associated default value.
  return browser.storage.local.get(PREF_DEFAULTS)
}