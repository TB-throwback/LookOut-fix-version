const PREF_PREFIX = "extensions.lookout.";
const PREF_NAMES = [
  "attach_raw_mapi",
  "direct_to_calendar",
  "disable_filename_character_set",
  "remove_winmail_dat",
  "strict_contenttype",
  "debug_enabled",
]

async function update(event) {
  let name = event.target.dataset.preference;
  let value = event.target.checked;
  browser.LegacyPrefs.setPref( `${PREF_PREFIX}${name}`, value);
}

async function init() {
  i18n.updateDocument();

  for (let name of PREF_NAMES) {
    let value = await browser.LegacyPrefs.getPref( `${PREF_PREFIX}${name}`);
    let element = document.getElementById(`${name}_check`);
    element.checked = value;
    element.dataset.preference = name;
    element.addEventListener("change", update);
  }
}

init();
