import * as storage from "../scripts/storage.mjs";

const USER_OPTIONS = [
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
  await browser.storage.local.set({ [name]: value })
}

async function init() {
  i18n.updateDocument();
  let prefs = await storage.getPrefs();

  for (let name of USER_OPTIONS) {
    let value = prefs[name]; 
    let element = document.getElementById(`${name}_check`);
    element.checked = value;
    element.dataset.preference = name;
    element.addEventListener("change", update);
  }
}

init();
