import { TnefExtractor } from "../src/scripts/lookout.mjs";

const fileInput = document.getElementById("fileInput");
const extractBtn = document.getElementById("extractBtn");
const downloadAllBtn = document.getElementById("downloadAllBtn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const dropzone = document.getElementById("dropzone");

let selectedFile = null;
let extractedFiles = [];

const PREF_DEFAULTS = {
  debug_enabled: false,
  attach_raw_mapi: false,
  disable_filename_character_set: false,
};

const PREF_STORAGE_PREFIX = "lookout.pref.";
const USER_OPTIONS = Object.keys(PREF_DEFAULTS);

const prefs = { ...PREF_DEFAULTS };

function readPref(name) {
  try {
    const value = localStorage.getItem(`${PREF_STORAGE_PREFIX}${name}`);
    if (value === null) {
      return PREF_DEFAULTS[name];
    }
    return value === "true";
  } catch {
    return PREF_DEFAULTS[name];
  }
}

function writePref(name, value) {
  try {
    localStorage.setItem(
      `${PREF_STORAGE_PREFIX}${name}`,
      String(Boolean(value)),
    );
  } catch {
    // Ignore storage failures (private mode / disabled storage).
  }
}

function setupPreferences() {
  USER_OPTIONS.forEach((name) => {
    const checkbox = document.getElementById(`${name}_check`);
    if (!checkbox) {
      return;
    }

    const currentValue = readPref(name);
    prefs[name] = currentValue;
    checkbox.checked = currentValue;

    checkbox.addEventListener("change", (event) => {
      const nextValue = Boolean(event.target.checked);
      prefs[name] = nextValue;
      writePref(name, nextValue);
    });
  });
}

function formatBytes(size) {
  const units = ["B", "KB", "MB", "GB"];
  let idx = 0;
  let val = size;
  while (val >= 1024 && idx < units.length - 1) {
    val /= 1024;
    idx += 1;
  }
  return `${val.toFixed(idx === 0 ? 0 : 1)} ${units[idx]}`;
}

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", isError);
}

function resetResults() {
  extractedFiles.forEach((item) => URL.revokeObjectURL(item.url));
  extractedFiles = [];
  resultsEl.innerHTML = "";
  downloadAllBtn.disabled = true;
}

function renderResults(files) {
  resetResults();

  if (!files.length) {
    setStatus("No embedded files were found in this winmail.dat.", true);
    return;
  }

  const frag = document.createDocumentFragment();
  extractedFiles = files.map((file) => {
    const url = URL.createObjectURL(file);

    const li = document.createElement("li");
    li.className = "result-item";

    const meta = document.createElement("div");
    meta.className = "result-meta";
    const name = document.createElement("strong");
    if (file.name === "NaN.html") {
      name.textContent = "message.html";
    } else {
      name.textContent = file.name || "attachment.bin";
    }
    const details = document.createElement("small");
    details.textContent = `${file.type || "application/octet-stream"} • ${formatBytes(file.size)}`;
    meta.append(name, details);

    const link = document.createElement("a");
    link.className = "download-link";
    link.href = url;
    link.download = file.name || "attachment.bin";
    link.textContent = "Download";

    li.append(meta, link);
    frag.append(li);

    return { file, url };
  });

  resultsEl.append(frag);
  downloadAllBtn.disabled = false;
  setStatus(
    `Extracted ${files.length} attachment${files.length === 1 ? "" : "s"}.`,
  );
}

function setSelectedFile(file) {
  selectedFile = file || null;
  extractBtn.disabled = !selectedFile;
  resetResults();

  if (selectedFile) {
    setStatus(
      `Selected: ${selectedFile.name} (${formatBytes(selectedFile.size)})`,
    );
  } else {
    setStatus("No file selected.");
  }
}

async function extractFromSelectedFile() {
  if (!selectedFile) {
    return;
  }

  setStatus("Extracting attachments...");
  extractBtn.disabled = true;

  try {
    const extractor = new TnefExtractor();
    const files = await extractor.parse(selectedFile, {}, { ...prefs });
    renderResults(files || []);
  } catch (error) {
    setStatus(`Extraction failed: ${error.message || error}`, true);
  } finally {
    extractBtn.disabled = !selectedFile;
  }
}

function downloadAll() {
  extractedFiles.forEach((item, idx) => {
    const a = document.createElement("a");
    a.href = item.url;
    a.download = item.file.name || `attachment-${idx + 1}.bin`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  });
}

function setupDropzone() {
  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });

  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      setSelectedFile(file);
    }
  });
}

async function setupPwaIntegrations() {
  if ("serviceWorker" in navigator) {
    try {
      await navigator.serviceWorker.register("./sw.js", { scope: "./" });
    } catch (error) {
      setStatus(
        `Service worker registration failed: ${error.message || error}`,
        true,
      );
    }
  }

  if (
    "launchQueue" in window &&
    typeof LaunchParams !== "undefined" &&
    "files" in LaunchParams.prototype
  ) {
    launchQueue.setConsumer(async (launchParams) => {
      const handle = launchParams.files?.[0];
      if (!handle) {
        return;
      }
      const file = await handle.getFile();
      setSelectedFile(file);
      await extractFromSelectedFile();
    });
  }
}

fileInput.addEventListener("change", (event) => {
  setSelectedFile(event.target.files?.[0] || null);
});

extractBtn.addEventListener("click", extractFromSelectedFile);
downloadAllBtn.addEventListener("click", downloadAll);

setupDropzone();
setupPreferences();
setupPwaIntegrations();
