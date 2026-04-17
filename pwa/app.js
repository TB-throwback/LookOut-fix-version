import { TnefExtractor } from "../src/scripts/lookout.mjs";

const fileInput = document.getElementById("fileInput");
const downloadAllBtn = document.getElementById("downloadAllBtn");
const shareAllBtn = document.getElementById("shareAllBtn");
const statusEl = document.getElementById("status");
const resultsEl = document.getElementById("results");
const dropzone = document.getElementById("dropzone");
const isAndroidApp =
  new URLSearchParams(window.location.search).get("android") === "1";
const isDevServer = ["127.0.0.1", "localhost"].includes(
  window.location.hostname,
);

if (isAndroidApp) {
  document.body.classList.add("android-app");
}

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

function hasAndroidShareSupport() {
  const bridge = window.AndroidBridge;
  return isAndroidApp && bridge && typeof bridge.shareFiles === "function";
}

function updateShareAllButtonState() {
  if (!shareAllBtn) {
    return;
  }

  const canUseShareAll = hasAndroidShareSupport();
  shareAllBtn.hidden = !canUseShareAll;

  if (!canUseShareAll) {
    shareAllBtn.disabled = true;
    return;
  }

  const files = extractedFiles.map((item) => item.file);
  shareAllBtn.disabled = files.length === 0 || !hasAndroidShareSupport();
}

function resetResults() {
  extractedFiles.forEach((item) => URL.revokeObjectURL(item.url));
  extractedFiles = [];
  resultsEl.innerHTML = "";
  downloadAllBtn.disabled = true;
  updateShareAllButtonState();
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
      // TODO update also file struct
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
    link.addEventListener("click", async (event) => {
      if (!isAndroidApp) {
        return;
      }
      event.preventDefault();
      try {
        await saveFileViaAndroid(file);
      } catch (error) {
        setStatus(`Download failed: ${error.message || error}`, true);
      }
    });

    const actions = document.createElement("div");
    actions.className = "result-actions";
    actions.append(link);

    if (isAndroidApp) {
      const openBtn = document.createElement("button");
      openBtn.className = "open-btn";
      openBtn.type = "button";
      openBtn.textContent = "Open";
      openBtn.addEventListener("click", async () => {
        try {
          await openFileViaAndroid(file);
        } catch (error) {
          setStatus(`Open failed: ${error.message || error}`, true);
        }
      });

      actions.append(openBtn);
    }

    if (hasAndroidShareSupport()) {
      const shareBtn = document.createElement("button");
      shareBtn.className = "share-btn";
      shareBtn.type = "button";
      shareBtn.textContent = "Share";
      shareBtn.addEventListener("click", async () => {
        try {
          await shareFiles([file]);
        } catch (error) {
          if (error && error.name === "AbortError") {
            return;
          }
          setStatus(`Share failed: ${error.message || error}`, true);
        }
      });

      actions.append(shareBtn);
    }

    li.append(meta, actions);
    frag.append(li);

    return { file, url };
  });

  resultsEl.append(frag);
  downloadAllBtn.disabled = false;
  updateShareAllButtonState();
  setStatus(
    `Extracted ${files.length} attachment${files.length === 1 ? "" : "s"}.`,
  );
}

function setSelectedFile(file) {
  selectedFile = file || null;
  resetResults();

  if (selectedFile) {
    setStatus(
      `Selected: ${selectedFile.name} (${formatBytes(selectedFile.size)})`,
    );
    void extractFromSelectedFile();
  } else {
    setStatus("No file selected.");
  }
}

async function extractFromSelectedFile() {
  if (!selectedFile) {
    return;
  }

  setStatus("Extracting attachments...");

  try {
    const extractor = new TnefExtractor();
    const files = await extractor.parse(selectedFile, {}, { ...prefs });
    renderResults(files || []);
  } catch (error) {
    setStatus(`Extraction failed: ${error.message || error}`, true);
  }
}

async function openFromAndroid(fileName, mimeType) {
  try {
    const response = await fetch("./input", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Unable to read the source file (${response.status}).`);
    }

    const bytes = new Uint8Array(await response.arrayBuffer());
    const file = new File([bytes], fileName || "winmail.dat", {
      type:
        mimeType ||
        response.headers.get("content-type") ||
        "application/octet-stream",
    });

    setSelectedFile(file);
  } catch (error) {
    setStatus(`Extraction failed: ${error.message || error}`, true);
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const comma = result.indexOf(",");
      resolve(comma >= 0 ? result.slice(comma + 1) : result);
    };
    reader.onerror = () => {
      reject(reader.error || new Error("Unable to encode file."));
    };
    reader.readAsDataURL(file);
  });
}

async function saveFileViaAndroid(file) {
  if (
    !window.AndroidBridge ||
    typeof window.AndroidBridge.downloadFile !== "function"
  ) {
    throw new Error("Android download bridge is unavailable.");
  }

  const base64Data = await fileToBase64(file);
  window.AndroidBridge.downloadFile(
    file.name || "attachment.bin",
    file.type || "application/octet-stream",
    base64Data,
  );
}

async function openFileViaAndroid(file) {
  if (
    !window.AndroidBridge ||
    typeof window.AndroidBridge.openFile !== "function"
  ) {
    throw new Error("Android open bridge is unavailable.");
  }

  const base64Data = await fileToBase64(file);
  window.AndroidBridge.openFile(
    file.name || "attachment.bin",
    file.type || "application/octet-stream",
    base64Data,
  );
}

async function shareFiles(files) {
  if (!files.length) {
    return;
  }

  if (hasAndroidShareSupport()) {
    const names = [];
    const mimeTypes = [];
    const base64DataList = [];

    for (const file of files) {
      names.push(file.name || "attachment.bin");
      mimeTypes.push(file.type || "application/octet-stream");
      base64DataList.push(await fileToBase64(file));
    }

    window.AndroidBridge.shareFiles(
      JSON.stringify(names),
      JSON.stringify(mimeTypes),
      JSON.stringify(base64DataList),
    );
    return;
  }

  throw new Error("Sharing is not supported on this device.");
}

async function downloadAll() {
  if (isAndroidApp) {
    try {
      for (const item of extractedFiles) {
        await saveFileViaAndroid(item.file);
      }
      return;
    } catch (error) {
      setStatus(`Download failed: ${error.message || error}`, true);
      return;
    }
  }

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
    const droppedFiles = event.dataTransfer && event.dataTransfer.files;
    const file = droppedFiles && droppedFiles[0];
    if (file) {
      setSelectedFile(file);
    }
  });
}

async function setupPwaIntegrations() {
  if (!isAndroidApp && !isDevServer && "serviceWorker" in navigator) {
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
      const launchFiles = launchParams.files;
      const handle = launchFiles && launchFiles[0];
      if (!handle) {
        return;
      }
      const file = await handle.getFile();
      setSelectedFile(file);
    });
  }
}

fileInput.addEventListener("change", (event) => {
  const files = event.target && event.target.files;
  setSelectedFile((files && files[0]) || null);
});

downloadAllBtn.addEventListener("click", downloadAll);

if (shareAllBtn) {
  shareAllBtn.addEventListener("click", async () => {
    try {
      await shareFiles(extractedFiles.map((item) => item.file));
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      setStatus(`Share failed: ${error.message || error}`, true);
    }
  });
}

setupDropzone();
setupPreferences();
setupPwaIntegrations();
updateShareAllButtonState();

window.Lookout = {
  openFromAndroid,
};
