import "./styles.css";
import {
  DEFAULT_PREFERENCES,
  buildManifestWorkbook,
  normalisePreferences,
  parseManifest
} from "./converter.js";

const SETTINGS_KEY = "manifest-converter-web-preferences-v1";

const elements = {
  fileInput: document.querySelector("#file-input"),
  fileTrigger: document.querySelector("#file-trigger"),
  fileSummary: document.querySelector("#file-summary"),
  fileName: document.querySelector("#file-name"),
  fileMeta: document.querySelector("#file-meta"),
  removeFile: document.querySelector("#remove-file"),
  convertButton: document.querySelector("#convert-button"),
  convertButtonLabel: document.querySelector("#convert-button-label"),
  statusPanel: document.querySelector("#status-panel"),
  statusTitle: document.querySelector("#status-title"),
  statusMessage: document.querySelector("#status-message"),
  preferencesForm: document.querySelector("#preferences-form"),
  bulkyKeywords: document.querySelector("#bulky-keywords"),
  resetPreferences: document.querySelector("#reset-preferences"),
  preferenceStatus: document.querySelector("#preference-status")
};

let selectedFile = null;
let parsedManifest = null;
let preferences = loadPreferences();
let preferenceStatusTimer = null;

applyPreferencesToForm(preferences);

elements.fileTrigger.addEventListener("click", () => elements.fileInput.click());
elements.fileInput.addEventListener("change", () => selectFile(elements.fileInput.files[0]));
elements.removeFile.addEventListener("click", clearSelectedFile);
elements.convertButton.addEventListener("click", convertSelectedFile);

for (const eventName of ["dragenter", "dragover"]) {
  elements.fileTrigger.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.fileTrigger.classList.add("is-dragging");
  });
}

for (const eventName of ["dragleave", "drop"]) {
  elements.fileTrigger.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.fileTrigger.classList.remove("is-dragging");
  });
}

elements.fileTrigger.addEventListener("drop", (event) => {
  selectFile(event.dataTransfer.files[0]);
});

elements.preferencesForm.addEventListener("input", () => {
  preferences = preferencesFromForm();
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(preferences));
  announcePreferencesSaved();

  if (selectedFile) {
    reviewSelectedFile(selectedFile);
  }
});

elements.resetPreferences.addEventListener("click", () => {
  preferences = normalisePreferences(DEFAULT_PREFERENCES);
  applyPreferencesToForm(preferences);
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(preferences));
  announcePreferencesSaved("Defaults restored");

  if (selectedFile) {
    reviewSelectedFile(selectedFile);
  }
});

async function selectFile(file) {
  if (!file) return;

  if (!file.name.toLocaleLowerCase("en-GB").endsWith(".csv")) {
    setStatus("error", "CSV required", "Select a file with a .csv extension.");
    return;
  }

  selectedFile = file;
  elements.fileName.textContent = file.name;
  elements.fileMeta.textContent = formatFileSize(file.size);
  elements.fileSummary.hidden = false;
  await reviewSelectedFile(file);
}

async function reviewSelectedFile(file) {
  setBusy(true, "Checking manifest…");

  try {
    const csvText = await file.text();
    if (file !== selectedFile) return;
    parsedManifest = parseManifest(csvText, preferences);
    elements.fileMeta.textContent = [
      formatFileSize(file.size),
      `${parsedManifest.orderCount} ${pluralise(parsedManifest.orderCount, "order")}`,
      `${parsedManifest.vehicleCount} ${pluralise(parsedManifest.vehicleCount, "vehicle")}`
    ].join(" · ");
    setStatus(
      "ready",
      "Ready to convert",
      `${parsedManifest.orderCount} orders across ${parsedManifest.vehicleCount} vehicles.`
    );
    elements.convertButton.disabled = false;
  } catch (error) {
    parsedManifest = null;
    elements.convertButton.disabled = true;
    setStatus("error", "Manifest could not be read", messageFromError(error));
  } finally {
    setBusy(false);
  }
}

async function convertSelectedFile() {
  if (!selectedFile || !parsedManifest) return;

  setBusy(true, "Creating workbook…");
  setStatus(
    "working",
    "Creating the Excel workbook",
    "Building vehicle sheets and Code128 barcodes on this device."
  );

  try {
    await nextPaint();
    const result = await buildManifestWorkbook(
      parsedManifest,
      selectedFile.name,
      preferences
    );
    downloadBuffer(result.buffer, result.outputFileName);
    setStatus(
      "success",
      "Workbook downloaded",
      `${result.outputFileName} is ready in your Downloads folder.`
    );
  } catch (error) {
    setStatus("error", "Conversion failed", messageFromError(error));
  } finally {
    setBusy(false);
  }
}

function clearSelectedFile() {
  selectedFile = null;
  parsedManifest = null;
  elements.fileInput.value = "";
  elements.fileSummary.hidden = true;
  elements.convertButton.disabled = true;
  setStatus("neutral", "No manifest selected", "Choose a CSV to review it before conversion.");
  elements.fileTrigger.focus();
}

function setBusy(isBusy, label) {
  elements.convertButton.classList.toggle("is-busy", isBusy);
  elements.convertButton.disabled = isBusy || !parsedManifest;
  elements.fileTrigger.disabled = isBusy;
  elements.removeFile.disabled = isBusy;
  elements.preferencesForm.querySelectorAll("input, textarea").forEach((control) => {
    control.disabled = isBusy;
  });
  elements.convertButtonLabel.textContent = isBusy ? label : "Convert and download";
}

function setStatus(kind, title, message) {
  elements.statusPanel.dataset.kind = kind;
  elements.statusTitle.textContent = title;
  elements.statusMessage.textContent = message;
}

function preferencesFromForm() {
  const data = new FormData(elements.preferencesForm);
  return normalisePreferences({
    bulkyItemsFirst: data.has("bulkyItemsFirst"),
    bulkyKeywords: elements.bulkyKeywords.value.split(/\r?\n/),
    includeBulkPickNote: data.has("includeBulkPickNote"),
    includeTradeRedeliveryNote: data.has("includeTradeRedeliveryNote"),
    includeItems: data.has("includeItems"),
    addDeliveryBorders: data.has("addDeliveryBorders")
  });
}

function applyPreferencesToForm(value) {
  const options = normalisePreferences(value);
  for (const name of [
    "bulkyItemsFirst",
    "includeBulkPickNote",
    "includeTradeRedeliveryNote",
    "includeItems",
    "addDeliveryBorders"
  ]) {
    elements.preferencesForm.elements[name].checked = options[name];
  }
  elements.bulkyKeywords.value = options.bulkyKeywords.join("\n");
}

function loadPreferences() {
  try {
    const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY));
    return normalisePreferences(stored ?? DEFAULT_PREFERENCES);
  } catch {
    return normalisePreferences(DEFAULT_PREFERENCES);
  }
}

function announcePreferencesSaved(message = "Preferences saved") {
  window.clearTimeout(preferenceStatusTimer);
  elements.preferenceStatus.textContent = message;
  preferenceStatusTimer = window.setTimeout(() => {
    elements.preferenceStatus.textContent = "";
  }, 1800);
}

function downloadBuffer(buffer, fileName) {
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} bytes`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function pluralise(count, word) {
  return count === 1 ? word : `${word}s`;
}

function messageFromError(error) {
  return error instanceof Error ? error.message : "An unexpected error occurred.";
}

function nextPaint() {
  return new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
}
