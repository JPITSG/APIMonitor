export interface ConfigData {
  url: string;
  healthyInterval: number;
  downInterval: number;
  loggingEnabled: boolean;
  historyLimit: number;
  logPath?: string;
  autoCheckForUpdates: boolean;
  updateCheckPending: boolean;
  updatePromptPending: boolean;
}

export interface ValidationResult {
  valid: boolean;
}

export interface HistoryEntry {
  time: string;
  from: string;
  to: string;
  message: string;
}

export interface InitData {
  view: "config" | "history";
  config?: ConfigData;
  history?: HistoryEntry[];
  webView2Version?: string;
  updateCompletedVersion?: string;
}

export interface UpdateResult {
  status:
    | "newer"
    | "same"
    | "older"
    | "cancelled"
    | "error"
    | "completed";
  title: string;
  message: string;
  currentVersion: string;
  remoteVersion: string;
  automatic: boolean;
}

export interface UpdateProgress {
  kilobytesPerSecond: number;
}

type InitCallback = (data: InitData) => void;
type ValidationCallback = (result: ValidationResult) => void;
type HistoryUpdateCallback = (entries: HistoryEntry[]) => void;

let initCallback: InitCallback | null = null;
let validationCallback: ValidationCallback | null = null;
let historyUpdateCallback: HistoryUpdateCallback | null = null;
let updateResultCallback: ((result: UpdateResult) => void) | null = null;
let updateProgressCallback: ((progress: UpdateProgress) => void) | null = null;

// Extend window for C <-> JS bridge
declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onValidationResult: (result: ValidationResult) => void;
    onHistoryUpdate: (entries: HistoryEntry[]) => void;
    onUpdateResult: (result: UpdateResult) => void;
    onUpdateProgress: (progress: UpdateProgress) => void;
    chrome?: {
      webview?: {
        postMessage: (s: string) => void;
      };
    };
  }
}

// Called by C via ExecuteScript
window.onInit = (data: InitData) => {
  initCallback?.(data);
};

window.onValidationResult = (result: ValidationResult) => {
  validationCallback?.(result);
};

window.onHistoryUpdate = (entries: HistoryEntry[]) => {
  historyUpdateCallback?.(entries);
};

window.onUpdateResult = (result: UpdateResult) => {
  updateResultCallback?.(result);
};

window.onUpdateProgress = (progress: UpdateProgress) => {
  updateProgressCallback?.(progress);
};

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

export function onValidation(cb: ValidationCallback) {
  validationCallback = cb;
}

export function onHistoryUpdate(cb: HistoryUpdateCallback) {
  historyUpdateCallback = cb;
}

export function onUpdateResult(cb: (result: UpdateResult) => void) {
  updateResultCallback = cb;
  return () => {
    if (updateResultCallback === cb) updateResultCallback = null;
  };
}

export function onUpdateProgress(cb: (progress: UpdateProgress) => void) {
  updateProgressCallback = cb;
  return () => {
    if (updateProgressCallback === cb) updateProgressCallback = null;
  };
}

function postMessage(msg: Record<string, unknown>) {
  try {
    window.chrome?.webview?.postMessage(JSON.stringify(msg));
  } catch {
    console.log("postMessage (no WebView2):", msg);
  }
}

export function getInit() {
  postMessage({ action: "getInit" });
}

export function validateUrl(url: string) {
  postMessage({ action: "validateUrl", url });
}

export function saveSettings(config: ConfigData) {
  postMessage({
    action: "saveSettings",
    url: config.url,
    healthyInterval: config.healthyInterval,
    downInterval: config.downInterval,
    loggingEnabled: config.loggingEnabled,
    historyLimit: config.historyLimit,
    autoCheckForUpdates: config.autoCheckForUpdates,
  });
}

export function configReady(checkAutomatically = false) {
  postMessage({ action: "configReady", checkAutomatically });
}

export function checkForUpdate(automatic = false) {
  postMessage({ action: "checkUpdate", automatic });
}

export function cancelUpdateCheck() {
  postMessage({ action: "cancelUpdateCheck" });
}

export function installUpdate() {
  postMessage({ action: "installUpdate" });
}

export function dismissUpdate() {
  postMessage({ action: "dismissUpdate" });
}

export function ignoreUpdateVersion(version: string) {
  postMessage({ action: "ignoreUpdateVersion", version });
}

export function dismissUpdateConfirmation() {
  postMessage({ action: "dismissUpdateConfirmation" });
}

export function clearHistory() {
  postMessage({ action: "clearHistory" });
}

export function closeDialog() {
  postMessage({ action: "close" });
}

export function reportHeight(height: number) {
  postMessage({ action: "resize", height });
}
