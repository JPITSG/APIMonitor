export interface ConfigData {
  url: string;
  interval: number;
  loggingEnabled: boolean;
  historyLimit: number;
  logPath?: string;
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
}

type InitCallback = (data: InitData) => void;
type ValidationCallback = (result: ValidationResult) => void;
type HistoryUpdateCallback = (entries: HistoryEntry[]) => void;

let initCallback: InitCallback | null = null;
let validationCallback: ValidationCallback | null = null;
let historyUpdateCallback: HistoryUpdateCallback | null = null;

// Extend window for C <-> JS bridge
declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onValidationResult: (result: ValidationResult) => void;
    onHistoryUpdate: (entries: HistoryEntry[]) => void;
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

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

export function onValidation(cb: ValidationCallback) {
  validationCallback = cb;
}

export function onHistoryUpdate(cb: HistoryUpdateCallback) {
  historyUpdateCallback = cb;
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
    interval: config.interval,
    loggingEnabled: config.loggingEnabled,
    historyLimit: config.historyLimit,
  });
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
