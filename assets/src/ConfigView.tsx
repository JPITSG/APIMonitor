import { useState, useEffect, useRef, useCallback } from "react";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Switch } from "./components/ui/switch";
import { Badge } from "./components/ui/badge";
import { Label } from "./components/ui/label";
import {
  onValidation,
  validateUrl,
  saveSettings,
  closeDialog,
  configReady,
  checkForUpdate,
  cancelUpdateCheck,
  installUpdate,
  dismissUpdate,
  ignoreUpdateVersion,
  dismissUpdateConfirmation,
  onUpdateResult,
  onUpdateProgress,
  type ConfigData,
  type UpdateResult,
  type ValidationResult,
} from "./lib/bridge";

interface ConfigViewProps {
  config: ConfigData;
  webView2Version: string;
  updateCompletedVersion: string;
}

export default function ConfigView({
  config,
  webView2Version,
  updateCompletedVersion,
}: ConfigViewProps) {
  const [url, setUrl] = useState(config.url);
  const [interval, setInterval] = useState(config.interval);
  const [loggingEnabled, setLoggingEnabled] = useState(config.loggingEnabled);
  const [historyLimit, setHistoryLimit] = useState(String(config.historyLimit));
  const [autoCheckForUpdates, setAutoCheckForUpdates] = useState(
    config.autoCheckForUpdates ?? true
  );

  // 0=none, 1=checking, 2=valid, 3=invalid
  const [validationState, setValidationState] = useState<number>(0);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const [updateChecking, setUpdateChecking] = useState(
    config.updateCheckPending ?? false
  );
  const [updateCancelling, setUpdateCancelling] = useState(false);
  const [updateSpeedKbps, setUpdateSpeedKbps] = useState<number | null>(null);
  const [updateAlert, setUpdateAlert] = useState<UpdateResult | null>(() =>
    updateCompletedVersion
      ? {
          status: "completed",
          title: "Update complete",
          message: `APIMonitor has been updated to version ${updateCompletedVersion}.`,
          currentVersion: "",
          remoteVersion: "",
          automatic: false,
        }
      : null
  );
  const automaticUpdateStarted = useRef(false);

  const handleValidationResult = useCallback((result: ValidationResult) => {
    setValidationState(result.valid ? 2 : 3);
  }, []);

  useEffect(() => {
    onValidation(handleValidationResult);

    // Trigger initial validation if URL is non-empty
    if (config.url.trim()) {
      setValidationState(1);
      validateUrl(config.url.trim());
    }
  }, [config.url, handleValidationResult]);

  useEffect(() => {
    const removeResultListener = onUpdateResult((result) => {
      setUpdateChecking(false);
      setUpdateCancelling(false);
      setUpdateSpeedKbps(null);
      if (result.status === "cancelled") {
        setUpdateAlert((current) =>
          result.automatic && current?.status === "completed" ? current : null
        );
      } else if (result.automatic && result.status !== "newer") {
        setUpdateAlert((current) =>
          current?.status === "completed" ? current : null
        );
      } else {
        setUpdateAlert(result);
      }
    });
    const removeProgressListener = onUpdateProgress((progress) => {
      setUpdateSpeedKbps(Math.max(0, Math.round(progress.kilobytesPerSecond)));
    });

    const shouldCheckAutomatically =
      config.autoCheckForUpdates &&
      !updateCompletedVersion &&
      !config.updateCheckPending &&
      !config.updatePromptPending &&
      !automaticUpdateStarted.current;
    if (shouldCheckAutomatically) {
      automaticUpdateStarted.current = true;
      setUpdateChecking(true);
    }
    configReady(shouldCheckAutomatically);

    return () => {
      removeResultListener();
      removeProgressListener();
    };
  }, [
    config.autoCheckForUpdates,
    config.updateCheckPending,
    config.updatePromptPending,
    updateCompletedVersion,
  ]);

  const handleUpdate = () => {
    if (updateChecking) {
      setUpdateCancelling(true);
      cancelUpdateCheck();
      return;
    }
    setUpdateAlert(null);
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    checkForUpdate(false);
  };

  const handleInstallUpdate = () => {
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    installUpdate();
  };

  const handleDismissUpdate = () => {
    if (updateAlert?.status === "completed") {
      dismissUpdateConfirmation();
    } else {
      dismissUpdate();
    }
    setUpdateAlert(null);
  };

  const handleIgnoreUpdateVersion = () => {
    if (!updateAlert?.remoteVersion) return;
    ignoreUpdateVersion(updateAlert.remoteVersion);
    setUpdateAlert(null);
  };

  const handleUrlChange = (newUrl: string) => {
    setUrl(newUrl);

    if (debounceRef.current) clearTimeout(debounceRef.current);

    const trimmed = newUrl.trim();
    if (!trimmed) {
      setValidationState(0);
      return;
    }

    setValidationState(1);
    debounceRef.current = setTimeout(() => {
      validateUrl(trimmed);
    }, 500);
  };

  const handleSave = () => {
    const trimmedUrl = url.trim();
    if (!trimmedUrl) return;

    let hl = parseInt(historyLimit, 10);
    if (isNaN(hl) || hl < 10) hl = 10;
    if (hl > 10000) hl = 10000;

    saveSettings({
      url: trimmedUrl,
      interval,
      loggingEnabled,
      historyLimit: hl,
      autoCheckForUpdates,
      updateCheckPending: config.updateCheckPending,
      updatePromptPending: config.updatePromptPending,
    });
  };

  const validationBadge = () => {
    switch (validationState) {
      case 1:
        return <Badge variant="secondary">Checking...</Badge>;
      case 2:
        return <Badge variant="success">Valid</Badge>;
      case 3:
        return <Badge variant="destructive">Invalid</Badge>;
      default:
        return null;
    }
  };

  return (
    <div className="p-5 flex flex-col gap-3 max-w-md mx-auto text-xs">
      <div data-row className="flex items-center justify-between gap-3">
        <Label htmlFor="api-url" className="shrink-0">API URL</Label>
        <Input
          id="api-url"
          value={url}
          onChange={(e) => handleUrlChange(e.target.value)}
          placeholder="http://example.com/api/status"
          className="flex-1 min-w-0"
        />
      </div>

      <div data-row className="flex items-center justify-between">
        <Label>API URL Status</Label>
        {validationBadge()}
      </div>

      <div data-row className="flex items-center justify-between">
        <Label htmlFor="interval">Check Interval</Label>
        <select
          id="interval"
          value={interval}
          onChange={(e) => setInterval(Number(e.target.value))}
          className="w-40 h-8 rounded-md border border-neutral-300 bg-transparent px-3 py-1 text-xs shadow-sm focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-neutral-400"
        >
          <option value={60}>Every 1 minute</option>
          <option value={120}>Every 2 minutes</option>
          <option value={300}>Every 5 minutes</option>
        </select>
      </div>

      <div data-row className="flex items-start justify-between gap-3">
        <div className="flex flex-col">
          <Label htmlFor="logging">Enable Debug Logging</Label>
          <span className="text-[10px] text-neutral-500 break-all leading-tight mt-1.5">({config.logPath})</span>
        </div>
        <Switch
          id="logging"
          checked={loggingEnabled}
          onCheckedChange={setLoggingEnabled}
          className="shrink-0 mt-0.5"
        />
      </div>

      <div data-row className="flex items-center justify-between">
        <Label htmlFor="history-limit">History Limit (10-10000)</Label>
        <Input
          id="history-limit"
          type="number"
          min={10}
          max={10000}
          value={historyLimit}
          onChange={(e) => setHistoryLimit(e.target.value)}
          className="w-40"
        />
      </div>

      <div data-row className="flex items-start justify-between gap-3">
        <div className="flex flex-col">
          <Label htmlFor="auto-update">Automatically Check for Updates</Label>
          <span className="text-[10px] text-neutral-500 leading-tight mt-1.5">
            Checks at startup, when Configure opens, and every 60 minutes.
          </span>
        </div>
        <Switch
          id="auto-update"
          checked={autoCheckForUpdates}
          onCheckedChange={setAutoCheckForUpdates}
          className="shrink-0 mt-0.5"
        />
      </div>

      <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
        <span
          className="select-none whitespace-nowrap text-[10px] leading-none tabular-nums text-neutral-400"
          title="Application version / WebView2 version"
        >
          v{__APP_VERSION__} / {webView2Version}
        </span>
        <div className="flex items-center gap-2">
          <Button
            variant={updateChecking ? "destructive" : "outline"}
            size="sm"
            className="min-w-[5rem]"
            disabled={updateCancelling}
            aria-label={
              updateChecking ? "Stop update check and download" : undefined
            }
            title={
              updateChecking ? "Stop update check and download" : undefined
            }
            onClick={handleUpdate}
          >
            {updateCancelling
              ? "Stopping..."
              : updateChecking
                ? updateSpeedKbps === null
                  ? "Checking..."
                  : `Checking (${updateSpeedKbps.toLocaleString()} KB/s)...`
                : "Update"}
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="min-w-[5rem]"
            onClick={() => closeDialog()}
          >
            Cancel
          </Button>
          <Button size="sm" className="min-w-[5rem]" onClick={handleSave}>
            Save
          </Button>
        </div>
      </div>

      {updateAlert && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 p-4">
          <div
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="update-alert-title"
            aria-describedby="update-alert-message"
            className="w-full max-w-sm space-y-3 rounded-lg border border-neutral-200 bg-white p-4 shadow-xl"
          >
            <div className="space-y-1">
              <h2 id="update-alert-title" className="text-sm font-semibold">
                {updateAlert.title}
              </h2>
              <p
                id="update-alert-message"
                className="text-xs leading-relaxed text-neutral-600"
              >
                {updateAlert.message}
              </p>
            </div>
            {updateAlert.currentVersion && updateAlert.remoteVersion && (
              <dl className="grid grid-cols-[1fr_auto] gap-x-4 gap-y-1 rounded-md border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs">
                <dt className="text-neutral-500">Current version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.currentVersion}
                </dd>
                <dt className="text-neutral-500">Remote version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.remoteVersion}
                </dd>
              </dl>
            )}
            <div className="flex justify-end gap-2">
              {updateAlert.status === "newer" && updateAlert.automatic && (
                <Button
                  variant="outline"
                  size="sm"
                  disabled={updateChecking}
                  onClick={handleIgnoreUpdateVersion}
                >
                  Ignore this version
                </Button>
              )}
              {(updateAlert.status === "newer" ||
                updateAlert.status === "same") && (
                <Button
                  variant="outline"
                  size="sm"
                  autoFocus
                  disabled={updateChecking}
                  onClick={handleDismissUpdate}
                >
                  Cancel
                </Button>
              )}
              <Button
                size="sm"
                autoFocus={
                  updateAlert.status !== "newer" &&
                  updateAlert.status !== "same"
                }
                disabled={updateChecking}
                onClick={
                  updateAlert.status === "newer" ||
                  updateAlert.status === "same"
                    ? handleInstallUpdate
                    : handleDismissUpdate
                }
              >
                {updateChecking
                  ? "Starting..."
                  : updateAlert.status === "same"
                    ? "Force update"
                    : updateAlert.status === "newer"
                      ? "Update"
                      : "OK"}
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
