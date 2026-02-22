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
  type ConfigData,
  type ValidationResult,
} from "./lib/bridge";

interface ConfigViewProps {
  config: ConfigData;
}

export default function ConfigView({ config }: ConfigViewProps) {
  const [url, setUrl] = useState(config.url);
  const [interval, setInterval] = useState(config.interval);
  const [loggingEnabled, setLoggingEnabled] = useState(config.loggingEnabled);
  const [historyLimit, setHistoryLimit] = useState(String(config.historyLimit));

  // 0=none, 1=checking, 2=valid, 3=invalid
  const [validationState, setValidationState] = useState<number>(0);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

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

      <div className="flex justify-end gap-2 pt-1">
        <Button variant="outline" size="sm" className="min-w-[5rem]" onClick={() => closeDialog()}>
          Cancel
        </Button>
        <Button size="sm" className="min-w-[5rem]" onClick={handleSave}>
          Save
        </Button>
      </div>
    </div>
  );
}
