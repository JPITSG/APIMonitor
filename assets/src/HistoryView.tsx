import { useState, useEffect, useCallback } from "react";
import { Button } from "./components/ui/button";
import {
  onHistoryUpdate,
  clearHistory,
  closeDialog,
  type HistoryEntry,
} from "./lib/bridge";

interface HistoryViewProps {
  initialHistory: HistoryEntry[];
}

export default function HistoryView({ initialHistory }: HistoryViewProps) {
  const [history, setHistory] = useState<HistoryEntry[]>(initialHistory);
  const [selectedIndex, setSelectedIndex] = useState<number>(-1);

  const handleHistoryUpdate = useCallback((entries: HistoryEntry[]) => {
    setHistory(entries);
    setSelectedIndex(-1);
  }, []);

  useEffect(() => {
    onHistoryUpdate(handleHistoryUpdate);
  }, [handleHistoryUpdate]);

  const handleCopy = useCallback(() => {
    if (selectedIndex < 0 || selectedIndex >= history.length) return;
    const entry = history[selectedIndex];
    const text = `${entry.time}\nFrom ${entry.from} to ${entry.to}\n${entry.message}`;
    navigator.clipboard.writeText(text).catch(() => {});
  }, [selectedIndex, history]);

  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.key === "c") {
        handleCopy();
      }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [handleCopy]);

  return (
    <div className="p-4 flex flex-col gap-3" style={{ minHeight: "100%" }}>
      <div className="flex-1 overflow-y-auto border border-neutral-200 rounded-md" style={{ maxHeight: "400px" }}>
        {history.length === 0 ? (
          <div className="p-6 text-center text-sm text-neutral-400">
            No status changes recorded yet.
          </div>
        ) : (
          <table className="w-full text-xs">
            <thead className="bg-neutral-50 sticky top-0">
              <tr className="border-b border-neutral-200">
                <th className="text-left px-3 py-2 font-medium text-neutral-600">Time</th>
                <th className="text-left px-3 py-2 font-medium text-neutral-600">From</th>
                <th className="text-left px-3 py-2 font-medium text-neutral-600">To</th>
                <th className="text-left px-3 py-2 font-medium text-neutral-600">Message</th>
              </tr>
            </thead>
            <tbody>
              {history.map((entry, i) => (
                <tr
                  key={i}
                  className={`border-b border-neutral-100 cursor-pointer transition-colors ${
                    selectedIndex === i
                      ? "bg-neutral-900 text-white"
                      : "hover:bg-neutral-50"
                  }`}
                  onClick={() => setSelectedIndex(i)}
                >
                  <td className="px-3 py-1.5 whitespace-nowrap">{entry.time}</td>
                  <td className="px-3 py-1.5 whitespace-nowrap">{entry.from}</td>
                  <td className="px-3 py-1.5 whitespace-nowrap">{entry.to}</td>
                  <td className="px-3 py-1.5">{entry.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="flex justify-end gap-2">
        <Button variant="outline" size="sm" onClick={() => clearHistory()}>
          Clear
        </Button>
        <Button size="sm" onClick={() => closeDialog()}>
          Close
        </Button>
      </div>
    </div>
  );
}
