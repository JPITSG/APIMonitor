import { useEffect, useRef, useState } from "react";
import { onInit, getInit, reportHeight, type InitData } from "./lib/bridge";
import ConfigView from "./ConfigView";
import HistoryView from "./HistoryView";

export default function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const [initData, setInitData] = useState<InitData | null>(null);

  useEffect(() => {
    onInit((data) => setInitData(data));
    getInit();
  }, []);

  // Report content height to C for window resizing
  useEffect(() => {
    const el = containerRef.current;
    if (!el || !initData) return;

    const report = () => {
      reportHeight(Math.ceil(el.scrollHeight));
    };

    // Fire immediately once content is rendered
    requestAnimationFrame(report);

    const observer = new ResizeObserver(report);
    observer.observe(el);
    return () => observer.disconnect();
  }, [initData]);

  if (!initData) return null;

  return (
    <div ref={containerRef}>
      {initData.view === "config" ? (
        <ConfigView config={initData.config!} />
      ) : (
        <HistoryView initialHistory={initData.history ?? []} />
      )}
    </div>
  );
}
