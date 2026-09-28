"use client";

import { useEffect, useState } from "react";

export type OfficeState =
  /** Waiting for Office.onReady. */
  | { status: "loading" }
  /** Office.js is ready. host is null when the page is opened outside Office (e.g. a plain browser tab). */
  | { status: "ready"; host: Office.HostType | null }
  /** office.js never loaded (offline, blocked CDN…). */
  | { status: "unavailable" };

/** Resolves once Office.js is ready and reports which host (if any) the page runs in. */
export function useOffice(): OfficeState {
  const [state, setState] = useState<OfficeState>(() =>
    typeof Office === "undefined" ? { status: "unavailable" } : { status: "loading" },
  );

  useEffect(() => {
    if (typeof Office === "undefined") return;
    let cancelled = false;
    Office.onReady().then((info) => {
      if (!cancelled) setState({ status: "ready", host: info.host ?? null });
    });
    return () => {
      cancelled = true;
    };
  }, []);

  return state;
}

/** Office theme when the host exposes one; undefined otherwise. */
function isOfficeThemeDark(): boolean | undefined {
  try {
    const background = Office.context?.officeTheme?.bodyBackgroundColor;
    const hex = background?.replace("#", "");
    if (!hex || hex.length < 6) return undefined;
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    return (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
  } catch {
    return undefined;
  }
}

/**
 * Whether the task pane should use the dark theme: follows the Office theme
 * when available, otherwise the system/WebView color scheme.
 */
export function useDarkMode(officeReady: boolean): boolean {
  const [systemDark, setSystemDark] = useState(
    () => window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false,
  );

  useEffect(() => {
    const media = window.matchMedia?.("(prefers-color-scheme: dark)");
    if (!media) return;
    const onChange = (event: MediaQueryListEvent) => setSystemDark(event.matches);
    media.addEventListener("change", onChange);
    return () => media.removeEventListener("change", onChange);
  }, []);

  const officeDark = officeReady ? isOfficeThemeDark() : undefined;
  return officeDark ?? systemDark;
}
