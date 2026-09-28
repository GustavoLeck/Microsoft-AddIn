"use client";

import dynamic from "next/dynamic";

/*
 * The task pane only makes sense inside Outlook (it needs the Office global),
 * so it is rendered on the client only. The static HTML just carries this
 * lightweight placeholder until the app loads.
 */
const TaskpaneApp = dynamic(() => import("@/components/TaskpaneApp"), {
  ssr: false,
  loading: () => (
    <p style={{ fontFamily: "Segoe UI, sans-serif", padding: 20 }}>Loading…</p>
  ),
});

export default function TaskpaneLoader() {
  return <TaskpaneApp />;
}
