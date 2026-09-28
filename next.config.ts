import type { NextConfig } from "next";

/**
 * The add-in is a static site: Outlook only loads HTML/JS from a URL,
 * so we export to plain files (out/) that can be hosted anywhere with HTTPS.
 *
 * trailingSlash makes every route a folder with an index.html
 * (out/taskpane/index.html), so the same URL — https://<host>/taskpane/ —
 * works in `next dev` and on any static host.
 */
const nextConfig: NextConfig = {
  output: "export",
  trailingSlash: true,
  images: { unoptimized: true },
  // Don't let `next dev` write AGENTS.md / CLAUDE.md into the repo.
  agentRules: false,
};

export default nextConfig;
