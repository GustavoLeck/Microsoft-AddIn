# BeInc — Outlook add-in (Next.js + React)

Task pane add-in for Outlook, built with Next.js (App Router, static export), React 19 and Fluent UI v9.
The add-in runs in **compose mode**: its buttons show up while writing an email.

## Requirements

- Node.js LTS
- Outlook with a Microsoft account (new Outlook for Windows, Outlook on the web, or classic Outlook with WebView2)

## Scripts

| Command | What it does |
| --- | --- |
| `npm run dev` (or `npm start`) | Dev server on **https://localhost:3000** with the Office Add-in dev certificates |
| `npm run build` | Static export to `out/` + `out/manifest.xml` (see *Production*) |
| `npm run lint` / `npm run typecheck` | ESLint / TypeScript |
| `npm run validate` | Validates `manifest.xml` with Microsoft's validator |
| `npm run certs:install` | (Re)installs the localhost dev certificates |

## Sideload in Outlook (development)

1. `npm run dev` and keep the terminal open.
2. Open https://aka.ms/olksideload → **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from File** → pick `manifest.xml`.
3. Write a new email → **Apps** → **BeInc** → **Show Task Pane**.

Opening https://localhost:3000/taskpane/ in a normal browser only shows a notice: the pane needs to run inside Outlook.

## Project layout

```
manifest.xml                 Add-in manifest (add-in only / XML)
public/
  assets/                    Icons referenced by the manifest
  commands.html, commands.js Function file for ribbon commands ("Perform an action")
scripts/
  dev.mjs                    next dev over HTTPS with office-addin-dev-certs
  build-manifest.mjs         Copies manifest to out/, swapping localhost for ADDIN_URL
src/
  app/layout.tsx             Loads office.js (+ History API workaround)
  app/taskpane/              Task pane route (/taskpane/)
  components/                React UI (Fluent UI)
  lib/office.ts              useOffice() / useDarkMode() hooks
  lib/mailbox.ts             Mailbox helpers (getItemSubject)
```

### Notes

- **office.js** is loaded from Microsoft's CDN in `layout.tsx`, never bundled, and must stay the last
  element in `<head>` (it injects its host script right after itself). A small script before it protects
  browser APIs office.js breaks: `history.pushState/replaceState` (nulled by office.js, needed by the
  Next.js router) and `String.prototype.startsWith/endsWith` (replaced by MicrosoftAjax in Outlook,
  which breaks Turbopack's chunk loader).
- The task pane is rendered **client-side only** (`next/dynamic` with `ssr: false`), since it depends on the `Office` global.
- The function file (`public/commands.*`) is plain HTML/JS on purpose: Outlook calls the registered
  functions right after the page loads, before React would hydrate.
- The theme follows Office's theme when available, otherwise the system color scheme.

## Production

The build is a static site: host `out/` on any HTTPS static host (e.g. Azure Storage static website).

```powershell
$env:ADDIN_URL = "https://<your-host>"
npm run build
```

`out/manifest.xml` then points to that host. Install that manifest for production users.
