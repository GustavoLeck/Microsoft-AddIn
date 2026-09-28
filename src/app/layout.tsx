import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "BeInc",
  description: "BeInc add-in for Outlook",
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
};

const OFFICE_JS_URL = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

/*
 * Runs before office.js and protects browser APIs that office.js breaks:
 *
 * - history.pushState/replaceState: office.js sets them to null, which breaks
 *   the Next.js router. Null assignments are ignored; real functions (e.g.
 *   Next.js wrapping them) are still accepted. Same effect as Microsoft's
 *   cache/restore workaround, without needing a script after office.js.
 *   https://learn.microsoft.com/office/dev/add-ins/develop/connect-to-javascript-frameworks
 *
 * - String.prototype.startsWith/endsWith: in Outlook, office.js loads
 *   MicrosoftAjax, which replaces them with versions that ignore the position
 *   argument. That breaks Turbopack's chunk loader ("can't infer type of chunk
 *   from URL") and any library relying on the standard behaviour. The overrides
 *   are ignored; MicrosoftAjax's own calls work the same with the native ones.
 *
 * office.js must stay the LAST element in <head>: it injects its host script
 * (e.g. outlook-web-16.01.js) right after itself, and anything React rendered
 * after it would no longer match during hydration.
 */
const protectBrowserApis = `(function(){
function keep(o,n,accept){var f=o[n];Object.defineProperty(o,n,{configurable:true,enumerable:false,get:function(){return f},set:function(v){if(accept&&typeof v==="function")f=v}})}
keep(window.history,"pushState",true);keep(window.history,"replaceState",true);
keep(String.prototype,"startsWith",false);keep(String.prototype,"endsWith",false);
})();`;

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en">
      <head>
        <script dangerouslySetInnerHTML={{ __html: protectBrowserApis }} />
        {/* Office.js must come from Microsoft's CDN (never bundled) and load before the app. Keep it last. */}
        {/* eslint-disable-next-line @next/next/no-sync-scripts */}
        <script src={OFFICE_JS_URL} />
      </head>
      <body>{children}</body>
    </html>
  );
}
