// Copies manifest.xml into out/ after `next build`.
// If ADDIN_URL is set (e.g. https://beinc.z13.web.core.windows.net), every
// https://localhost:3000 in the manifest is replaced with it — same idea as
// urlDev/urlProd in the old webpack config.
import { mkdir, readFile, writeFile } from "node:fs/promises";

const DEV_URL = "https://localhost:3000";
const prodUrl = process.env.ADDIN_URL?.trim().replace(/\/+$/, "");

const manifest = await readFile("manifest.xml", "utf8");
const output = prodUrl ? manifest.replaceAll(DEV_URL, prodUrl) : manifest;

await mkdir("out", { recursive: true });
await writeFile("out/manifest.xml", output);

console.log(
  prodUrl
    ? `out/manifest.xml -> ${prodUrl}`
    : "out/manifest.xml keeps https://localhost:3000 (set ADDIN_URL for a production manifest)",
);
