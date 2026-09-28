// Starts `next dev` on https://localhost:3000 using the Office Add-in dev
// certificates (~/.office-addin-dev-certs). Outlook only loads add-ins over
// trusted HTTPS, so this replaces the webpack dev server from the Yeoman template.
import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import devCerts from "office-addin-dev-certs";

const require = createRequire(import.meta.url);
const port = process.env.PORT ?? "3000";

// Generates and trusts the certificates on first run (Windows may ask for confirmation).
await devCerts.ensureCertificatesAreInstalled();

const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const nextBin = require.resolve("next/dist/bin/next");

const args = [
  nextBin,
  "dev",
  "--port",
  port,
  "--experimental-https",
  "--experimental-https-key",
  path.join(certDir, "localhost.key"),
  "--experimental-https-cert",
  path.join(certDir, "localhost.crt"),
  "--experimental-https-ca",
  path.join(certDir, "ca.crt"),
];

const child = spawn(process.execPath, args, { stdio: "inherit" });
child.on("exit", (code) => process.exit(code ?? 0));
for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}
