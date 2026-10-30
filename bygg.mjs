#!/usr/bin/env node
// Bygger Klickern.exe: ett enfilsprogram (Nodes single executable application) med sidorna och bryggan
// inbakade. Mottagaren behöver alltså varken Node eller något annat installerat.
//
//   npm run bygg
//
// Resultatet hamnar i dist/ (står utanför git). esbuild, postject och resedit-cli hämtas med npx när
// skriptet körs och blir inga beroenden i projektet. Programmet är inte kodsignerat, så Windows SmartScreen
// varnar första gången. Mer: README.md här bredvid.

import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const här = path.dirname(fileURLToPath(import.meta.url));
const dist = path.join(här, "dist");
const exe = path.join(dist, "Klickern.exe");
const npx = process.platform === "win32" ? "npx.cmd" : "npx";

const kör = (kommando, args, val = {}) => execFileSync(kommando, args, { stdio: "inherit", shell: kommando.endsWith(".cmd"), cwd: här, ...val });

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(dist, { recursive: true });

console.log("[bygg] paketerar programmet");
kör(npx, ["--yes", "esbuild@0.25.11", "program.mjs", "--bundle", "--platform=node", "--format=cjs", "--target=node24", "--outfile=dist/klickern.cjs", "--log-level=warning", "--define:import.meta.url=undefined"]);

console.log("[bygg] bakar in sidorna och bryggan");
const konfig = {
  main: path.join(dist, "klickern.cjs"),
  output: path.join(dist, "klickern.blob"),
  disableExperimentalSEAWarning: true,
  useCodeCache: false,
  useSnapshot: false,
  assets: Object.fromEntries(["telefon.html", "dator.html", "powerpoint.ps1"].map((f) => [f, path.join(här, f)])),
};
fs.writeFileSync(path.join(dist, "sea.json"), JSON.stringify(konfig, null, 2));
kör(process.execPath, ["--experimental-sea-config", path.join(dist, "sea.json")]);

console.log("[bygg] skapar Klickern.exe");
// Programmet först och ikonen sist. I omvänd ordning klagar postject på relokeringstabellen.
const kopia = path.join(dist, "node.exe");
fs.copyFileSync(process.execPath, kopia);
kör(npx, ["--yes", "postject@1.0.0-alpha.6", "dist/node.exe", "NODE_SEA_BLOB", "dist/klickern.blob", "--sentinel-fuse", "NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2", "--overwrite"]);
kör(npx, ["--yes", "resedit-cli@3.1.1", "--in", "dist/node.exe", "--out", "dist/Klickern.exe", "--ignore-signed", "--icon", "1,klickern.ico",
  "--product-name", "Klickern", "--file-description", "Klickern", "--original-filename", "Klickern.exe", "--internal-name", "Klickern",
  "--product-version", "2026.10.30.0", "--file-version", "2026.10.30.0"]);
fs.rmSync(kopia);

for (const f of ["klickern.cjs", "klickern.blob", "sea.json"]) fs.rmSync(path.join(dist, f), { force: true });
console.log(`[bygg] klart: ${exe} (${Math.round(fs.statSync(exe).size / 1024 / 1024)} MB)`);
