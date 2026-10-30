#!/usr/bin/env node
// Klickern: telefonen och hörlurarna som clicker till PowerPoint. Telefonen når datorn på samma wifi eller
// genom Tailscale.
//
//   npm start        som exe-filen, men direkt ur mappen
//   npm run bygg     bygger Klickern.exe (ett enfilsprogram med Node inuti)
//
// Programmet lyssnar på alla nätverkskort (port 4466), så att telefonen når det på samma wifi. Allt som
// telefonen behöver ligger under /k/<nyckel>/, och nyckeln är 128 slumpade bitar som bara QR-koden bär.
// Datorsidan, med QR-koden och Tailscale, har en egen port (4467) som bara lyssnar på 127.0.0.1.
// Tailscale lägger klickern på https://<dator>.<tailnet>.ts.net:8466 och pekar bara på telefonens port:
// Tailscale ansluter från datorn själv, så datorsidan får inte ligga på samma port. Mer: README.md här bredvid.

import { execFile, spawn } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import sea from "node:sea";
import { fileURLToPath } from "node:url";
import QRCode from "qrcode";
import { skapaKlicker } from "./klickern.mjs";

const PORT = Number(process.env.KLICKERN_PORT ?? 4466);
const TS_PORT = Number(process.env.KLICKERN_TS_PORT ?? 8466);
const DATOR_PORT = Number(process.env.KLICKERN_DATOR_PORT ?? PORT + 1);
const VERSION = "2026.10.30";
const mapp = path.join(process.env.APPDATA ?? os.homedir(), "Klickern");
fs.mkdirSync(mapp, { recursive: true });
// Windows letar först i arbetsmappen efter program. Dubbelklickar man på Klickern i Hämtade filer är det den
// mappen, så ett utplacerat cmd.exe där skulle köras. Därför en egen arbetsmapp och fullständiga sökvägar.
process.chdir(mapp);
const SYSTEM32 = path.join(process.env.SystemRoot ?? "C:\\Windows", "System32");

// I exe-filen ligger sidorna och bryggan som tillgångar, annars bredvid skriptet.
const iExe = sea.isSea();
const fil = (namn) => (iExe ? sea.getAsset(namn, "utf8") : fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), namn), "utf8"));

// PowerShell kör bara bryggan från en fil, och utan BOM läser PowerShell 5.1 å, ä och ö fel.
const bryggfil = path.join(mapp, "powerpoint.ps1");
fs.writeFileSync(bryggfil, `\ufeff${fil("powerpoint.ps1").replace(/^\ufeff/, "")}`);

function läsNyckel(ny = false) {
  const f = path.join(mapp, "nyckel.txt");
  if (!ny) {
    try {
      const k = fs.readFileSync(f, "utf8").trim();
      if (/^[A-Za-z0-9_-]{22}$/.test(k)) return k;
    } catch {}
  }
  const k = crypto.randomBytes(16).toString("base64url");
  fs.writeFileSync(f, k);
  return k;
}

let nyckel = läsNyckel();
const nyKlicker = () => skapaKlicker({ bas: `/k/${nyckel}`, brygga: bryggfil, telefon: () => fil("telefon.html"), logg: (...d) => console.log("  [powerpoint]", ...d) });
let klicker = nyKlicker();

// ── Adresserna ───────────────────────────────────────────────────────────

/** Datorns adresser på det lokala nätet. Virtuella kort och Tailscale hoppas över, wifi först. */
function wifiAdresser() {
  const ut = [];
  for (const [namn, lista] of Object.entries(os.networkInterfaces())) {
    if (/vethernet|virtual|vmware|vbox|hyper-v|tailscale|loopback|bluetooth|wsl|docker/i.test(namn)) continue;
    for (const a of lista ?? []) {
      if (a.family === "IPv4" && !a.internal && !a.address.startsWith("169.254.") && !a.address.startsWith("100.")) ut.push({ namn, adress: a.address });
    }
  }
  const wifi = (n) => /wi-?fi|wlan|trådlös|wireless/i.test(n);
  return ut.sort((a, b) => Number(wifi(b.namn)) - Number(wifi(a.namn)));
}

const TAILSCALE = [path.join(process.env.ProgramFiles ?? "C:\\Program Files", "Tailscale", "tailscale.exe"), "tailscale"];

/** Kör tailscale med första programmet som finns. */
async function ts(args, ms = 8000) {
  for (const exe of TAILSCALE) {
    const svar = await new Promise((klar) => {
      execFile(exe, args, { timeout: ms, windowsHide: true, encoding: "utf8" }, (fel, ut, err) => klar({ fel, ut: `${ut ?? ""}${err ?? ""}` }));
    });
    if (svar.fel?.code === "ENOENT") continue;
    return { ok: !svar.fel, ut: svar.ut, fanns: true };
  }
  return { ok: false, ut: "", fanns: false };
}

let tsCache = { tid: 0, lage: null };
async function tailscaleLage(färskt = false) {
  if (!färskt && tsCache.lage && Date.now() - tsCache.tid < 10_000) return tsCache.lage;
  const s = await ts(["status", "--json"]);
  let lage = { finns: s.fanns, igang: false };
  if (s.fanns) {
    try {
      const j = JSON.parse(s.ut.slice(s.ut.indexOf("{")));
      const dns = j.Self?.DNSName?.replace(/\.$/, "");
      lage = { finns: true, igang: j.BackendState === "Running", dns, url: dns ? `https://${dns}:${TS_PORT}` : null, aktiv: false };
      if (lage.igang) {
        const v = await ts(["serve", "status", "--json"]);
        try {
          lage.aktiv = Boolean(JSON.parse(v.ut.slice(v.ut.indexOf("{")) || "{}")?.TCP?.[String(TS_PORT)]);
        } catch {}
      }
    } catch {}
  }
  tsCache = { tid: Date.now(), lage };
  return lage;
}

/**
 * Lägger klickern på Tailscale Serve, bara för de egna enheterna. Har nätverket inte https påslaget
 * skriver tailscale ut en länk för att slå på det och väntar; länken skickas då vidare till datorsidan.
 */
function tailscalePå() {
  return new Promise((klar) => {
    const exe = fs.existsSync(TAILSCALE[0]) ? TAILSCALE[0] : TAILSCALE[1];
    const p = spawn(exe, ["serve", "--bg", `--https=${TS_PORT}`, `http://127.0.0.1:${PORT}`], { windowsHide: true });
    let ut = "";
    const avsluta = (svar) => {
      clearTimeout(tid);
      p.kill();
      tsCache.tid = 0;
      klar(svar);
    };
    const titta = (bit) => {
      ut += bit;
      const länk = ut.match(/https:\/\/login\.tailscale\.com\/\S+/)?.[0];
      if (länk) avsluta({ ok: false, länk });
    };
    p.stdout.on("data", titta);
    p.stderr.on("data", titta);
    p.on("error", () => avsluta({ ok: false, fel: "Tailscale hittades inte." }));
    p.on("exit", (kod) => avsluta(kod === 0 ? { ok: true } : { ok: false, fel: ut.trim() || `tailscale avslutades (${kod})` }));
    const tid = setTimeout(() => avsluta({ ok: false, fel: "Tailscale svarade inte." }), 20_000);
  });
}

async function tailscaleAv() {
  await ts(["serve", `--https=${TS_PORT}`, "off"]);
  tsCache.tid = 0;
}

const qrMinne = new Map();
async function qr(text) {
  if (!qrMinne.has(text)) qrMinne.set(text, await QRCode.toString(text, { type: "svg", margin: 0, errorCorrectionLevel: "M", color: { dark: "#131211", light: "#0000" } }));
  return qrMinne.get(text);
}

// ── Servern ──────────────────────────────────────────────────────────────

// Namnet stoppar DNS-rebinding: en webbplats som pekar sitt namn på 127.0.0.1 kommer inte åt datorsidan.
const lokalVärd = (req) => /^(127\.0\.0\.1|localhost)(:\d+)?$/i.test(req.headers.host ?? "");

function svara(res, status, typ, kropp) {
  res.writeHead(status, { "content-type": `${typ}; charset=utf-8`, "cache-control": "no-store", "referrer-policy": "no-referrer" });
  res.end(kropp);
}

async function datorLage() {
  const t = await tailscaleLage();
  const wifi = await Promise.all(wifiAdresser().map(async (a) => {
    const url = `http://${a.adress}:${PORT}/k/${nyckel}/`;
    return { ...a, url, qr: await qr(url) };
  }));
  const tailscale = { ...t };
  if (t.aktiv && t.url) {
    tailscale.telefon = `${t.url}/k/${nyckel}/`;
    tailscale.qr = await qr(tailscale.telefon);
  }
  let powerpoint = null;
  try {
    powerpoint = await klicker.läs();
  } catch {}
  return { version: VERSION, dator: os.hostname(), port: PORT, wifi, tailscale, powerpoint, telefoner: klicker.telefoner() };
}

/** Datorsidan och dess knappar. Knapparna kräver samma ursprung och klickerns rubrik, som telefonens steg. */
async function dator(req, res, url) {
  if (req.method === "GET" && url.pathname === "/") return svara(res, 200, "text/html", fil("dator.html").replace('"%VERSION%"', JSON.stringify(VERSION)));
  if (req.method === "GET" && url.pathname === "/dator/lage") return svara(res, 200, "application/json", JSON.stringify(await datorLage()));
  if (req.method !== "POST" || req.headers.origin !== `http://${req.headers.host}` || req.headers["x-klickern"] !== "1") return svara(res, 403, "application/json", '{"fel":"fel ursprung"}');
  if (url.pathname === "/dator/tailscale/pa") return svara(res, 200, "application/json", JSON.stringify(await tailscalePå()));
  if (url.pathname === "/dator/tailscale/av") {
    await tailscaleAv();
    return svara(res, 200, "application/json", '{"ok":true}');
  }
  if (url.pathname === "/dator/ny-nyckel") {
    klicker.stäng();
    nyckel = läsNyckel(true);
    klicker = nyKlicker();
    return svara(res, 200, "application/json", '{"ok":true}');
  }
  return svara(res, 404, "application/json", '{"fel":"finns inte"}');
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, "http://klickern");
    if (await klicker.hantera(req, res, url)) return;
    svara(res, 404, "text/plain", "Finns inte. Skanna QR-koden på datorn igen.");
  } catch (e) {
    console.error("  [klickern]", e?.stack ?? e);
    if (!res.headersSent) svara(res, 500, "text/plain", "Något gick fel i Klickern.");
  }
});

// Datorsidan visar nyckeln, så den har en egen server som bara lyssnar på den egna datorn.
const datorServer = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, "http://klickern");
    if (!lokalVärd(req) || !(url.pathname === "/" || url.pathname.startsWith("/dator/"))) return svara(res, 404, "text/plain", "Finns inte.");
    await dator(req, res, url);
  } catch (e) {
    console.error("  [klickern]", e?.stack ?? e);
    if (!res.headersSent) svara(res, 500, "text/plain", "Något gick fel i Klickern.");
  }
});

function öppna() {
  if (process.env.KLICKERN_INGEN_WEBBLASARE === "1") return;
  spawn(path.join(SYSTEM32, "cmd.exe"), ["/c", "start", "", `http://127.0.0.1:${DATOR_PORT}/`], { detached: true, stdio: "ignore", windowsHide: true }).unref();
}

const vidFel = (e) => {
  if (e.code === "EADDRINUSE") {
    console.log("\n  Klickern är redan igång. Datorsidan öppnas i webbläsaren.\n");
    öppna();
    setTimeout(() => process.exit(0), 1500);
  } else {
    console.error("  [klickern]", e?.stack ?? e);
    process.exit(1);
  }
};
server.on("error", vidFel);
datorServer.on("error", vidFel);

datorServer.listen(DATOR_PORT, "127.0.0.1", () => {
  server.listen(PORT, "0.0.0.0", () => {
    console.log(`
  Klickern ${VERSION}
  Telefonen som clicker till PowerPoint.

  Datorsidan med QR-koden:  http://127.0.0.1:${DATOR_PORT}/
  Låt fönstret vara öppet medan du föreläser. Stäng det för att stänga av Klickern.
`);
    öppna();
  });
});

const slut = () => {
  klicker.stäng();
  process.exit(0);
};
process.on("SIGINT", slut);
process.on("SIGTERM", slut);
