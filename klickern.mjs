#!/usr/bin/env node
// Klickerns server: telefonsidan, läget som ström och stegen. Monteras under en bas av program.mjs.
//
//   node klickern.mjs [port]   servern ensam på 127.0.0.1 (prov och felsökning)
//
// Telefonen läser läget som en ström (GET <bas>/lage) och skickar steg med POST <bas>/steg
// (fram, bak, start, svart, bilder eller ga N som text). bilder svarar också med avsnitten och
// bildernas rubriker, till menyn. Bryggan är en PowerShell-process som startar med
// första telefonen och stängs två minuter efter den sista.

import { spawn } from "node:child_process";
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Mappen med bryggan och telefonsidan. Programmet (program.mjs) ger i stället egna sökvägar, eftersom
// filerna där ligger inbakade i exe-filen.
const här = () => path.dirname(fileURLToPath(import.meta.url));
// Fullständig sökväg: Windows letar annars först i arbetsmappen, och där kan ett annat program ligga.
const POWERSHELL = path.join(process.env.SystemRoot ?? "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
const STEG = new Set(["fram", "bak", "start", "svart", "bilder"]);
// ga N: hoppa till bild N från menyn med avsnitten.
const tillåtet = (s) => STEG.has(s) || /^ga \d{1,4}$/.test(s);
const TAKT = 400; // ms mellan avläsningarna medan en telefon är ansluten
const TYST_EFTER = 120_000;
// Så länge PowerPoint får på sig att svara innan bryggan startas om. Att starta bildspelet och att
// läsa rubrikerna i en stor presentation kan ta tid.
const TÅLAMOD = { långsamt: 20_000, annat: 8_000 };
const långsamt = (kommando) => kommando === "start" || kommando === "bilder" || kommando.startsWith("ga ");

export function skapaKlicker({
  bas = "/_powerpoint",
  logg = (...delar) => console.log("[klickern]", ...delar),
  brygga: bryggfil = path.join(här(), "powerpoint.ps1"),
  telefon = () => fs.readFileSync(path.join(här(), "telefon.html"), "utf8"),
} = {}) {
  let ps = null;
  let kö = [];
  let vakt = null;
  const telefoner = new Set();
  let senast = "";
  let takt = null;
  let tystnad = null;

  const vakta = () => {
    clearTimeout(vakt);
    const först = kö[0];
    if (!först) return;
    const p = ps;
    vakt = setTimeout(() => {
      logg(`PowerPoint svarade inte på ${först.kommando}; bryggan startas om`);
      p?.kill();
    }, långsamt(först.kommando) ? TÅLAMOD.långsamt : TÅLAMOD.annat);
  };

  function brygga() {
    if (ps) return ps;
    const p = spawn(POWERSHELL, ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", bryggfil], {
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });
    ps = p;
    let rest = "";
    p.stdout.setEncoding("utf8");
    p.stdout.on("data", (bit) => {
      rest += bit;
      for (let i; (i = rest.indexOf("\n")) >= 0; ) {
        const rad = rest.slice(0, i).trim();
        rest = rest.slice(i + 1);
        if (!rad || ps !== p) continue;
        let lage;
        try {
          lage = JSON.parse(rad);
        } catch {
          lage = { lage: "fel", fel: rad };
        }
        kö.shift()?.klar(lage);
        vakta();
      }
    });
    p.stderr.setEncoding("utf8");
    p.stderr.on("data", (d) => logg(d.trim()));
    p.stdin.on("error", () => {});
    const slut = (fel) => {
      if (ps !== p) return;
      ps = null;
      clearTimeout(vakt);
      const kvar = kö;
      kö = [];
      for (const v of kvar) v.klar({ lage: "fel", fel });
    };
    p.on("exit", (kod) => slut(`bryggan till PowerPoint stängdes (${kod})`));
    p.on("error", (e) => slut(`bryggan till PowerPoint startade inte: ${e.message}`));
    return p;
  }

  /** Ett kommando till bryggan. Svaret är läget efter kommandot. */
  function fråga(kommando) {
    return new Promise((klar) => {
      const p = brygga();
      kö.push({ kommando, klar });
      if (kö.length === 1) vakta();
      p.stdin.write(`${kommando}\n`);
    });
  }

  function sänd(lage) {
    const { svar: _svar, tangent: _tangent, innehall: _innehall, ...rest } = lage;
    const json = JSON.stringify(rest);
    if (json === senast) return;
    senast = json;
    for (const res of telefoner) res.write(`data: ${json}\n\n`);
  }

  // Läget läses om hela tiden, så att telefonen följer med också när någon klickar på datorn.
  const avläs = () => {
    if (!kö.length) void fråga("lage").then(sänd);
  };

  const somna = () => {
    clearTimeout(tystnad);
    if (!telefoner.size) tystnad = setTimeout(() => ps?.stdin.end(), TYST_EFTER);
  };

  function lyssna(req, res) {
    res.writeHead(200, { "content-type": "text/event-stream; charset=utf-8", "cache-control": "no-store, no-transform", "x-accel-buffering": "no" });
    res.write("retry: 2000\n\n");
    if (senast) res.write(`data: ${senast}\n\n`);
    telefoner.add(res);
    clearTimeout(tystnad);
    takt ??= setInterval(avläs, TAKT);
    avläs();
    const puls = setInterval(() => res.write(": puls\n\n"), 15_000);
    req.on("close", () => {
      clearInterval(puls);
      telefoner.delete(res);
      if (telefoner.size) return;
      clearInterval(takt);
      takt = null;
      somna();
    });
  }

  const json = (res, status, data) => {
    res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
    res.end(JSON.stringify(data));
  };

  async function steg(req, res) {
    let kropp = "";
    for await (const bit of req) {
      kropp += bit;
      if (kropp.length > 64) return json(res, 413, { fel: "för långt" });
    }
    const s = kropp.trim();
    if (!tillåtet(s)) return json(res, 400, { fel: "okänt steg" });
    clearTimeout(tystnad);
    const lage = await fråga(s);
    sänd(lage);
    somna();
    json(res, 200, lage);
  }

  function sida(res) {
    const html = telefon()
      .replace('"%BAS%"', JSON.stringify(bas))
      .replace('"%DATOR%"', JSON.stringify(os.hostname()).replace(/</g, "\\u003c"));
    res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store", "referrer-policy": "no-referrer" });
    res.end(html);
  }

  /** Svarar på klickerns adresser och ger false för allt annat. */
  async function hantera(req, res, url = new URL(req.url, "http://klickern")) {
    if (url.pathname !== bas && !url.pathname.startsWith(`${bas}/`)) return false;
    // Bara sidan själv får skicka steg, aldrig en annan webbplats i samma webbläsare. Ett steg måste
    // ha sidans eget ursprung och klickerns rubrik. Rubriken gör att webbläsaren först frågar om lov
    // innan en annan sida får skicka, och den frågan besvaras aldrig. Ett anrop utan ursprung, eller
    // med ursprunget null, stoppas också (säkerhetsgranskningen 30 oktober 2026).
    const origin = req.headers.origin;
    let värd = null;
    try {
      värd = origin ? new URL(origin).host : null;
    } catch {}
    const främmande = origin !== undefined && värd !== req.headers.host;
    const oäkta = req.method !== "GET" && (värd !== req.headers.host || req.headers["x-klickern"] !== "1");
    if (främmande || oäkta) {
      json(res, 403, { fel: "fel ursprung" });
      return true;
    }
    const del = url.pathname.slice(bas.length) || "/";
    if (req.method === "GET" && (del === "/" || del === "")) sida(res);
    else if (req.method === "GET" && del === "/lage") lyssna(req, res);
    else if (req.method === "POST" && del === "/steg") await steg(req, res);
    else json(res, 404, { fel: "finns inte" });
    return true;
  }

  function stäng() {
    clearInterval(takt);
    clearTimeout(tystnad);
    for (const res of telefoner) res.end();
    telefoner.clear();
    ps?.kill();
  }

  return {
    bas,
    hantera,
    stäng,
    /** Läser av PowerPoint nu, till programmets datorsida. Bryggan somnar som vanligt när ingen frågar. */
    läs: async () => {
      clearTimeout(tystnad);
      const l = kö.length && senast ? JSON.parse(senast) : await fråga("lage");
      sänd(l);
      somna();
      return l;
    },
    /** Hur många telefoner som är anslutna. */
    telefoner: () => telefoner.size,
  };
}

const direkt = Boolean(import.meta.url) && process.argv[1] && path.resolve(process.argv[1]).toLowerCase() === fileURLToPath(import.meta.url).toLowerCase();
if (direkt) {
  const port = Number(process.argv[2] ?? 4466);
  const klicker = skapaKlicker();
  const värdar = new Set([`127.0.0.1:${port}`, `localhost:${port}`]);
  http
    .createServer(async (req, res) => {
      try {
        if (!värdar.has(req.headers.host ?? "")) {
          res.writeHead(403);
          return res.end();
        }
        const url = new URL(req.url, "http://klickern");
        if (url.pathname === "/") {
          res.writeHead(302, { location: `${klicker.bas}/` });
          return res.end();
        }
        if (!(await klicker.hantera(req, res, url))) {
          res.writeHead(404);
          res.end();
        }
      } catch (e) {
        console.error("[klickern]", e?.stack ?? e);
        if (!res.headersSent) res.writeHead(500);
        res.end();
      }
    })
    .listen(port, "127.0.0.1", () => console.log(`[klickern] http://127.0.0.1:${port}${klicker.bas}/`));
}
