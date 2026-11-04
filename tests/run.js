#!/usr/bin/env node
/*
 * Testsuite voor de doublurecontrole (index.html).
 *
 * Gebruik (eenmalig installeren, daarna testen):
 *   cd tests
 *   npm install
 *   npx playwright install chromium
 *   npm test                 # alle tests behalve het grote bestand
 *   npm test -- --groot      # ook een test met twee bestanden van 130 MB
 *
 * De tests maken testmappen aan in een tijdelijke map van het besturingssysteem en ruimen die
 * daarna weer op. Er wordt niets in de repository geschreven.
 */
"use strict";
const { chromium } = require("playwright");
const fs = require("fs"), path = require("path"), os = require("os"), crypto = require("crypto"), zlib = require("zlib");
const { pathToFileURL } = require("url");

const ROOT = path.resolve(__dirname, "..");
const HTML_PAD = path.join(ROOT, "index.html");
const HTML_URL = pathToFileURL(HTML_PAD).href;
const BRON = fs.readFileSync(HTML_PAD, "utf8");
const SCRIPT = BRON.slice(BRON.indexOf("<script>") + 8, BRON.lastIndexOf("</script>"));
const VERSIE = (SCRIPT.match(/const VERSIE = "([^"]+)"/) || [])[1];
const MET_GROOT = process.argv.includes("--groot");
const WERK = fs.mkdtempSync(path.join(os.tmpdir(), "doucon-test-"));

/* ---------- Mini-testraamwerk ---------- */
const resultaten = [];
let huidige = null;
function ok(voorwaarde, melding) { if (!voorwaarde) throw new Error(melding); }
function gelijk(werkelijk, verwacht, melding) {
  if (JSON.stringify(werkelijk) !== JSON.stringify(verwacht))
    throw new Error(`${melding}\n      verwacht: ${JSON.stringify(verwacht)}\n      werkelijk: ${JSON.stringify(werkelijk)}`);
}
async function test(naam, fn) {
  huidige = naam;
  const t0 = Date.now();
  try { await fn(); resultaten.push({ naam, ok: true }); console.log(`  ✓ ${naam} (${Date.now() - t0} ms)`); }
  catch (e) { resultaten.push({ naam, ok: false, fout: e }); console.log(`  ✗ ${naam}\n      ${String(e.message || e).split("\n").join("\n      ")}`); }
}

/* ---------- Hulpfuncties: bestanden en Office-bestanden ---------- */
function schrijf(rel, inhoud, mtimeMin) {
  const p = path.join(WERK, rel);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  fs.writeFileSync(p, inhoud);
  if (mtimeMin !== undefined) { const t = new Date(Date.UTC(2024, 0, 1) + mtimeMin * 60000); fs.utimesSync(p, t, t); }
  return p;
}
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = b => { let c = 0xFFFFFFFF; for (const x of b) c = CRC[(c ^ x) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
// Minimale ZIP-schrijver met deflate (methode 8), zodat ook het uitpakken van gecomprimeerde Office-bestanden wordt getest
function zip(bestanden) {
  const delen = [], centraal = []; let off = 0;
  for (const [naam, tekst] of Object.entries(bestanden)) {
    const n = Buffer.from(naam), data = Buffer.from(tekst), comp = zlib.deflateRawSync(data), crc = crc32(data);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(n.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(n.length, 28); ch.writeUInt32LE(off, 42);
    delen.push(lh, n, comp); centraal.push(ch, n); off += 30 + n.length + comp.length;
  }
  const cd = Buffer.concat(centraal), eo = Buffer.alloc(22);
  eo.writeUInt32LE(0x06054b50, 0); eo.writeUInt16LE(centraal.length / 2, 8); eo.writeUInt16LE(centraal.length / 2, 10); eo.writeUInt32LE(cd.length, 12); eo.writeUInt32LE(off, 16);
  return Buffer.concat([...delen, cd, eo]);
}
const docx = alineas => zip({
  "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  "word/document.xml": `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${alineas.map(a => `<w:p><w:r><w:t>${a}</w:t></w:r></w:p>`).join("")}</w:body></w:document>`
});
// Minimale Excel-werkmap; cellen als { ref: "<c ...>...</c>" }, met optioneel 1904-datumsysteem
const xlsx = (cellen, is1904) => zip({
  "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  "xl/workbook.xml": `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${is1904 ? '<workbookPr date1904="1"/>' : ""}<sheets><sheet name="Blad1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  "xl/_rels/workbook.xml.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  "xl/styles.xml": '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
  "xl/worksheets/sheet1.xml": `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${cellen.join("")}</row></sheetData></worksheet>`
});
// Leest de bestandsnamen uit een ZIP (voor de controle van de Excel- en Word-exports)
function zipInhoud(buf) {
  ok(buf.readUInt32LE(0) === 0x04034b50, "bestand begint niet met een ZIP-kop");
  const uit = {}; let p = 0;
  while (buf.readUInt32LE(p) === 0x04034b50) {
    const methode = buf.readUInt16LE(p + 8), cl = buf.readUInt32LE(p + 18), nl = buf.readUInt16LE(p + 26), xl = buf.readUInt16LE(p + 28);
    const naam = buf.slice(p + 30, p + 30 + nl).toString(), data = buf.slice(p + 30 + nl + xl, p + 30 + nl + xl + cl);
    uit[naam] = (methode === 8 ? zlib.inflateRawSync(data) : data).toString("utf8");
    p += 30 + nl + xl + cl;
  }
  return uit;
}

/* ---------- Functies uit index.html laden voor eenheidstests ---------- */
function haalFunctie(naam) {
  const start = SCRIPT.indexOf(`function ${naam}(`); ok(start >= 0, `functie ${naam} niet gevonden`);
  return SCRIPT.slice(start, SCRIPT.indexOf("\n}\n", start) + 2);
}
function haalRegel(naam) { const a = SCRIPT.indexOf(`function ${naam}(`); ok(a >= 0, `functie ${naam} niet gevonden`); return SCRIPT.slice(a, SCRIPT.indexOf("\n", a) + 1); }
function haalTussen(begin, eind) { const a = SCRIPT.indexOf(begin), b = SCRIPT.indexOf(eind, a); ok(a >= 0 && b > a, `codeblok '${begin}' niet gevonden`); return SCRIPT.slice(a, b); }

/* ---------- Testmappen ---------- */
function maakBasismap() {
  const m = "Testmap";
  schrijf(`${m}/rapport.txt`, "hallo wereld", 0);
  schrijf(`${m}/a/rapport.txt`, "hallo wereld", 10);
  schrijf(`${m}/c/Rapport.TXT`, "hallo wereld", 20);
  schrijf(`${m}/a/b/rapport.txt`, "andere inhoud", 30);
  schrijf(`${m}/uniek.txt`, "uniek", 0);
  schrijf(`${m}/a/kopie van iets.txt`, "hallo wereld", 40);
  schrijf(`${m}/Jaarverslag (2023).txt`, "Jaar 2023", 0);
  schrijf(`${m}/Jaarverslag (2024).txt`, "Jaar 2024", 0);
  schrijf(`${m}/(1).txt`, "een", 0);
  schrijf(`${m}/(2).txt`, "twee", 0);
  schrijf(`${m}/leeg1.txt`, "", 0); schrijf(`${m}/a/leeg1.txt`, "", 5);
  schrijf(`${m}/Thumbs.db`, "x"); schrijf(`${m}/~$rapport.docx`, "x"); schrijf(`${m}/.~lock.rapport.docx#`, "x"); schrijf(`${m}/~WRL0001.tmp`, "x");
  schrijf(`${m}/notulen.docx`, docx(["Eerste alinea", "Tweede alinea", "Derde"]), 0);
  schrijf(`${m}/c/notulen.docx`, docx(["Eerste alinea", "Tweede alinea aangepast", "Derde", "Vierde"]), 50);
  schrijf(`${m}/Archief/oud.txt`, "archief", 0);
  schrijf(`${m}/oud.txt`, "archief", 60);
  return path.join(WERK, m);
}
function maakGroteMap(n) {
  const m = "Veel";
  for (let i = 0; i < n; i++) { schrijf(`${m}/f${i}.txt`, `inhoud ${i}`, 0); schrijf(`${m}/kopie/f${i}.txt`, `inhoud ${i}`, 10); }
  return path.join(WERK, m);
}

/* ---------- Browserhulpen ---------- */
let browser, paginafouten = [];
async function nieuwePagina(viewport) {
  const p = await browser.newPage({ acceptDownloads: true, viewport: viewport || { width: 1400, height: 900 } });
  p.on("pageerror", e => paginafouten.push(`${huidige}: ${e.message}`));
  p.on("console", m => { if (m.type() === "error") paginafouten.push(`${huidige}: ${m.text()}`); });
  await p.goto(HTML_URL);
  await p.evaluate(() => { window.showSaveFilePicker = undefined; });   // opslaan via download, zonder dialoog
  return p;
}
async function zetOpties(p, o) {
  await p.evaluate(o => {
    for (const [id, v] of Object.entries(o)) {
      const el = document.getElementById(id);
      if (el.type === "checkbox") el.checked = v; else el.value = v;
      el.dispatchEvent(new Event("input", { bubbles: true }));
    }
  }, o);
}
async function wachtKlaar(p) {
  await p.waitForFunction(() => document.querySelector("#stop").disabled && document.querySelector("#resultaat").style.display !== "none", null, { timeout: 120000 });
}
async function analyseer(p, map) { await p.setInputFiles("#kiezer", map); await wachtKlaar(p); }
const tegels = p => p.$$eval(".kaart", ks => Object.fromEntries(ks.map(k => [k.querySelector(".lbl").textContent, k.querySelector(".getal").textContent])));
async function groepen(p, tab) {
  await p.click(`.tab[data-tab="${tab}"]`);
  return p.$$eval("details.groep .gnaam", g => g.map(x => x.textContent));
}
const status = p => p.$eval("#exportStatus", e => e.textContent);
async function download(p, knop) {
  const [dl] = await Promise.all([p.waitForEvent("download"), p.click("#" + knop)]);
  const doel = path.join(WERK, "downloads", dl.suggestedFilename());
  fs.mkdirSync(path.dirname(doel), { recursive: true });
  await dl.saveAs(doel);
  return doel;
}

/* ---------- Tests ---------- */
(async () => {
  console.log(`Doublurecontrole ${VERSIE}: tests (werkmap ${WERK})\n`);

  console.log("Eenheidstests");
  await test("SHA-256 in blokken is gelijk aan Node crypto (randgevallen in grootte en blokgrens)", () => {
    const sha256Stroom = new Function(haalTussen("const hex =", "// Hashen in drie stappen") + "; return sha256Stroom;")();
    for (const n of [0, 1, 55, 56, 63, 64, 65, 119, 120, 1000, 65536, 1048583])
      for (const stuk of [1, 7, 64, 4096, 1 << 20]) {
        const buf = crypto.randomBytes(n), h = sha256Stroom();
        for (let i = 0; i < n; i += stuk) h.update(new Uint8Array(buf.subarray(i, i + stuk)));
        gelijk(h.digest(), crypto.createHash("sha256").update(buf).digest("hex"), `hash wijkt af bij ${n} bytes in stukken van ${stuk}`);
      }
  });
  await test("Naamvergelijking: kopie-aanduidingen, jaartallen, lege namen en Unicode", () => {
    const naamSleutel = new Function(haalFunctie("splitExt") + haalFunctie("naamSleutel") + "; return naamSleutel;")();
    const o = { kopie: true, extensie: true, hoofdletter: false };
    const gevallen = {
      "Rapport (1).docx": "rapport.docx", "Rapport - Kopie (2).docx": "rapport.docx", "Kopie van Rapport.docx": "rapport.docx",
      "Copy of Rapport.docx": "rapport.docx", "Rapport kopie.docx": "rapport.docx", "Rapport kopie 2.docx": "rapport.docx",
      "Jaarverslag (2023).pdf": "jaarverslag (2023).pdf", "(1).txt": "(1).txt", "Offerte (v2).docx": "offerte (v2).docx",
      "café.txt": "café.txt"
    };
    for (const [in_, uit] of Object.entries(gevallen)) gelijk(naamSleutel(in_, o), uit, `naamSleutel("${in_}")`);
    gelijk(naamSleutel("Rapport (1).docx", { kopie: false, extensie: true, hoofdletter: false }), "rapport (1).docx", "zonder kopie-optie blijft (1) staan");
    gelijk(naamSleutel("Rapport.DOCX", { kopie: false, extensie: false, hoofdletter: true }), "Rapport", "hoofdlettergevoelig zonder extensie");
  });
  await test("Uitsluitlijst: mapnaam, pad, extensie en jokerteken", () => {
    const f = new Function(haalFunctie("splitExt") + haalTussen("function parseUitsluiten", "/* ---------- ") + "; return { parseUitsluiten, maakUitsluiter };")();
    const u = f.maakUitsluiter(f.parseUitsluiten("Archief; Projecten/Oud, .tmp, Concept*"));
    const x = pad => ({ pad, naam: pad.split("/").pop() });
    gelijk(u(x("Top/Archief/a.txt")), true, "map Archief op elk niveau");
    gelijk(u(x("Top/a/archief/b.txt")), true, "hoofdletters maken niet uit");
    gelijk(u(x("Top/Projecten/Oud/a.txt")), true, "pad Projecten/Oud");
    gelijk(u(x("Top/Projecten/Oude dingen/a.txt")), false, "Projecten/Oude dingen valt er niet onder");
    gelijk(u(x("Top/a/werk.TMP")), true, "extensie .tmp");
    gelijk(u(x("Top/Concepten/a.txt")), true, "jokerteken Concept*");
    gelijk(u(x("Top/Archief.txt")), false, "een bestand met de naam Archief.txt valt er niet onder");
    gelijk(f.maakUitsluiter(f.parseUitsluiten("  ")), null, "lege lijst sluit niets uit");
  });

  await test("Gedeelde Excel-formules: celverwijzingen verschuiven", () => {
    const verschuif = new Function(haalRegel("refDelen") + haalRegel("kolLetter") + haalFunctie("verschuifFormule") + "; return verschuifFormule;")();
    gelijk(verschuif("A2*2", "B2", "B5"), "A5*2", "relatieve rij");
    gelijk(verschuif("SUM(A2:C2)", "D2", "E3"), "SUM(B3:D3)", "bereik, rij en kolom");
    gelijk(verschuif("$A$1+A$1+$A1", "B2", "C4"), "$A$1+B$1+$A3", "absolute delen blijven staan");
    gelijk(verschuif('IF(A2="B2";Blad1!A2;LOG10(A2))', "B2", "B3"), 'IF(A3="B2";Blad1!A3;LOG10(A3))', "tekst en functienamen blijven staan");
  });
  await test("Excel-datums: 1900- en 1904-datumsysteem", () => {
    const f = new Function(haalFunctie("serieelNaarTekst") + "; return serieelNaarTekst;")();
    gelijk(f("45292", "datum", false), "01-01-2024", "1900-systeem");
    gelijk(f("43830", "datum", true), "01-01-2024", "1904-systeem");
  });
  await test("Oude versienummers worden als v0.x getoond", () => {
    const toonVersie = new Function(haalTussen("const toonVersie", "\n") + "; return toonVersie;")();
    gelijk([toonVersie("v9.1"), toonVersie("v9"), toonVersie(undefined), toonVersie("v1.0")], ["v0.9.1", "v0.9", "v0.8 of ouder", "v1.0"], "versienummers");
  });

  browser = await chromium.launch();
  const basis = maakBasismap();

  console.log("\nAnalyse en weergave");
  let p = await nieuwePagina();
  await test("Analyse: samenvatting, groepen, systeembestanden en jaartallen", async () => {
    await zetOpties(p, { optKopie: true });
    await analyseer(p, basis);
    const t = await tegels(p);
    gelijk(t["geanalyseerd"], "16", "aantal geanalyseerde bestanden (4 systeembestanden overgeslagen)");
    gelijk(await groepen(p, "afw"), ["notulen.docx", "rapport.txt"], "afwijkingen");
    gelijk(await groepen(p, "ident"), ["leeg1.txt", "oud.txt"], "identieke doublures");
    const inh = await groepen(p, "inh");
    ok(inh.length === 1 && inh[0].includes("kopie van iets.txt"), "zelfde inhoud, andere naam: " + JSON.stringify(inh));
    ok(!JSON.stringify(await groepen(p, "afw")).includes("Jaarverslag"), "jaartallen mogen niet als kopie gelden");
    ok(!JSON.stringify(await groepen(p, "afw")).includes("(1).txt") && !JSON.stringify(await groepen(p, "afw")).includes("(2).txt"), "(1).txt en (2).txt mogen niet als dezelfde naam gelden");
    ok(!(await p.$eval("#instellingen", d => d.open)), "instellingen klappen in na de analyse");
  });
  await test("Groepen worden pas opgebouwd bij uitklappen", async () => {
    await p.click('.tab[data-tab="ident"]');
    const d = 'details.groep[data-id="I1"]';
    gelijk(await p.$$eval(`${d} table`, t => t.length), 0, "ingeklapte groep heeft nog geen tabel");
    await p.click(`${d} > summary`);
    await p.waitForSelector(`${d} table`);
    await p.click("#alleOpen");
    gelijk(await p.$$eval("details.groep:not([data-geladen])", d => d.length), 0, "na Alles uitklappen zijn alle groepen opgebouwd");
  });
  await test("Filter wacht tot er niet meer getypt wordt", async () => {
    await p.click('.tab[data-tab="afw"]');
    await p.fill("#filter", "notulen");
    gelijk(await p.$$eval("details.groep", d => d.length), 2, "direct na typen is de lijst nog niet gefilterd");
    await p.waitForFunction(() => document.querySelectorAll("details.groep").length === 1, null, { timeout: 2000 });
    await p.fill("#filter", "");
    await p.waitForFunction(() => document.querySelectorAll("details.groep").length === 2, null, { timeout: 2000 });
  });
  await test("Word-bestanden inhoudelijk vergelijken", async () => {
    await p.click('details.groep[data-id="A1"] .vergelijkKnop');
    await p.waitForFunction(() => /Tekstverschillen/.test(document.querySelector("#diff_A1").textContent));
    const t = await p.textContent("#diff_A1");
    ok(t.includes("Vierde") && t.includes("aangepast"), "verschil niet gevonden: " + t.slice(0, 200));
  });
  await test("Tegel in de samenvatting werkt met muis en toetsenbord", async () => {
    await p.click('.kaart[data-tab="ident"]');
    gelijk(await p.$eval(".tab.actief", t => t.dataset.tab), "ident", "klik op tegel");
    await p.focus('.kaart[data-tab="afw"]'); await p.keyboard.press("Enter");
    gelijk(await p.$eval(".tab.actief", t => t.dataset.tab), "afw", "Enter op tegel");
    gelijk(await p.$eval(".tab.actief", t => t.getAttribute("aria-selected")), "true", "aria-selected op actief tabblad");
  });
  await test("Vast blok blijft bij scrollen bovenaan", async () => {
    await p.setViewportSize({ width: 1366, height: 400 });
    await p.click("#alleOpen");
    await p.mouse.wheel(0, 2000); await p.waitForTimeout(200);
    ok(await p.evaluate(() => window.scrollY) > 100, "pagina is niet gescrold");
    gelijk(await p.$eval("#vastblok", e => Math.round(e.getBoundingClientRect().top)), 0, "vast blok staat niet bovenaan");
    await p.setViewportSize({ width: 1400, height: 900 });
  });

  console.log("\nVoorstel, exports en opmerkingenbestand");
  let jsonPad, data;
  await test("Voorstel toepassen en exports (Excel, Word, opmerkingenbestand)", async () => {
    await p.click('.kaart[data-tab="ident"]');
    await p.click("#voorstelToepassen");
    ok((await status(p)).startsWith("Voorstel toegepast"), "statusmelding: " + await status(p));
    const overbodig = (await tegels(p))["overbodig"];
    const xlsx = zipInhoud(fs.readFileSync(await download(p, "exportXlsx")));
    ok(xlsx["xl/workbook.xml"].includes('name="Toelichting"'), "Excel mist tabblad Toelichting");
    const toel = Object.entries(xlsx).filter(([n]) => /sheet\d+\.xml$/.test(n)).map(([, v]) => v).find(v => v.includes("Doublurecontrole " + VERSIE));
    ok(toel, "Excel-toelichting bevat geen versienummer " + VERSIE);
    ok(/<c r="B\d+" s="2"><v>16<\/v><\/c>/.test(toel), "aantal geanalyseerde bestanden staat niet als getal in Excel");
    const word = zipInhoud(fs.readFileSync(await download(p, "exportRapport")))["word/document.xml"];
    ok(word.includes("Doublurecontrole " + VERSIE), "Word-rapport bevat geen versienummer");
    ok(word.includes(">" + overbodig + "<"), `overbodige opslag (${overbodig}) in Word wijkt af van het scherm`);
    jsonPad = await download(p, "voortgangOpslaan");
    data = JSON.parse(fs.readFileSync(jsonPad, "utf8"));
    gelijk([data.versie, data.formaat], [VERSIE, 4], "versie en formaat van het opmerkingenbestand");
    ok(data.acties.length > 0 && data.acties.every(a => typeof a.grootte === "number" && a.hash), "keuzes missen grootte of hash");
    ok(data.weergave && Array.isArray(data.weergave.groepen), "opmerkingenbestand mist de weergave");
  });
  await p.close();

  await test("Weergave en instellingen terugzetten; instellingen overnemen", async () => {
    ok(jsonPad, "geen opmerkingenbestand: vorige test mislukt");
    p = await nieuwePagina();
    await analyseer(p, basis);                                   // nu zonder kopie-optie
    await p.setInputFiles("#importKiezer", jsonPad);
    await p.waitForSelector("#instellingenOvernemen");
    ok((await p.textContent("#instellingenVoorstel")).includes("Kopie-aanduidingen"), "verschil in instellingen niet getoond");
    await p.click("#instellingenOvernemen");
    await p.waitForFunction(() => document.querySelector("#instellingenVoorstel").style.display === "none" && document.querySelector("#stop").disabled);
    gelijk(await p.$eval("#optKopie", e => e.checked), true, "kopie-optie is niet overgenomen");
    gelijk(await p.$eval(".tab.actief", t => t.dataset.tab), data.weergave.tab, "actief tabblad is niet teruggezet");
    ok(/\d+ keuze\(s\) overgenomen/.test(await status(p)), "keuzes niet opnieuw geladen: " + await status(p));
    await p.close();
  });

  console.log("\nHercontrole");
  // Opschonen zoals een mapbeheerder dat zou doen (alleen als het opmerkingenbestand is aangemaakt)
  const verwijderd = data ? data.acties.filter(a => a.actie === "verwijderen").map(a => a.pad) : [];
  for (const pad of verwijderd) fs.unlinkSync(path.join(WERK, pad));
  await test("Hercontrole na opschonen (nieuw gekozen map)", async () => {
    ok(jsonPad, "geen opmerkingenbestand: eerdere test mislukt");
    p = await nieuwePagina();
    await zetOpties(p, { optKopie: true });
    await analyseer(p, basis);
    await p.setInputFiles("#importKiezer", jsonPad);
    await p.waitForSelector('.tab[data-tab="herc"]');
    await p.click('.tab[data-tab="herc"]');
    const rijen = await p.$$eval("#lijst tbody tr", r => r.map(x => [x.cells[0].textContent, x.cells[2].textContent]));
    for (const pad of verwijderd) ok(rijen.some(([pd, st]) => pd === pad && st === "Uitgevoerd"), `${pad} staat niet op Uitgevoerd`);
    ok(rijen.filter(([, st]) => st === "Aanwezig").length >= 1, "behouden bestanden niet als Aanwezig getoond");
    await p.close();
  });
  await test("Hercontrole uitzetten: geen tabblad, wel keuzes laden", async () => {
    ok(jsonPad, "geen opmerkingenbestand: eerdere test mislukt");
    p = await nieuwePagina();
    await zetOpties(p, { optKopie: true });
    await analyseer(p, basis);
    await p.uncheck("#optHercontrole");
    await p.setInputFiles("#importKiezer", jsonPad);
    await p.waitForFunction(() => /Geladen/.test(document.querySelector("#exportStatus").textContent));
    gelijk(await p.$$eval('.tab[data-tab="herc"]', t => t.length), 0, "tabblad Hercontrole staat er toch");
    await p.check("#optHercontrole");
    await p.waitForSelector('.tab[data-tab="herc"]');
    await p.close();
  });
  await test("Opnieuw analyseren met verouderde bestandslijst geeft een waarschuwing", async () => {
    // Map kiezen, daarna pas opschonen, dan "Opnieuw analyseren"
    const m = "Verouderd";
    schrijf(`${m}/a.txt`, "zelfde", 0); schrijf(`${m}/x/a.txt`, "zelfde", 10);
    p = await nieuwePagina();
    await analyseer(p, path.join(WERK, m));
    fs.unlinkSync(path.join(WERK, m, "x/a.txt"));
    await p.click("#opnieuw"); await wachtKlaar(p);
    ok((await status(p)).includes("sinds het kiezen van de map"), "geen waarschuwing: " + await status(p));
    await p.close();
  });

  await test("Excel-bestanden vergelijken: gedeelde formules en 1904-datums", async () => {
    const m = "Excel";
    const c = (ref, inhoud, s) => `<c r="${ref}"${s ? ` s="${s}"` : ""}>${inhoud}</c>`;
    schrijf(`${m}/plan.xlsx`, xlsx([c("A1", "<v>1</v>"), c("B1", '<f t="shared" ref="B1:C1" si="0">A1*2</f><v>2</v>'), c("C1", '<f t="shared" si="0"/><v>4</v>'), c("D1", "<v>43830</v>", 1)], true), 0);
    schrijf(`${m}/kopie/plan.xlsx`, xlsx([c("A1", "<v>1</v>"), c("B1", '<f t="shared" ref="B1:C1" si="0">A1*3</f><v>3</v>'), c("C1", '<f t="shared" si="0"/><v>9</v>'), c("D1", "<v>43830</v>", 1)], true), 10);
    p = await nieuwePagina();
    await analyseer(p, path.join(WERK, m));
    await p.click('details.groep[data-id="A1"] .vergelijkKnop');
    await p.waitForFunction(() => /Celverschillen|Geen verschillen/.test(document.querySelector("#diff_A1").textContent));
    const rijen = await p.$$eval("#diff_A1 tbody tr", r => r.map(x => [...x.cells].map(c => c.textContent)));
    const cel = ref => (rijen.find(r => r[1] === ref) || []).slice(2);
    gelijk(cel("C1"), ["=B1*2  →  4", "=B1*3  →  9"], "gedeelde formule in C1 niet goed afgeleid");
    ok(!rijen.some(r => r[1] === "D1"), "datum in D1 hoort gelijk te zijn");
    await p.close();
  });
  await test("Grote bestanden met dezelfde grootte: eerst begin en eind vergelijken", async () => {
    const m = "Voorfilter", a = crypto.randomBytes(3 * 1024 * 1024), b = crypto.randomBytes(3 * 1024 * 1024);
    schrijf(`${m}/een.bin`, a, 0); schrijf(`${m}/x/twee.bin`, a, 10); schrijf(`${m}/drie.bin`, b, 0);
    p = await nieuwePagina();
    // Lage grenzen, zodat met kleine bestanden ook het hashen één voor één en in blokken wordt doorlopen
    await p.evaluate(() => { window.DOUCON_TEST = { exclusiefVanaf: 1024 * 1024, stroomVanaf: 2 * 1024 * 1024 }; });
    await analyseer(p, path.join(WERK, m));
    gelijk(await p.evaluate(() => state.res.gehasht), 2, "alleen de twee bestanden met gelijke vingerafdruk horen volledig gehasht te worden");
    await p.click('.tab[data-tab="inh"]'); await p.click("#alleOpen");
    const verwacht = crypto.createHash("sha256").update(a).digest("hex");
    gelijk(await p.$$eval("td.hash", t => t.map(x => x.title)), [verwacht, verwacht], "hash (in blokken berekend) wijkt af");
    await p.close();
  });

  console.log("\nPrestaties");
  await test("3000 groepen: tabblad openen en filteren binnen redelijke tijd", async () => {
    p = await nieuwePagina();
    await analyseer(p, maakGroteMap(3000));
    const ms = await p.evaluate(() => { const t0 = performance.now(); document.querySelector('.tab[data-tab="ident"]').click(); return performance.now() - t0; });
    const knopen = await p.evaluate(() => document.querySelectorAll("#lijst *").length);
    console.log(`      tabblad openen: ${Math.round(ms)} ms, ${knopen} elementen in de lijst`);
    ok(knopen < 40000, `te veel elementen in de lijst (${knopen}); worden ingeklapte groepen toch opgebouwd?`);
    ok(ms < 5000, `tabblad openen duurde ${Math.round(ms)} ms`);
    await p.close();
  });

  if (MET_GROOT) {
    await test("Groot bestand (130 MB, ingebouwde hashfunctie, één tegelijk) wordt correct gehasht", async () => {
      const m = "Groot", data = crypto.randomBytes(130 * 1024 * 1024);
      schrijf(`${m}/groot.bin`, data, 0); schrijf(`${m}/kopie/groot.bin`, data, 10);
      const verwacht = crypto.createHash("sha256").update(data).digest("hex");
      p = await nieuwePagina();
      await analyseer(p, path.join(WERK, m));
      await p.click('.tab[data-tab="ident"]'); await p.click("#alleOpen");
      const hashes = await p.$$eval("td.hash", t => t.map(x => x.title));
      gelijk(hashes, [verwacht, verwacht], "hash van het grote bestand");
      await p.close();
    });
  } else console.log("  - groot bestand overgeslagen (gebruik --groot)");

  await test("Geen JavaScript-fouten in de pagina", () => { gelijk(paginafouten, [], "fouten in de console"); });

  await browser.close();
  fs.rmSync(WERK, { recursive: true, force: true });
  const mis = resultaten.filter(r => !r.ok).length;
  console.log(`\n${resultaten.length - mis} geslaagd, ${mis} mislukt`);
  process.exit(mis ? 1 : 0);
})().catch(async e => {
  console.error(e);
  if (browser) await browser.close();
  process.exit(1);
});
