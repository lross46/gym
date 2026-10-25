// Turns the trainer's Word files (one per training day) into a cycle.
// Reads .doc (Word 97-2003) and .docx. Runs in the browser; also loadable from Node for tests.
(function (root) {
  "use strict";
  const W1252 = "€\u0081‚ƒ„…†‡ˆ‰Š‹Œ\u008dŽ\u008f\u0090‘’“”•–—˜™š›œ\u009džŸ";
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

  // Text of a .doc, using its piece table. Table cells end with \x07; a row ends with one more \x07.
  function docText(CFB, bytes) {
    const cfb = CFB.read(bytes, { type: "array" });
    const get = (n) => { const e = CFB.find(cfb, n); return e && e.content; };
    const wd = get("WordDocument");
    if (!wd) throw new Error("not a Word file");
    const tbl = get((u16(wd, 0x0a) & 0x0200) ? "1Table" : "0Table");
    if (!tbl) throw new Error("damaged Word file");
    const ccpText = u32(wd, 76);
    let p = u32(wd, 0x1a2);
    while (tbl[p] === 0x01) p += 3 + u16(tbl, p + 1);
    if (tbl[p] !== 0x02) throw new Error("unsupported Word layout");
    const n = (u32(tbl, p + 1) - 4) / 12;
    p += 5;
    let out = "";
    for (let i = 0; i < n && out.length < ccpText; i++) {
      const cp0 = u32(tbl, p + i * 4), cp1 = u32(tbl, p + (i + 1) * 4);
      let fc = u32(tbl, p + (n + 1) * 4 + i * 8 + 2);
      const compressed = (fc & 0x40000000) !== 0;
      fc &= 0x3fffffff;
      const len = Math.min(cp1, ccpText) - cp0;
      if (compressed) {
        const o = fc / 2;
        for (let k = 0; k < len; k++) { const c = wd[o + k]; out += c >= 0x80 && c < 0xa0 ? W1252[c - 0x80] : String.fromCharCode(c); }
      } else {
        for (let k = 0; k < len; k++) out += String.fromCharCode(u16(wd, fc + k * 2));
      }
    }
    out = out.replace(/\x13[^\x13\x14\x15]*\x14([^\x15]*)\x15/g, "$1").replace(/\x13[^\x15]*\x15/g, "");
    return out.replace(/[\x01\x08\x0c]/g, "").replace(/\x0b/g, "\n");
  }

  // Same shape of text from a .docx, so one parser handles both.
  function docxText(CFB, bytes, DOMParserImpl) {
    const zip = CFB.read(bytes, { type: "array" });
    const e = CFB.find(zip, "word/document.xml");
    if (!e) throw new Error("not a Word file");
    const xml = new TextDecoder("utf-8").decode(e.content);
    const dom = new DOMParserImpl().parseFromString(xml, "application/xml");
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const kids = (el, name) => Array.from(el.childNodes).filter((c) => c.localName === name && c.namespaceURI === W);
    const para = (pEl) => {
      let s = "";
      pEl.getElementsByTagNameNS(W, "*");
      const walk = (el) => {
        for (const c of Array.from(el.childNodes)) {
          if (c.namespaceURI !== W) continue;
          if (c.localName === "t") s += c.textContent;
          else if (c.localName === "tab") s += "\t";
          else if (c.localName === "br" || c.localName === "cr") s += "\n";
          else if (c.localName !== "pPr" && c.localName !== "rPr") walk(c);
        }
      };
      walk(pEl);
      return s;
    };
    let out = "";
    const block = (el) => {
      for (const c of Array.from(el.childNodes)) {
        if (c.namespaceURI !== W) continue;
        if (c.localName === "p") out += para(c) + "\r";
        else if (c.localName === "tbl") {
          for (const tr of kids(c, "tr")) {
            for (const tc of kids(tr, "tc")) out += kids(tc, "p").map(para).join("\r") + "\x07";
            out += "\x07";
          }
        } else if (c.localName === "sdt" || c.localName === "sdtContent") block(c);
      }
    };
    block(dom.getElementsByTagNameNS(W, "body")[0]);
    return out;
  }

  const tidy = (s) => (s || "").replace(/\r/g, "\n").replace(/\u2028/g, "\n").replace(/\xa0/g, " ")
    .replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  const oneLine = (s) => tidy(s).split(/\n+/).map((x) => x.trim()).filter(Boolean).join(" / ");
  const weekNo = (s) => { const t = oneLine(s); return /^\d+(\.0)?$/.test(t) ? parseInt(t, 10) : t; };

  // One day's file -> { num, duration, day, warmup, exercises }.
  function parseDay(text) {
    const cells = text.split("\x07");
    const h = cells.findIndex((c) => /ESERCIZIO\s*$/i.test(c));
    if (h < 0) throw new Error("couldn't find the exercise table");
    const intro = cells[h].replace(/ESERCIZIO\s*$/i, "");
    const all = (intro + "\n" + cells.slice(0, h).join("\n")).replace(/\r/g, "\n");
    const num = (all.match(/allenamento\s+(\d+\w*)/i) || [])[1] || "";
    const duration = ((all.match(/Durata[^:]*:\s*([^\n]*)/i) || [])[1] || "").trim();
    const day = ((all.match(/GIORNO\s*["“”']?\s*([A-Z])/i) || [])[1] || "").toUpperCase();
    const warmup = tidy((all.match(/RISCALDAMENTO:?\s*([\s\S]*)$/i) || [])[1] || "");
    let ncols = 0;
    while (h + ncols < cells.length && cells[h + ncols].trim() !== "") ncols++;
    const exercises = [];
    let cur = null, i = h + ncols + 1;
    while (i + ncols <= cells.length) {
      const row = cells.slice(i, i + ncols);
      i += ncols + 1;
      if (row.every((c) => !c.trim())) continue;
      const [name, wk, sets, reps, rest, note] = row.map(tidy);
      if (name) {
        cur = { name: oneLine(name), desc: "", weeks: [] };
        exercises.push(cur);
      }
      if (!cur) continue;
      if (note) cur.desc = cur.desc ? cur.desc + "\n" + note : note;
      cur.weeks.push({ w: weekNo(wk), sets: oneLine(sets), reps: oneLine(reps), rest: oneLine(rest), log: "" });
    }
    if (!exercises.length) throw new Error("no exercises found");
    return { num, duration, day, warmup, exercises };
  }

  // Several day files -> a cycle in the app's format.
  function buildCycle(days) {
    days = days.slice().sort((a, b) => (a.day || "Z").localeCompare(b.day || "Z"));
    const num = (days.find((d) => d.num) || {}).num || "";
    const n = parseInt(num, 10);
    const suffix = num.replace(/^\d+/, "").toLowerCase();
    return {
      id: num ? "c" + n + suffix : "c" + Date.now(),
      num: isNaN(n) ? null : n,
      label: num ? "Cycle " + n + suffix : "New cycle",
      when: "", tag: "", title: "",
      duration: (days.find((d) => d.duration) || {}).duration || "",
      period: "",
      days: days.map((d, k) => ({ id: d.day || "ABCDEFG"[k], warmup: d.warmup, cooldown: "", exercises: d.exercises })),
      source: "import",
    };
  }

  async function importFiles(CFB, files) {
    const days = [];
    for (const f of files) {
      const bytes = new Uint8Array(await f.arrayBuffer());
      const isZip = bytes[0] === 0x50 && bytes[1] === 0x4b;
      const text = isZip ? docxText(CFB, bytes, DOMParser) : docText(CFB, bytes);
      try { days.push(parseDay(text)); } catch (e) { throw new Error(f.name + ": " + e.message); }
    }
    return buildCycle(days);
  }

  const api = { docText, docxText, parseDay, buildCycle, importFiles };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.GymImport = api;
})(typeof window !== "undefined" ? window : this);
