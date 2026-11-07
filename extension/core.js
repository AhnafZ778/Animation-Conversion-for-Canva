/* Canva -> PPTX animation transfer: core logic (browser + node).
 *
 * Pipeline:  Canva page HTML --extractFromHtml--> IR --injectIntoPptx(IR, canva-exported .pptx)--> animated .pptx
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CanvaPptx = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // ------------------------------------------------------------------ Canva decoding
  // Element animation ids, taken from Canva's front-end bundle.
  const ELEMENT_ANIMS = {
    1: "Baseline", 2: "Breathe", 3: "Drift", 4: "Fade", 5: "Neon", 6: "Pan", 7: "Pop", 8: "Rise",
    11: "Stomp", 13: "Tumble", 14: "Photo Flow", 15: "Photo Zoom", 16: "Photo Rise", 17: "Block",
    18: "Bounce", 19: "Burst", 20: "Merge", 21: "Shift", 22: "Skate", 23: "Spread", 24: "Typewriter",
    25: "Ascend", 26: "Wipe", 27: "Roll", 29: "Blur", 31: "Succession", 32: "Clarify",
    33: "Caption Highlight", 35: "Caption Reveal",
  };
  // Canva effect -> [PowerPoint effect, default entry side, approximate?]
  const TO_PPT = {
    Fade: ["fade", null, false], Rise: ["fly", "bottom", false], Pan: ["fly", "left", false],
    Wipe: ["wipe", "left", false], Pop: ["zoom", null, false],
    Baseline: ["wipe", "top", true], Block: ["wipe", "left", true], Typewriter: ["wipe", "left", true],
    Ascend: ["fly", "bottom", true], Drift: ["fly", "bottom", true], Shift: ["fly", "right", true],
    Tumble: ["zoom", null, true], Stomp: ["zoom", null, true], Burst: ["zoom", null, true],
    Spread: ["zoom", null, true], Breathe: ["zoom", null, true], Bounce: ["zoom", null, true],
    Blur: ["fade", null, true], Neon: ["fade", null, true], Merge: ["fade", null, true],
    Succession: ["fade", null, true], Clarify: ["fade", null, true], Roll: ["fly", "left", true],
    "Photo Flow": ["fade", null, true], "Photo Zoom": ["zoom", null, true], "Photo Rise": ["fly", "bottom", true],
    Skate: ["fly", "left", true], "Caption Reveal": ["fade", null, true], "Caption Highlight": ["fade", null, true],
  };
  // UNVERIFIED: Canva direction enum letters (A..E = 1..5) -> side the element enters from.
  const DIR_ENUM = { B: "bottom", C: "top", D: "right", E: "left" };
  const ELEMENT_TYPES = { K: "text", I: "image", H: "group", J: "shape", U: "video" };

  const KNOWN_ELEMENT_KEYS = new Set(["A?", "A", "B", "C", "D", "_", "a", "b", "c", "d", "e", "f", "g", "h", "j", "E", "F", "N", "X", "Z", "W", "Q", "R", "2"]);
  const KNOWN_OPTION_KEYS = new Set(["A", "C", "D", "G", "H"]);
  const KNOWN_PAGE_KEYS = new Set(["A?", "Z", "a", "W", "C", "D", "E", "G", "w", "O"]);

  const DEFAULTS = { staggerMs: 70, defaultDurationMs: 500 };

  function jsUnescape(s) {
    return s.replace(/\\(x[0-9a-fA-F]{2}|u[0-9a-fA-F]{4}|[\s\S])/g, (_, t) => {
      if (t[0] === "x" && t.length === 3) return String.fromCharCode(parseInt(t.slice(1), 16));
      if (t[0] === "u" && t.length === 5) return String.fromCharCode(parseInt(t.slice(1), 16));
      const map = { n: "\n", t: "\t", r: "\r", b: "\b", f: "\f", 0: "\0" };
      return t in map ? map[t] : t;
    });
  }

  function loadBootstrap(html) {
    const m = html.match(/window\['bootstrap'\] = JSON\.parse\('((?:[^'\\]|\\[\s\S])*)'\)/);
    if (!m) throw new Error("This page doesn't contain a Canva design (open a design in the editor and try again).");
    return JSON.parse(jsUnescape(m[1]));
  }

  function findPages(boot) {
    const page = boot && boot.page ? boot.page : {};
    for (const v of Object.values(page)) {
      const a = v && v.E && v.E.draft && v.E.draft.content && v.E.draft.content.A;
      if (Array.isArray(a)) return a;
    }
    throw new Error("Couldn't find the design's pages. Canva may have changed its page format.");
  }

  function textOf(e) {
    try { return e.a.C.A.join("").trim(); } catch (_) { return null; }
  }

  function designTitle(boot) {
    try {
      for (const v of Object.values(boot.page)) {
        const t = v && v.E && v.E.draft && v.E.draft.content && v.E.draft.title;
        if (typeof t === "string") return t;
      }
    } catch (_) { /* ignore */ }
    return null;
  }

  /**
   * Trigger rules (verified against a hand-built reference deck):
   *  - element has animation X and no step id Z      -> automatic, plays when the slide appears
   *  - element has X and Z, Z listed in page order W -> on click, click order = position of Z in W
   *  - elements that share a Z play together on one click
   */
  function convert(boot, opts) {
    opts = Object.assign({}, DEFAULTS, opts || {});
    const pages = findPages(boot);
    const ir = { version: "0.2", source: "canva", title: designTitle(boot), deck: { width: 1920, height: 1080 }, slides: [] };
    const warnings = [];
    pages.forEach((p, pi) => {
      const sn = pi + 1;
      const c = p.C || {};
      ir.deck.width = c.A || 1920; ir.deck.height = c.B || 1080;
      const slide = { id: "s" + sn, elements: [], animations: [] };
      if (p.O) {
        slide.transition = { type: "fade", duration: (p.O.B || 500000) / 1e6 };
        warnings.push(`Slide ${sn}: has a page transition; its exact type can't be read yet, so a fade is used.`);
      }
      const W = Array.isArray(p.W) ? p.W : [];
      const worder = new Map();
      W.forEach((it, wi) => { if (!worder.has(it.A)) worder.set(it.A, wi); }); // stale/duplicate ids: first wins
      const autos = [], clicks = new Map(), unknownKeys = [];
      (p.E || []).forEach((e, ei) => {
        // Canva omits coordinates that are exactly 0, so a missing top/left means 0; size must be present.
        if (typeof e.D !== "number" || typeof e.C !== "number") return;
        const id = e._ || `p${sn}e${ei}`;
        const el = { id, type: ELEMENT_TYPES[e["A?"]] || "other", x: e.B || 0, y: e.A || 0, w: e.D, h: e.C, z: ei };
        const t = textOf(e);
        if (t) el.text = t;
        slide.elements.push(el);
        const x = e.X;
        if (!x) return;
        const aid = x.A;
        const name = ELEMENT_ANIMS[aid] || `unknown-${aid}`;
        const o = x.B || {};
        let tgt = TO_PPT[name];
        if (!tgt) {
          warnings.push(`Slide ${sn}: Canva animation "${name}" isn't supported yet; a fade is used.`);
          tgt = ["fade", null, true];
        } else if (tgt[2]) {
          warnings.push(`Slide ${sn}: Canva "${name}" has no PowerPoint equivalent; approximated as ${tgt[0]}.`);
        }
        const dur = ((o.G && typeof o.G.A === "number") ? o.G.A / 1000 : opts.defaultDurationMs);
        const a = { target: id, effect: tgt[0], durationMs: Math.round(dur), canva: { id: aid, name, step: e.Z } };
        const dir = DIR_ENUM[o.C] || tgt[1];
        if (dir) a.direction = dir;
        // A click step id is stored either directly (Z) or via a per-element link map (W: {group: stepId}).
        let z = e.Z;
        if (z == null && e.W && typeof e.W === "object") { const v = Object.values(e.W)[0]; if (v != null) z = v; }
        a.canva.step = z;
        if (z != null) {                       // has a step id => it is an on-click step
          if (!clicks.has(z)) clicks.set(z, { order: worder.has(z) ? worder.get(z) : 1e6 + ei, items: [] });
          clicks.get(z).items.push(a);
          if (!worder.has(z)) warnings.push(`Slide ${sn}: a click step isn't in Canva's step order list; it was placed by stacking order.`);
        } else autos.push(a);
        const odd = Object.keys(e).filter((k) => !KNOWN_ELEMENT_KEYS.has(k));
        const oddX = Object.keys(x).filter((k) => !["A?", "A", "B"].includes(k));
        const oddO = Object.keys(o).filter((k) => !KNOWN_OPTION_KEYS.has(k));
        if (odd.length || oddX.length || oddO.length) unknownKeys.push(...odd.map((k) => "element." + k), ...oddX.map((k) => "animation." + k), ...oddO.map((k) => "options." + k));
      });
      autos.forEach((a, k) => {
        a.trigger = "withPrevious";
        if (k && opts.staggerMs) a.delayMs = k * opts.staggerMs;
        slide.animations.push(a);
      });
      [...clicks.values()].sort((a, b) => a.order - b.order).forEach((step) => {
        step.items.forEach((a, k) => { a.trigger = k === 0 ? "onClick" : "withPrevious"; slide.animations.push(a); });
      });
      const odd = [...new Set(unknownKeys)];
      if (odd.length) warnings.push(`Slide ${sn}: Canva stores animation settings this tool doesn't understand yet (${odd.join(", ")}). If clicks look wrong, use "Download support file".`);
      const oddPage = Object.keys(p).filter((k) => !KNOWN_PAGE_KEYS.has(k));
      if (oddPage.length && slide.animations.length) warnings.push(`Slide ${sn}: unrecognised slide settings (${oddPage.join(", ")}).`);
      slide.debug = { pageKeys: Object.keys(p), W: W.map((w) => w.A), animatedKeys: [...new Set(unknownKeys)] };
      ir.slides.push(slide);
    });
    return { ir, warnings: [...new Set(warnings)] };
  }

  function extractFromHtml(html, opts) {
    const r = convert(loadBootstrap(html), opts);
    const skip = opts && opts.skip ? new Set(opts.skip) : null;
    if (skip) r.ir.slides.forEach((s) => { s.animations = s.animations.filter((a) => !skip.has(s.id + ":" + a.target)); });
    return r;
  }

  // ------------------------------------------------------------------ PowerPoint timing XML
  const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
  const DIRS = { bottom: [4, 4, "down"], top: [1, 1, "up"], left: [8, 8, "left"], right: [2, 2, "right"] };
  const SUPPORTED = new Set(["appear", "fade", "fly", "wipe", "zoom"]);

  function resolveEffect(a, warnings) {
    let eff = a.effect, d = a.direction;
    if (eff === "rise") { eff = "fly"; d = d || "bottom"; }
    if (eff === "slide") eff = "fly";
    if (!SUPPORTED.has(eff)) { warnings.push(`effect "${a.effect}" unsupported; fade used`); eff = "fade"; }
    return [eff, d && DIRS[d] ? d : "bottom"];
  }

  function effectXml(a, spid, ids, nodeType, warnings) {
    const [eff, d] = resolveEffect(a, warnings);
    const dur = Math.max(1, Math.round(a.durationMs != null ? a.durationMs : 500));
    const delay = Math.max(0, Math.round(a.delayMs || 0));
    const tgt = `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;
    const parId = ids.next();
    const vis = () => `<p:set><p:cBhvr><p:cTn id="${ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
    const fx = (f) => `<p:animEffect transition="in" filter="${f}"><p:cBhvr><p:cTn id="${ids.next()}" dur="${dur}"/>${tgt}</p:cBhvr></p:animEffect>`;
    const attr = (name, from, to) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${ids.next()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${name}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${from}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${to}"/></p:val></p:tav></p:tavLst></p:anim>`;
    let pid, sub, body;
    if (eff === "appear") { pid = 1; sub = 0; body = vis(); }
    else if (eff === "fade") { pid = 10; sub = 0; body = vis() + fx("fade"); }
    else if (eff === "fly") {
      pid = 2; sub = DIRS[d][0];
      const from = { bottom: ["ppt_y", "1+#ppt_h/2"], top: ["ppt_y", "0-#ppt_h/2"], left: ["ppt_x", "0-#ppt_w/2"], right: ["ppt_x", "1+#ppt_w/2"] }[d];
      const other = from[0] === "ppt_y" ? "ppt_x" : "ppt_y";
      body = vis() + attr(from[0], from[1], "#" + from[0]) + attr(other, "#" + other, "#" + other);
    } else if (eff === "wipe") { pid = 22; sub = DIRS[d][1]; body = vis() + fx(`wipe(${DIRS[d][2]})`); }
    else { pid = 53; sub = 16; body = vis() + attr("ppt_w", "0", "#ppt_w") + attr("ppt_h", "0", "#ppt_h") + fx("fade"); }
    return {
      xml: `<p:par><p:cTn id="${parId}" presetID="${pid}" presetClass="entr" presetSubtype="${sub}" fill="hold" nodeType="${nodeType}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`,
      endMs: dur + delay,
    };
  }

  /** animations: [{target, effect, direction, durationMs, delayMs, trigger}]; spids: {target -> shape id} */
  function buildTimingXml(animations, spids, warnings) {
    const groups = []; // {auto, steps:[[anim,...],...]}
    for (const a of animations) {
      if (!(a.target in spids)) continue;
      const t = a.trigger || "onClick";
      if (t === "onClick" || !groups.length) groups.push({ auto: t !== "onClick", steps: [[a]] });
      else if (t === "afterPrevious") groups[groups.length - 1].steps.push([a]);
      else groups[groups.length - 1].steps[groups[groups.length - 1].steps.length - 1].push(a);
    }
    if (!groups.length) return null;
    let n = 2;
    const ids = { next: () => ++n };
    let out = "";
    for (const g of groups) {
      const outerId = ids.next();
      let stepsXml = "", cursor = 0;
      g.steps.forEach((step, si) => {
        const stepId = ids.next();
        let effs = "", stepEnd = 0;
        step.forEach((a, i) => {
          const first = si === 0 && i === 0;
          const nodeType = first && !g.auto ? "clickEffect"
            : ((i === 0 && !first) || (first && g.auto && a.trigger === "afterPrevious")) ? "afterEffect" : "withEffect";
          const r = effectXml(a, spids[a.target], ids, nodeType, warnings);
          effs += r.xml; stepEnd = Math.max(stepEnd, r.endMs);
        });
        stepsXml += `<p:par><p:cTn id="${stepId}" fill="hold"><p:stCondLst><p:cond delay="${Math.round(cursor)}"/></p:stCondLst><p:childTnLst>${effs}</p:childTnLst></p:cTn></p:par>`;
        cursor += stepEnd;
      });
      const cond = '<p:cond delay="indefinite"/>' + (g.auto ? '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : "");
      out += `<p:par><p:cTn id="${outerId}" fill="hold"><p:stCondLst>${cond}</p:stCondLst><p:childTnLst>${stepsXml}</p:childTnLst></p:cTn></p:par>`;
    }
    return `<p:timing xmlns:p="${NS_P}"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${out}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  }

  function buildTransitionXml(t) {
    if (!t) return null;
    const dur = t.duration != null ? t.duration : 0.7;
    const spd = dur > 1 ? "slow" : dur > 0.4 ? "med" : "fast";
    const dmap = { left: "l", right: "r", up: "u", down: "d" };
    const dd = dmap[t.direction] || "l";
    const inner = { fade: "<p:fade/>", push: `<p:push dir="${dd}"/>`, wipe: `<p:wipe dir="${dd}"/>`, none: "" }[t.type];
    return `<p:transition xmlns:p="${NS_P}" spd="${spd}">${inner == null ? "<p:fade/>" : inner}</p:transition>`;
  }

  // ------------------------------------------------------------------ PPTX reading / injection
  const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const SHAPE_TAGS = new Set(["sp", "pic", "grpSp", "graphicFrame", "cxnSp"]);

  const kids = (el, local, ns) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === local && (!ns || n.namespaceURI === ns));
  const firstDesc = (el, local) => el.getElementsByTagNameNS("*", local)[0] || null;

  function xfrmOf(shape) {
    let x = null;
    const tag = shape.localName;
    if (tag === "graphicFrame") x = kids(shape, "xfrm")[0];
    else if (tag === "grpSp") { const pr = kids(shape, "grpSpPr")[0]; x = pr && kids(pr, "xfrm")[0]; }
    else { const pr = kids(shape, "spPr")[0]; x = pr && kids(pr, "xfrm")[0]; }
    if (!x) return null;
    const off = kids(x, "off")[0], ext = kids(x, "ext")[0];
    if (!off || !ext) return null;
    return [+off.getAttribute("x"), +off.getAttribute("y"), +ext.getAttribute("cx"), +ext.getAttribute("cy")];
  }

  function shapeInfo(doc) {
    const tree = doc.getElementsByTagNameNS(NS_P, "spTree")[0];
    const out = [];
    if (!tree) return out;
    for (const s of Array.from(tree.childNodes)) {
      if (s.nodeType !== 1 || !SHAPE_TAGS.has(s.localName)) continue;
      const c = firstDesc(s, "cNvPr");
      const bbox = xfrmOf(s);
      if (!c || !bbox) continue;
      const paras = Array.from(s.getElementsByTagNameNS(NS_A, "p")).map((p) => Array.from(p.getElementsByTagNameNS(NS_A, "t")).map((t) => t.textContent).join(""));
      const text = paras.join(" ").trim();
      out.push({ id: c.getAttribute("id"), bbox, text: text || null, hasText: s.localName === "sp" && s.getElementsByTagNameNS(NS_P, "txBody").length > 0 });
    }
    return out;
  }

  const norm = (s) => (s || "").split(/\s+/).filter(Boolean).join(" ").toLowerCase();

  function matchShape(el, shapes, scale, used) {
    const e = [el.x, el.y, el.w, el.h].map((v) => v * scale);
    let best = null, bestCost = Infinity;
    for (const s of shapes) {
      if (used.has(s.id)) continue;
      let cost = Math.abs(s.bbox[0] - e[0]) + Math.abs(s.bbox[1] - e[1]) + Math.abs(s.bbox[2] - e[2]) + Math.abs(s.bbox[3] - e[3]);
      if (el.text && s.hasText) cost = norm(s.text) === norm(el.text) ? cost * 0.1 : cost + 2e6;
      if (cost < bestCost) { best = s; bestCost = cost; }
    }
    return { shape: best, cost: bestCost };
  }

  async function slideParts(zip, dom) {
    const pres = await zip.file("ppt/presentation.xml").async("string");
    const rels = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
    const pdoc = dom.parse(pres), rdoc = dom.parse(rels);
    const rmap = {};
    for (const r of Array.from(rdoc.getElementsByTagName("Relationship"))) rmap[r.getAttribute("Id")] = r.getAttribute("Target");
    const sz = pdoc.getElementsByTagNameNS(NS_P, "sldSz")[0];
    const lst = pdoc.getElementsByTagNameNS(NS_P, "sldId");
    const paths = Array.from(lst).map((s) => {
      let t = rmap[s.getAttributeNS(NS_R, "id")] || "";
      t = t.replace(/^\//, "");
      return t.startsWith("ppt/") ? t : "ppt/" + t;
    });
    return { paths, width: +sz.getAttribute("cx"), height: +sz.getAttribute("cy") };
  }

  /**
   * Adds animations/transitions from `ir` to a PPTX (a Buffer/ArrayBuffer/Uint8Array/Blob).
   * deps: { JSZip, DOMParser, XMLSerializer }
   * returns { data: Uint8Array, report: {slides:[...], warnings:[...]} }
   */
  // ---- font metrics -------------------------------------------------------------------------------
  // Standard Helvetica/Times advance widths (1/1000 em) for ASCII 32..126. Arial and Times New Roman are
  // metric-compatible with these (and so are Liberation Sans/Serif on Linux), so they are predictable fallbacks.
  const AFM = {
    sans: [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584],
    sansBold: [278,333,474,556,556,889,722,238,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,333,333,584,584,584,611,975,722,722,722,722,667,611,778,722,278,556,722,611,833,722,778,667,778,722,667,611,722,667,944,667,667,611,333,278,333,584,556,333,556,611,556,611,556,333,611,611,278,278,556,278,889,611,611,611,611,389,556,333,611,556,778,556,556,500,389,280,389,584],
    serif: [250,333,408,500,500,833,778,180,333,333,500,564,250,333,250,278,500,500,500,500,500,500,500,500,500,500,278,278,564,564,564,444,921,722,667,667,722,611,556,722,722,333,389,722,611,889,722,722,556,722,667,556,611,722,722,944,722,722,611,333,278,333,469,500,333,444,500,444,500,444,333,500,500,278,278,500,278,778,500,500,500,500,333,389,278,500,500,722,500,500,444,480,200,480,541],
    serifBold: [250,333,555,500,500,1000,833,278,333,333,500,570,250,333,250,278,500,500,500,500,500,500,500,500,500,500,333,333,570,570,570,500,930,722,667,722,722,667,611,778,778,389,500,778,667,944,722,778,611,778,722,556,667,722,722,1000,722,722,667,333,278,333,581,500,333,500,556,444,556,444,333,500,556,278,333,556,278,833,556,500,556,556,444,389,333,556,500,722,500,500,444,394,220,394,520],
  };
  function afmWidth(kind, text) {
    const t = AFM[kind]; let w = 0;
    for (const ch of text) { const c = ch.codePointAt(0); w += (c >= 32 && c <= 126) ? t[c - 32] : (kind.startsWith("serif") ? 500 : 556); }
    return w / 1000; // in em
  }

  /** Minimal sfnt (TrueType/OpenType) reader: advance widths, cmap, OS/2 class info. */
  function parseSfnt(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const n = dv.getUint16(4), t = {};
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      t[String.fromCharCode(u8[o], u8[o + 1], u8[o + 2], u8[o + 3])] = { off: dv.getUint32(o + 8), len: dv.getUint32(o + 12) };
    }
    if (!t.head || !t.hhea || !t.hmtx || !t.cmap) return null;
    const upm = dv.getUint16(t.head.off + 18), nh = dv.getUint16(t.hhea.off + 34);
    let sub = null;
    const cm = t.cmap.off, nt = dv.getUint16(cm + 2), recs = [];
    for (let i = 0; i < nt; i++) recs.push({ p: dv.getUint16(cm + 4 + i * 8), e: dv.getUint16(cm + 6 + i * 8), o: cm + dv.getUint32(cm + 8 + i * 8) });
    const pick = recs.find((r) => r.p === 3 && r.e === 10) || recs.find((r) => r.p === 3 && r.e === 1) || recs.find((r) => r.p === 0);
    if (pick) {
      const f = dv.getUint16(pick.o), o = pick.o;
      if (f === 4) {
        const sc = dv.getUint16(o + 6) / 2;
        sub = (cp) => {
          if (cp > 0xffff) return 0;
          for (let s = 0; s < sc; s++) {
            if (cp > dv.getUint16(o + 14 + s * 2)) continue;
            const start = dv.getUint16(o + 16 + sc * 2 + s * 2);
            if (cp < start) return 0;
            const delta = dv.getInt16(o + 16 + sc * 4 + s * 2), ro = dv.getUint16(o + 16 + sc * 6 + s * 2);
            if (!ro) return (cp + delta) & 0xffff;
            const g = dv.getUint16(o + 16 + sc * 6 + s * 2 + ro + (cp - start) * 2);
            return g ? (g + delta) & 0xffff : 0;
          }
          return 0;
        };
      } else if (f === 12) {
        const ng = dv.getUint32(o + 12);
        sub = (cp) => {
          for (let g = 0; g < ng; g++) {
            const b = o + 16 + g * 12, s = dv.getUint32(b), e = dv.getUint32(b + 4);
            if (cp >= s && cp <= e) return dv.getUint32(b + 8) + (cp - s);
          }
          return 0;
        };
      }
    }
    if (!sub) return null;
    const adv = (g) => dv.getUint16(t.hmtx.off + Math.min(g, nh - 1) * 4);
    const font = { upm, glyph: sub, width: (cp) => adv(sub(cp)) / upm, weight: 400, familyClass: 0, panose: null };
    if (t["OS/2"]) {
      const o = t["OS/2"].off;
      font.weight = dv.getUint16(o + 4); font.familyClass = dv.getInt16(o + 30) >> 8;
      font.panose = Array.from(u8.subarray(o + 32, o + 42));
    }
    // all-caps display font? ("a" and "A" share a glyph)
    const ga = sub(97), gA = sub(65);
    font.capsOnly = !!ga && ga === gA && sub(98) === sub(66);
    if (font.panose && font.panose[0] === 2 && font.panose[3] === 9) font.mono = true;
    const sc = font.familyClass, ps = font.panose && font.panose[0] === 2 ? font.panose[1] : 0;
    font.serif = (sc >= 1 && sc <= 5) || sc === 7 ? true : sc === 8 ? false : (ps >= 2 && ps <= 10);
    font.script = sc === 10;
    font.classified = sc !== 0 || ps !== 0;
    return font;
  }

  /** Parse every embedded font Canva put in the PPTX: { typeface: {faces:Set, font:parsed|null} }. */
  async function embeddedFaces(zip, dom) {
    const out = {};
    const pres = zip.file("ppt/presentation.xml");
    if (!pres) return out;
    const doc = dom.parse(await pres.async("string"));
    const rmap = {};
    const rf = zip.file("ppt/_rels/presentation.xml.rels");
    if (rf) for (const r of Array.from(dom.parse(await rf.async("string")).getElementsByTagName("Relationship"))) rmap[r.getAttribute("Id")] = r.getAttribute("Target");
    for (const ef of Array.from(doc.getElementsByTagNameNS(NS_P, "embeddedFont"))) {
      const f = Array.from(ef.childNodes).find((c) => c.nodeType === 1 && c.localName === "font");
      if (!f) continue;
      const name = f.getAttribute("typeface");
      const ent = (out[name] = { faces: new Set(), font: null, bytes: {} });
      for (const c of Array.from(ef.childNodes)) {
        if (c.nodeType !== 1 || c.localName === "font") continue;
        ent.faces.add(c.localName);
        try {
          let t = rmap[c.getAttributeNS(NS_R, "id")] || ""; t = t.replace(/^\//, "");
          const file = zip.file(t.startsWith("ppt/") ? t : "ppt/" + t);
          if (!file) continue;
          const d = await file.async("uint8array");
          const dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
          const eot = dv.getUint32(0, true), fsz = dv.getUint32(4, true), flags = dv.getUint32(12, true);
          if (flags & (0x4 | 0x10000000)) continue; // compressed/encrypted EOT: can't read
          const sfnt = d.subarray(eot - fsz, eot);
          ent.bytes[c.localName] = sfnt;
          if (c.localName !== "regular" || ent.font) continue;
          ent.font = parseSfnt(sfnt);
          if (ent.font && !ent.font.classified) {
            if (/roca|slab|serif(?!.*sans)|garamond|georgia|times|playfair|merriweather|lora|baskerville|bodoni|didot|caslon|cormorant|libre franklin(?!)|crimson|bitter|zilla|rockwell|cambria|palatino|minion|cheltenham|dm serif|abril|chunk/i.test(name) && !/sans/i.test(name)) ent.font.serif = true;
          }
        } catch (e) { /* unreadable font: fall back to heuristics */ }
      }
    }
    return out;
  }

  /**
   * Text fidelity fixes on a slide DOM (Canva's export is the source of the text; we only undo two things
   * that make PowerPoint render it wider than Canva did):
   *  1. "doubled bold": a run is b="true" but its font (e.g. "Telegraf Bold", "Roca One Heavy") is embedded
   *     with only a regular face, so PowerPoint fakes an extra bold on top -> wider glyphs -> extra wrapping.
   *  2. a single-line text box may wrap if the font renders a bit wider; keep it on one line (wrap="none").
   * Returns the number of changes.
   */
  function fixTextLayout(doc, faces, opts) {
    opts = opts || {};
    let n = 0;
    const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const isOn = (v) => v === "true" || v === "1";
    const kids = (el, ln) => Array.from(el.childNodes).filter((c) => c.nodeType === 1 && c.localName === ln);
    for (const tag of ["rPr", "endParaRPr"]) {
      for (const r of Array.from(doc.getElementsByTagNameNS(A, tag))) {
        const lat = kids(r, "latin")[0];
        const ent = lat && faces[lat.getAttribute("typeface")];
        if (!ent) continue;
        const face = ent.faces, font = ent.font;
        const run = tag === "rPr" && r.parentNode && r.parentNode.localName === "r" ? r.parentNode : null;
        const tEl = run && kids(run, "t")[0];
        const text = tEl ? tEl.textContent : "";
        const wasBold = isOn(r.getAttribute("b"));
        if (opts.compat && font && tEl && text) {
          // Swap to Arial / Times New Roman (predictable everywhere) and rescale so the run is exactly as wide as in Canva.
          const sz = +r.getAttribute("sz") || 1800, spc = +r.getAttribute("spc") || 0;
          const upper = font.capsOnly || r.getAttribute("cap") === "all";
          const shown = upper ? text.toUpperCase() : text;
          let w0 = 0; for (const ch of shown) w0 += font.width(ch.codePointAt(0)) * sz / 100;
          w0 += spc / 100 * Array.from(shown).length;
          const bold = wasBold || font.weight >= 600 || /bold|heavy|black/i.test(lat.getAttribute("typeface"));
          const kind = (font.serif ? "serif" : "sans") + (bold ? "Bold" : "");
          const w1 = afmWidth(kind, shown) * sz / 100;
          if (w0 > 0 && w1 > 0) {
            const k = Math.max(0.5, Math.min(1.15, w0 / w1));
            r.setAttribute("sz", String(Math.round(sz * k)));
            r.removeAttribute("spc");
            r.setAttribute("b", bold ? "true" : "false");
            if (upper) r.setAttribute("cap", "all");
            const fam = font.serif ? "Times New Roman" : "Arial";
            for (const t of ["latin", "ea", "cs", "sym"]) for (const e of kids(r, t)) e.setAttribute("typeface", fam);
            const l2 = kids(r, "latin")[0]; if (l2) l2.setAttribute("pitchFamily", font.serif ? "18" : "34");
            n++;
            continue;
          }
        }
        if (wasBold && !face.has("bold") && !face.has("boldItalic")) { r.setAttribute("b", "false"); n++; }
        if (isOn(r.getAttribute("i")) && !face.has("italic") && !face.has("boldItalic")) { r.setAttribute("i", "false"); n++; }
        // all-caps display fonts (e.g. Bebas): if a viewer substitutes the font, lowercase would show as lowercase
        if (font && font.capsOnly && !r.hasAttribute("cap")) { r.setAttribute("cap", "all"); n++; }
        // tell substituting viewers whether this is a serif / sans / script / mono face
        if (font && !lat.hasAttribute("pitchFamily")) { lat.setAttribute("pitchFamily", font.script ? "66" : font.mono ? "49" : font.serif ? "18" : "34"); n++; }
      }
    }
    for (const sp of Array.from(doc.getElementsByTagNameNS(NS_P, "sp"))) {
      let anc = sp.parentNode, grouped = false;
      while (anc && anc.nodeType === 1) { if (anc.localName === "grpSp") { grouped = true; break; } anc = anc.parentNode; }
      if (grouped) continue; // child coordinates are scaled; leave them alone
      const tx = Array.from(sp.childNodes).find((c) => c.nodeType === 1 && c.localName === "txBody");
      if (!tx) continue;
      const body = Array.from(tx.childNodes).find((c) => c.nodeType === 1 && c.localName === "bodyPr");
      const paras = Array.from(tx.childNodes).filter((c) => c.nodeType === 1 && c.localName === "p");
      if (!body || body.hasAttribute("wrap")) continue;
      const hasText = (tx.textContent || "").trim().length > 0;
      if (!hasText) continue;
      const ext = sp.getElementsByTagNameNS(A, "ext")[0];
      if (!ext) continue;
      const cy = +ext.getAttribute("cy");
      let single = false;
      if (paras.length === 1 && !paras[0].getElementsByTagNameNS(A, "br").length) {
        const p = paras[0];
        const sp100 = p.getElementsByTagNameNS(A, "spcPts")[0];
        let maxSz = 0;
        for (const r of Array.from(p.getElementsByTagNameNS(A, "rPr"))) maxSz = Math.max(maxSz, +r.getAttribute("sz") || 0);
        const lineEmu = sp100 ? (+sp100.getAttribute("val") / 100) * 12700 : (maxSz / 100) * 12700 * 1.2;
        single = !!lineEmu && cy <= lineEmu * 1.5;
      }
      if (single) { body.setAttribute("wrap", "none"); n++; continue; }
      // Multi-line: Canva sized the box exactly to its text. If a viewer substitutes a wider font the box would
      // grow downward over the next element; shrink-on-overflow keeps the text inside Canva's box instead.
      const grow = kids(body, "spAutoFit")[0];
      if (grow) { body.replaceChild(doc.createElementNS(A, "a:normAutofit"), grow); n++; }
    }
    return n;
  }


  // ---- exact look: draw text with Canva's own embedded fonts and place it as a picture -----------------
  const REL_IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

  function runStyle(r) {
    const A = NS_A;
    const rp = Array.from(r.childNodes).find((c) => c.nodeType === 1 && c.localName === "rPr");
    const st = { sz: 18, b: false, i: false, u: false, strike: false, spc: 0, cap: false, color: "#000", alpha: 1, face: null, bad: false };
    if (!rp) return st;
    st.sz = (+rp.getAttribute("sz") || 1800) / 100;
    st.b = rp.getAttribute("b") === "true" || rp.getAttribute("b") === "1";
    st.i = rp.getAttribute("i") === "true" || rp.getAttribute("i") === "1";
    st.u = !!rp.getAttribute("u") && rp.getAttribute("u") !== "none";
    st.strike = !!rp.getAttribute("strike") && rp.getAttribute("strike") !== "noStrike";
    st.spc = (+rp.getAttribute("spc") || 0) / 100;
    st.cap = rp.getAttribute("cap") === "all";
    for (const c of Array.from(rp.childNodes)) {
      if (c.nodeType !== 1) continue;
      if (c.localName === "solidFill") {
        const clr = Array.from(c.childNodes).find((x) => x.nodeType === 1);
        if (clr && clr.localName === "srgbClr") {
          st.color = "#" + clr.getAttribute("val");
          const al = Array.from(clr.childNodes).find((x) => x.nodeType === 1 && x.localName === "alpha");
          if (al) st.alpha = (+al.getAttribute("val")) / 100000;
        } else st.bad = true; // theme colours etc.
      } else if (c.localName === "latin") st.face = c.getAttribute("typeface");
      else if (["gradFill", "blipFill", "pattFill", "ln", "effectLst", "highlight"].includes(c.localName)) st.bad = true;
    }
    return st;
  }

  async function rasterizeText(zip, doc, path, faces, env, tag) {
    const A = NS_A, out = { done: 0, skipped: 0, why: {} };
    const skip = (w) => { out.skipped++; out.why[w] = (out.why[w] || 0) + 1; };
    const kids = (el, ln) => Array.from(el.childNodes).filter((c) => c.nodeType === 1 && c.localName === ln);
    const D = env.document;
    const relPath = path.replace(/slides\/(slide\d+\.xml)$/, "slides/_rels/$1.rels");
    let relDoc = null, relRoot = null, nextRel = 1, nextImg = 1;
    if (!env.fontsLoaded) { // register every embedded face once, each under its own family so no faux-bold is ever applied
      env.fontsLoaded = true;
      for (const ent of Object.values(faces)) {
        ent.fam = {};
        for (const slot of Object.keys(ent.bytes || {})) {
          try {
            const name = "cvx" + (env.fontSeq = (env.fontSeq || 0) + 1);
            const ff = new env.FontFace(name, ent.bytes[slot].slice().buffer);
            await ff.load(); D.fonts.add(ff); ent.fam[slot] = name;
          } catch (e) { /* unusable face: the run falls back to the regular face */ }
        }
      }
    }
    for (const sp of Array.from(doc.getElementsByTagNameNS(NS_P, "sp"))) {
      const tx = kids(sp, "txBody")[0];
      if (!tx || !(tx.textContent || "").trim()) continue;
      let anc = sp.parentNode, grouped = false, sx = 1, sy = 1, badGroup = false;
      while (anc && anc.nodeType === 1) {
        if (anc.localName === "grpSp") {
          grouped = true;
          const gx = kids(kids(anc, "grpSpPr")[0] || anc, "xfrm")[0];
          if (gx) {
            if ((+gx.getAttribute("rot") || 0) || gx.getAttribute("flipH") === "true" || gx.getAttribute("flipV") === "true" || gx.getAttribute("flipH") === "1" || gx.getAttribute("flipV") === "1") badGroup = true;
            const ge = kids(gx, "ext")[0], gc = kids(gx, "chExt")[0];
            if (ge && gc && +gc.getAttribute("cx") && +gc.getAttribute("cy")) { sx *= +ge.getAttribute("cx") / +gc.getAttribute("cx"); sy *= +ge.getAttribute("cy") / +gc.getAttribute("cy"); }
          }
        }
        anc = anc.parentNode;
      }
      const spPr = kids(sp, "spPr")[0], nv = kids(sp, "nvSpPr")[0];
      const cnv = nv && kids(nv, "cNvPr")[0];
      const xf = spPr && kids(spPr, "xfrm")[0];
      const bodyPr = kids(tx, "bodyPr")[0];
      const visible = spPr && Array.from(spPr.childNodes).some((c) => c.nodeType === 1 &&
        (["solidFill", "gradFill", "blipFill", "pattFill"].includes(c.localName) || (c.localName === "ln" && !kids(c, "noFill").length)));
      const geom = spPr && kids(spPr, "prstGeom")[0];
      if (badGroup) { skip("inside a rotated/flipped group"); continue; }
      if (!xf || !cnv || !bodyPr) { skip("placeholder/inherited"); continue; }
      if (visible) { skip("shape with fill/outline"); continue; }
      if ((geom && geom.getAttribute("prst") !== "rect") || kids(spPr, "custGeom").length) { skip("non-rectangular shape"); continue; }
      const off = kids(xf, "off")[0], ext = kids(xf, "ext")[0];
      const X = +off.getAttribute("x"), Y = +off.getAttribute("y"), CX = +ext.getAttribute("cx"), CY = +ext.getAttribute("cy");
      const rot = +xf.getAttribute("rot") || 0;
      const insL = +(bodyPr.getAttribute("lIns") ?? 91440) / 12700, insR = +(bodyPr.getAttribute("rIns") ?? 91440) / 12700;
      const insT = +(bodyPr.getAttribute("tIns") ?? 45720) / 12700, insB = +(bodyPr.getAttribute("bIns") ?? 45720) / 12700;
      const wrap = bodyPr.getAttribute("wrap") !== "none";
      const boxW = CX * sx / 12700, boxH = CY * sy / 12700;
      const R = Math.max(1.5, Math.min(4, 4096 / Math.max(boxW, boxH, 1)));
      const cv = D.createElement("canvas"); const g = cv.getContext("2d");
      let bad = false;

      const setFont = (st) => {
        const ent = faces[st.face];
        let fam, weight = "400", style = "normal";
        if (ent && ent.fam) {
          const slot = st.b && st.i && ent.fam.boldItalic ? "boldItalic" : st.i && ent.fam.italic ? "italic" : st.b && ent.fam.bold ? "bold" : "regular";
          fam = '"' + (ent.fam[slot] || ent.fam.regular) + '"';
        } else { fam = '"' + (st.face || "Arial") + '", sans-serif'; weight = st.b ? "700" : "400"; style = st.i ? "italic" : "normal"; }
        g.font = `${style} ${weight} ${st.sz * R}px ${fam}`;
        g.letterSpacing = `${st.spc * R}px`;
        return g;
      };
      // tokenise paragraphs
      const paras = kids(tx, "p").map((p) => {
        const pPr = kids(p, "pPr")[0];
        const algn = (pPr && pPr.getAttribute("algn")) || "l";
        const marL = pPr ? (+pPr.getAttribute("marL") || 0) / 12700 : 0, indent = pPr ? (+pPr.getAttribute("indent") || 0) / 12700 : 0;
        let lnPts = null, lnPct = null, bef = 0, aft = 0, bu = null;
        if (pPr) {
          const ls = kids(pPr, "lnSpc")[0];
          if (ls) { const pt = kids(ls, "spcPts")[0], pc = kids(ls, "spcPct")[0]; if (pt) lnPts = +pt.getAttribute("val") / 100; else if (pc) lnPct = +pc.getAttribute("val") / 100000; }
          const sb = kids(pPr, "spcBef")[0], sa = kids(pPr, "spcAft")[0];
          if (sb && kids(sb, "spcPts")[0]) bef = +kids(sb, "spcPts")[0].getAttribute("val") / 100;
          if (sa && kids(sa, "spcPts")[0]) aft = +kids(sa, "spcPts")[0].getAttribute("val") / 100;
          const bc = kids(pPr, "buChar")[0]; if (bc) bu = bc.getAttribute("char");
          if (kids(pPr, "buAutoNum")[0]) bad = true;
        }
        const items = []; let first = null;
        for (const c of Array.from(p.childNodes)) {
          if (c.nodeType !== 1) continue;
          if (c.localName === "br") { items.push({ br: true }); continue; }
          if (c.localName === "fld") { bad = true; continue; }
          if (c.localName !== "r") continue;
          const st = runStyle(c); if (st.bad) bad = true;
          if (!first) first = st;
          let t = (kids(c, "t")[0] || { textContent: "" }).textContent;
          if (st.cap) t = t.toUpperCase();
          for (const part of t.split(/(\s+)/)) if (part) items.push({ text: part, space: /^\s+$/.test(part), st });
        }
        const end = kids(p, "endParaRPr")[0];
        return { algn, marL, indent, lnPts, lnPct, bef, aft, bu, items, first };
      });
      if (bad) { skip("unsupported text style"); continue; }
      // layout
      const lines = []; let y = insT;
      let minX = 0, maxX = boxW;
      for (let pi = 0; pi < paras.length; pi++) {
        const P = paras[pi];
        const ref = P.first || { sz: 18, face: null };
        y += pi ? P.bef : 0;
        const avail = Math.max(1, boxW - insL - insR - P.marL - (P.bu ? 0 : Math.min(0, P.indent)));
        const startX = (n) => insL + P.marL + (!P.bu && n === 0 ? P.indent : 0);
        let cur = [], curW = 0, n = 0;
        const flush = (forced) => {
          while (cur.length && cur[cur.length - 1].space) { curW -= cur[cur.length - 1].w; cur.pop(); }
          const maxSz = cur.reduce((m, it) => Math.max(m, it.st.sz), ref.sz);
          const lh = P.lnPts != null ? P.lnPts : (P.lnPct != null ? P.lnPct : 1) * maxSz * 1.2;
          const m = setFont(cur.length ? cur.reduce((a, b) => (b.st.sz > a.st.sz ? b : a)).st : ref).measureText("Hg");
          const asc = m.fontBoundingBoxAscent / R, desc = m.fontBoundingBoxDescent / R;
          const base = y + (lh - (asc + desc)) / 2 + asc;
          const availW = boxW - insL - insR - P.marL - (!P.bu && n === 0 ? P.indent : 0);
          let x0 = startX(n);
          if (P.algn === "ctr") x0 += (availW - curW) / 2; else if (P.algn === "r") x0 += availW - curW;
          lines.push({ items: cur, x: x0, base, lh, bullet: P.bu && n === 0 ? { ch: P.bu, st: ref, x: insL + P.marL + P.indent } : null, w: curW });
          minX = Math.min(minX, x0); maxX = Math.max(maxX, x0 + curW);
          y += lh; n++; cur = []; curW = 0;
        };
        for (const it of P.items) {
          if (it.br) { flush(true); continue; }
          it.w = setFont(it.st).measureText(it.text).width / R;
          const avl = boxW - insL - insR - P.marL - (!P.bu && n === 0 ? P.indent : 0);
          if (wrap && !it.space && cur.length && curW + it.w > avl + 0.01) flush(false);
          if (!cur.length && it.space) continue; // no leading spaces on wrapped lines
          cur.push(it); curW += it.w;
        }
        if (cur.length || !P.items.length) flush(false);
        y += P.aft;
      }
      const contentH = y - insT + insB;
      let top = 0, bottom = Math.max(boxH, insT + contentH);
      let shiftY = 0;
      const anchor = bodyPr.getAttribute("anchor") || "t";
      if (anchor === "ctr") shiftY = (boxH - insT - insB - (y - insT)) / 2; else if (anchor === "b") shiftY = boxH - insB - y;
      if (shiftY < 0 && !wrap) shiftY = shiftY; // keep
      const maxSzAll = Math.max(...paras.map((p) => (p.first ? p.first.sz : 18)));
      const pad = maxSzAll * 0.4;
      let L = Math.min(0, minX) - pad, T = Math.min(0, shiftY) - pad, RR = Math.max(boxW, maxX) + pad, BB = Math.max(boxH, y + shiftY + insB) + pad;
      if (rot) { L = -pad; T = -pad; RR = boxW + pad; BB = boxH + pad; }
      const W = Math.ceil((RR - L) * R), H = Math.ceil((BB - T) * R);
      if (W < 2 || H < 2 || W > 8192 || H > 8192) { out.skipped++; continue; }
      cv.width = W; cv.height = H;
      g.clearRect(0, 0, W, H);
      g.textBaseline = "alphabetic";
      const draw = (it, x, base) => {
        const st = it.st; setFont(st);
        g.globalAlpha = st.alpha; g.fillStyle = st.color;
        g.fillText(it.text, (x - L) * R, (base - T + shiftY) * R);
        if (st.u || st.strike) {
          const th = Math.max(1, st.sz * 0.065 * R);
          if (st.u) g.fillRect((x - L) * R, (base - T + shiftY + st.sz * 0.12) * R, it.w * R, th);
          if (st.strike) g.fillRect((x - L) * R, (base - T + shiftY - st.sz * 0.3) * R, it.w * R, th);
        }
        g.globalAlpha = 1;
      };
      for (const ln of lines) {
        let x = ln.x;
        if (ln.bullet) draw({ text: ln.bullet.ch, st: ln.bullet.st, w: 0 }, ln.bullet.x, ln.base);
        for (const it of ln.items) { draw(it, x, ln.base); x += it.w; }
      }
      const blob = await new Promise((res) => cv.toBlob(res, "image/png"));
      if (!blob) { out.skipped++; continue; }
      const data = new Uint8Array(await blob.arrayBuffer());
      const mediaName = `cvtext_${tag}_${nextImg++}.png`;
      zip.file("ppt/media/" + mediaName, data, { createFolders: false });
      if (!relDoc) {
        const rf = zip.file(relPath);
        relDoc = new DOMParser().parseFromString(rf ? await rf.async("string") : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>', "application/xml");
        relRoot = relDoc.documentElement;
      }
      const rid = "rIdCvt" + nextRel++;
      const rel = relDoc.createElementNS(relRoot.namespaceURI, "Relationship");
      rel.setAttribute("Id", rid); rel.setAttribute("Type", REL_IMG); rel.setAttribute("Target", "../media/" + mediaName);
      relRoot.appendChild(rel);
      const e = (v) => Math.round(v * 12700);
      const picXml = `<p:pic xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><p:nvPicPr><p:cNvPr id="${cnv.getAttribute("id")}" name="${(cnv.getAttribute("name") || "Text").replace(/"/g, "&quot;")}" descr="${(tx.textContent || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;").slice(0, 200)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm${rot ? ` rot="${rot}"` : ""}><a:off x="${Math.round(X + e(L) / sx)}" y="${Math.round(Y + e(T) / sy)}"/><a:ext cx="${Math.round(e(RR - L) / sx)}" cy="${Math.round(e(BB - T) / sy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      const pic = doc.importNode(new DOMParser().parseFromString(picXml, "application/xml").documentElement, true);
      sp.parentNode.replaceChild(pic, sp);
      out.done++;
    }
    if (relDoc) zip.file(relPath, new XMLSerializer().serializeToString(relDoc), { createFolders: false });
    return out;
  }

  async function injectIntoPptx(ir, pptxData, deps, options) {
    const { JSZip, DOMParser, XMLSerializer } = deps;
    const fixText = !options || options.fixText !== false;
    const compat = !!(options && options.compat);
    const dom = { parse: (s) => new DOMParser().parseFromString(s, "application/xml") };
    const ser = new XMLSerializer();
    const zip = await JSZip.loadAsync(pptxData, { createFolders: false });
    const info = await slideParts(zip, dom);
    const raster = !!(options && options.rasterText);
    const env = { document: deps.document || (typeof document !== "undefined" ? document : null), FontFace: deps.FontFace || (typeof FontFace !== "undefined" ? FontFace : null) };
    const canRaster = raster && env.document && env.FontFace;
    const faces = (fixText || raster) ? await embeddedFaces(zip, dom) : {};
    const warnings = [];
    if (raster && !canRaster) warnings.push("Exact-look text needs a browser; it was skipped.");
    if (info.paths.length !== ir.slides.length)
      warnings.push(`The PPTX has ${info.paths.length} slides but the Canva design has ${ir.slides.length}. Slides are matched in order, so make sure the export is of the same design.`);
    const scale = info.width / ir.deck.width;
    const report = { slides: [], warnings };
    const n = Math.min(info.paths.length, ir.slides.length);
    for (let i = 0; i < n; i++) {
      const s = ir.slides[i], path = info.paths[i];
      const doc = dom.parse(await zip.file(path).async("string"));
      const root = doc.documentElement;
      let removed = 0;
      // drop existing timing/transition
      for (const el of Array.from(root.childNodes)) {
        if (el.nodeType !== 1) continue;
        const isTr = el.localName === "transition" && el.namespaceURI === NS_P;
        const isAlt = el.localName === "AlternateContent" && el.getElementsByTagNameNS(NS_P, "transition").length > 0;
        const isTm = el.localName === "timing" && el.namespaceURI === NS_P;
        if (isTr || isAlt || isTm) { root.removeChild(el); removed++; }
      }
      const textFixes = fixText ? fixTextLayout(doc, faces, { compat }) : 0;
      const shapes = shapeInfo(doc);
      const byId = new Map(s.elements.map((e) => [e.id, e]));
      const used = new Set(), spids = {};
      let unmatched = 0;
      for (const a of s.animations) {
        const el = byId.get(a.target);
        if (!el) continue;
        const { shape, cost } = matchShape(el, shapes, scale, used);
        if (!shape || cost > 100 * scale) { unmatched++; warnings.push(`Slide ${i + 1}: couldn't find the PPTX shape for an animated ${el.type}${el.text ? ` ("${el.text.slice(0, 30)}")` : ""}; its animation was skipped.`); continue; }
        used.add(shape.id); spids[a.target] = shape.id;
      }
      const sub = [];
      const trXml = buildTransitionXml(s.transition);
      if (trXml) sub.push(trXml);
      const wl = [];
      const tmXml = buildTimingXml(s.animations, spids, wl);
      wl.forEach((w) => warnings.push(`Slide ${i + 1}: ${w}`));
      if (tmXml) sub.push(tmXml);
      const extLst = Array.from(root.childNodes).find((c) => c.nodeType === 1 && c.localName === "extLst");
      for (const x of sub) {
        const node = doc.importNode(dom.parse(x).documentElement, true);
        extLst ? root.insertBefore(node, extLst) : root.appendChild(node);
      }
      let rz = { done: 0, skipped: 0 };
      if (canRaster) {
        rz = await rasterizeText(zip, doc, path, faces, env, i + 1);
        if (rz.skipped) warnings.push(`Slide ${i + 1}: ${rz.skipped} text box(es) stayed live text (${Object.entries(rz.why).map(([k, v]) => v + " " + k).join(", ")}).`);
      }
      if (sub.length || removed || textFixes || rz.done) { // leave untouched slides byte-identical
        const body = ser.serializeToString(doc).replace(/^\s*<\?xml[^>]*\?>\s*/, ""); // some serializers add their own declaration
        zip.file(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body, { createFolders: false });
      }
      report.slides.push({
        textFixes, textPictures: rz.done,
        slide: i + 1, animations: s.animations.length, applied: Object.keys(spids).length, unmatched,
        clicks: s.animations.filter((a) => a.trigger === "onClick").length,
        auto: s.animations.filter((a) => a.trigger !== "onClick").length,
        transition: !!s.transition,
      });
    }
    const data = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { data, report };
  }

  function breakdown(ir) {
    return ir.slides.map((s, i) => ({
      slide: i + 1,
      auto: s.animations.filter((a) => a.canva.step == null).length,
      clickSteps: s.animations.filter((a) => a.trigger === "onClick").length,
      total: s.animations.length,
    })).filter((r) => r.total);
  }

  /** One row per animation, described so people can recognise it and untick unwanted ones. */
  function listAnimations(ir) {
    const rows = [];
    ir.slides.forEach((s, si) => {
      const byId = new Map(s.elements.map((e) => [e.id, e]));
      s.animations.forEach((a) => {
        const el = byId.get(a.target) || {};
        const full = el.w >= ir.deck.width * 0.9 && el.h >= ir.deck.height * 0.9;
        const what = el.text ? `text "${el.text.replace(/\s+/g, " ").slice(0, 28)}"`
          : `${el.type || "element"}${full ? " (full slide)" : el.w ? ` (${Math.round(el.w)}×${Math.round(el.h)})` : ""}`;
        rows.push({ key: s.id + ":" + a.target, slide: si + 1, what, effect: a.canva.name, trigger: a.trigger === "onClick" ? "on click" : (a.canva.step != null ? "with click" : "automatic") });
      });
    });
    return rows;
  }

  /** Support file: everything about the design's slides (no images), so odd cases can be diagnosed. */
  function supportData(html) {
    const boot = loadBootstrap(html);
    return { tool: "canva-animations-to-pptx", at: new Date().toISOString(), pages: findPages(boot) };
  }

  function summarize(ir) {
    const slides = ir.slides.length;
    const anims = ir.slides.reduce((n, s) => n + s.animations.length, 0);
    const clicks = ir.slides.reduce((n, s) => n + s.animations.filter((a) => a.trigger === "onClick").length, 0);
    return { slides, animations: anims, clickSteps: clicks, title: ir.title };
  }

  return { _embeddedFaces: embeddedFaces, DEFAULTS, ELEMENT_ANIMS, TO_PPT, jsUnescape, loadBootstrap, findPages, convert, extractFromHtml, buildTimingXml, buildTransitionXml, injectIntoPptx, summarize, breakdown, listAnimations, supportData };
});
