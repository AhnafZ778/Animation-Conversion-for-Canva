const test = require("node:test"), assert = require("node:assert");
const fs = require("fs"), path = require("path");
const JSZip = require("jszip");
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
const C = require("../core.js");

const el = (o) => Object.assign({ "A?": "K", A: 100, B: 100, D: 300, C: 80, _: "id" + Math.random() }, o);
const design = (pages) => ({ page: { Pm: { E: { draft: { content: { A: pages } } } } } });
const page = (E, extra) => Object.assign({ "A?": "i", C: { A: 1920, B: 1080 }, E }, extra || {});

test("auto vs click vs shared step", () => {
  const E = [
    el({ _: "auto1", X: { "A?": "B", A: 4 } }),
    el({ _: "c1", Z: "0", X: { "A?": "B", A: 8 } }),
    el({ _: "c2", Z: "1", X: { "A?": "B", A: 26, B: { C: "E" } } }),
    el({ _: "c3", Z: "1", X: { "A?": "B", A: 7 } }),
  ];
  const { ir } = C.convert(design([page(E, { W: [{ A: "1" }, { A: "0" }] })]));
  const a = ir.slides[0].animations;
  // auto first, then click steps in W order ("1" before "0"), members of a step together
  assert.deepStrictEqual(a.map((x) => [x.target, x.trigger]), [["auto1", "withPrevious"], ["c2", "onClick"], ["c3", "withPrevious"], ["c1", "onClick"]]);
});

test("step id with no/partial W list is still a click, ordered by W then stack", () => {
  const E = [el({ _: "x", Z: "9", X: { "A?": "B", A: 4 } }), el({ _: "y", Z: "0", X: { "A?": "B", A: 4 } }), el({ _: "z", Z: "7", X: { "A?": "B", A: 4 } })];
  const { ir, warnings } = C.convert(design([page(E, { W: [{ A: "0" }] })]));
  assert.deepStrictEqual(ir.slides[0].animations.map((a) => [a.target, a.trigger]), [["y", "onClick"], ["x", "onClick"], ["z", "onClick"]]);
  assert.ok(warnings.some((w) => /isn't in Canva's step order list/.test(w)));
  const none = C.convert(design([page([el({ _: "a", Z: "3", X: { "A?": "B", A: 4 } })])]));
  assert.strictEqual(none.ir.slides[0].animations[0].trigger, "onClick");
});

test("unrecognised animation fields are surfaced", () => {
  const E = [el({ _: "a", Z: "0", Y9: 1, X: { "A?": "B", A: 4, K: 2, B: { Z: 1 } } })];
  const { warnings } = C.convert(design([page(E, { W: [{ A: "0" }] })]));
  assert.ok(warnings.some((w) => /element\.Y9/.test(w) && /animation\.K/.test(w) && /options\.Z/.test(w)));
});

test("breakdown + support data", () => {
  const E = [el({ _: "a", X: { "A?": "B", A: 4 } }), el({ _: "b", Z: "0", X: { "A?": "B", A: 4 } })];
  const { ir } = C.convert(design([page(E, { W: [{ A: "0" }] })]));
  assert.deepStrictEqual(C.breakdown(ir), [{ slide: 1, auto: 1, clickSteps: 1, total: 2 }]);
});

test("unknown effect id falls back to fade with a note", () => {
  const { ir, warnings } = C.convert(design([page([el({ _: "a", X: { "A?": "B", A: 999 } })])]));
  assert.strictEqual(ir.slides[0].animations[0].effect, "fade");
  assert.ok(warnings.some((w) => /unknown-999/.test(w)));
});

test("slides with no animations / non-geometric elements don't crash", () => {
  const { ir } = C.convert(design([page([]), page([{ "A?": "K" }]), { "A?": "i" }]));
  assert.strictEqual(ir.slides.length, 3);
});

test("durations: explicit in microseconds, otherwise default", () => {
  const E = [el({ _: "a", X: { "A?": "B", A: 4, B: { G: { A: 1000000 } } } }), el({ _: "b", X: { "A?": "B", A: 4 } })];
  const a = C.convert(design([page(E)]), { defaultDurationMs: 600 }).ir.slides[0].animations;
  assert.deepStrictEqual(a.map((x) => x.durationMs), [1000, 600]);
});

test("stagger option", () => {
  const E = [0, 1, 2].map((i) => el({ _: "e" + i, X: { "A?": "B", A: 4 } }));
  const a = C.convert(design([page(E)]), { staggerMs: 100 }).ir.slides[0].animations;
  assert.deepStrictEqual(a.map((x) => x.delayMs || 0), [0, 100, 200]);
});

test("wrong page gives a friendly error", () => {
  assert.throws(() => C.extractFromHtml("<html>nothing</html>"), /doesn't contain a Canva design/);
});

test("js string unescape", () => {
  assert.strictEqual(C.jsUnescape("a\\\\nb\\'c\\x41\\u0042"), "a\\nb'cAB");
});

test("mismatched slide count warns, and works on the real export", async (t) => {
  const root = path.resolve(__dirname, "..", "..");
  const html = path.join(root, "samples", "design-page.html"), pptx = path.join(root, "samples", "export.pptx");
  if (!fs.existsSync(html) || !fs.existsSync(pptx)) return t.skip("needs private sample files");
  const { ir } = C.extractFromHtml(fs.readFileSync(html, "utf8"));
  ir.slides.pop();
  const { report } = await C.injectIntoPptx(ir, fs.readFileSync(pptx), { JSZip, DOMParser, XMLSerializer });
  assert.ok(report.warnings.some((w) => /slides but the Canva design has/.test(w)));
});

test("shape that can't be found is reported, not silently dropped", async (t) => {
  const root = path.resolve(__dirname, "..", "..");
  const html = path.join(root, "samples", "design-page.html"), pptx = path.join(root, "samples", "export.pptx");
  if (!fs.existsSync(html) || !fs.existsSync(pptx)) return t.skip("needs private sample files");
  const { ir } = C.extractFromHtml(fs.readFileSync(html, "utf8"));
  const s = ir.slides[5]; s.elements.forEach((e) => { e.x += 5000; e.y += 5000; });   // nothing will match
  const { report } = await C.injectIntoPptx(ir, fs.readFileSync(pptx), { JSZip, DOMParser, XMLSerializer });
  assert.ok(report.warnings.some((w) => /Slide 6: couldn't find/.test(w)));
  assert.strictEqual(report.slides[5].applied, 0);
});

test("real-world structure: step linked via Z, via per-element W map, and plain automatic", () => {
  const pages = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "real-world-pages.json"), "utf8"));
  const { ir } = C.convert(design(pages));
  const trig = ir.slides.map((s) => s.animations.map((a) => a.trigger));
  assert.deepStrictEqual(trig.slice(0, 4), [["onClick"], ["onClick"], ["withPrevious"], ["withPrevious"]]);
});

test("elements at x=0 or y=0 (coordinate omitted by Canva) are kept", () => {
  const e = { "A?": "I", B: 5, D: 100, C: 100, _: "top", X: { "A?": "B", A: 4 } };   // no A => y = 0
  const f = { "A?": "I", A: 5, D: 100, C: 100, _: "left", X: { "A?": "B", A: 4 } };  // no B => x = 0
  const { ir } = C.convert(design([page([e, f])]));
  assert.strictEqual(ir.slides[0].animations.length, 2);
  assert.deepStrictEqual([ir.slides[0].elements[0].y, ir.slides[0].elements[1].x], [0, 0]);
});

test("listAnimations describes elements; skip option removes unticked ones", () => {
  const E = [el({ _: "t", text: undefined, a: { C: { A: ["Hello world"] } }, X: { "A?": "B", A: 4 } }), el({ _: "bg", D: 1920, C: 1080, "A?": "I", X: { "A?": "B", A: 29 } })];
  const { ir } = C.convert(design([page(E)]));
  const rows = C.listAnimations(ir);
  assert.ok(/Hello world/.test(rows[0].what)); assert.ok(/full slide/.test(rows[1].what));
  const kept = C.extractFromHtml.length && C.convert(design([page(E)])).ir;
  const skipped = (() => { const r = C.convert(design([page(E)])); const sk = new Set([rows[1].key]); r.ir.slides.forEach((s) => { s.animations = s.animations.filter((a) => !sk.has(s.id + ":" + a.target)); }); return r.ir; })();
  assert.strictEqual(skipped.slides[0].animations.length, 1); assert.strictEqual(kept.slides[0].animations.length, 2);
});

test("text fix: removes doubled bold only for regular-only embedded fonts; one-line boxes get wrap=none", async (t) => {
  const root = path.resolve(__dirname, "..", "..");
  const html = path.join(root, "samples", "design-page.html"), pptx = path.join(root, "samples", "export.pptx");
  if (!fs.existsSync(html) || !fs.existsSync(pptx)) return t.skip("needs private sample files");
  const { ir } = C.extractFromHtml(fs.readFileSync(html, "utf8"));
  const deps = { JSZip, DOMParser, XMLSerializer };
  const on = await C.injectIntoPptx(ir, fs.readFileSync(pptx), deps);
  const off = await C.injectIntoPptx(ir, fs.readFileSync(pptx), deps, { fixText: false });
  assert.ok(on.report.slides.reduce((a, s) => a + s.textFixes, 0) > 0);
  const z1 = await JSZip.loadAsync(on.data), z0 = await JSZip.loadAsync(off.data);
  const x1 = await z1.file("ppt/slides/slide3.xml").async("string"), x0 = await z0.file("ppt/slides/slide3.xml").async("string");
  assert.ok(/b="true"/.test(x0) && !/b="true"/.test(x1));
  assert.deepStrictEqual(x1.match(/<a:t>[^<]*<\/a:t>/g), x0.match(/<a:t>[^<]*<\/a:t>/g)); // text itself untouched
  assert.ok(!Object.keys(z1.files).some((n) => n.endsWith("/") && n !== "")); // no stray folder entries
});

test("compat mode swaps to Arial/Times with width-matched sizes; default keeps Canva fonts and adds hints", async (t) => {
  const root = path.resolve(__dirname, "..", "..");
  const html = path.join(root, "samples", "design-page.html"), pptx = path.join(root, "samples", "export.pptx");
  if (!fs.existsSync(html) || !fs.existsSync(pptx)) return t.skip("needs private sample files");
  const { ir } = C.extractFromHtml(fs.readFileSync(html, "utf8"));
  const deps = { JSZip, DOMParser, XMLSerializer };
  const def = await JSZip.loadAsync((await C.injectIntoPptx(ir, fs.readFileSync(pptx), deps)).data);
  const cmp = await JSZip.loadAsync((await C.injectIntoPptx(ir, fs.readFileSync(pptx), deps, { compat: true })).data);
  const d = await def.file("ppt/slides/slide3.xml").async("string"), c = await cmp.file("ppt/slides/slide3.xml").async("string");
  assert.ok(/typeface="Roca One Heavy" pitchFamily="18"/.test(d));
  assert.ok(/typeface="Times New Roman"/.test(c) && /typeface="Arial"/.test(c) && !/Telegraf/.test(c));
  assert.deepStrictEqual(c.match(/<a:t>[^<]*<\/a:t>/g), d.match(/<a:t>[^<]*<\/a:t>/g));
});


// ---- synthetic PPTX (no private files needed) -------------------------------------------------------
async function syntheticPptx() {
  const z = new JSZip();
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>');
  z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  z.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8"?><p:presentation ${NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="18288000" cy="10287000"/></p:presentation>`);
  z.file("ppt/_rels/presentation.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="x" Target="slides/slide2.xml"/></Relationships>');
  const sp = (id, x, y, cx, cy, body, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="true"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr anchor="t" lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${text}</p:txBody></p:sp>`;
  const para = (t) => `<a:p><a:pPr algn="l"><a:lnSpc><a:spcPts val="5000"/></a:lnSpc></a:pPr><a:r><a:rPr lang="en-US" b="true" sz="4000"><a:latin typeface="Some Font"/></a:rPr><a:t>${t}</a:t></a:r></a:p>`;
  const slide = (shapes) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld></p:sld>`;
  z.file("ppt/slides/slide1.xml", slide(
    sp(2, 952500, 476250, 7620000, 635000, "", para("One line")) +
    sp(3, 952500, 2000000, 7620000, 1270000, "", para("First paragraph") + para("Second paragraph"))));
  z.file("ppt/slides/slide2.xml", slide(sp(2, 952500, 476250, 7620000, 635000, "", para("Untouched slide"))));
  return z.generateAsync({ type: "uint8array" });
}
const px = (emu) => emu / 9525;
const synthDesign = () => {
  const mk = (id, y, h, extra) => el(Object.assign({ _: id, "A?": "K", A: px(y), B: px(952500), D: px(7620000), C: px(h) }, extra));
  return C.convert(design([
    page([mk("t1", 476250, 635000, { X: { "A?": "B", A: 4 } }), mk("t2", 2000000, 1270000, { Z: "0", X: { "A?": "B", A: 8 } })]),
    page([mk("u1", 476250, 635000)]),
  ])).ir;
};

test("injects timing into the right shapes of a synthetic PPTX", async () => {
  const { data, report } = await C.injectIntoPptx(synthDesign(), await syntheticPptx(), { JSZip, DOMParser, XMLSerializer });
  assert.deepStrictEqual(report.slides.map((s) => s.applied), [2, 0]);
  const x = await (await JSZip.loadAsync(data)).file("ppt/slides/slide1.xml").async("string");
  assert.match(x, /<p:timing/);
  assert.match(x, /<p:spTgt spid="2"/);
  assert.match(x, /<p:spTgt spid="3"/);
  assert.match(x, /delay="indefinite"/); // t2 is an on-click step
});

test("slides with nothing to change are left byte-identical (text fix off)", async () => {
  const src = await syntheticPptx();
  const { data } = await C.injectIntoPptx(synthDesign(), src, { JSZip, DOMParser, XMLSerializer }, { fixText: false });
  const a = await JSZip.loadAsync(src), b = await JSZip.loadAsync(data);
  for (const f of ["ppt/slides/slide2.xml", "ppt/presentation.xml", "[Content_Types].xml"])
    assert.strictEqual(await b.file(f).async("string"), await a.file(f).async("string"), f);
});

test("text fix: one-line box stays on one line, multi-line box shrinks instead of growing; opt-out works", async () => {
  const deps = { JSZip, DOMParser, XMLSerializer };
  const on = await (await JSZip.loadAsync((await C.injectIntoPptx(synthDesign(), await syntheticPptx(), deps)).data)).file("ppt/slides/slide1.xml").async("string");
  assert.strictEqual((on.match(/wrap="none"/g) || []).length, 1);
  assert.strictEqual((on.match(/<a:normAutofit/g) || []).length, 1);
  const off = await (await JSZip.loadAsync((await C.injectIntoPptx(synthDesign(), await syntheticPptx(), deps, { fixText: false })).data)).file("ppt/slides/slide1.xml").async("string");
  assert.ok(!/wrap="none"/.test(off) && !/normAutofit/.test(off));
  assert.deepStrictEqual(on.match(/<a:t>[^<]*<\/a:t>/g), off.match(/<a:t>[^<]*<\/a:t>/g)); // text itself is never touched
});
