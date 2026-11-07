const $ = (id) => document.getElementById(id);
if (typeof chrome !== "undefined" && chrome.runtime && chrome.runtime.getManifest) $("ver").textContent = "v" + chrome.runtime.getManifest().version;
let html = null, pptxFile = null, srcInfo = "";

function setSource(h, label) {
  html = h;
  try {
    const { ir } = CanvaPptx.extractFromHtml(h);
    const s = CanvaPptx.summarize(ir);
    $("source").className = "";
    const auto = s.animations - s.clickSteps;
    $("source").innerHTML = `<div class="figs"><div class="fig"><b>${s.slides}</b><span>slides</span></div><div class="fig"><b>${s.animations}</b><span>animations</span></div><div class="fig"><b>${s.clickSteps}</b><span>on click</span></div></div><p class="src">${label}</p>`;
    const rows = CanvaPptx.listAnimations(ir);
    $("listhint").hidden = !rows.length;
    $("breakdown").innerHTML = rows.length ? `<div class="tablewrap"><table id="anims"><thead><tr><th><span class="sr">Include</span></th><th>Slide</th><th>Element</th><th>Effect</th><th>Plays</th></tr></thead><tbody>${rows.map((r) => `<tr><td><input class="tick" type="checkbox" checked data-key="${r.key.replace(/"/g, "")}"></td><td>${r.slide}</td><td class="what" title="${r.what.replace(/"/g, "&quot;").replace(/</g, "&lt;")}">${r.what.replace(/</g, "&lt;")}</td><td>${r.effect}</td><td><span class="how ${/click/i.test(r.trigger) ? "click" : ""}">${r.trigger}</span></td></tr>`).join("")}</tbody></table></div>` : `<div class="empty"><svg viewBox="0 0 24 24"><path d="M5 19c0-8 5-14 14-14 0 9-6 14-14 14zM5 19l7-7"/></svg><span>No animations found in this design. Add some in Canva (Animate), wait for it to save, then read again.</span></div>`;
    $("support").disabled = false;
  } catch (e) { html = null; $("source").className = "status err"; $("source").textContent = e.message; $("breakdown").innerHTML = ""; }
  update();
}
function update() { $("go").disabled = !(html && pptxFile); }

(async () => {
  if (typeof chrome !== "undefined" && chrome.storage) {
    const { saved } = await chrome.storage.local.get("saved");
    if (saved) setSource(saved.html, "Read from Canva");
  }
})();

$("support").onclick = () => {
  const data = CanvaPptx.supportData(html);
  const url = URL.createObjectURL(new Blob([JSON.stringify(data)], { type: "application/json" }));
  const a = document.createElement("a"); a.href = url; a.download = "canva-support-file.json"; a.click();
};
$("htmlfile").onchange = async (e) => { const f = e.target.files[0]; if (f) setSource(await f.text(), `Loaded ${f.name}`); };
const takePptx = (f) => {
  if (!f) return; pptxFile = f;
  $("dropText").textContent = f.name; $("dzSub").textContent = `${(f.size / 1048576).toFixed(1)} MB · click to change`;
  $("dzIco").innerHTML = '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>'; $("drop").classList.add("has"); update();
};
$("pptx").onchange = (e) => takePptx(e.target.files[0]);
const drop = $("drop");
["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", (e) => takePptx(e.dataTransfer.files[0]));

$("go").onclick = async () => {
  const out = $("result");
  const goBtn = $("go"); goBtn.disabled = true; goBtn.innerHTML = '<span class="spin"></span>Converting';
  out.innerHTML = "";
  try {
    const skip = [...document.querySelectorAll("#anims input[type=checkbox]")].filter((c) => !c.checked).map((c) => c.dataset.key);
    const opts = { staggerMs: +$("stagger").value, defaultDurationMs: +$("dur").value, skip };
    const { ir, warnings } = CanvaPptx.extractFromHtml(html, opts);
    const buf = await pptxFile.arrayBuffer();
    const { data, report } = await CanvaPptx.injectIntoPptx(ir, buf, { JSZip, DOMParser, XMLSerializer }, { fixText: $("fixtext").checked, compat: $("compat").checked, rasterText: $("exact").checked });
    const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
    const url = URL.createObjectURL(blob);
    const name = pptxFile.name.replace(/\.pptx$/i, "") + " (animated).pptx";
    const all = [...report.warnings, ...warnings];
    const tf = report.slides.reduce((a, s) => a + (s.textFixes || 0), 0);
    const tp = report.slides.reduce((a, s) => a + (s.textPictures || 0), 0);
    if (tp) all.unshift(`Exact look: ${tp} text boxes drawn with Canva's fonts.`);
    if (tf) all.unshift(`Text layout: ${tf} text fixes applied (doubled bold removed / single-line boxes kept on one line).`);
    const rows = report.slides.filter((s) => s.animations || s.transition).map((s) =>
      `<tr><td>${s.slide}</td><td>${s.applied}/${s.animations}</td><td>${s.clicks}</td><td>${s.auto}</td></tr>`).join("");
    const applied = report.slides.reduce((a, s) => a + s.applied, 0), total = report.slides.reduce((a, s) => a + s.animations, 0);
    out.innerHTML = `<div class="done">
      <div><h3>Your deck is <em>ready.</em></h3><p>${applied} of ${total} animations applied${tp ? `, ${tp} text boxes drawn with Canva's fonts` : ""}.</p></div>
      <a class="dl" href="${url}" download="${name.replace(/"/g, "")}"><svg viewBox="0 0 24 24"><path d="M12 4v12M7 11l5 5 5-5M4 20h16"/></svg><span>Download ${name.replace(/</g, "&lt;")}</span></a>
      <div class="tablewrap"><table><thead><tr><th>Slide</th><th>Applied</th><th>Click steps</th><th>Automatic</th></tr></thead><tbody>${rows}</tbody></table></div>
      ${all.length ? `<details><summary>${all.length} note${all.length > 1 ? "s" : ""}</summary><ul>${all.map((w) => `<li>${w.replace(/</g, "&lt;")}</li>`).join("")}</ul></details>` : ""}</div>`;
    out.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (e) {
    out.innerHTML = `<div class="status err">${String(e.message || e).replace(/</g, "&lt;")}</div>`;
  } finally { goBtn.innerHTML = 'Convert to PowerPoint <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12h14M13 6l6 6-6 6"/></svg>'; update(); }
};
