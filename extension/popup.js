const $ = (id) => document.getElementById(id);
$("ver").textContent = "v" + chrome.runtime.getManifest().version;
const describe = (html) => { const { ir } = CanvaPptx.extractFromHtml(html); return { summary: CanvaPptx.summarize(ir), breakdown: CanvaPptx.breakdown(ir) }; };
const status = (html, cls) => { $("status").className = "status " + (cls || ""); $("status").innerHTML = html; };

async function refreshSaved() {
  const { saved } = await chrome.storage.local.get("saved");
  if (saved) { try { const d = describe(saved.html); $("open").disabled = false; status(`Last read: <b>${d.summary.slides}</b> slides · <b>${d.summary.animations}</b> animations · ${d.summary.clickSteps} on click`, "ok"); } catch (_) { /* old/invalid data */ } }
}

$("read").onclick = async () => {
  status("Reading your design");
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !/^https:\/\/www\.canva\.com\/design\//.test(tab.url || "")) {
    status("Open a Canva design in the editor first, then click this again.", "err"); return;
  }
  chrome.tabs.sendMessage(tab.id, { type: "canva-read" }, async (r) => {
    if (chrome.runtime.lastError || !r) { status("Couldn't reach the page. Reload the Canva tab and try again.", "err"); return; }
    if (!r.ok) { status(r.error, "err"); return; }
    let d;
    try { d = describe(r.html); } catch (e) { status(e.message, "err"); return; }
    await chrome.storage.local.set({ saved: { html: r.html, url: r.url, at: Date.now() } });
    $("open").disabled = false;
    const none = d.summary.animations === 0;
    const per = d.breakdown.map((b) => `Slide ${b.slide}: ${b.auto} auto · ${b.clickSteps} click`).join("<br>");
    status(none ? "This design has no animations. Add some in Canva (Animate), wait for it to save, and read again."
      : `Found <b>${d.summary.animations}</b> animations on <b>${d.summary.slides}</b> slides (<b>${d.summary.clickSteps} on click</b>).<br><small>${per}</small><br>Now download the PPTX from Canva, then open the converter.`, none ? "err" : "ok");
  });
};
$("open").onclick = () => chrome.tabs.create({ url: chrome.runtime.getURL("converter.html") });
refreshSaved();
