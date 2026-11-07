// Runs on canva.com/design/*. Returns the page HTML; all decoding happens in the (always up-to-date) popup/converter.
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg && msg.type === "canva-read") {
    fetch(location.href, { credentials: "include", cache: "no-store" })
      .then((r) => r.text())
      .then((html) => sendResponse({ ok: true, html, url: location.href }))
      .catch((e) => sendResponse({ ok: false, error: String((e && e.message) || e) }));
    return true;
  }
});
