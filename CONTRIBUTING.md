# Contributing

Thanks for helping! The most valuable contributions are **verifying how Canva effects map to PowerPoint** and **keeping up with Canva's format**.

## Setup

```bash
cd extension
npm ci
npm test
```

Load `extension/` as an unpacked extension in Chrome to try it live.

## Where things are

| File | What |
| --- | --- |
| `extension/core.js` | Decoding Canva's page data, the PowerPoint timing/transition XML, shape matching, text fixes, exact-look text. No build step; works in the browser and in Node. |
| `extension/converter.*` | The converter page. |
| `extension/popup.*`, `content.js` | The extension popup and the content script that reads the page. |
| `extension/test/core.test.js` | Unit tests. |

## Tests

- Tests run with Node's built-in runner (`npm test`). Most use a synthetic PPTX and need nothing else.
- Four extra tests use a real Canva export. Put a saved Canva design page at `samples/design-page.html` and its PPTX export at `samples/export.pptx` to enable them. `samples/` is gitignored; **never commit your own designs or exports.**
- Add a test for every bug you fix.

## Reporting a bug

Use the bug template and attach the **support file**: in the converter, press *Download support file*. It contains only the animation and layout structure Canva sent for your design, but **check it for text you don't want to share** before posting.

## Adding or fixing an effect

Effect ids and the PowerPoint mapping are in `ELEMENT_ANIMS` and `TO_PPT` at the top of `core.js`. If you can confirm what an effect's entry direction or option means in Canva, please say how you verified it.

## Style

Plain JavaScript, no dependencies at runtime, no build tooling. Keep the extension fully local: no network calls and no remote code.
