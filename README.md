# Slide Timer for PowerPoint

A simple countdown timer you place on a slide. You set the time while building the deck and click it to start during the presentation. The digits scale to whatever size you make the box.

## Using it

| Where | What |
|---|---|
| **Editing** | Insert it through **Insert → Add-ins → My Add-ins → Slide Timer**. Set the time in the row under the digits (min / sec, or the 5·10·15·20 presets) and pick Light or Dark. Drag the corners to resize it; the digits follow. |
| **Presenting** | **Click** the timer to start, pause or resume. **Double-click** resets it. |
| **Keys** (after you click the timer once) | **Space / Enter / S** start or pause. **R** resets. **Arrows, PageUp/PageDown and presenter clickers** still move the slides; the timer passes them back to PowerPoint. |
| **At zero** | The timer stops at `0:00` and turns red. |

Each timer saves its own duration and style inside the `.pptx`, so different slides can have different timers.

### Presenter View

In Presenter View, PowerPoint keeps the keyboard for itself, so the timer's keys don't work there. The slide preview in Presenter View is a picture, not the live timer; clicking it just advances the slide. To start a timer without touching the audience screen, tick **Auto-start** in the timer's settings row. The countdown then starts from the full time every time that slide appears in a slideshow. You can still click the timer on the audience screen (for example on the second display) to pause or reset it.

> **Note:** PowerPoint may reload the add-in when you move away from its slide and come back. If that happens, the timer resets to its full duration.

## Development (localhost)

```bash
npm install
npm run certs      # one-time: installs a trusted localhost HTTPS certificate
npm start          # serves docs/ at https://localhost:3000
npm run sideload   # copies manifest.dev.xml into PowerPoint's add-in folder (Mac)
```
If `sideload` fails with *Operation not permitted*, macOS is protecting PowerPoint's container folder. Either give your terminal app **Full Disk Access** (System Settings → Privacy & Security), or drag `manifest.dev.xml` into `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` with Finder (**⇧⌘G** to go there; create `wef` if it doesn't exist).

Restart PowerPoint. **Slide Timer (dev)** then appears under **Insert → Add-ins → My Add-ins**, under the *Developer Add-ins* heading. The dev server has to be running whenever you use the dev version.

To test in a plain browser, open `https://localhost:3000/timer.html?view=read&s=10`. `view=read` previews the slideshow look and `s` sets the duration in seconds. Add `&auto=1` to test auto-start.

## Publishing (GitHub Pages)

`manifest.xml` points to **https://simonvanherweghe.github.io/devine-timer/**.

1. Create a **public** GitHub repo named `devine-timer` and push this folder to it:
   ```bash
   gh repo create devine-timer --public --source . --push
   ```
2. In the repo, go to **Settings → Pages → Deploy from a branch → `main` / `/docs`**, then wait a minute until `https://simonvanherweghe.github.io/devine-timer/timer.html` loads.
3. Run `npm run sideload:prod` (or drag `manifest.xml` into the `wef` folder, see above) and restart PowerPoint. **Slide Timer** now works without a local server.

If you use a different repo name or account, update the 4 URLs in `manifest.xml` to match.

To share it with colleagues, give them `manifest.xml`. On Windows, they add it through a shared-folder catalog; alternatively, an admin can deploy it through the Microsoft 365 admin center.

## Files

- `docs/timer.html`, `timer.css`, `timer.js`: the add-in itself (no build step)
- `manifest.dev.xml`: points to `https://localhost:3000`
- `manifest.xml`: points to GitHub Pages
- `scripts/make-icons.js`: regenerates the icons in `docs/assets/`
