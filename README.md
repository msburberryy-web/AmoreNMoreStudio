# Amorè N' More — Wedding Studio

## Current State

The live site at `https://msburberryy-web.github.io/AmoreNMoreStudio/` shows a **Coming Soon** page while the full site is being prepared.

---

## File Structure

```
/
├── index.html          ← Live: Coming Soon page (public)
├── config.json         ← Google Apps Script URL for the RSVP form
├── grand-opening.jpg   ← (TODO) Add the Grand Opening flyer image here
├── README.md           ← You are here
│
└── _archive/
    └── index.html      ← Full wedding studio service page + RSVP section
                           (hidden from public, stored here temporarily)
```

---

## What's in `_archive/`

`_archive/index.html` is the complete website containing:
- Full wedding salon service page (original content)
- Studio Opening RSVP section with curtain reveal animation
- Scroll-reveal form with Google Apps Script submission
- Time slot selection (၁၁–၁၃ · ၁၃–၁၅ · ၁၅–၁၇)

**Do not delete this file** — it will become the live site once ready.

---

## To Go Live

When ready to launch, do the following:

### 1. Add the Grand Opening image
Place the flyer image in the repo root as `grand-opening.jpg`.

### 2. Set the Google Apps Script URL
Open `config.json` and replace `YOUR_SCRIPT_ID_HERE` with your actual script ID:
```json
{
  "googleScriptUrl": "https://script.google.com/macros/s/YOUR_SCRIPT_ID_HERE/exec"
}
```

### 3. Swap the pages
```bash
cp _archive/index.html index.html
git add index.html
git commit -m "Launch full site"
git push
```

GitHub Pages will update within ~1 minute.

---

## Branches

| Branch | Purpose |
|---|---|
| `main` | Live (currently shows Coming Soon) |
| `claude/add-studio-opening-rsvp-n84Iz` | Development branch for the RSVP feature (merged into main, archived) |
