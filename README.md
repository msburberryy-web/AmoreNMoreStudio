# Amorè N' More — Wedding Studio

## Live URLs

| URL | What it shows |
|---|---|
| `msburberryy-web.github.io/AmoreNMoreStudio/` | **"Studio is opened!"** page with RSVP link |
| `msburberryy-web.github.io/AmoreNMoreStudio/studioOpensNow/` | **Grand Opening** — curtain reveal + RSVP form + event details |

---

## File Structure

```
/
├── index.html                  ← "Studio is opened!" landing page (links to studioOpensNow)
├── config.json                 ← Google Apps Script URL for the RSVP form
├── grand-opening.jpg           ← (TODO) Add the Grand Opening flyer image here
├── README.md                   ← You are here
│
├── studioOpensNow/
│   └── index.html              ← Full event page:
│                                    • Silk curtain reveal over grand-opening.jpg
│                                    • RSVP form (name, email/phone, time slot, message, questions)
│                                    • Event details section (date, time, address, quote)
│
└── _archive/
    └── index.html              ← Original full wedding studio service page
                                   (preserved, not publicly linked)
```

---

## To activate the RSVP form

### 1. Add the Grand Opening flyer image
Place the image in the **repo root** as `grand-opening.jpg`.
It is referenced in `studioOpensNow/index.html` as `../grand-opening.jpg`.

### 2. Set the Google Apps Script URL
Open `config.json` and replace the placeholder with your real script ID:
```json
{
  "googleScriptUrl": "https://script.google.com/macros/s/YOUR_SCRIPT_ID_HERE/exec"
}
```

---

## Time slots (RSVP form)
- ၁၁ – ၁၃ &nbsp; (11:00 AM – 1:00 PM)
- ၁၃ – ၁၅ &nbsp; (1:00 PM – 3:00 PM)
- ၁၅ – ၁၇ &nbsp; (3:00 PM – 5:00 PM)

---

## Event Details (hard-coded in studioOpensNow/index.html)
- **Date**: Saturday, 27 June 2026
- **Start**: 11:00 AM
- **Address**: 112-0011 Tokyo, Bunkyoku, Sengoku, 4-chome-26-2, SANSAN Sengoku Bldg. 301
