# Fit for Life Coach Feedback Forms — Project Context

## Objective
Collect feedback from ~150 Fit for Life Coach users segmented into 3 groups. Each user receives a personalized link generated from Google Sheets with pre-filled `?email=` parameter. Responses are submitted to Google Apps Script and written to Google Sheets.

## Files
```
form-active.html              → Users actively using Fit for Life Coach (≥ 20% training ratio)
form-inactive.html            → Users who completed onboarding but stopped (< 20% training ratio)
form-onboarding-dropout.html  → Users who never completed onboarding
google-apps-script.js         → Apps Script code to paste in Google Sheets editor
send-mails.ps1                → PowerShell script to send personalized emails (Office 365 SMTP)
preview-email.html            → Email template for local preview
```

## Infrastructure
- **Google Sheet:** `1O-dJv9CNcgidMwbd23dw5fErdzZuRW1MWrIDFtMFBsA`
- **Apps Script URL:** `https://script.google.com/macros/s/AKfycbyu_G3-Q1JanWo0mQMZ7EbowcPE2oNB0HC7RqjUEwyi2y2mA7IUprJ6Irnry0Ne8ZaEDw/exec`
- **GitHub Pages base:** `https://datasport-ai.github.io/Milo_Formular/`
- **Responses:** 3 separate tabs — `Responses_active`, `Responses_inactive`, `Responses_dropout`

## User Segmentation

Segmentation was done manually from the Google Sheet:
- **Active** (`form-active.html`): registered < 2 weeks ago AND logged > 50% of planned activities
- **Inactive** (`form-inactive.html`): registered < 2 weeks ago AND logged ≤ 50% of planned activities
- **Onboarding dropout** (`form-onboarding-dropout.html`): never completed onboarding (`onboarding_completed = No`)

## Personalized Link Formula (Users tab, column I)
```
=IF(E2="No","https://datasport-ai.github.io/Milo_Formular/form-onboarding-dropout.html?email="&ENCODEURL(A2)&"&step="&F2,IF(AND(E2="Yes",IFERROR(H2/G2,0)<0.5),"https://datasport-ai.github.io/Milo_Formular/form-inactive.html?email="&ENCODEURL(A2),"https://datasport-ai.github.io/Milo_Formular/form-active.html?email="&ENCODEURL(A2)))
```
Columns assumed: A=email, E=onboarding_completed, F=onboarding_step, G=planned_trainings, H=logged_trainings

## Tech Rules
- Standalone HTML files, zero external dependencies (except Open Sans via Google Fonts)
- No framework, pure vanilla JS — no Node.js
- Mobile-first, responsive
- Segmented progress bar at top (colored fill per step)
- `submitForm(data)` sends POST to Apps Script with `Content-Type: text/plain;charset=utf-8` + `mode: no-cors`
- All forms are in German (Swiss German tone — use "du", avoid "Sie")
- Email sending: PowerShell + Office 365 SMTP (`smtp.office365.com:587`)

## Design System
- Background: red `#E70E22` (Fit for Life Coach brand color)
- Card: `#F7F7F7`, rounded, bottom drawer on mobile / centered card on desktop
- Primary color / CTA: `#E70E22`
- Text: `#141414`
- Font: Open Sans (Google Fonts), fallback Arial
- Logo: Fit for Life Coach SVG logo in header

## Query Parameters (all three forms)
- `?email=user@example.com` → read on load, stored in `answers.email`, never displayed
- `?step=XXXX` → only for `form-onboarding-dropout.html`, maps to German label

### Step Mapping (onboarding dropout)
```
RESULTS_PERMISSION                   → "Freigabe deiner Resultate"
WEARABLE_CONNECTION                  → "Verbindung deines Geräts (GPS-Uhr etc.)"
CONTACT_PREFERENCES                  → "Kommunikationseinstellungen"
WEEKLY_SCHEDULE                      → "Wochenplan"
ARE_YOU_READY_FOR_THE_NEXT_CHALLENGE → "Nächste Herausforderung"
HOW_FAST_CAN_YOU_RUN                 → "Lauftempo"
START_DATE                           → "Startdatum"
GOAL_SETTING                         → "Zielsetzung"
TRAINING_FREQUENCY                   → "Trainingsfrequenz"
```
Fallback if unknown or missing: `"einem bestimmten Schritt"`

---

## Form Structures

### form-active.html — 4 steps

**Step 1** — Hilft dir der Coach, deine Leistungen zu verbessern?
- Single choice (required): Ja, deutlich / Noch zu früh, um das zu sagen / Noch nicht / Nein
- Conditional inline textarea (optional, no step change):
  - "Ja, deutlich" → "Was gefällt dir besonders gut? Nenn uns ein oder zwei Dinge."
  - "Noch zu früh, um das zu sagen" → "Was hat dich bisher überzeugt — oder was fehlt noch?"
  - "Noch nicht" → "Was hält dich zurück? Was würde dir helfen?"
  - "Nein" → "Was hat dich enttäuscht? Was hättest du erwartet?"

**Step 2** — Was muss dringend verbessert werden?
- "Nenn uns ein oder zwei Dinge, die wir dringend verbessern müssen." — textarea (optional)

**Step 3** — NPS
- Slider 0–10: "Auf einer Skala von 0 bis 10: Würdest du Fit for Life Coach jemandem in deinem Umfeld empfehlen?"
- No referral email block

**Step 4** — Abschluss (open question + incentive + submit)
- "Wenn du Fit for Life Coach einem Freund erklären müsstest, was würdest du sagen?" — textarea (required)
- Incentive box: CHF 20.– Gutschein bei Datasport
- Absenden button

**submitForm payload:** `{ formType: 'active', email, nps, q1, q1_follow, q2, q4 }`

---

### form-inactive.html — 3 steps

**Step 1** — Warum nutzt du Fit for Life Coach nicht mehr?
- Single choice (required): Ich habe nicht mehr trainiert / Fit for Life Coach hat meine Erwartungen/Bedürfnisse nicht erfüllt / Ich nutze ein anderes Tool / Ich hatte es vergessen / Technische Probleme
- Conditional inline textarea (optional, no step change):
  - "Fit for Life Coach hat meine Erwartungen/Bedürfnisse nicht erfüllt" → "Was hat dich enttäuscht? Was hättest du erwartet?"
  - "Ich nutze ein anderes Tool" → "Welches Tool nutzt du? Was bietet es dir, das Fit for Life Coach nicht bietet?"
  - "Technische Probleme" → "Bitte beschreibe das Problem."

**Step 2** — Was würde dich zur Rückkehr bewegen?
- Textarea (optional)

**Step 3** — Abschluss (feedback + incentive + submit)
- "Möchtest du noch etwas hinzufügen?" — textarea (optional)
- Incentive box: CHF 20.– Gutschein bei Datasport
- Absenden button

**submitForm payload:** `{ formType: 'inactive', email, q1, q1_follow, q2, q3 }`

---

### form-onboarding-dropout.html — 3 steps

**Step 1** — Warum hast du aufgehört?
- Question: "Du hast dein Fit for Life Coach-Setup bei **[stepLabel]** unterbrochen — was hat dich dazu bewogen?"
- Single choice (required): Keine Zeit / Onboarding zu lang / Dieser Schritt war unklar / Ich wollte keine Daten weitergeben / Technische Probleme / Anderes
- Conditional inline textarea (optional, no step change):
  - "Dieser Schritt war unklar" → "Was genau war unklar?"
  - "Ich wollte keine Daten weitergeben" → "Welche Daten haben dich gestört?"
  - "Technische Probleme" → "Bitte beschreibe das Problem."
  - "Anderes" → "Bitte beschreibe kurz."

**Step 2** — Planst du abzuschliessen?
- "Planst du, dein Fit for Life Coach-Profil noch abzuschliessen?" — Ja / Nein / Vielleicht (required)

**Step 3** — Abschluss (feedback + incentive + submit)
- "Möchtest du noch etwas hinzufügen?" — textarea (optional)
- Incentive box: CHF 20.– Gutschein bei Datasport
- Absenden button

**submitForm payload:** `{ formType: 'dropout', email, stepRaw, stepLabel, reason, reasonDetail, completeLater, freeComment }`

---

## Email sending

Script: `send-mails.ps1` — requires Outlook desktop (Outlook COM object), sent from `mydatasport@datasport.com`.

```powershell
.\send-mails.ps1 -Preview                        # preview in browser (no send)
.\send-mails.ps1 -DryRun                         # list recipients without sending
.\send-mails.ps1 -Limit 3                        # send to first 3 contacts
.\send-mails.ps1                                 # send to all (asks confirmation)
.\send-mails.ps1 -CsvPath .\contacts_test.csv    # test send to rmalet@datasport.com
```

`contacts_test.csv` — single-row CSV for test sends (rmalet@datasport.com, form-active link).

## Pending
- [ ] Incentive fulfillment — process to send CHF 20.– voucher (all 3 forms) within 48h
