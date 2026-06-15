# Milo Feedback Forms — Project Context

## Objective
Collect feedback from ~150 Milo users segmented into 3 groups. Each segment receives a distinct link to a standalone HTML form. Results are submitted via `submitForm(data)` placeholder function (to be connected to Google Apps Script later).

## Files
```
form-active.html           → Users actively using Milo
form-inactive.html         → Users who completed onboarding but stopped using Milo
form-onboarding-dropout.html → Users who never completed onboarding
```

## Tech Rules
- Standalone HTML files, zero external dependencies (except Open Sans via Google Fonts)
- No framework, pure vanilla JS
- Mobile-first, responsive
- One question visible per step, navigation via Weiter / Absenden buttons
- Progress bar at top showing current step (e.g. "Frage 2 von 3") with colored fill
- `submitForm(data)` is an empty placeholder function at the end of each file — do not implement backend logic
- All forms are in German (Swiss German tone — use "du", avoid "Sie")

## Design System
- Background: red (Milo brand color — do not change, already implemented)
- Card: white, centered, rounded
- Primary color / CTA: `#E70E22`
- Text: `#141414`
- Font: Open Sans (Google Fonts), fallback Arial
- Logo: Milo logo in header

## Query Parameters (all three forms)
- `?email=user@example.com` → read on load, store in `answers.email`, never display on screen
- `?step=XXXX` → only for `form-onboarding-dropout.html`, maps to German label (see mapping below)

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

### form-onboarding-dropout.html — 3 steps

**Step 1** — Warum hast du aufgehört?
- Question: "Du hast dein Milo-Setup bei **[stepLabel]** unterbrochen — was hat dich dazu bewogen?"
- Single choice (required): Keine Zeit / Onboarding zu lang / Dieser Schritt war unklar / Ich wollte keine Daten weitergeben / Technische Probleme / Anderes
- Conditional inline textarea (optional, no step change):
  - "Dieser Schritt war unklar" → "Was genau war unklar?"
  - "Ich wollte keine Daten weitergeben" → "Welche Daten haben dich gestört?"
  - "Technische Probleme" → "Bitte beschreibe das Problem."
  - "Anderes" → "Bitte beschreibe kurz."

**Step 2** — Planst du abzuschliessen?
- "Planst du, dein Milo-Profil noch abzuschliessen?" — Ja / Nein / Vielleicht (required)

**Step 3** — Abschluss (feedback + incentive + submit)
- "Möchtest du noch etwas hinzufügen?" — textarea (optional)
- Incentive box: "Als Dankeschön für dein Feedback schenken wir dir einen zusätzlichen Monat Milo — du erhältst deinen Code per E-Mail innerhalb von 48 Stunden."
- Absenden button

**submitForm data:** `{ email, stepRaw, stepLabel, reason, reasonDetail, completeLater, freeComment }`

---

### form-inactive.html — 3 steps

**Step 1** — Warum nutzt du Milo nicht mehr?
- Single choice (required): Ich habe nicht mehr trainiert / Milo hat meine Erwartungen/Bedürfnisse nicht erfüllt / Ich nutze ein anderes Tool / Ich hatte es vergessen / Technische Probleme
- Conditional inline textarea (optional, no step change):
  - "Milo hat meine Erwartungen/Bedürfnisse nicht erfüllt" → "Was hat dich enttäuscht? Was hättest du erwartet?"
  - "Ich nutze ein anderes Tool" → "Welches Tool nutzt du? Was bietet es dir, das Milo nicht bietet?"
  - "Technische Probleme" → "Bitte beschreibe das Problem."

**Step 2** — Was würde dich zur Rückkehr bewegen?
- Textarea (optional)

**Step 3** — Abschluss (feedback + incentive + submit)
- "Möchtest du noch etwas hinzufügen?" — textarea (optional)
- Incentive box: "Als Dankeschön für dein Feedback schenken wir dir einen zusätzlichen Monat Milo — du erhältst deinen Code per E-Mail innerhalb von 48 Stunden."
- Absenden button

**submitForm data:** `{ email, reason, reasonDetail, returnMotivation, freeComment }`

---

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

**Step 3** — NPS + Referral
- NPS slider 0–10: "Auf einer Skala von 0 bis 10: Würdest du Milo jemandem in deinem Umfeld empfehlen?"
- If score ≥ 7: referral block appears dynamically
  - Text: "Kennst du jemanden, der Milo ausprobieren sollte? Trag seine E-Mail-Adresse ein — er erhält einen Monat Milo gratis, und du bekommst einen CHF 20.– Gutschein für deine nächste Datasport-Anmeldung."
  - Repeatable email field (+ button to add more, no limit)

**Step 4** — Abschluss (open question + incentive + submit)
- "Wenn du Milo einem Freund erklären müsstest, was würdest du sagen?" — textarea (required)
- Incentive box: "Als Dankeschön für dein Feedback schenken wir dir einen zusätzlichen Monat Milo — du erhältst deinen Code per E-Mail innerhalb von 48 Stunden."
- Absenden button

**submitForm data:** `{ email, performanceRating, performanceComment, improvements, nps, referralEmails, openDescription }`

---

## Pending
- [ ] Google Apps Script endpoint — replace `submitForm()` placeholder onceSheet is ready
- [ ] Incentive email sending process (48h turnaround mentioned in forms)
- [ ] Personalized links to generate per user from Google Sheets (email + step query params)
