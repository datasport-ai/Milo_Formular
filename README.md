# Fit for Life Coach — Feedback Forms

Personalized feedback forms for ~150 Fit for Life Coach users, split into 3 segments.

## User Segmentation

Segmentation was done manually from the Google Sheet:

| Segment | Form | Criteria |
|---|---|---|
| **Active** | `form-active.html` | Registered < 2 weeks ago AND logged > 50% of planned activities |
| **Inactive** | `form-inactive.html` | Registered < 2 weeks ago AND logged ≤ 50% of planned activities |
| **Onboarding dropout** | `form-onboarding-dropout.html` | Never completed onboarding |

## Files

| File | Description |
|---|---|
| `form-active.html` | Form for active users |
| `form-inactive.html` | Form for inactive users |
| `form-onboarding-dropout.html` | Form for onboarding dropouts |
| `google-apps-script.js` | Code to paste in the Google Sheets Apps Script editor |
| `send-mails.ps1` | PowerShell script to send personalized emails |
| `preview-email.html` | Email template (local preview only) |

## Infrastructure

- **Google Sheet:** `1O-dJv9CNcgidMwbd23dw5fErdzZuRW1MWrIDFtMFBsA`
- **Apps Script URL:** `https://script.google.com/macros/s/AKfycbyu_G3-Q1JanWo0mQMZ7EbowcPE2oNB0HC7RqjUEwyi2y2mA7IUprJ6Irnry0Ne8ZaEDw/exec`
- **GitHub Pages:** `https://datasport-ai.github.io/Fit for Life Coach_Formular/`
- **Responses:** 3 tabs — `Responses_active`, `Responses_inactive`, `Responses_dropout`

## Stack

- Vanilla HTML/CSS/JS — no external dependencies (except Open Sans via Google Fonts)
- Email sending: PowerShell + Office 365 SMTP (`smtp.office365.com:587`)
- No Node.js

## See also

→ `CONTEXT.md` for full documentation (form structure, design system, URL parameters)
