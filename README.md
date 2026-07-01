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
| `contacts.csv` | Full contact list (all segments) |
| `contacts-reminder.csv` | Filtered list for reminder (non-respondents only) — regénérer avant usage |
| `responded.txt` | Emails of users who already responded (used to generate reminder list) |
| `preview-email.html` | Email template (local preview only) |

## Infrastructure

- **Google Sheet:** `1O-dJv9CNcgidMwbd23dw5fErdzZuRW1MWrIDFtMFBsA`
- **Apps Script URL:** `https://script.google.com/macros/s/AKfycbyu_G3-Q1JanWo0mQMZ7EbowcPE2oNB0HC7RqjUEwyi2y2mA7IUprJ6Irnry0Ne8ZaEDw/exec`
- **GitHub Pages:** `https://datasport-ai.github.io/Milo_Formular/`
- **Responses:** 3 tabs — `Responses_active`, `Responses_inactive`, `Responses_dropout`

## Usage

### Send emails

Requires Outlook desktop (uses Outlook COM object). `contacts.csv` must be present at the root.

```powershell
# Preview the email in the browser (first contact in CSV)
.\send-mails.ps1 -Preview

# Dry run — list all recipients without sending
.\send-mails.ps1 -DryRun

# Send to first 3 contacts only (test)
.\send-mails.ps1 -Limit 3

# Send to all contacts (asks for confirmation)
.\send-mails.ps1
```

### Send reminder (non-respondents only)

1. Copy the emails of people who already responded from the Google Sheet into `responded.txt` (one email per line)
2. Run the script below to generate `contacts-reminder.csv` (excludes respondents + deduplicates)
3. Send as usual with `-CsvPath contacts-reminder.csv`

```powershell
# Generate contacts-reminder.csv
$responded = Get-Content ".\responded.txt" | ForEach-Object { $_.Trim().ToLower() } | Where-Object { $_ -ne '' } | Sort-Object -Unique
$all = Import-Csv ".\contacts.csv"
$reminder = $all | Where-Object { $responded -notcontains $_.email.ToLower() }
$reminder | Export-Csv ".\contacts-reminder.csv" -NoTypeInformation -Encoding UTF8
Write-Host "Rappel : $($reminder.Count) contacts (sur $($all.Count) total, $($responded.Count) ont repondu)"

# Then send
.\send-mails.ps1 -CsvPath contacts-reminder.csv -DryRun  # vérifier d'abord
.\send-mails.ps1 -CsvPath contacts-reminder.csv
```

> `responded.txt` est conservé dans le repo pour garder la trace des répondants.
> Note : si un email de `responded.txt` n'existe pas dans `contacts.csv`, le total ne sera pas exactement `total - répondants` (cas normal pour les tests internes).

### Test send (single recipient)

`contacts_test.csv` contains a single test entry (rmalet@datasport.com, form-active).

```powershell
# Preview test email in browser
.\send-mails.ps1 -Preview -CsvPath .\contacts_test.csv

# Send test email to rmalet@datasport.com
.\send-mails.ps1 -CsvPath .\contacts_test.csv
```

### Deploy Google Apps Script

Paste `google-apps-script.js` into the Apps Script editor of the Google Sheet and deploy as a web app.

### Forms (GitHub Pages)

The three HTML forms are served directly via GitHub Pages — no build step needed.

## Stack

- Vanilla HTML/CSS/JS — no external dependencies (except Open Sans via Google Fonts)
- Email sending: PowerShell + Outlook COM (Outlook desktop required)
- No Node.js

## See also

→ `CONTEXT.md` for full documentation (form structure, design system, URL parameters)
