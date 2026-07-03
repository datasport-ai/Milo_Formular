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
| `responded.txt` | Emails of users who already responded (used to generate reminder list + thank-you list) |
| `coupons.csv` | Raw voucher list (code, password, amount, validity) provided by Datasport |
| `generate-thankyou-csv.ps1` | Generates `contacts-thankyou.csv` by assigning one voucher per respondent |
| `contacts-thankyou.csv` | Generated: respondent + assigned voucher (input for `send-thankyou-mails.ps1`) |
| `send-thankyou-mails.ps1` | PowerShell script to send thank-you emails with voucher to respondents |
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

### Send thank-you emails (with voucher)

Chaque répondant reçoit un mail de remerciement avec un coupon de réduction unique (numéro + mot de passe).

1. S'assurer que `responded.txt` est à jour (voir section rappel ci-dessus) et que `coupons.csv` contient assez de coupons non utilisés (un par répondant unique)
2. Générer `contacts-thankyou.csv` (associe un coupon unique à chaque répondant, récupère prénom/nom depuis `contacts.csv`) :

```powershell
.\generate-thankyou-csv.ps1
```

Le script signale les emails absents de `contacts.csv` (prénom/nom à compléter manuellement dans le CSV généré) et le nombre de coupons non utilisés.

3. Vérifier puis envoyer :

```powershell
.\send-thankyou-mails.ps1 -Preview   # apercu dans le navigateur
.\send-thankyou-mails.ps1 -DryRun    # verifier la liste des destinataires
.\send-thankyou-mails.ps1            # envoi reel (demande confirmation)
```

Chaque envoi réussi est noté dans `sent-thankyou.txt`. Au lancement suivant, ces contacts sont automatiquement sautés — pas besoin de gérer un index manuellement, et pas de risque de double envoi (donc double consommation de coupon) si le script est relancé après un envoi partiel ou un test avec `-Limit`. Utiliser `-Resend` pour forcer un renvoi malgré le log.

> Les coupons sont une ressource limitée et non réutilisable : ne pas relancer `generate-thankyou-csv.ps1` après l'envoi des mails, sous peine de réassigner un coupon déjà communiqué à quelqu'un d'autre si `responded.txt` ou `coupons.csv` changent entretemps.

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
