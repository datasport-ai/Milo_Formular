# ============================================================
#  send-thankyou-mails.ps1  -  Envoi mails de remerciement
#  + coupon FIT for LIFE Coach (aux repondants du formulaire)
#  Utilise Outlook COM (Outlook desktop requis)
# ============================================================
#
#  CSV attendu (genere par generate-thankyou-csv.ps1) :
#  email, first_name, last_name, coupon_code, coupon_password,
#  coupon_currency, coupon_amount, coupon_valid_until
#
#  USAGE :
#    .\send-thankyou-mails.ps1            -> Envoie a tous les contacts pas encore envoyes
#    .\send-thankyou-mails.ps1 -DryRun   -> Previsu sans envoyer
#    .\send-thankyou-mails.ps1 -Preview  -> Ouvre apercu dans le navigateur
#    .\send-thankyou-mails.ps1 -Limit 3  -> Envoie seulement aux 3 premiers (non deja envoyes)
#    .\send-thankyou-mails.ps1 -Resend   -> Renvoie a tout le monde, y compris deja envoyes
#
#  SUIVI DES ENVOIS :
#  Chaque envoi reussi est note dans sent-thankyou.txt (un email par ligne).
#  Au prochain lancement, ces contacts sont automatiquement sautes (pas de
#  double envoi / double consommation de coupon). Utiliser -Resend pour
#  forcer un renvoi malgre le log.
# ============================================================

param(
    [switch]$DryRun,
    [switch]$Preview,
    [switch]$Force,
    [switch]$Resend,
    [int]$Limit = 0,
    [string]$CsvPath = ".\contacts-thankyou.csv",
    [string]$SentLogPath = ".\sent-thankyou.txt"
)

# --- Configuration ---
$LogoPath  = ".\FIT for LIFE Coach Logo white background.png"
$FromEmail = "mydatasport@datasport.com"
$Subject   = "FIT for LIFE Coach - Danke fuer dein Feedback! Dein Gutschein / Thank you for your feedback! Your voucher"

# =====================================================================
# Fonction : genere le HTML personalise pour un contact
# =====================================================================
function Get-ThankYouEmailHtml {
    param(
        [string]$FirstName,
        [string]$CouponCode,
        [string]$CouponPin,
        [string]$CouponCurrency,
        [string]$CouponAmount,
        [string]$CouponValidUntil,
        [string]$LogoSrc = "cid:fflc-logo"
    )

    return @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0; padding:0; background-color:#F2F2F2; font-family:'Segoe UI',Arial,sans-serif;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#F2F2F2;">
<tr><td align="center" style="padding:32px 16px;">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="600" style="max-width:600px; width:100%; background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,0.10);">

  <!-- HEADER rouge avec logo + texte -->
  <tr>
    <td align="center" style="background-color:#E70E22; padding:24px 40px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td style="vertical-align:middle; padding-right:14px;">
            <img src="$LogoSrc" alt="" width="40"
                 style="display:block; width:40px; height:40px;">
          </td>
          <td style="vertical-align:middle;">
            <span style="font-size:22px; font-weight:700; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; letter-spacing:0.3px;">FIT for LIFE Coach</span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ======================================================
       VERSION ALLEMANDE
  ====================================================== -->
  <tr>
    <td style="padding:36px 48px 28px;">

      <p style="margin:0 0 20px; font-size:22px; font-weight:700; color:#141414; line-height:1.3;">
        Hallo $FirstName,<br>
        herzlichen Dank f&uuml;r dein Feedback!
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        Du hast dir Zeit genommen, um uns deine Meinung zu <strong style="color:#141414;">FIT for LIFE Coach</strong>
        mitzuteilen &ndash; das hilft uns enorm, den Service weiterzuentwickeln. Vielen Dank daf&uuml;r!
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        Als Dankesch&ouml;n erh&auml;ltst du wie versprochen einen Gutschein &uuml;ber
        <strong style="color:#141414;">$CouponCurrency $CouponAmount</strong>,
        den du bei deiner n&auml;chsten Anmeldung zu einem Event bei Datasport einl&ouml;sen kannst.
      </p>

      <!-- Coupon DE -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0;">
        <tr>
          <td style="background-color:#FFF5F5; border-left:4px solid #E70E22; border-radius:6px; padding:20px 24px;">
            <p style="margin:0 0 12px; font-size:15px; font-weight:700; color:#141414;">Dein Gutschein</p>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666; width:140px;">Gutscheinnummer</td>
                <td style="padding:4px 0; font-size:16px; font-weight:700; color:#141414; font-family:'Courier New',monospace; letter-spacing:0.5px;">$CouponCode</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">Passwort</td>
                <td style="padding:4px 0; font-size:16px; font-weight:700; color:#141414; font-family:'Courier New',monospace; letter-spacing:0.5px;">$CouponPin</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">Wert</td>
                <td style="padding:4px 0; font-size:15px; color:#141414;">$CouponCurrency $CouponAmount</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">G&uuml;ltig bis</td>
                <td style="padding:4px 0; font-size:15px; color:#141414;">$CouponValidUntil</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>

      <!-- Redemption info DE -->
      <p style="margin:0 0 8px; font-size:15px; font-weight:700; color:#141414;">Wie kann ich einen Gutschein einl&ouml;sen?</p>
      <p style="margin:0 0 10px; font-size:14px; color:#666666; line-height:1.6;">
        Wenn du dich f&uuml;r eine Veranstaltung anmeldest, kannst du im Zahlungsprozess bei der &laquo;Kasse&raquo;
        den Gutschein als Zahlungsmittel einsetzen. Wenn du den gesamten Betrag mit dem Gutschein bezahlst,
        sparst du &uuml;berdies das Auftragshandling.
      </p>
      <p style="margin:0 0 24px; font-size:14px; color:#666666; line-height:1.6;">
        Der Gutschein kann flexibel eingel&ouml;st werden: Restbetr&auml;ge verfallen nicht und du kannst
        mehrere Gutscheine gleichzeitig einl&ouml;sen.
      </p>

      <!-- CTA DE -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 24px;">
        <tr>
          <!--[if mso]>
          <td align="center" style="background-color:#e70e22; border-radius:8px;">
          <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word"
            href="https://datasport.com" style="height:52px; v-text-anchor:middle; width:240px;"
            arcsize="12%" strokecolor="#141414" fillcolor="#141414">
            <w:anchorlock/>
            <center style="color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700;">
              Zu den Events &#8594;
            </center>
          </v:roundrect>
          </td>
          <![endif]-->
          <!--[if !mso]><!-->
          <td align="center" style="background-color:#141414; border-radius:8px;">
            <a href="https://datasport.com"
               style="display:inline-block; padding:15px 36px; background-color:#e70e22; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700; text-decoration:none; border-radius:8px; letter-spacing:0.3px;">
              Zu den Events &#8594;
            </a>
          </td>
          <!--<![endif]-->
        </tr>
      </table>

      <p style="margin:0; font-size:13px; color:#ADADAD; line-height:1.5;">
        Gib die Gutscheinnummer und das Passwort bei der Anmeldung zu einem Event deiner Wahl ein.
      </p>

      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0 20px;">
        <tr><td style="border-top:1px solid #EEEEEE;"></td></tr>
      </table>

      <p style="margin:0 0 4px; font-size:15px; color:#444444; line-height:1.6;">Herzliche Gr&uuml;sse,</p>
      <p style="margin:0; font-size:15px; font-weight:700; color:#141414;">Das Datasport-Team</p>

    </td>
  </tr>

  <!-- SEPARATEUR BILINGUE -->
  <tr>
    <td style="padding:0 48px;">
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
        <tr>
          <td style="border-top:2px dashed #DEDEDE; padding:16px 0; text-align:center;">
            <span style="font-size:11px; color:#ADADAD; letter-spacing:1px; text-transform:uppercase;">
              &#9670;&nbsp; English version below &nbsp;&#9670;
            </span>
          </td>
        </tr>
      </table>
    </td>
  </tr>

  <!-- ======================================================
       VERSION ANGLAISE
  ====================================================== -->
  <tr>
    <td style="padding:28px 48px 36px;">

      <p style="margin:0 0 20px; font-size:22px; font-weight:700; color:#141414; line-height:1.3;">
        Hi $FirstName,<br>
        thank you so much for your feedback!
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        You took the time to share your thoughts on <strong style="color:#141414;">FIT for LIFE Coach</strong>
        &ndash; this helps us enormously in improving the service. Thank you!
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        As promised, here's your thank-you voucher worth
        <strong style="color:#141414;">$CouponCurrency $CouponAmount</strong>,
        which you can redeem the next time you register for an event at Datasport.
      </p>

      <!-- Coupon EN -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0;">
        <tr>
          <td style="background-color:#FFF5F5; border-left:4px solid #E70E22; border-radius:6px; padding:20px 24px;">
            <p style="margin:0 0 12px; font-size:15px; font-weight:700; color:#141414;">Your voucher</p>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666; width:140px;">Voucher number</td>
                <td style="padding:4px 0; font-size:16px; font-weight:700; color:#141414; font-family:'Courier New',monospace; letter-spacing:0.5px;">$CouponCode</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">Password</td>
                <td style="padding:4px 0; font-size:16px; font-weight:700; color:#141414; font-family:'Courier New',monospace; letter-spacing:0.5px;">$CouponPin</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">Value</td>
                <td style="padding:4px 0; font-size:15px; color:#141414;">$CouponCurrency $CouponAmount</td>
              </tr>
              <tr>
                <td style="padding:4px 0; font-size:14px; color:#666666;">Valid until</td>
                <td style="padding:4px 0; font-size:15px; color:#141414;">$CouponValidUntil</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>

      <!-- Redemption info EN -->
      <p style="margin:0 0 8px; font-size:15px; font-weight:700; color:#141414;">How can I redeem a voucher?</p>
      <p style="margin:0 0 10px; font-size:14px; color:#666666; line-height:1.6;">
        When you register for an event, you can use the voucher as a means of payment during the payment
        process at the 'Checkout'. If you pay the entire amount with the voucher, you also save on order handling.
      </p>
      <p style="margin:0 0 24px; font-size:14px; color:#666666; line-height:1.6;">
        The voucher can be redeemed flexibly: remaining amounts do not expire and you can redeem several
        vouchers at the same time.
      </p>

      <!-- CTA EN -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 24px;">
        <tr>
          <!--[if mso]>
          <td align="center" style="background-color:#e70e22; border-radius:8px;">
          <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word"
            href="https://datasport.com" style="height:52px; v-text-anchor:middle; width:240px;"
            arcsize="12%" strokecolor="#e70e22" fillcolor="#e70e22">
            <w:anchorlock/>
            <center style="color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700;">
              Go to the events &#8594;
            </center>
          </v:roundrect>
          </td>
          <![endif]-->
          <!--[if !mso]><!-->
          <td align="center" style="background-color:#141414; border-radius:8px;">
            <a href="https://datasport.com"
               style="display:inline-block; padding:15px 36px; background-color:#e70e22; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700; text-decoration:none; border-radius:8px; letter-spacing:0.3px;">
              Go to the events &#8594;
            </a>
          </td>
          <!--<![endif]-->
        </tr>
      </table>

      <p style="margin:0; font-size:13px; color:#ADADAD; line-height:1.5;">
        Enter the voucher number and password when registering for an event of your choice.
      </p>

      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0 20px;">
        <tr><td style="border-top:1px solid #EEEEEE;"></td></tr>
      </table>

      <p style="margin:0 0 4px; font-size:15px; color:#444444; line-height:1.6;">Best regards,</p>
      <p style="margin:0; font-size:15px; font-weight:700; color:#141414;">The Datasport Team</p>

    </td>
  </tr>

  <!-- FOOTER -->
  <tr>
    <td style="background-color:#F7F7F7; padding:20px 48px; border-top:1px solid #EEEEEE;">
      <p style="margin:0 0 4px; font-size:12px; color:#ADADAD; line-height:1.6; text-align:center;">
        Du erh&auml;ltst / You receive this email because you completed the FIT for LIFE Coach feedback survey on
        <a href="https://datasport.com" style="color:#E70E22; text-decoration:none;">datasport.com</a>.
      </p>
      <p style="margin:0; font-size:12px; color:#ADADAD; text-align:center;">
        &copy; 2026 Datasport AG &nbsp;|&nbsp; Schweiz
      </p>
    </td>
  </tr>

</table>
</td></tr>
</table>
</body>
</html>
"@
}

# =====================================================================

if (-not (Test-Path $CsvPath)) {
    Write-Host "ERREUR : CSV introuvable : $CsvPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $LogoPath)) {
    Write-Host "ATTENTION : Logo introuvable, email envoye sans logo." -ForegroundColor Yellow
    $LogoPath = $null
}

# --- Mode Preview ---
if ($Preview) {
    $row         = Import-Csv $CsvPath | Select-Object -First 1
    $previewPath = "$PSScriptRoot\preview-email-thankyou.html"
    $logoSrcPreview = if ($LogoPath) { "file:///" + ((Resolve-Path $LogoPath).Path -replace '\\', '/') } else { "" }
    Get-ThankYouEmailHtml -FirstName $row.first_name -CouponCode $row.coupon_code -CouponPin $row.coupon_password `
        -CouponCurrency $row.coupon_currency -CouponAmount $row.coupon_amount -CouponValidUntil $row.coupon_valid_until `
        -LogoSrc $logoSrcPreview |
        Out-File -FilePath $previewPath -Encoding UTF8
    Write-Host "Preview genere : $previewPath" -ForegroundColor Green
    Write-Host "Contact test   : $($row.first_name) ($($row.email)) - coupon $($row.coupon_code)" -ForegroundColor Gray
    Start-Process $previewPath
    exit 0
}

# =====================================================================

$alreadySent = New-Object System.Collections.Generic.HashSet[string]
if (Test-Path $SentLogPath) {
    Get-Content $SentLogPath | ForEach-Object {
        $e = $_.Trim().ToLower()
        if ($e -ne '') { [void]$alreadySent.Add($e) }
    }
}

$allContacts = Import-Csv $CsvPath
$skipped     = @()
if (-not $Resend) {
    $contacts = $allContacts | Where-Object { -not $alreadySent.Contains($_.email.ToLower()) }
    $skipped  = $allContacts | Where-Object { $alreadySent.Contains($_.email.ToLower()) }
} else {
    $contacts = $allContacts
}

if ($Limit -gt 0) { $contacts = $contacts | Select-Object -First $Limit }
$total = ($contacts | Measure-Object).Count

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  FIT for LIFE Coach - Mails de remerciement" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Contacts   : $total"
if ($skipped.Count -gt 0) {
    Write-Host "  Deja envoyes (sautes) : $($skipped.Count) (voir $SentLogPath)" -ForegroundColor DarkGray
}
Write-Host "  From       : $FromEmail"
Write-Host "  Sujet      : $Subject"
if ($DryRun) {
    Write-Host "  MODE       : DRY RUN (aucun email envoye)" -ForegroundColor Yellow
}
if ($Resend) {
    Write-Host "  MODE       : RESEND (log d'envoi ignore)" -ForegroundColor Red
}
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

if ($total -eq 0) {
    Write-Host "Rien a envoyer (tous les contacts du CSV sont deja dans $SentLogPath)." -ForegroundColor Yellow
    exit 0
}

if (-not $DryRun -and -not $Force) {
    $confirm = Read-Host "Confirmer l'envoi a $total destinataires ? (O/N)"
    if ($confirm -notmatch '^[Oo]$') {
        Write-Host "Annule." -ForegroundColor Yellow
        exit 0
    }
}

$outlook = New-Object -ComObject Outlook.Application
$sent    = 0
$errors  = 0

foreach ($row in $contacts) {

    if (-not $row.first_name -or -not $row.coupon_code) {
        Write-Host "  [SKIP] $($row.email) - prenom ou coupon manquant, a completer manuellement dans le CSV" -ForegroundColor Yellow
        continue
    }

    $htmlBody = Get-ThankYouEmailHtml -FirstName $row.first_name -CouponCode $row.coupon_code -CouponPin $row.coupon_password `
        -CouponCurrency $row.coupon_currency -CouponAmount $row.coupon_amount -CouponValidUntil $row.coupon_valid_until

    if ($DryRun) {
        Write-Host "  [DRY RUN] $($row.email) ($($row.first_name)) - coupon $($row.coupon_code)" -ForegroundColor Gray
        $sent++
        continue
    }

    try {
        $mail = $outlook.CreateItem(0)
        $mail.SentOnBehalfOfName = $FromEmail
        $mail.To      = $row.email
        $mail.Subject = $Subject

        if ($LogoPath) {
            $logoFull = (Resolve-Path $LogoPath).Path
            $att = $mail.Attachments.Add($logoFull)
            $att.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                "fflc-logo"
            )
        }

        $mail.HTMLBody = $htmlBody
        $mail.Send()

        Add-Content -Path $SentLogPath -Value $row.email.ToLower()

        $sent++
        Write-Host "  [OK] $($row.email) ($($row.first_name)) - coupon $($row.coupon_code)  [$sent/$total]" -ForegroundColor Green
        Start-Sleep -Milliseconds 500
    }
    catch {
        $errors++
        Write-Host "  [ERREUR] $($row.email) - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
if ($DryRun) {
    Write-Host "  DRY RUN termine : $sent contacts listes" -ForegroundColor Yellow
} else {
    Write-Host "  Termine : $sent envoyes, $errors erreurs" -ForegroundColor Green
}
Write-Host "=======================================" -ForegroundColor Cyan
