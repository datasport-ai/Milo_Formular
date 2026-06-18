# ============================================================
#  send-mails.ps1  -  Envoi emails FIT for LIFE Coach
#  Utilise Outlook COM (Outlook desktop requis)
# ============================================================
#
#  LIENS : Verifier que les URLs dans contacts.csv pointent
#  vers le bon repo GitHub avant d'envoyer.
#
#  USAGE :
#    .\send-mails.ps1            -> Envoie a tous les contacts
#    .\send-mails.ps1 -DryRun   -> Previsu sans envoyer
#    .\send-mails.ps1 -Preview  -> Ouvre apercu dans le navigateur
#    .\send-mails.ps1 -Limit 3  -> Envoie seulement aux 3 premiers
# ============================================================

param(
    [switch]$DryRun,
    [switch]$Preview,
    [int]$Limit = 0
)

# --- Configuration ---
$CsvPath   = ".\contacts.csv"
$LogoPath  = ".\FIT for LIFE Coach Logo white background.png"
$FromEmail = "rmalet@datasport.com"
$Subject   = "FIT for LIFE Coach - Dein Feedback / Your Feedback"

# --- Image Canva (decommenter quand disponible) ---
# $CanvaImagePath = ".\canva-image.jpg"

# =====================================================================
# Fonction : genere le HTML personalise pour un contact
# =====================================================================
function Get-EmailHtml {
    param([string]$FirstName, [string]$Link, [string]$LogoSrc = "cid:fflc-logo")

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

  <!-- IMAGE CANVA (decommenter quand disponible) -->
  <!--
  <tr>
    <td align="center" style="padding:0;">
      <img src="cid:canva-image" alt="" width="600"
           style="display:block; width:100%; max-width:600px; height:auto;">
    </td>
  </tr>
  -->

  <!-- ======================================================
       VERSION ALLEMANDE
  ====================================================== -->
  <tr>
    <td style="padding:36px 48px 28px;">

      <p style="margin:0 0 20px; font-size:22px; font-weight:700; color:#141414; line-height:1.3;">
        Hallo $FirstName,<br>
        du geh&ouml;rst zu den Ersten &ndash; danke, dass du dabei bist.
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        Als einer unserer ersten Nutzer von <strong style="color:#141414;">FIT for LIFE Coach</strong>
        bist du Teil von etwas Besonderem. Wir entwickeln diesen Service gemeinsam mit Menschen wie
        dir &ndash; und dein Feedback ist f&uuml;r uns unglaublich wertvoll.
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        Wir w&uuml;rden uns sehr freuen, wenn du dir <strong>2 Minuten</strong> Zeit nimmst,
        um uns zu sagen, wie du FIT for LIFE Coach bisher erlebt hast.
      </p>

      <!-- Incentive DE -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0;">
        <tr>
          <td style="background-color:#FFF5F5; border-left:4px solid #E70E22; border-radius:6px; padding:16px 20px;">
            <p style="margin:0; font-size:15px; color:#141414; line-height:1.6;">
              <strong>Als Dankesh&ouml;n: CHF 20.&ndash; Gutschein bei Datasport</strong><br>
              Nach dem Ausf&uuml;llen erh&auml;ltst du einen Gutschein &uuml;ber CHF 20.&ndash;,
              den du f&uuml;r die Anmeldung zu einem Event deiner Wahl einl&ouml;sen kannst.
            </p>
          </td>
        </tr>
      </table>

      <!-- CTA DE -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 24px;">
        <tr>
          <!--[if mso]>
          <td align="center" style="background-color:#141414; border-radius:8px;">
          <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word"
            href="$Link" style="height:52px; v-text-anchor:middle; width:240px;"
            arcsize="12%" strokecolor="#141414" fillcolor="#141414">
            <w:anchorlock/>
            <center style="color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700;">
              Zum Fragebogen &#8594;
            </center>
          </v:roundrect>
          </td>
          <![endif]-->
          <!--[if !mso]><!-->
          <td align="center" style="background-color:#141414; border-radius:8px;">
            <a href="$Link"
               style="display:inline-block; padding:15px 36px; background-color:#141414; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700; text-decoration:none; border-radius:8px; letter-spacing:0.3px;">
              Zum Fragebogen &#8594;
            </a>
          </td>
          <!--<![endif]-->
        </tr>
      </table>

      <p style="margin:0; font-size:13px; color:#ADADAD; line-height:1.5;">
        Der Fragebogen dauert ca. 2 Minuten. Du kannst die Sprache im Formular selbst w&auml;hlen.
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
        you're one of the first &ndash; thank you for being part of this.
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        As one of our first <strong style="color:#141414;">FIT for LIFE Coach</strong> users,
        you're part of something special. We're building this service together with people like
        you &ndash; and your feedback means the world to us.
      </p>

      <p style="margin:0 0 16px; font-size:16px; color:#444444; line-height:1.65;">
        We'd love for you to take <strong>2 minutes</strong> to share how you've experienced
        FIT for LIFE Coach so far.
      </p>

      <!-- Incentive EN -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin:24px 0;">
        <tr>
          <td style="background-color:#FFF5F5; border-left:4px solid #E70E22; border-radius:6px; padding:16px 20px;">
            <p style="margin:0; font-size:15px; color:#141414; line-height:1.6;">
              <strong>As a thank-you: CHF 20.&ndash; voucher at Datasport</strong><br>
              After completing the survey, you'll receive a CHF 20.&ndash; voucher to use
              for registering for an event of your choice.
            </p>
          </td>
        </tr>
      </table>

      <!-- CTA EN -->
      <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin:8px 0 24px;">
        <tr>
          <!--[if mso]>
          <td align="center" style="background-color:#141414; border-radius:8px;">
          <v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w="urn:schemas-microsoft-com:office:word"
            href="$Link" style="height:52px; v-text-anchor:middle; width:240px;"
            arcsize="12%" strokecolor="#141414" fillcolor="#141414">
            <w:anchorlock/>
            <center style="color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700;">
              Go to the survey &#8594;
            </center>
          </v:roundrect>
          </td>
          <![endif]-->
          <!--[if !mso]><!-->
          <td align="center" style="background-color:#141414; border-radius:8px;">
            <a href="$Link"
               style="display:inline-block; padding:15px 36px; background-color:#141414; color:#ffffff; font-family:'Segoe UI',Arial,sans-serif; font-size:16px; font-weight:700; text-decoration:none; border-radius:8px; letter-spacing:0.3px;">
              Go to the survey &#8594;
            </a>
          </td>
          <!--<![endif]-->
        </tr>
      </table>

      <p style="margin:0; font-size:13px; color:#ADADAD; line-height:1.5;">
        The survey takes approx. 2 minutes. You can choose the language directly in the form.
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
        Du erh&auml;ltst / You receive this email because you activated FIT for LIFE Coach on
        <a href="https://www.datasport.com" style="color:#E70E22; text-decoration:none;">datasport.com</a>.
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
    $row       = Import-Csv $CsvPath | Select-Object -First 1
    $previewPath = "$PSScriptRoot\preview-email.html"
    $logoSrcPreview = if ($LogoPath) { "file:///" + ((Resolve-Path $LogoPath).Path -replace '\\', '/') } else { "" }
    Get-EmailHtml -FirstName $row.first_name -Link $row.link -LogoSrc $logoSrcPreview |
        Out-File -FilePath $previewPath -Encoding UTF8
    Write-Host "Preview genere : $previewPath" -ForegroundColor Green
    Write-Host "Contact test   : $($row.first_name) ($($row.email))" -ForegroundColor Gray
    Start-Process $previewPath
    exit 0
}

# =====================================================================

$contacts = Import-Csv $CsvPath
if ($Limit -gt 0) { $contacts = $contacts | Select-Object -First $Limit }
$total = ($contacts | Measure-Object).Count

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  FIT for LIFE Coach - Envoi emails"    -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "  Contacts   : $total"
Write-Host "  From       : $FromEmail"
Write-Host "  Sujet      : $Subject"
if ($DryRun) {
    Write-Host "  MODE       : DRY RUN (aucun email envoye)" -ForegroundColor Yellow
}
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

if (-not $DryRun) {
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

    $htmlBody = Get-EmailHtml -FirstName $row.first_name -Link $row.link

    if ($DryRun) {
        Write-Host "  [DRY RUN] $($row.email) ($($row.first_name))" -ForegroundColor Gray
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

        # Image Canva (decommenter quand disponible)
        # if ($CanvaImagePath -and (Test-Path $CanvaImagePath)) {
        #     $canvaFull = (Resolve-Path $CanvaImagePath).Path
        #     $attC = $mail.Attachments.Add($canvaFull)
        #     $attC.PropertyAccessor.SetProperty(
        #         "http://schemas.microsoft.com/mapi/proptag/0x3712001E",
        #         "canva-image"
        #     )
        # }

        $mail.HTMLBody = $htmlBody
        $mail.Send()

        $sent++
        Write-Host "  [OK] $($row.email) ($($row.first_name))  [$sent/$total]" -ForegroundColor Green
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
