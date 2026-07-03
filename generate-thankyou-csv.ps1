# ============================================================
#  generate-thankyou-csv.ps1  -  Genere contacts-thankyou.csv
#  Associe a chaque repondant (responded.txt) un coupon unique
#  pris dans coupons.csv, en recuperant le prenom/nom depuis
#  contacts.csv.
# ============================================================
#
#  USAGE :
#    .\generate-thankyou-csv.ps1
#
#  Regenere ce fichier UNIQUEMENT si responded.txt ou coupons.csv
#  ont change. Les coupons sont une ressource limitee et non
#  reutilisable : ne pas relancer sans raison une fois les mails
#  envoyes, sous peine de reassigner un coupon deja communique
#  a quelqu'un d'autre si l'ordre des sources change.
# ============================================================

param(
    [string]$RespondedPath = ".\responded.txt",
    [string]$ContactsPath  = ".\contacts.csv",
    [string]$CouponsPath   = ".\coupons.csv",
    [string]$OutPath       = ".\contacts-thankyou.csv"
)

if (-not (Test-Path $RespondedPath)) { Write-Host "ERREUR : introuvable $RespondedPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $ContactsPath))  { Write-Host "ERREUR : introuvable $ContactsPath" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $CouponsPath))   { Write-Host "ERREUR : introuvable $CouponsPath" -ForegroundColor Red; exit 1 }

# Dedup en conservant l'ordre d'apparition dans responded.txt
$seen = New-Object System.Collections.Generic.HashSet[string]
$responded = @()
foreach ($line in Get-Content $RespondedPath) {
    $email = $line.Trim().ToLower()
    if ($email -eq '') { continue }
    if ($seen.Add($email)) { $responded += $email }
}

$contacts = Import-Csv $ContactsPath
$contactsByEmail = @{}
foreach ($row in $contacts) { $contactsByEmail[$row.email.ToLower()] = $row }

$coupons = @(Import-Csv $CouponsPath)

Write-Host "Repondants uniques : $($responded.Count)"
Write-Host "Coupons disponibles : $($coupons.Count)"

if ($coupons.Count -lt $responded.Count) {
    Write-Host "ERREUR : pas assez de coupons ($($coupons.Count)) pour $($responded.Count) repondants." -ForegroundColor Red
    exit 1
}

$missing = @()
$result = @()
for ($i = 0; $i -lt $responded.Count; $i++) {
    $email  = $responded[$i]
    $coupon = $coupons[$i]
    $contact = $contactsByEmail[$email]

    if (-not $contact) {
        $missing += $email
    }

    $result += [PSCustomObject]@{
        email               = $email
        first_name          = if ($contact) { $contact.first_name } else { '' }
        last_name           = if ($contact) { $contact.last_name } else { '' }
        coupon_code         = $coupon.coupon_code
        coupon_password     = $coupon.password
        coupon_currency     = $coupon.currency
        coupon_amount       = $coupon.amount
        coupon_valid_until  = $coupon.valid_until
    }
}

$result | Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Genere : $OutPath ($($result.Count) lignes)" -ForegroundColor Green
$unusedCoupons = $coupons.Count - $responded.Count
if ($unusedCoupons -gt 0) {
    Write-Host "Coupons non utilises (restants dans $CouponsPath, index $($responded.Count)..$($coupons.Count - 1)) : $unusedCoupons" -ForegroundColor Yellow
}
if ($missing.Count -gt 0) {
    Write-Host "ATTENTION : $($missing.Count) email(s) absent(s) de $ContactsPath - prenom/nom vides, a completer manuellement :" -ForegroundColor Yellow
    $missing | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }
}
