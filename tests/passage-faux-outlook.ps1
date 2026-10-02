# Fait tourner la vraie boucle du passage d'adresses-outlook.ps1 (interface
# comprise, fenêtre jamais affichée) contre un faux Outlook. Aucune frappe ne
# part vers Windows et les réglages enregistrés ne sont pas touchés.
#
# Le faux Outlook montre encore l'ancienne liste à la première lecture qui
# suit chaque frappe, comme le vrai quand il est lent : la lecture doit
# l'ignorer et ne noter que ce qui répond à la recherche tapée.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -STA -File outils\tests\passage-faux-outlook.ps1
#   ... -File outils\tests\passage-faux-outlook.ps1 -Source gmail
#
# Avec -Source gmail, le faux montre les suggestions sous les noms que Gmail
# leur donne (« Nom adresse », plus un élément qui les colle toutes bout à
# bout), et c'est la vraie lecture Gmail de l'app qui les démêle.

param([ValidateSet("outlook", "gmail")][string]$Source = "outlook")

$ErrorActionPreference = "Stop"
$out = Join-Path $env:TEMP "adresses-outlook-test"
New-Item -ItemType Directory -Force -Path $out | Out-Null

$sim = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "simulation-adresses.ps1"), [System.Text.Encoding]::UTF8)
$sim = $sim.Substring(0, $sim.IndexOf('$rows = @()')).Replace('$PSScriptRoot', "`"$PSScriptRoot`"")
. ([scriptblock]::Create($sim))

$src = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "..\adresses-outlook.ps1"), [System.Text.Encoding]::UTF8)
$src = $src.Replace("Add-Type -TypeDefinition `$PlanifSource", "")
$src = $src.Substring(0, $src.LastIndexOf("Apply-Config"))
. ([scriptblock]::Create($src))

# --- faux Outlook
$faux = [FauxOutlook]::new(1500, 8, 7)
$script:buffer = ""; $script:shown = @(); $script:staleOnce = $false; $script:staleHits = 0; $script:typed = 0
function Outlook-Process { [pscustomobject]@{ Id = 1; MainWindowHandle = [IntPtr]1 } }
function Au-Premier-Plan($p) { $true }
function Focus-Dans-Champ-A { $true }
function Mettre-Devant($p) { }
function Ouvrir-Gmail { $script:targetPid = 1; [pscustomobject]@{ Id = 1; MainWindowHandle = [IntPtr]1 } }
function Wait([int]$ms) { Pump }
function Save-Config { }
function Taper([string]$keys) {
  if ($keys -match '^\{BACKSPACE (\d+)\}$') {
    $n = [Math]::Min([int]$matches[1], $script:buffer.Length); $script:buffer = $script:buffer.Substring(0, $script:buffer.Length - $n)
  } elseif ($keys -ne "^n") {
    $script:buffer += ($keys -replace '\{(.)\}', '$1'); $script:staleOnce = $true; $script:typed++
  }
}
function Lire-Suggestions($p) {
  if ($script:staleOnce) { $script:staleOnce = $false; if ($script:shown.Count) { $script:staleHits++; return $script:shown } }
  if (-not $script:buffer) { return @() }
  $hits = @($faux.Ask($script:buffer))
  if ($Source -eq "gmail") {
    $names = @(($hits | ForEach-Object { $_.Name + $_.Mail }) -join "") + @($hits | ForEach-Object { $_.Name + " " + $_.Mail })
    $script:shown = @(Parse-Gmail ([string[]]$names))
  } else {
    $script:shown = @($hits | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Mail = $_.Mail.ToLower() } })
  }
  return $script:shown
}

# --- réglages du test
Apply-Config
$ui.SmartBudget.Text = "300"; $ui.SmartD3.IsChecked = $true
$ui.Prefix.Text = ""; $ui.Keep.Text = ""; $ui.Exclude.Text = ""; $ui.Dedupe.IsChecked = $true
$ui.SpeedNormal.IsChecked = $true; $ui.FmtBoth.IsChecked = $true; $ui.Folder.Text = $out
$ui.OpenAtEnd.IsChecked = $false; $ui.CopyAtEnd.IsChecked = $false
if ($Source -eq "gmail") { $ui.SrcGmail.IsChecked = $true } else { $ui.SrcOutlook.IsChecked = $true }

$t0 = Get-Date
Passage
$secs = [int]((Get-Date) - $t0).TotalSeconds

# --- vérifications
$fails = @()
$csv = Get-ChildItem $out -Filter "adresses-$Source-*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$rows = @(Import-Csv $csv.FullName -Delimiter ";" -Encoding UTF8)
$byMail = @{}; foreach ($c in $faux.All) { $byMail[$c.Mail.ToLower()] = $c }
$wrong = @($rows | Where-Object { -not [Planif]::Matches($byMail[$_.Adresse].Name + " " + $_.Adresse, $_.Recherche) })
if ($wrong.Count) { $fails += "$($wrong.Count) adresses attribuées à une recherche qui ne les montre pas (liste périmée lue)" }
$dupes = @($rows | Group-Object Adresse | Where-Object Count -gt 1)
if ($dupes.Count) { $fails += "$($dupes.Count) adresses notées deux fois malgré le dédoublonnage" }
# Le planificateur seul en trouve environ 700 dans ce carnet en 300 recherches.
if ($rows.Count -lt 500) { $fails += "seulement $($rows.Count) adresses : la lecture reste bloquée sur une ancienne liste ?" }
if ($script:typed -gt 300) { $fails += "$script:typed recherches pour une limite de 300" }
if ($script:staleHits -lt 20) { $fails += "le faux Outlook n'a presque pas servi de liste périmée ($script:staleHits)" }
if ($ui.NoticeText.Text -notmatch "^Terminé") { $fails += "message de fin inattendu : $($ui.NoticeText.Text)" }

"recherches tapées : $script:typed (lectures de liste : $($faux.Queries))"
"adresses notées : $($rows.Count) (le carnet en compte $($faux.All.Count))"
"listes périmées présentées puis ignorées : $script:staleHits"
"durée du test : $secs s"
"message : $($ui.NoticeText.Text)"
if ($fails.Count) { "ÉCHEC"; $fails | ForEach-Object { " - $_" }; exit 1 } else { "OK" }
