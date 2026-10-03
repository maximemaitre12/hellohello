# Fait tourner la vraie boucle du passage d'adresses-outlook.ps1 (interface
# comprise, fenêtre jamais affichée) contre un faux Outlook. Aucune frappe ne
# part vers Windows, et ni les réglages ni le journal de reprise enregistrés ne
# sont touchés.
#
# Trois passages : un d'une traite, jusqu'au bout ; un arrêté en route ; puis
# sa reprise, qui doit finir exactement comme le premier sans retaper une
# seule recherche. Les fichiers doivent se remplir pendant le passage.
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

# Le journal de reprise et les réglages restent dans le dossier du test.
$cfgDir = $out
$resumePath = Join-Path $out "reprise.jsonl"
Remove-Item $resumePath -ErrorAction SilentlyContinue
Get-ChildItem $out -Filter "adresses-*" | Remove-Item

# --- faux Outlook
$script:faux = $null
$script:buffer = ""; $script:shown = @(); $script:staleOnce = $false; $script:staleHits = 0; $script:typed = 0
$script:stopAt = 0; $script:midCount = -1
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
    # À mi-course, ce qui est trouvé doit déjà être sur le disque.
    if ($script:typed -eq 60) { $script:midCount = (Csv-Rows).Count }
    if ($script:stopAt -and $script:typed -ge $script:stopAt) { $script:stop = $true }
  }
}
function Lire-Suggestions($p) {
  if ($script:staleOnce) { $script:staleOnce = $false; if ($script:shown.Count) { $script:staleHits++; return $script:shown } }
  if (-not $script:buffer) { return @() }
  $hits = @($script:faux.Ask($script:buffer))
  if ($Source -eq "gmail") {
    $names = @(($hits | ForEach-Object { $_.Name + $_.Mail }) -join "") + @($hits | ForEach-Object { $_.Name + " " + $_.Mail })
    $script:shown = @(Parse-Gmail ([string[]]$names))
  } else {
    $script:shown = @($hits | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Mail = $_.Mail.ToLower() } })
  }
  return $script:shown
}

# Le CSV du passage en cours, lu sans gêner l'app qui l'a ouvert en écriture.
function Csv-Rows {
  $f = Get-ChildItem $out -Filter "adresses-$Source-*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $f) { return @() }
  $fs = [System.IO.File]::Open($f.FullName, "Open", "Read", "ReadWrite")
  $text = (New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)).ReadToEnd(); $fs.Close()
  return @($text | ConvertFrom-Csv -Delimiter ";")
}

function Run([bool]$resume, [int]$stopAt) {
  $script:faux = [FauxOutlook]::new(400, 8, 7)
  $script:buffer = ""; $script:shown = @(); $script:typed = 0; $script:stopAt = $stopAt; $script:midCount = -1
  $t0 = Get-Date
  Passage $resume
  Refresh-Resume
  [pscustomobject]@{ Typed = $script:typed; Mid = $script:midCount; Rows = @(Csv-Rows); Notice = $ui.NoticeText.Text
                     Resume = $ui.ResumeBox.Visibility -eq "Visible"; Secs = [int]((Get-Date) - $t0).TotalSeconds }
}

# --- réglages du test
Apply-Config
$ui.Prefix.Text = ""; $ui.Keep.Text = ""; $ui.Exclude.Text = ""; $ui.Dedupe.IsChecked = $true
$ui.SpeedNormal.IsChecked = $true; $ui.FmtBoth.IsChecked = $true; $ui.Folder.Text = $out
$ui.OpenAtEnd.IsChecked = $false; $ui.CopyAtEnd.IsChecked = $false
if ($Source -eq "gmail") { $ui.SrcGmail.IsChecked = $true } else { $ui.SrcOutlook.IsChecked = $true }

$fails = @()

# 1. D'une traite.
$a = Run $false 0
$byMail = @{}; foreach ($c in $script:faux.All) { $byMail[$c.Mail.ToLower()] = $c }
$wrong = @($a.Rows | Where-Object { -not [Planif]::Matches($byMail[$_.Adresse].Name + " " + $_.Adresse, $_.Recherche) })
if ($wrong.Count) { $fails += "$($wrong.Count) adresses attribuées à une recherche qui ne les montre pas (liste périmée lue)" }
$dupes = @($a.Rows | Group-Object Adresse | Where-Object Count -gt 1)
if ($dupes.Count) { $fails += "$($dupes.Count) adresses notées deux fois malgré le dédoublonnage" }
if ($a.Rows.Count -lt 0.85 * $script:faux.All.Count) { $fails += "seulement $($a.Rows.Count) adresses sur $($script:faux.All.Count)" }
if ($a.Typed -le 300) { $fails += "$($a.Typed) recherches : le passage s'est arrêté comme avec l'ancienne limite" }
if ($a.Mid -le 0) { $fails += "rien sur le disque après 60 recherches ($($a.Mid) lignes) : l'enregistrement n'est pas au fur et à mesure" }
if ($script:staleHits -lt 20) { $fails += "le faux Outlook n'a presque pas servi de liste périmée ($script:staleHits)" }
if ($a.Notice -notmatch "^Terminé") { $fails += "message de fin inattendu : $($a.Notice)" }
if ($a.Resume -or (Test-Path $resumePath)) { $fails += "un passage allé au bout laisse une reprise proposée" }
"1. d'une traite : $($a.Typed) recherches, $($a.Rows.Count) adresses sur $($script:faux.All.Count), $($a.Mid) lignes sur le disque après 60 recherches, $($a.Secs) s"

# 2. Arrêté en route.
Start-Sleep -Seconds 61   # nouveau nom de fichier, à la minute
$b = Run $false 150
if ($b.Notice -notmatch "^Arrêté") { $fails += "arrêt : message inattendu : $($b.Notice)" }
if (-not $b.Resume) { $fails += "arrêt : la reprise n'est pas proposée" }
if ($ui.BtnRun.Content -ne "Reprendre le passage") { $fails += "arrêt : le bouton dit « $($ui.BtnRun.Content) »" }
"2. arrêté : $($b.Typed) recherches, $($b.Rows.Count) adresses enregistrées"
"   proposé : $($ui.ResumeText.Text)"

# 3. Reprise : finit comme le passage d'une traite, sans retaper.
$c = Run $true 0
if ($c.Typed + $b.Typed -ne $a.Typed) { $fails += "reprise : $($b.Typed) + $($c.Typed) recherches au lieu de $($a.Typed)" }
$sa = ($a.Rows | ForEach-Object { $_.Recherche + "|" + $_.Adresse }) -join "`n"
$sc = ($c.Rows | ForEach-Object { $_.Recherche + "|" + $_.Adresse }) -join "`n"
if ($sa -ne $sc) { $fails += "reprise : le fichier final diffère du passage d'une traite ($($c.Rows.Count) lignes contre $($a.Rows.Count))" }
if ($c.Notice -notmatch "^Terminé" -or $c.Notice -notmatch "reprise a ajouté") { $fails += "reprise : message inattendu : $($c.Notice)" }
if ($c.Resume -or (Test-Path $resumePath)) { $fails += "reprise : le journal n'est pas effacé à la fin" }
"3. reprise : $($c.Typed) recherches de plus, $($c.Rows.Count) adresses au total"
"   message : $($c.Notice)"

if ($fails.Count) { "ÉCHEC"; $fails | ForEach-Object { " - $_" }; exit 1 } else { "OK" }
