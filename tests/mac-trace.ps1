# Fait tourner le planificateur de la version Windows sur plusieurs faux
# carnets d'adresses, et enregistre les carnets et la suite exacte des
# recherches tapées. outils\tests\mac-planif.mjs rejoue les mêmes carnets avec
# le planificateur de la version Mac et doit trouver la même suite.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File outils\tests\mac-trace.ps1 <fichier.json>

param([Parameter(Mandatory)][string]$Out)

$sim = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "simulation-adresses.ps1"), [System.Text.Encoding]::UTF8)
$sim = $sim.Substring(0, $sim.IndexOf('$rows = @()')).Replace('$PSScriptRoot', "`"$PSScriptRoot`"")
. ([scriptblock]::Create($sim))

$cases = @()
foreach ($c in @(
  @{ Size = 400;  Cap = 5; Budget = 1000000; Depth = 12; Prefix = "" },
  @{ Size = 400;  Cap = 8; Budget = 1000000; Depth = 12; Prefix = "" },
  @{ Size = 1500; Cap = 8; Budget = 1000000; Depth = 12; Prefix = "" },
  @{ Size = 1500; Cap = 5; Budget = 300;     Depth = 3;  Prefix = "" },
  @{ Size = 5000; Cap = 5; Budget = 3000;    Depth = 12; Prefix = "" },
  @{ Size = 1500; Cap = 8; Budget = 1000000; Depth = 12; Prefix = "ma" },
  @{ Size = 1500; Cap = 8; Budget = 1000000; Depth = 12; Prefix = "jean d" }
)) {
  $o = [FauxOutlook]::new($c.Size, $c.Cap, 7)
  # Un filtre de domaine sur un cas sur deux, pour couvrir keep[].
  $keepDomain = if ($c.Size -eq 1500) { "test.com" } else { "" }
  $p = New-Object Planif -ArgumentList $c.Prefix, $c.Depth, $c.Budget
  $queries = New-Object System.Collections.Generic.List[string]
  while ($true) {
    $q = $p.Next()
    if (-not $q) { break }
    $queries.Add($q)
    $l = $o.Ask($q)
    $keep = [bool[]]@($l | ForEach-Object { -not $keepDomain -or $_.Mail.EndsWith("@" + $keepDomain) })
    $p.Report($q, [string[]]@($l | ForEach-Object { $_.Name }), [string[]]@($l | ForEach-Object { $_.Mail }), $keep) | Out-Null
  }
  $cases += [ordered]@{
    size = $c.Size; cap = $c.Cap; budget = $c.Budget; depth = $c.Depth; prefix = $c.Prefix; keepDomain = $keepDomain
    contacts = @($o.All | ForEach-Object { ,@($_.Name, $_.Mail, $_.W.ToString("R", [Globalization.CultureInfo]::InvariantCulture)) })
    queries = @($queries); stop = $p.StopReason
  }
}
[System.IO.File]::WriteAllText($Out, (ConvertTo-Json $cases -Depth 6 -Compress), [System.Text.UTF8Encoding]::new($false))
"$($cases.Count) cas ecrits"
