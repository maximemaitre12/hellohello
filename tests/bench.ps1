param([string]$Planner = "..\adresses-outlook.ps1", [int[]]$Sizes = @(1500, 40000), [int[]]$Caps = @(5, 8), [int[]]$Budgets = @(300, 1000, 3000), [int]$Depth = 4)
$src = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot $Planner), [System.Text.Encoding]::UTF8)
$m = [regex]::Match($src, "(?s)\`$PlanifSource = @'\r?\n(.*?)\r?\n'@")
$carnet = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "carnet-geant.cs"), [System.Text.Encoding]::UTF8)
$run = @'
public static class Run {
  public static int[] Go(FauxOutlook o, int depth, int budget, out string why) {
    Planif p = new Planif("", depth, budget);
    HashSet<string> found = new HashSet<string>();
    string q; int full = 0;
    while ((q = p.Next()) != null) {
      var l = o.Ask(q);
      if (l.Count == o.Cap) full++;
      p.Report(q, l.Select(c => c.Name).ToArray(), l.Select(c => c.Mail).ToArray(), l.Select(c => true).ToArray());
      foreach (var c in l) found.Add(c.Mail);
    }
    why = p.StopReason;
    return new[] { found.Count, o.Queries, o.Empty, full };
  }
}
'@
Add-Type -ReferencedAssemblies System.Core -TypeDefinition ("using System.Linq;`n" + $m.Groups[1].Value + "`n" + $carnet + "`n" + $run)
$rows = @()
foreach ($size in $Sizes) { foreach ($cap in $Caps) { foreach ($b in $Budgets) {
  $why = ""; $t = [Diagnostics.Stopwatch]::StartNew()
  $o = [FauxOutlook]::new($size, $cap, 7)
  $r = [Run]::Go($o, $Depth, $b, [ref]$why)
  $rows += [pscustomobject]@{ Carnet = $size; Plafond = $cap; Budget = $b; Adresses = $r[0]; Recherches = $r[1]; Vides = $r[2]; Pleines = $r[3]
    "Par rech" = "{0:0.00}" -f ($r[0] / [Math]::Max(1,$r[1])); "Max theo" = [Math]::Min($size, $cap * $r[1]); Arret = $why; ms = $t.ElapsedMilliseconds }
}}}
$rows | Format-Table -AutoSize | Out-String -Width 200
