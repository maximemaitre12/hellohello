# Courbe du planificateur : adresses trouvées après N recherches, rapportées
# au total réel du faux carnet. Sert à comparer deux versions de l'algorithme.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File outils\tests\courbe.ps1 -Sizes 1500 -Caps 5 -Max 3000
#
# -Planner pointe vers le script à tester (par défaut celui du dossier parent),
# ou vers un fichier .cs qui contient seulement la classe Planif.
# PLANIF_CTOR (variable d'environnement) donne l'appel du constructeur, au cas où il change.
param([string]$Planner = "..\adresses-outlook.ps1", [string]$Sizes = "400,1500,5000,40000", [string]$Caps = "5,8",
      [int]$Max = 3000, [string]$Ctor = $env:PLANIF_CTOR)
if (-not $Ctor) { $Ctor = 'new Planif("")' }
$src = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot $Planner), [System.Text.Encoding]::UTF8)
if ($Planner.EndsWith(".cs")) { $code = $src } else {
  $m = [regex]::Match($src, "(?s)\`$PlanifSource = @'\r?\n(.*?)\r?\n'@")
  if (-not $m.Success) { throw "Bloc du planificateur introuvable" }
  $code = $m.Groups[1].Value
}
$carnet = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "carnet-geant.cs"), [System.Text.Encoding]::UTF8)
$run = @'
public static class Courbe {
  static int[] Marks = { 26, 100, 300, 1000, 3000, 10000, 30000 };
  public static string Go(FauxOutlook o, int budget) {
    Planif p = __CTOR__;
    HashSet<string> found = new HashSet<string>();
    var sw = System.Diagnostics.Stopwatch.StartNew();
    string q; StringBuilder b = new StringBuilder(); int mi = 0;
    while ((q = p.Next()) != null) {
      var l = o.Ask(q);
      p.Report(q, l.Select(c => c.Name).ToArray(), l.Select(c => c.Mail).ToArray(), l.Select(c => true).ToArray());
      foreach (var c in l) found.Add(c.Mail);
      while (mi < Marks.Length && o.Queries == Marks[mi]) { b.AppendFormat("{0}:{1} ", Marks[mi], found.Count); mi++; }
      if (o.Queries >= budget) break;
    }
    return string.Format("total={0} trouve={1} ({2:0}%) rech={3} vides={4} arret={5} ms={6} | {7}",
      o.All.Count, found.Count, 100.0 * found.Count / o.All.Count, o.Queries, o.Empty, p.StopReason, sw.ElapsedMilliseconds, b);
  }
}
'@.Replace("__CTOR__", $Ctor)
Add-Type -ReferencedAssemblies System.Core -TypeDefinition ("using System.Linq;`n" + $code + "`n" + $carnet + "`n" + $run)
foreach ($size in ($Sizes -split ",")) { foreach ($cap in ($Caps -split ",")) {
  "carnet=$size plafond=$cap  " + [Courbe]::Go([FauxOutlook]::new([int]$size, [int]$cap, 7), $Max)
}}
