# Banc d'essai du planificateur d'adresses-outlook.ps1, sans Outlook.
#
# Un faux carnet d'adresses répond comme le nouvel Outlook : un contact sort
# dès qu'un mot de son nom ou de son adresse commence par la recherche, les
# contacts les plus fréquents en tête, liste coupée au plafond. On y fait
# tourner l'ancien algorithme (largeur d'abord, alphabétique, 26 suites à
# chaque fois) et le planificateur actuel, extrait tel quel du script.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File outils\tests\simulation-adresses.ps1

$src = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot "..\adresses-outlook.ps1"), [System.Text.Encoding]::UTF8)
$m = [regex]::Match($src, "(?s)\`$PlanifSource = @'\r?\n(.*?)\r?\n'@")
if (-not $m.Success) { throw "Bloc du planificateur introuvable" }
$banc = @'

public class FauxOutlook {
  public class Contact { public string Name, Mail; public double W; public List<string> Toks; }
  public List<Contact> All = new List<Contact>();
  public int Cap; public int Queries = 0, Empty = 0, Seconds = 0;

  static string[] First = ("maxime thomas samuel gildas sabrina larisa noe yelena qinkunze irene youriy vladyslav william xavier mason " +
    "camille lea manon chloe ines jade louise emma alice lucas hugo louis jules arthur paul nathan theo raphael " +
    "olena oksana dmytro andriy iryna wei jing xiaoming yuxuan zihan haoran lei chen ana lucia pablo javier carmen " +
    "sophie julie claire marie anne pierre jean nicolas antoine julien romain kevin quentin yann zoe victor").Split(' ');
  static string[] Last = ("bernier poirel maitre jimenez konovalova neuvelt ouidir tian rodriguez strashnyi graff galezowski zhang " +
    "martin bernard dubois durand lefebvre leroy moreau simon laurent michel garcia david bertrand roux vincent fournier " +
    "morel girard andre mercier dupont lambert bonnet francois martinez shevchenko kovalenko bondarenko tkachenko kravchenko " +
    "wang li liu yang huang zhao wu zhou xu sun ma zhu hu guo lopez sanchez perez gomez fernandez gonzalez blanc " +
    "chevalier robin clement morin nicolas henry roussel mathieu gauthier masson marchand duval denis lemaire fabre").Split(' ');
  static string[] Domains = { "test.com", "test.com", "test.com", "test.com", "deloitte.fr",
    "groupeonepoint.com", "gmail.com", "gmail.com", "hotmail.com", "uv.es", "farmasoft.ua" };
  static string[] Services = { "comptabilite", "iael", "bba2en.studies", "scolarite", "admissions", "career.center",
    "library", "it.support", "international", "alumni", "housing", "exams", "student.life" };

  public FauxOutlook(int n, int cap, int seed) {
    Cap = cap;
    Random r = new Random(seed);
    HashSet<string> used = new HashSet<string>();
    foreach (string s in Services) Add(s.Replace('.', ' ').ToUpper(), s + "@test.com", used);
    while (All.Count < n) {
      string f = First[r.Next(First.Length)], l = Last[r.Next(Last.Length)], d = Domains[r.Next(Domains.Length)];
      int fmt = r.Next(4);
      string local = fmt == 0 ? f + "." + l : fmt == 1 ? f[0] + l : fmt == 2 ? l + "." + f : f + l;
      if (used.Contains(local + "@" + d)) local += r.Next(2, 99);
      string name = r.Next(2) == 0 ? Cap1(f) + " " + l.ToUpper() : l.ToUpper() + " " + Cap1(f);
      Add(name, local + "@" + d, used);
    }
    // Fréquence de contact : quelques personnes très fréquentes, une longue traîne.
    var order = All.OrderBy(c => r.Next()).ToList();
    for (int i = 0; i < order.Count; i++) order[i].W = 1.0 / Math.Pow(i + 1, 1.1);
  }
  static string Cap1(string s) { return char.ToUpper(s[0]) + s.Substring(1); }
  void Add(string name, string mail, HashSet<string> used) {
    if (!used.Add(mail)) return;
    All.Add(new Contact { Name = name, Mail = mail, Toks = Planif.Tokens(name + " " + mail) });
  }

  public List<Contact> Ask(string q) {
    Queries++;
    List<string> qw = Planif.Tokens(q);
    var hits = All.Where(c => qw.All(w => c.Toks.Any(t => t.StartsWith(w, StringComparison.Ordinal))))
                  .OrderByDescending(c => c.W).Take(Cap).ToList();
    // Durées du réglage Normal : environ 2,8 s pour une liste, 2,5 s pour une recherche vide.
    if (hits.Count == 0) { Empty++; Seconds += 3; } else Seconds += 3;
    return hits;
  }
}

public static class Banc {
  public static int[] Nouveau(FauxOutlook o, int depth, int budget, out string why) {
    Planif p = new Planif("", depth, budget);
    HashSet<string> found = new HashSet<string>();
    string q;
    while ((q = p.Next()) != null) {
      var l = o.Ask(q);
      p.Report(q, l.Select(c => c.Name).ToArray(), l.Select(c => c.Mail).ToArray(), l.Select(c => true).ToArray());
      foreach (var c in l) found.Add(c.Mail);
    }
    why = p.StopReason;
    return new[] { found.Count, o.Queries, o.Empty };
  }

  // L'ancien algorithme, tel qu'il était dans le script.
  public static int[] Ancien(FauxOutlook o, int depth, int budget) {
    HashSet<string> found = new HashSet<string>();
    List<string> level = new List<string>();
    for (char c = 'a'; c <= 'z'; c++) level.Add(c.ToString());
    int planned = 26, cap = 0;
    while (level.Count > 0) {
      var counts = new List<KeyValuePair<string, int>>();
      foreach (string q in level) {
        var l = o.Ask(q);
        if (l.Count > cap) cap = l.Count;
        counts.Add(new KeyValuePair<string, int>(q, l.Count));
        foreach (var c in l) found.Add(c.Mail);
      }
      var next = new List<string>();
      foreach (var kv in counts) {
        if (kv.Key.Length >= depth || kv.Value < 3 || kv.Value < cap) continue;
        if (planned + 26 > budget) break;
        for (char c = 'a'; c <= 'z'; c++) next.Add(kv.Key + c);
        planned += 26;
      }
      level = next;
    }
    return new[] { found.Count, o.Queries, o.Empty };
  }
}
'@
Add-Type -ReferencedAssemblies System.Core -TypeDefinition ("using System.Linq;`n" + $m.Groups[1].Value + "`n" + $banc)

$rows = @()
foreach ($size in 400, 1500, 5000) {
  foreach ($cap in 5, 8) {
    foreach ($budget in 100, 300, 1000) {
      foreach ($depth in 3, 4) {
        $a = [Banc]::Ancien(([FauxOutlook]::new($size, $cap, 7)), $depth, $budget)
        $why = ""
        $n = [Banc]::Nouveau(([FauxOutlook]::new($size, $cap, 7)), $depth, $budget, [ref]$why)
        $rows += [pscustomobject]@{
          Carnet = $size; Plafond = $cap; Budget = $budget; Prof = $depth
          "Ancien adr" = $a[0]; "Ancien rech" = $a[1]; "Ancien vides" = $a[2]
          "Nouveau adr" = $n[0]; "Nouveau rech" = $n[1]; "Nouveau vides" = $n[2]; Arret = $why
          "Gain" = "{0:+0;-0}%" -f (100.0 * ($n[0] - $a[0]) / [Math]::Max(1, $a[0]))
        }
      }
    }
  }
}
$rows | Format-Table -AutoSize | Out-String -Width 250
