# Adresses Outlook
#
# Ouvre un nouveau message dans le nouvel Outlook, tape les lettres de A à Z
# dans le champ « À » et note toutes les adresses que propose la liste de
# suggestions. Là où la liste est pleine, des adresses restent cachées : la
# recherche est approfondie (a devient aa, ab..., martin devient martin j...),
# jusqu'à ce qu'il n'y ait plus rien à trouver. Chaque adresse est écrite dans
# le fichier texte et/ou CSV dès qu'elle est trouvée, et un passage interrompu
# reprend là où il s'était arrêté.
#
# Rien n'est jamais envoyé. Le brouillon reste ouvert, champ « À » vide.
#
# Pendant le passage, les frappes vont à la fenêtre active. Avant chaque
# recherche, le script vérifie qu'Outlook est au premier plan et que le curseur
# est dans le champ « À », et s'arrête sinon, pour ne jamais taper ailleurs.
# L'app se réduit en un petit panneau toujours visible qui ne prend jamais le
# clavier. Cliquer dessus arrête le passage proprement.
#
# Les réglages sont gardés dans %APPDATA%\AetherOutils\adresses-outlook.json.

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, UIAutomationClient, UIAutomationTypes
Add-Type @"
using System; using System.Text; using System.Runtime.InteropServices;
public class Fen {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr e);
  [DllImport("user32.dll")] static extern IntPtr SendMessageTimeout(IntPtr h, uint msg, IntPtr w, IntPtr l, uint flags, uint timeout, out IntPtr result);
  delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int n);
  // Chrome et Edge ne montrent le contenu de la page à l'accessibilité de
  // Windows qu'une fois qu'on le leur demande : on le leur demande.
  public static void ReveillerPage(IntPtr top) {
    if (top == IntPtr.Zero) return;
    EnumChildWindows(top, (h, l) => {
      StringBuilder sb = new StringBuilder(256); GetClassName(h, sb, 256);
      if (sb.ToString() == "Chrome_RenderWidgetHostHWND") { IntPtr r; SendMessageTimeout(h, 0x003D, IntPtr.Zero, (IntPtr)(-4), 2, 500, out r); }
      return true; }, IntPtr.Zero);
  }
}
"@

# ------------------------------------------------------------ planificateur
#
# Décide quelle recherche taper ensuite. Il ne connaît d'Outlook que ce qu'il
# en voit : la liste de suggestions, plafonnée à quelques lignes, où un contact
# apparaît dès qu'un mot de son nom ou de son adresse commence par la recherche.
#
# - Les 26 lettres passent d'abord : elles donnent la taille de la liste
#   d'Outlook (le plafond) et un premier vocabulaire.
# - Une recherche dont la liste est pleine cache des contacts : ses suites
#   (a devient aa, ab...) deviennent des pistes. Une liste moins que pleine est
#   complète : rien à creuser dessous.
# - Les pistes ne sont pas jouées dans l'ordre alphabétique mais par
#   rendement attendu, sur tout l'alphabet à la fois : combien d'adresses
#   inconnues la liste devrait montrer. Le calcul s'appuie sur les suites de
#   lettres réellement vues dans les noms (« qi » existe chez vous, « qx »
#   non), et retire les adresses déjà notées qui occuperaient la liste. Il
#   apprend en route ce que chaque genre de piste rapporte vraiment en
#   adresses nouvelles, ce qui écarte vite les recherches vides.
# - Pas de profondeur maximale : on creuse tant que la liste est pleine.
# - Un nom trop courant (martin, wang, jean) n'est pas départagé par une
#   lettre de plus. Quand les lettres ne rapportent plus, un deuxième mot
#   prend le relais (« martin j », « martin s »).
# - Le passage s'arrête seul quand il n'y a plus de piste rentable, ou quand
#   une longue série de recherches n'a plus rien apporté.
#
# Testé hors Outlook par outils\tests\courbe.ps1 (adresses trouvées au fil des
# recherches, sur des carnets simulés de 400 à 40 000 contacts) et
# outils\tests\simulation-adresses.ps1, qui extraient ce bloc tel quel.
$PlanifSource = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public class Planif {
  // Une recherche tapée, ou seulement envisagée.
  class Node {
    public List<Node> Kids;
    public string Q; public string[] Words; public Node Parent; public int Depth; public bool Split, Root;
    public double Est, EstRaw, Exact = 1.0, Stale, ExpNew; public int Kind;
    public bool Answered, Virtual; public int Count; public double Novelty;
    public int KCount, KSeen;
  }

  public int Cap = 0, MaxDepth, Budget, Done = 0, SinceNew = 0, Pending = 0, Saturation = 60;
  public string StopReason = "";
  public const double MinScore = 0.1;

  Queue<Node> roots = new Queue<Node>();
  Heap heap = new Heap();
  // Pistes qui ne valent presque rien pour l'instant : revues de loin en loin seulement.
  List<Node> cold = new List<Node>();
  const double Cold = 0.01;
  Dictionary<string, Node> nodes = new Dictionary<string, Node>();
  // Pistes à un seul mot libre, par ce mot : pour savoir lesquelles passent le plafond de connus.
  Dictionary<string, Node> byWord = new Dictionary<string, Node>();
  List<Node> promote = new List<Node>();
  string[] fixedWords = new string[0];
  HashSet<string> mails = new HashSet<string>();
  List<string[]> contacts = new List<string[]>();
  Dictionary<string, List<int>> index = new Dictionary<string, List<int>>();
  double[,] big = new double[27, 26]; double[] bigRow = new double[27];
  double[,] start = new double[26, 26]; double[] startRow = new double[26];
  double[] predicted = new double[12], obtained = new double[12];
  double[] predNew = new double[12], obtNew = new double[12];
  int lastRebuild = 0, lastDeep = 0;
  // Les recherches à deux mots ne viennent qu'une fois les lettres épuisées.
  bool splitPhase = false;
  // Branches qui viennent de rapporter : leurs suites sont à revoir à la hausse.
  List<Node> bump = new List<Node>();

  public Planif(string prefix) : this(prefix, 12, int.MaxValue) { }
  public Planif(string prefix, int maxDepth, int budget) {
    MaxDepth = maxDepth; Budget = budget;
    List<string> fw = Tokens(prefix);
    if (fw.Count > 0 && !(prefix ?? "").EndsWith(" ")) fw.RemoveAt(fw.Count - 1);
    fixedWords = fw.ToArray();
    for (char c = 'a'; c <= 'z'; c++) {
      Node n = Make((prefix ?? "") + c, null, false);
      n.Root = true;
      roots.Enqueue(n);
    }
  }

  Node Make(string q, Node parent, bool split) {
    Node n = new Node { Q = q, Words = Tokens(q).ToArray(), Parent = parent, Split = split, Depth = parent == null ? 1 : parent.Depth + 1 };
    nodes[q] = n;
    if (n.Words.Length > 1) { n.KCount = Known(n.Words); n.KSeen = contacts.Count; }
    if (!split && (parent == null || !parent.Split)) byWord[n.Words[n.Words.Length - 1]] = n;
    return n;
  }

  public static string Norm(string s) {
    if (s == null) return "";
    string d = s.ToLowerInvariant().Normalize(NormalizationForm.FormD);
    StringBuilder b = new StringBuilder(d.Length);
    foreach (char c in d) if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) b.Append(c);
    return b.ToString();
  }

  public static List<string> Tokens(string s) {
    List<string> t = new List<string>();
    StringBuilder w = new StringBuilder();
    foreach (char c in Norm(s) + " ") {
      if ((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')) w.Append(c);
      else if (w.Length > 0) { t.Add(w.ToString()); w.Clear(); }
    }
    return t;
  }

  static bool MatchToks(IList<string> toks, IList<string> words) {
    foreach (string qw in words) {
      bool ok = false;
      foreach (string t in toks) if (t.StartsWith(qw, StringComparison.Ordinal)) { ok = true; break; }
      if (!ok) return false;
    }
    return true;
  }

  // Le contact correspond-il à la recherche, à la façon d'Outlook ? Chaque mot
  // de la recherche doit commencer un mot du nom ou de l'adresse.
  public static bool Matches(string text, string query) { return MatchToks(Tokens(text), Tokens(query)); }

  // Contacts déjà connus qui répondent à ces mots.
  int Known(string[] words) {
    if (words.Length == 0) return contacts.Count;
    List<int> best = null;
    foreach (string w in words) {
      List<int> l;
      if (!index.TryGetValue(w, out l)) return 0;
      if (best == null || l.Count < best.Count) best = l;
    }
    if (words.Length == 1) return best.Count;
    int k = 0;
    foreach (int i in best) if (MatchToks(contacts[i], words)) k++;
    return k;
  }

  // Même chose pour une piste, en ne regardant que les contacts appris depuis
  // la dernière fois (les numéros de contact ne font que croître).
  int Known(Node n) {
    if (n.Words.Length == 1) return Known(n.Words);
    // Si l'index d'un de ses mots est plus court que ce qui reste à lire, on recompte par lui.
    int smallest = int.MaxValue;
    foreach (string w in n.Words) { List<int> l; smallest = Math.Min(smallest, index.TryGetValue(w, out l) ? l.Count : 0); }
    if (smallest < contacts.Count - n.KSeen) { n.KCount = Known(n.Words); n.KSeen = contacts.Count; return n.KCount; }
    for (; n.KSeen < contacts.Count; n.KSeen++) if (MatchToks(contacts[n.KSeen], n.Words)) n.KCount++;
    return n.KCount;
  }

  void Learn(string name, string mail) {
    List<string> toks = Tokens(name + " " + mail);
    int id = contacts.Count;
    contacts.Add(toks.ToArray());
    HashSet<string> prefixes = new HashSet<string>();
    foreach (string t in toks) {
      int prev = 26;
      for (int i = 0; i < t.Length; i++) {
        int c = t[i] - 'a';
        if (c < 0 || c > 25) { prev = -1; continue; }
        if (prev >= 0) { big[prev, c]++; bigRow[prev]++; }
        if (i == 1 && prev >= 0 && prev < 26) { start[prev, c]++; startRow[prev]++; }
        prev = c;
      }
      for (int l = 1; l <= t.Length; l++) prefixes.Add(t.Substring(0, l));
    }
    bool inScope = MatchToks(toks, fixedWords);
    foreach (string p in prefixes) {
      List<int> li;
      if (!index.TryGetValue(p, out li)) index[p] = li = new List<int>();
      li.Add(id);
      Node n;
      if (!inScope || !byWord.TryGetValue(p, out n)) continue;
      if (Cap >= 3 && !n.Answered && n.Parent != null && Known(n) == Cap) promote.Add(n);
      // Une branche qui rapporte rend ses suites plus prometteuses.
      if (n.Kids != null) bump.Add(n);
    }
  }

  // Probabilité qu'un mot qui commence par lw[..-1] continue par sa dernière lettre.
  double Continuation(string lw) {
    int c = lw[lw.Length - 1] - 'a';
    if (c < 0 || c > 25) return 0.03;
    int p = 26;
    if (lw.Length >= 2) { p = lw[lw.Length - 2] - 'a'; if (p < 0 || p > 25) return 0.03; }
    double g = (big[p, c] + 0.2) / (bigRow[p] + 5.2);
    if (lw.Length != 2) return g;
    // Deuxième lettre d'un mot : soit la suite d'un prénom ou d'un nom, soit
    // une initiale collée au nom (gpoirel, tmaitre).
    double s = (start[p, c] + 0.2) / (startRow[p] + 5.2);
    double initial = (big[26, c] + 0.2) / (bigRow[26] + 5.2);
    return 0.45 * s + 0.35 * initial + 0.2 * g;
  }

  bool Full(Node s) { return s.Virtual || (s.Answered && s.Count >= 3 && s.Count >= Cap); }

  double TotalEst(Node s) {
    if (!Full(s)) return s.Count;
    double floor = s.Root ? Cap * 6.0 : Cap * 2.0;
    return Math.Max(floor, Math.Max(Known(s) * 2.0, s.Est));
  }

  double Calibration(int kind) { return (obtained[kind] + 3.0) / (predicted[kind] + 3.0); }
  // Rendement réel en adresses nouvelles, sur rendement prévu, par genre de
  // piste. Une liste pleine de contacts déjà connus est longue mais n'apporte
  // rien : c'est ce facteur qui l'apprend.
  double Yield(int kind) { return (obtNew[kind] + 2.0) / (predNew[kind] + 2.0); }

  // Adresses inconnues que la liste de n devrait montrer, pondérées par ce que
  // la branche a rapporté jusqu'ici. Une recherche dont on connaît déjà de quoi
  // remplir la liste passe en tête : on ne la tapera pas, on ouvrira ses suites.
  double Score(Node n) {
    Node p = n.Parent;
    if (!Full(p)) return 0;
    int k = Known(n);
    if (Cap >= 3 && k >= Cap && n.Depth < MaxDepth) return double.MaxValue;
    string lw = n.Words[n.Words.Length - 1];
    double share = (k + 3.0 * n.Exact * Continuation(lw)) / (Known(p) + 3.0);
    n.Kind = n.Split ? 8 + (k == 0 ? 1 : 0) : Math.Min(n.Depth, 4) - 1 + (k == 0 ? 4 : 0);
    n.EstRaw = TotalEst(p) * Math.Min(1.0, share);
    n.Est = n.EstRaw * Calibration(n.Kind);
    double expNew = Math.Min(Cap, Math.Max(n.Est, k)) - k;
    if (expNew <= 0) return 0;
    n.ExpNew = expNew;
    return expNew * (0.3 + p.Novelty) * Yield(n.Kind);
  }

  // Les suites d'une recherche pleine : une lettre de plus au dernier mot, et,
  // pour ceux qui portent exactement ce mot (martin, wang, jean), qu'aucune
  // lettre de plus ne départage, un deuxième mot.
  void Expand(Node p) {
    if (p.Depth >= MaxDepth) return;
    string lw = p.Words[p.Words.Length - 1];
    for (char c = 'a'; c <= 'z'; c++) Push(p.Q + c, p, false, 1.0);
    double exact = ExactShare(p.Words, lw);
    if (splitPhase && exact > 0) for (char c = 'a'; c <= 'z'; c++) SplitWord(p, c.ToString(), exact);
  }

  // Part des contacts connus de la recherche qui n'ont que ce mot exact.
  double ExactShare(string[] words, string lw) {
    List<int> l;
    if (lw.Length < 2 || !index.TryGetValue(lw, out l)) return 0;
    int all = 0, ex = 0;
    foreach (int i in l) {
      if (!MatchToks(contacts[i], words)) continue;
      all++;
      bool longer = false, same = false;
      foreach (string t in contacts[i]) {
        if (t == lw) same = true;
        else if (t.StartsWith(lw, StringComparison.Ordinal)) longer = true;
      }
      if (same && !longer) ex++;
    }
    return all == 0 ? 0 : (double)ex / all;
  }

  // Un nouveau mot qui commence un mot déjà dans la recherche ne trie rien
  // (« martin ma » répond comme « martin ») : on l'allonge jusqu'à ce qu'il
  // s'en distingue. Un mot qu'un autre commence déjà est redondant.
  void SplitWord(Node p, string s, double exact) {
    foreach (string w in p.Words) {
      if (w == s || s.StartsWith(w, StringComparison.Ordinal)) return;
      if (w.StartsWith(s, StringComparison.Ordinal)) {
        for (char c = 'a'; c <= 'z'; c++) SplitWord(p, s + c, exact);
        return;
      }
    }
    Push(p.Q + " " + s, p, true, exact);
  }

  void Push(string q, Node parent, bool split, double exact) {
    if (nodes.ContainsKey(q)) return;
    Node n = Make(q, parent, split);
    n.Exact = exact;
    if (parent.Kids == null) parent.Kids = new List<Node>();
    parent.Kids.Add(n);
    n.Stale = Score(n);
    if (n.Stale < Cold) cold.Add(n); else heap.Push(n.Stale, n);
  }

  // Deuxième temps : les lettres ne rapportent plus. Les noms trop courants
  // pour être départagés par une lettre de plus le seront par un deuxième mot.
  void OpenSplits() {
    splitPhase = true;
    foreach (Node p in new List<Node>(nodes.Values)) {
      if (!Full(p) || p.Depth >= MaxDepth) continue;
      string lw = p.Words[p.Words.Length - 1];
      double exact = ExactShare(p.Words, lw);
      if (exact > 0) for (char c = 'a'; c <= 'z'; c++) SplitWord(p, c.ToString(), exact);
    }
    lastRebuild = Done;
  }

  // Les scores ne font en général que baisser (on connaît de plus en plus de
  // contacts), d'où le tas paresseux. Mais une recherche peut aussi passer le
  // seuil du plafond de connus, ou profiter d'un recalibrage : on recalcule
  // tout de temps en temps.
  void Rebuild(bool deep) {
    Heap h = new Heap();
    List<Node> c2 = deep ? new List<Node>() : cold;
    int pending = 0;
    IEnumerable<Node> all = heap.Items();
    if (deep) { List<Node> l = new List<Node>(all); l.AddRange(cold); all = l; }
    HashSet<Node> once = new HashSet<Node>();
    foreach (Node n in all) {
      if (!once.Add(n)) continue;
      // Parent qui a tout montré, ou piste déjà traitée : elle ne servira plus.
      if (n.Answered || (n.Parent.Answered && !n.Parent.Virtual && !Full(n.Parent) && n.Parent.Count < 3)) continue;
      n.Stale = Score(n);
      if (n.Stale >= MinScore) pending++;
      if (n.Stale < Cold) c2.Add(n); else h.Push(n.Stale, n);
    }
    heap = h; cold = c2; if (deep) lastDeep = Done; Pending = pending; lastRebuild = Done;
  }

  public string Next() {
    if (Done >= Budget) { StopReason = "budget"; return null; }
    if (roots.Count > 0) return roots.Dequeue().Q;
    Saturation = Math.Max(60, Done / 4);
    if (SinceNew >= Saturation) { StopReason = "saturation"; return null; }
    if (Done - lastRebuild >= 500) Rebuild(Done - lastDeep >= 3000);
    foreach (Node pn in promote) if (!pn.Answered) heap.Push(double.MaxValue, pn);
    promote.Clear();
    foreach (Node b in bump)
      foreach (Node kid in b.Kids) {
        if (kid.Answered) continue;
        double sc = Score(kid);
        if (sc > kid.Stale + 1e-9) { kid.Stale = sc; heap.Push(sc, kid); }
      }
    bump.Clear();
    bool rebuilt = false;
    while (true) {
      if (heap.Count == 0 || heap.TopKey < MinScore) {
        // Avant de conclure, on s'assure que ce n'est pas un score périmé.
        // Un recalcul récent compte : sans cela, chaque fin de file en relancerait un.
        if (rebuilt || Done - lastRebuild < 50) {
          if (!splitPhase) { OpenSplits(); rebuilt = false; continue; }
          Pending = 0; StopReason = "epuise"; return null;
        }
        Rebuild(Done - lastDeep >= 500); rebuilt = true; continue;
      }
      Node n = (Node)heap.Pop();
      if (n.Answered) continue;
      double s = Score(n);
      if (heap.Count > 0 && s < heap.TopKey - 1e-9) { n.Stale = s; heap.Push(s, n); continue; }
      if (s < MinScore) { n.Stale = s; heap.Push(s, n); continue; }
      // Une recherche dont on connaît déjà de quoi remplir la liste ne
      // montrerait que du connu : on ne la tape pas, on passe à ses suites.
      if (s == double.MaxValue) {
        n.Virtual = true; n.Answered = true; n.Novelty = n.Parent.Novelty;
        Expand(n);
        continue;
      }
      if (Pending > 0) Pending--;
      return n.Q;
    }
  }

  // Remet en file une recherche proposée mais pas tapée.
  public void Requeue(string q) {
    Node n;
    if (nodes.TryGetValue(q, out n) && !n.Answered && n.Parent != null) heap.Push(n.Stale, n);
  }

  // Ce qu'Outlook a montré pour q. keep[i] dit si l'adresse passe les filtres.
  // Rend le nombre d'adresses nouvelles et gardées.
  public int Report(string q, string[] names, string[] mailList, bool[] keep) {
    Done++;
    if (mailList.Length > Cap) Cap = mailList.Length;
    int fresh = 0, newAll = 0;
    for (int i = 0; i < mailList.Length; i++) {
      string m = mailList[i].ToLowerInvariant();
      if (mails.Add(m)) { Learn(names[i], m); newAll++; if (keep[i]) fresh++; }
    }
    SinceNew = fresh > 0 ? 0 : SinceNew + 1;
    Node n;
    if (!nodes.TryGetValue(q, out n)) { n = Make(q, null, false); n.Root = true; }
    if (n.Parent != null) {
      predicted[n.Kind] += Math.Min(Cap, n.EstRaw); obtained[n.Kind] += mailList.Length;
      predNew[n.Kind] += n.ExpNew; obtNew[n.Kind] += newAll;
    }
    n.Answered = true; n.Count = mailList.Length;
    n.Novelty = (fresh + 0.5) / (mailList.Length + 1.0);
    if (mailList.Length >= 3) Expand(n);
    return fresh;
  }

  // Tas binaire, plus grande clé en tête.
  class Heap {
    List<double> k = new List<double>(); List<object> v = new List<object>();
    public int Count { get { return k.Count; } }
    public double TopKey { get { return k[0]; } }
    public IEnumerable<Node> Items() { foreach (object o in v) yield return (Node)o; }
    public void Push(double key, object val) {
      k.Add(key); v.Add(val);
      int i = k.Count - 1;
      while (i > 0) { int p = (i - 1) / 2; if (k[p] >= k[i]) break; Swap(i, p); i = p; }
    }
    public object Pop() {
      object top = v[0]; int last = k.Count - 1;
      Swap(0, last); k.RemoveAt(last); v.RemoveAt(last);
      int i = 0;
      while (true) {
        int l = 2 * i + 1, r = l + 1, m = i;
        if (l < k.Count && k[l] > k[m]) m = l;
        if (r < k.Count && k[r] > k[m]) m = r;
        if (m == i) break;
        Swap(i, m); i = m;
      }
      return top;
    }
    void Swap(int a, int b) { double t = k[a]; k[a] = k[b]; k[b] = t; object o = v[a]; v[a] = v[b]; v[b] = o; }
  }
}
'@
Add-Type -TypeDefinition $PlanifSource

$AE = [System.Windows.Automation.AutomationElement]
$Scope = [System.Windows.Automation.TreeScope]
$CT = [System.Windows.Automation.ControlType]
$ChampA = [string][char]0x00C0

# ---------------------------------------------------------------- réglages

$cfgDir = Join-Path $env:APPDATA "AetherOutils"
$cfgPath = Join-Path $cfgDir "adresses-outlook.json"
function Default-Config {
  [ordered]@{
    Source = "outlook"; GmailAccount = ""
    Prefix = ""
    Keep = ""; Exclude = ""; Dedupe = $true
    Speed = "normal"
    Format = "txt"; Folder = [Environment]::GetFolderPath("Desktop"); OpenAtEnd = $false; CopyAtEnd = $false
    # Le fichier du dernier passage, pour Exporter et Copier au prochain lancement.
    Dernier = ""
  }
}
# Journal du passage en cours : chaque recherche et ce qu'elle a montré. Il
# permet de reprendre un passage interrompu sans retaper ce qui est fait. Il
# est effacé quand un passage va au bout.
$resumePath = Join-Path $cfgDir "reprise.jsonl"
$cfg = Default-Config
if (Test-Path $cfgPath) {
  try {
    $saved = Get-Content $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in @($cfg.Keys)) { if ($null -ne $saved.$k) { $cfg[$k] = $saved.$k } }
  } catch { }
}
if (-not (Test-Path $cfg.Folder)) { $cfg.Folder = [Environment]::GetFolderPath("Desktop") }

# Secondes par recherche, pour l'estimation affichée.
$SpeedInfo = @{
  fast    = @{ Sec = 1.7; Settle = 300;  Poll = 250; Tries = 8;  EmptyTries = 3; Stable = 1; Text = "Rapide : lit la liste dès qu'elle répond à la recherche. Peut rater des résultats de l'annuaire, qui arrivent en second." }
  normal  = @{ Sec = 2.8; Settle = 450;  Poll = 350; Tries = 10; EmptyTries = 4; Stable = 2; Text = "Normal : attend que la liste soit la même deux fois de suite. Le bon équilibre pour la plupart des recherches." }
  careful = @{ Sec = 4.5; Settle = 1200; Poll = 450; Tries = 14; EmptyTries = 6; Stable = 3; Text = "Prudent : attend trois lectures identiques, annuaire compris. Le plus fiable, le plus lent." }
}

# --------------------------------------------------------------- interface

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Adresses Outlook" Width="600" SizeToContent="Height"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13" Foreground="#0F172A">
  <Window.Resources>
    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Foreground" Value="White"/><Setter Property="Background" Value="#1E4D8C"/>
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="FontSize" Value="14"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="b" Background="{TemplateBinding Background}" CornerRadius="10" Padding="20,11">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#173E73"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Secondary" TargetType="Button">
      <Setter Property="Foreground" Value="#334155"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="b" Background="White" BorderBrush="#DCE3EC" BorderThickness="1" CornerRadius="10" Padding="14,9">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#F3F6FB"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Chrome" TargetType="Button">
      <Setter Property="Foreground" Value="#5B6A80"/><Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="10"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="b" Background="Transparent" CornerRadius="8" Width="32" Height="32">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#EEF2F8"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Segment" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#5B6A80"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="b" Background="Transparent" CornerRadius="9" Padding="16,7">
            <ContentPresenter HorizontalAlignment="Center"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="White"/><Setter Property="Foreground" Value="#1E4D8C"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Pill" TargetType="RadioButton">
      <Setter Property="Foreground" Value="#334155"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="b" Background="White" BorderBrush="#DCE3EC" BorderThickness="1" CornerRadius="9" Padding="13,7">
            <ContentPresenter HorizontalAlignment="Center"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#9FB6D6"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="#1E4D8C"/><Setter TargetName="b" Property="BorderBrush" Value="#1E4D8C"/>
              <Setter Property="Foreground" Value="White"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Card" TargetType="RadioButton">
      <Setter Property="Cursor" Value="Hand"/><Setter Property="Margin" Value="0,0,8,8"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="b" Background="White" BorderBrush="#DCE3EC" BorderThickness="1" CornerRadius="11" Padding="12,10">
            <ContentPresenter/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#9FB6D6"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="#F1F6FD"/><Setter TargetName="b" Property="BorderBrush" Value="#1E4D8C"/>
              <Setter TargetName="b" Property="BorderThickness" Value="2"/><Setter TargetName="b" Property="Padding" Value="11,9"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Switch" TargetType="CheckBox">
      <Setter Property="Cursor" Value="Hand"/><Setter Property="Margin" Value="0,0,0,10"/><Setter Property="Foreground" Value="#334155"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Grid Background="Transparent">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Border x:Name="track" Width="38" Height="22" CornerRadius="11" Background="#CBD5E1" VerticalAlignment="Center">
              <Ellipse x:Name="thumb" Width="16" Height="16" Fill="White" HorizontalAlignment="Left" Margin="3,0"/></Border>
            <ContentPresenter Grid.Column="1" Margin="10,0,0,0" VerticalAlignment="Center"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="track" Property="Background" Value="#1E4D8C"/><Setter TargetName="thumb" Property="HorizontalAlignment" Value="Right"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="Input" TargetType="TextBox">
      <Setter Property="Padding" Value="9,7"/><Setter Property="FontSize" Value="13"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border x:Name="b" Background="White" BorderBrush="#DCE3EC" BorderThickness="1" CornerRadius="9">
            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/></Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#1E4D8C"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="Bold"/><Setter Property="FontSize" Value="14"/><Setter Property="Margin" Value="0,0,0,10"/>
    </Style>
    <Style x:Key="Label" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Margin" Value="0,4,0,6"/><Setter Property="Foreground" Value="#334155"/>
    </Style>
    <Style x:Key="Help" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#5B6A80"/><Setter Property="FontSize" Value="12"/><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="Margin" Value="0,5,0,12"/>
    </Style>
  </Window.Resources>

  <Border Margin="18" CornerRadius="18" Background="#F6F8FC" BorderBrush="#E3E9F2" BorderThickness="1">
    <Border.Effect><DropShadowEffect BlurRadius="28" ShadowDepth="6" Opacity="0.18" Color="#0F172A"/></Border.Effect>
    <StackPanel>

      <Grid x:Name="Header" Background="Transparent" Margin="22,18,14,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <Border Width="44" Height="44" CornerRadius="12" Background="#1E4D8C" VerticalAlignment="Top">
          <TextBlock Text="@" Foreground="White" FontSize="22" FontWeight="Bold" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="0,-3,0,0"/></Border>
        <StackPanel Grid.Column="1" Margin="14,1,0,0">
          <TextBlock x:Name="AppTitle" Text="Adresses Outlook" FontSize="19" FontWeight="Bold"/>
          <TextBlock x:Name="Subtitle" Text="Les adresses que propose votre messagerie, recherche par recherche" Foreground="#5B6A80" Margin="0,2,0,0"/>
        </StackPanel>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Top">
          <Button x:Name="BtnMin" Style="{StaticResource Chrome}" Content="&#xE921;" ToolTip="Réduire"/>
          <Button x:Name="BtnClose" Style="{StaticResource Chrome}" Content="&#xE8BB;" ToolTip="Fermer"/>
        </StackPanel>
      </Grid>

      <StackPanel x:Name="Full" Margin="22,16,22,22">

        <Border Background="#E9EEF6" CornerRadius="11" Padding="3" HorizontalAlignment="Left">
          <StackPanel Orientation="Horizontal">
            <RadioButton x:Name="TabRun" GroupName="tabs" Style="{StaticResource Segment}" Content="Recherche" IsChecked="True"/>
            <RadioButton x:Name="TabSet" GroupName="tabs" Style="{StaticResource Segment}" Content="Paramètres"/>
          </StackPanel>
        </Border>

        <!-- ============================ onglet Recherche -->
        <StackPanel x:Name="PageRun" Margin="0,14,0,0">
          <Grid Margin="0,0,0,12">
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <TextBlock Text="Chercher dans" FontWeight="SemiBold" Foreground="#334155" VerticalAlignment="Center" Margin="2,0,12,0"/>
            <WrapPanel Grid.Column="1" VerticalAlignment="Center">
              <RadioButton x:Name="SrcOutlook" GroupName="src" Style="{StaticResource Pill}" Content="Outlook"/>
              <RadioButton x:Name="SrcGmail" GroupName="src" Style="{StaticResource Pill}" Content="Gmail"/>
            </WrapPanel>
          </Grid>
          <Border Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16,12">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <StackPanel>
                <TextBlock Text="Ce qui va se passer" FontSize="12" Foreground="#5B6A80"/>
                <TextBlock x:Name="Plan" Text="" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,2,0,0"/>
              </StackPanel>
              <Button x:Name="BtnEdit" Grid.Column="1" Style="{StaticResource Secondary}" Content="Modifier" VerticalAlignment="Center" Margin="10,0,0,0"/>
            </Grid>
          </Border>

          <Border Margin="0,12,0,0" Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16">
            <StackPanel>
              <Grid>
                <TextBlock Text="Progression" FontWeight="SemiBold"/>
                <TextBlock x:Name="Count" Text="0 recherche" Foreground="#5B6A80" HorizontalAlignment="Right"/>
              </Grid>
              <UniformGrid x:Name="Chips" Columns="13" Margin="0,12,0,0"/>
              <TextBlock x:Name="Many" Visibility="Collapsed" Margin="0,10,0,0" Foreground="#5B6A80"
                         Text="Trop de recherches pour les afficher une par une : la barre suit l'avancée."/>
              <Grid Margin="0,14,0,0" Height="6">
                <Border Background="#E8EEF6" CornerRadius="3"/>
                <Border x:Name="Fill" Background="#1E4D8C" CornerRadius="3" HorizontalAlignment="Left" Width="0"/>
              </Grid>
            </StackPanel>
          </Border>

          <Border x:Name="ResumeBox" Visibility="Collapsed" Margin="0,12,0,0" CornerRadius="12" Background="White" BorderBrush="#9FB6D6" BorderThickness="1" Padding="14,10">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Passage interrompu" FontWeight="SemiBold"/>
                <TextBlock x:Name="ResumeText" TextWrapping="Wrap" Foreground="#5B6A80" FontSize="12" Margin="0,2,0,0"/>
              </StackPanel>
              <Button x:Name="BtnFresh" Grid.Column="1" Style="{StaticResource Secondary}" Content="Repartir de zéro" Margin="10,0,0,0" VerticalAlignment="Center"/>
            </Grid>
          </Border>

          <Grid Margin="0,12,0,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="10"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="BtnRun" Style="{StaticResource Primary}" Content="Lancer le passage"/>
            <Button x:Name="BtnExport" Grid.Column="2" Style="{StaticResource Secondary}" Content="Exporter" IsEnabled="False" ToolTip="Enregistrer les adresses dans un fichier texte, une par ligne"/>
            <Button x:Name="BtnCopy" Grid.Column="4" Style="{StaticResource Secondary}" Content="Copier" IsEnabled="False"/>
          </Grid>

          <Border x:Name="Notice" Margin="0,12,0,0" CornerRadius="10" Background="#EEF4FC" Padding="12,10">
            <TextBlock x:Name="NoticeText" TextWrapping="Wrap" Foreground="#1E4D8C"
                       Text="Gardez Outlook ouvert. Pendant le passage, ne touchez ni au clavier ni à la souris. Rien n'est envoyé."/>
          </Border>

          <Border Margin="0,12,0,0" Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1">
            <StackPanel>
              <Grid Margin="16,12,16,8">
                <TextBlock Text="Résultats" FontWeight="SemiBold"/>
                <TextBlock x:Name="Distinct" Text="" Foreground="#5B6A80" HorizontalAlignment="Right"/>
              </Grid>
              <ScrollViewer Height="230" VerticalScrollBarVisibility="Auto" Margin="0,0,0,10">
                <StackPanel x:Name="Rows" Margin="16,0,16,0">
                  <TextBlock x:Name="Empty" Text="Les adresses apparaîtront ici." Foreground="#94A3B8" Margin="0,18,0,18" HorizontalAlignment="Center"/>
                </StackPanel>
              </ScrollViewer>
            </StackPanel>
          </Border>
        </StackPanel>

        <!-- ============================ onglet Paramètres -->
        <ScrollViewer x:Name="PageSet" Visibility="Collapsed" Height="600" Margin="0,14,0,0" VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="0,0,8,0">

            <Border Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16">
              <StackPanel>
                <TextBlock Style="{StaticResource SectionTitle}" Text="1. Quoi chercher"/>
                <Border Background="#F1F6FD" CornerRadius="10" Padding="12,10" Margin="0,0,0,12">
                  <StackPanel>
                    <TextBlock TextWrapping="Wrap" Foreground="#1E4D8C" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,4"
                               Text="Le passage va jusqu'au bout, sans limite de recherches."/>
                    <TextBlock TextWrapping="Wrap" Foreground="#1E4D8C" FontSize="12"
                               Text="Il commence par les 26 lettres. Quand la liste de suggestions est pleine, des adresses restent cachées dessous : la recherche devient une piste à creuser (a devient al, am, an...). Les pistes passent par ordre de rendement, sur tout l'alphabet, et un nom trop courant pour être départagé par une lettre l'est par un deuxième mot (martin j, martin s...). Il s'arrête seul quand il n'y a plus rien à trouver. Chaque adresse est enregistrée dès qu'elle est trouvée : vous pouvez arrêter à tout moment, et reprendre plus tard là où vous en étiez."/>
                  </StackPanel>
                </Border>

                <StackPanel x:Name="PrefixBox">
                  <TextBlock Style="{StaticResource Label}" Text="Préfixe (facultatif)"/>
                  <TextBox x:Name="Prefix" Style="{StaticResource Input}"/>
                  <TextBlock Style="{StaticResource Help}" Text="Placé devant chaque recherche. Avec « ma », la messagerie reçoit maa, mab, mac... puis approfondit à partir de là. Utile pour creuser un nom ou une équipe."/>
                </StackPanel>

                <StackPanel x:Name="GmailBox">
                  <TextBlock Style="{StaticResource Label}" Text="Compte Gmail (facultatif)"/>
                  <TextBox x:Name="GmailAccount" Style="{StaticResource Input}"/>
                  <TextBlock Style="{StaticResource Help}" Text="Seulement si plusieurs comptes Google sont connectés dans votre navigateur : l'adresse du compte à utiliser. Vide : le compte principal. Gmail s'ouvre dans votre navigateur par défaut."/>
                </StackPanel>
              </StackPanel>
            </Border>

            <Border Margin="0,12,0,0" Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16">
              <StackPanel>
                <TextBlock Style="{StaticResource SectionTitle}" Text="2. Quoi garder"/>
                <TextBlock Style="{StaticResource Label}" Text="Garder seulement ces domaines (facultatif)" Margin="0,0,0,6"/>
                <TextBox x:Name="Keep" Style="{StaticResource Input}"/>
                <TextBlock Style="{StaticResource Help}" Text="Séparés par des virgules, par exemple : test.com, deloitte.fr. Les sous-domaines comptent (edu.test.com est gardé avec test.com)."/>

                <TextBlock Style="{StaticResource Label}" Text="Exclure (facultatif)"/>
                <TextBox x:Name="Exclude" Style="{StaticResource Input}"/>
                <TextBlock Style="{StaticResource Help}" Text="Domaines ou adresses exactes, séparés par des virgules. Pratique pour retirer votre propre adresse ou les boîtes de service."/>

                <CheckBox x:Name="Dedupe" Style="{StaticResource Switch}" Content="Ne noter chaque adresse qu'une seule fois"/>
              </StackPanel>
            </Border>

            <Border Margin="0,12,0,0" Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16">
              <StackPanel>
                <TextBlock Style="{StaticResource SectionTitle}" Text="3. Rythme"/>
                <WrapPanel>
                  <RadioButton x:Name="SpeedFast" GroupName="speed" Style="{StaticResource Pill}" Content="Rapide"/>
                  <RadioButton x:Name="SpeedNormal" GroupName="speed" Style="{StaticResource Pill}" Content="Normal"/>
                  <RadioButton x:Name="SpeedCareful" GroupName="speed" Style="{StaticResource Pill}" Content="Prudent"/>
                </WrapPanel>
                <TextBlock x:Name="SpeedHelp" Style="{StaticResource Help}" Margin="0,0,0,0"/>
              </StackPanel>
            </Border>

            <Border Margin="0,12,0,0" Background="White" CornerRadius="14" BorderBrush="#E6EBF2" BorderThickness="1" Padding="16">
              <StackPanel>
                <TextBlock Style="{StaticResource SectionTitle}" Text="4. Enregistrement"/>
                <TextBlock Style="{StaticResource Label}" Text="Format du fichier" Margin="0,0,0,8"/>
                <WrapPanel>
                  <RadioButton x:Name="FmtTxt" GroupName="fmt" Style="{StaticResource Pill}" Content="Texte : adresses seules"/>
                  <RadioButton x:Name="FmtCsv" GroupName="fmt" Style="{StaticResource Pill}" Content="CSV pour Excel, avec les noms"/>
                  <RadioButton x:Name="FmtBoth" GroupName="fmt" Style="{StaticResource Pill}" Content="Les deux"/>
                </WrapPanel>
                <TextBlock Style="{StaticResource Label}" Text="Dossier"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                  <Border Background="#F6F8FC" CornerRadius="9" BorderBrush="#E6EBF2" BorderThickness="1" Padding="10,8">
                    <TextBlock x:Name="Folder" TextTrimming="CharacterEllipsis" Foreground="#334155"/></Border>
                  <Button x:Name="BtnFolder" Grid.Column="1" Style="{StaticResource Secondary}" Content="Choisir" Margin="8,0,0,0"/>
                </Grid>
                <StackPanel Margin="0,14,0,0">
                  <CheckBox x:Name="OpenAtEnd" Style="{StaticResource Switch}" Content="Ouvrir le fichier à la fin"/>
                  <CheckBox x:Name="CopyAtEnd" Style="{StaticResource Switch}" Content="Copier les adresses dans le presse-papiers à la fin"/>
                </StackPanel>
              </StackPanel>
            </Border>

            <Grid Margin="2,12,0,4">
              <TextBlock Text="Vos réglages sont gardés pour la prochaine fois." Foreground="#5B6A80" VerticalAlignment="Center"/>
              <Button x:Name="BtnReset" Style="{StaticResource Secondary}" Content="Réglages d'origine" HorizontalAlignment="Right"/>
            </Grid>
          </StackPanel>
        </ScrollViewer>
      </StackPanel>

      <!-- ============================ panneau compact, pendant le passage -->
      <StackPanel x:Name="Compact" Margin="22,14,22,20" Visibility="Collapsed">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <TextBlock x:Name="LiveQuery" Text="A" FontSize="28" FontWeight="Bold" Foreground="#1E4D8C" TextTrimming="CharacterEllipsis"/>
          <StackPanel Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Center">
            <TextBlock x:Name="LiveCount" Text="0 recherche" FontWeight="SemiBold" HorizontalAlignment="Right"/>
            <TextBlock x:Name="LiveEta" Text="" Foreground="#5B6A80" FontSize="12" HorizontalAlignment="Right"/>
          </StackPanel>
        </Grid>
        <TextBlock x:Name="LiveMail" Text="Ouverture d'un nouveau message" Foreground="#5B6A80" Margin="0,4,0,10" TextTrimming="CharacterEllipsis"/>
        <Grid Height="6">
          <Border Background="#E8EEF6" CornerRadius="3"/>
          <Border x:Name="LiveFill" Background="#1E4D8C" CornerRadius="3" HorizontalAlignment="Left" Width="0"/>
        </Grid>
        <Grid Margin="0,12,0,0">
          <TextBlock Text="Ne touchez à rien" Foreground="#B45309" FontSize="12" VerticalAlignment="Center"/>
          <Button x:Name="BtnStop" Style="{StaticResource Secondary}" Content="Arrêter" HorizontalAlignment="Right"/>
        </Grid>
      </StackPanel>

    </StackPanel>
  </Border>
</Window>
'@

$win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$ui = @{}
$xaml.SelectNodes("//*[@*[local-name()='Name']]") | ForEach-Object {
  $name = $_.GetAttribute("Name", "http://schemas.microsoft.com/winfx/2006/xaml")
  if ($name) { $ui[$name] = $win.FindName($name) }
}

function Brush([string]$hex) { [System.Windows.Media.BrushConverter]::new().ConvertFromString($hex) }
function Pump { $win.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background) }
function Wait([int]$ms) { $end = (Get-Date).AddMilliseconds($ms); while ((Get-Date) -lt $end) { Start-Sleep -Milliseconds 40; Pump } }

# ------------------------------------------------------ réglages <-> écran

function Apply-Config {
  $script:loading = $true
  if ($cfg.Source -eq "gmail") { $ui.SrcGmail.IsChecked = $true } else { $ui.SrcOutlook.IsChecked = $true }
  $ui.GmailAccount.Text = $cfg.GmailAccount
  $ui.Prefix.Text = $cfg.Prefix
  $ui.Keep.Text = $cfg.Keep; $ui.Exclude.Text = $cfg.Exclude; $ui.Dedupe.IsChecked = [bool]$cfg.Dedupe
  switch ($cfg.Speed) { "fast" { $ui.SpeedFast.IsChecked = $true } "careful" { $ui.SpeedCareful.IsChecked = $true } default { $ui.SpeedNormal.IsChecked = $true } }
  switch ($cfg.Format) { "csv" { $ui.FmtCsv.IsChecked = $true } "both" { $ui.FmtBoth.IsChecked = $true } default { $ui.FmtTxt.IsChecked = $true } }
  $ui.Folder.Text = $cfg.Folder
  $ui.OpenAtEnd.IsChecked = [bool]$cfg.OpenAtEnd; $ui.CopyAtEnd.IsChecked = [bool]$cfg.CopyAtEnd
  $script:loading = $false
}

function Read-Config {
  $cfg.Source = if ($ui.SrcGmail.IsChecked) { "gmail" } else { "outlook" }
  $cfg.GmailAccount = $ui.GmailAccount.Text.Trim()
  # Une virgule ou un point-virgule ferait valider un destinataire.
  $cfg.Prefix = ($ui.Prefix.Text -replace '[;,]', '').Trim()
  $cfg.Keep = $ui.Keep.Text.Trim(); $cfg.Exclude = $ui.Exclude.Text.Trim(); $cfg.Dedupe = [bool]$ui.Dedupe.IsChecked
  $cfg.Speed = if ($ui.SpeedFast.IsChecked) { "fast" } elseif ($ui.SpeedCareful.IsChecked) { "careful" } else { "normal" }
  $cfg.Format = if ($ui.FmtCsv.IsChecked) { "csv" } elseif ($ui.FmtBoth.IsChecked) { "both" } else { "txt" }
  $cfg.Folder = $ui.Folder.Text
  $cfg.OpenAtEnd = [bool]$ui.OpenAtEnd.IsChecked; $cfg.CopyAtEnd = [bool]$ui.CopyAtEnd.IsChecked
}

function Save-Config {
  try {
    if (-not (Test-Path $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir | Out-Null }
    ($cfg | ConvertTo-Json) | Set-Content -Path $cfgPath -Encoding UTF8
  } catch { }
}

function Build-Queries {
  return @([char[]]'abcdefghijklmnopqrstuvwxyz' | ForEach-Object { $cfg.Prefix + [string]$_ })
}

function Duration-Text([double]$sec) {
  if ($sec -lt 60) { return "moins d'une minute" }
  $m = [Math]::Round($sec / 60)
  if ($m -lt 60) { return "environ $m min" }
  return ("environ {0} h {1:00}" -f [Math]::Floor($m / 60), ($m % 60))
}

function Update-Summary {
  if ($script:loading) { return }
  Read-Config
  $queries = Build-Queries
  $src = Source-Name
  $ui.AppTitle.Text = "Adresses $src"
  $ui.GmailBox.Visibility = if ($cfg.Source -eq "gmail") { "Visible" } else { "Collapsed" }
  # La consigne de départ, tant qu'aucun passage n'a affiché son propre message.
  if (-not $script:hasRun) {
    $ui.NoticeText.Text = if ($cfg.Source -eq "gmail") {
      "Gmail s'ouvrira dans votre navigateur par défaut, sur un nouveau message. Pendant le passage, ne touchez ni au clavier ni à la souris. Rien n'est envoyé."
    } else { "Gardez Outlook ouvert. Pendant le passage, ne touchez ni au clavier ni à la souris. Rien n'est envoyé." }
  }
  $start = if ($cfg.Prefix) { "Dans $src, 26 recherches de départ ($($cfg.Prefix)a à $($cfg.Prefix)z)" } else { "Dans $src, 26 recherches de départ (A à Z)" }
  $fmt = switch ($cfg.Format) { "csv" { "CSV" } "both" { "texte + CSV" } default { "texte" } }
  $filters = @()
  if ($cfg.Keep) { $filters += "domaines filtrés" }
  if ($cfg.Exclude) { $filters += "exclusions" }
  $ftxt = if ($filters.Count) { " · " + ($filters -join ", ") } else { "" }
  # La durée dépend de ce que la messagerie contient : on ne la connaît pas d'avance.
  $ui.Plan.Text = "$start, puis les pistes les plus rentables jusqu'à ce qu'il n'y ait plus rien à trouver$ftxt · $fmt, enregistré au fur et à mesure · de quelques minutes à plusieurs heures selon la taille de votre carnet, arrêt possible à tout moment"
  $ui.SpeedHelp.Text = $SpeedInfo[$cfg.Speed].Text
  Build-Chips $queries
  Save-Config
}

# ------------------------------------------------------------- progression

$chips = @{}
function Build-Chips($queries) {
  $ui.Chips.Children.Clear(); $chips.Clear()
  $ui.Chips.Columns = 13
  if ($true) {
    foreach ($q in $queries) {
      $t = New-Object System.Windows.Controls.TextBlock
      $t.Text = $q.Substring($q.Length - 1).ToUpper(); $t.FontWeight = "SemiBold"; $t.FontSize = 12; $t.HorizontalAlignment = "Center"; $t.VerticalAlignment = "Center"
      $b = New-Object System.Windows.Controls.Border
      $b.Height = 28; $b.Margin = "2"; $b.CornerRadius = "7"; $b.Child = $t
      $ui.Chips.Children.Add($b) | Out-Null
      $chips[$q] = @{ Box = $b; Text = $t }
      Chip $q "wait"
    }
  }
  $ui.Count.Text = "0 recherche"; $ui.Fill.Width = 0
}

function Chip([string]$q, [string]$state) {
  if (-not $chips.ContainsKey($q)) { return }
  $c = $chips[$q]
  switch ($state) {
    "wait" { $c.Box.Background = Brush "#EEF2F8"; $c.Box.BorderThickness = "0"; $c.Text.Foreground = Brush "#7A879A" }
    "now"  { $c.Box.Background = Brush "#FFFFFF"; $c.Box.BorderBrush = Brush "#1E4D8C"; $c.Box.BorderThickness = "2"; $c.Text.Foreground = Brush "#1E4D8C" }
    "ok"   { $c.Box.Background = Brush "#1E4D8C"; $c.Box.BorderThickness = "0"; $c.Text.Foreground = Brush "#FFFFFF" }
    "none" { $c.Box.Background = Brush "#FEF3C7"; $c.Box.BorderThickness = "0"; $c.Text.Foreground = Brush "#92400E" }
  }
}

# Le nombre de recherches à venir n'est connu qu'à peu près : on affiche celles
# qui sont faites, et la barre suit l'estimation.
function Progress([int]$done, [int]$total) {
  $ui.Count.Text = Plural $done "recherche" "recherches"; $ui.LiveCount.Text = $ui.Count.Text
  if ($total -gt 0) {
    $ui.Fill.Width = [Math]::Round(($ui.Fill.Parent.ActualWidth) * $done / $total)
    $ui.LiveFill.Width = [Math]::Round(($ui.LiveFill.Parent.ActualWidth) * $done / $total)
  }
}

function Notice([string]$text, [string]$kind) {
  $ui.NoticeText.Text = $text
  switch ($kind) {
    "info" { $ui.Notice.Background = Brush "#EEF4FC"; $ui.NoticeText.Foreground = Brush "#1E4D8C" }
    "ok"   { $ui.Notice.Background = Brush "#DCFCE7"; $ui.NoticeText.Foreground = Brush "#15803D" }
    "warn" { $ui.Notice.Background = Brush "#FEF3C7"; $ui.NoticeText.Foreground = Brush "#92400E" }
  }
}

function Add-Row([string]$q, [string]$mail, [string]$name) {
  $ui.Empty.Visibility = "Collapsed"
  $g = New-Object System.Windows.Controls.Grid
  $g.Margin = "0,0,0,10"
  $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = "Auto"
  $c1 = New-Object System.Windows.Controls.ColumnDefinition
  $g.ColumnDefinitions.Add($c0); $g.ColumnDefinitions.Add($c1)
  $badge = New-Object System.Windows.Controls.Border
  $badge.MinWidth = 30; $badge.Height = 30; $badge.CornerRadius = "8"; $badge.VerticalAlignment = "Top"; $badge.Padding = "7,0"
  $bt = New-Object System.Windows.Controls.TextBlock
  $label = if ($q.Length -gt 10) { $q.Substring(0, 9) + "..." } else { $q }
  $bt.Text = $label.ToUpper(); $bt.FontWeight = "Bold"; $bt.FontSize = 12; $bt.HorizontalAlignment = "Center"; $bt.VerticalAlignment = "Center"
  $badge.Child = $bt
  $sp = New-Object System.Windows.Controls.StackPanel
  $sp.Margin = "12,0,0,0"; [System.Windows.Controls.Grid]::SetColumn($sp, 1)
  $m = New-Object System.Windows.Controls.TextBlock; $m.FontWeight = "SemiBold"; $m.TextTrimming = "CharacterEllipsis"
  $n = New-Object System.Windows.Controls.TextBlock; $n.FontSize = 12; $n.TextTrimming = "CharacterEllipsis"; $n.Foreground = Brush "#5B6A80"
  if ($mail) {
    $badge.Background = Brush "#E8F0FA"; $bt.Foreground = Brush "#1E4D8C"
    $m.Text = $mail; $m.Foreground = Brush "#0F172A"; $n.Text = $name
  } else {
    $badge.Background = Brush "#FEF3C7"; $bt.Foreground = Brush "#92400E"
    $m.Text = $name; $m.Foreground = Brush "#92400E"; $n.Text = ""
  }
  $sp.Children.Add($m) | Out-Null; if ($n.Text) { $sp.Children.Add($n) | Out-Null }
  $g.Children.Add($badge) | Out-Null; $g.Children.Add($sp) | Out-Null
  $ui.Rows.Children.Add($g) | Out-Null
  # Un long passage trouve des milliers d'adresses : l'écran garde les
  # dernières, le fichier les a toutes. (La première ligne est le texte d'attente.)
  if ($ui.Rows.Children.Count -gt 301) { $ui.Rows.Children.RemoveAt(1) }
}

function Reset-Results {
  $ui.Rows.Children.Clear(); $ui.Rows.Children.Add($ui.Empty) | Out-Null; $ui.Empty.Visibility = "Visible"
  $ui.Distinct.Text = ""
}

$normal = @{}
function Compact-Mode([bool]$on) {
  if ($on) {
    $normal.Left = $win.Left; $normal.Top = $win.Top
    $ui.Full.Visibility = "Collapsed"; $ui.Subtitle.Visibility = "Collapsed"
    $ui.BtnMin.Visibility = "Collapsed"; $ui.BtnClose.Visibility = "Collapsed"
    $ui.Compact.Visibility = "Visible"
    $win.Width = 400; $win.Topmost = $true
    Pump
    $wa = [System.Windows.SystemParameters]::WorkArea
    $win.Left = $wa.Right - $win.ActualWidth - 8; $win.Top = $wa.Bottom - $win.ActualHeight - 8
  } else {
    $ui.Compact.Visibility = "Collapsed"
    $ui.Full.Visibility = "Visible"; $ui.Subtitle.Visibility = "Visible"
    $ui.BtnMin.Visibility = "Visible"; $ui.BtnClose.Visibility = "Visible"
    $win.Width = 600; $win.Topmost = $false
    $win.Left = $normal.Left; $win.Top = $normal.Top
  }
  Pump
}

# ------------------------------------------------------------- Outlook

function Outlook-Process {
  Get-Process olk -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
}
function Au-Premier-Plan($p) {
  $procId = 0
  [Fen]::GetWindowThreadProcessId([Fen]::GetForegroundWindow(), [ref]$procId) | Out-Null
  return $procId -eq $p.Id
}
function Source-Name { if ($cfg.Source -eq "gmail") { "Gmail" } else { "Outlook" } }

# Outlook : un groupe nommé « À » (ou « To »). Gmail : le champ des
# destinataires est la liste déroulante qui a le curseur, dans le navigateur
# ouvert sur Gmail. Son nom dépend de la langue de Gmail, on ne s'y fie pas.
# Dès que Gmail montre ses suggestions, Windows annonce le curseur sur la
# suggestion surlignée, alors que les frappes vont toujours dans « À » : une
# suggestion du même navigateur compte donc aussi.
$script:targetPid = 0
function Focus-Dans-Champ-A {
  $f = try { $AE::FocusedElement } catch { $null }
  if (-not $f) { return $false }
  if ($cfg.Source -eq "gmail") {
    $type = $f.Current.ControlType
    return (($type -eq $CT::ComboBox -or $type -eq $CT::ListItem) -and $f.Current.ProcessId -eq $script:targetPid)
  }
  return ($f.Current.ControlType -eq $CT::Group -and ($f.Current.Name -eq $ChampA -or $f.Current.Name -eq "To"))
}

# Gmail nomme chaque suggestion « Nom adresse » (ou « adresse adresse »). Il
# ajoute un élément qui colle toutes les suggestions bout à bout, sans
# l'espace avant l'adresse : « Jean Dupontjean.dupont@exemple.frMarie ... ». On
# l'écarte. Avec une seule suggestion il ne contient qu'une adresse, donc on le
# reconnaît à ce qu'il est : la suite exacte des autres, collées.
function Parse-Gmail([string[]]$names) {
  $cands = @()
  foreach ($n in $names) {
    # Contact sans nom : Gmail répète l'adresse.
    if ($n -match '^(\S+@\S+)\s(\S+@\S+)$' -and $matches[1] -eq $matches[2]) {
      $cands += [pscustomobject]@{ Raw = $n; Glued = $matches[1] + $matches[2]; Name = $matches[2].ToLower(); Mail = $matches[2].ToLower() }
      continue
    }
    if (([regex]::Matches($n, '@')).Count -ne 1) { continue }
    if ($n -match '^(.*)\s(\S+@\S+)$') {
      $nom = $matches[1].Trim(); $mail = $matches[2]
      $cands += [pscustomobject]@{ Raw = $n; Glued = $matches[1] + $mail; Name = $(if ($nom) { $nom } else { $mail.ToLower() }); Mail = $mail.ToLower() }
    } elseif ($n -match '^\S+@\S+$') { $cands += [pscustomobject]@{ Raw = $n; Glued = $n; Name = $n; Mail = $n.ToLower() } }
  }
  $out = @()
  for ($i = 0; $i -lt $cands.Count; $i++) {
    $others = (@($cands | Select-Object -Index (@(0..($cands.Count - 1)) -ne $i)) | ForEach-Object { $_.Glued }) -join ""
    if ($cands.Count -gt 1 -and $cands[$i].Raw -eq $others) { continue }
    $out += [pscustomobject]@{ Name = $cands[$i].Name; Mail = $cands[$i].Mail }
  }
  return $out
}

# Toutes les suggestions affichées, dans l'ordre : @{ Name; Mail }.
function Lire-Suggestions($p) {
  if ($p.MainWindowHandle -eq [IntPtr]::Zero) { return @() }
  $racine = try { $AE::FromHandle($p.MainWindowHandle) } catch { $null }
  if (-not $racine) { return @() }
  if ($cfg.Source -eq "gmail") {
    $items = $racine.FindAll($Scope::Descendants,
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::ListItem)))
    return @(Parse-Gmail ([string[]]@($items | ForEach-Object { $_.Current.Name })))
  }
  $liste = $racine.FindFirst($Scope::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, "FloatingSuggestionsList")))
  if (-not $liste) { return @() }
  $items = $liste.FindAll($Scope::Descendants,
    (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $CT::ListItem)))
  $out = @()
  foreach ($it in $items) {
    $n = $it.Current.Name
    if ($n -match '^(.*?)\s+-\s+(\S+@\S+)$') { $out += [pscustomobject]@{ Name = $matches[1]; Mail = $matches[2].ToLower() } }
    elseif ($n -match '(\S+@\S+)') { $out += [pscustomobject]@{ Name = $n; Mail = $matches[1].ToLower() } }
  }
  return $out
}

# SendKeys donne un sens à + ^ % ~ ( ) { } [ ] : on les protège.
function Escape-Keys([string]$s) {
  ($s.ToCharArray() | ForEach-Object { if ('+^%~(){}[]'.Contains([string]$_)) { "{$_}" } else { [string]$_ } }) -join ""
}

function Domain-Match([string]$mail, [string]$entry) {
  $d = $entry.Trim().ToLower().TrimStart("@")
  if (-not $d) { return $false }
  $domaine = ($mail -split "@")[-1]
  return ($domaine -eq $d -or $domaine.EndsWith("." + $d))
}

function Passes-Filters([string]$mail) {
  if ($cfg.Keep) {
    $ok = $false
    foreach ($e in ($cfg.Keep -split ",")) { if (Domain-Match $mail $e) { $ok = $true; break } }
    if (-not $ok) { return $false }
  }
  if ($cfg.Exclude) {
    foreach ($e in ($cfg.Exclude -split ",")) {
      $x = $e.Trim().ToLower()
      if (-not $x) { continue }
      if ($x.Contains("@") -and -not $x.StartsWith("@")) { if ($mail -eq $x) { return $false } }
      elseif (Domain-Match $mail $x) { return $false }
    }
  }
  return $true
}

# Outlook au premier plan. Une frappe sur Alt autorise ce changement de fenêtre.
function Mettre-Devant($p) {
  [Fen]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero); [Fen]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
  [Fen]::ShowWindow($p.MainWindowHandle, 3) | Out-Null
  [Fen]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
}
# Ouvre un nouveau message Gmail par son adresse directe (le raccourci « c »
# de Gmail est souvent désactivé), dans le navigateur par défaut, et attend
# qu'il soit devant avec le curseur dans « À ». Rend la fenêtre, ou un texte
# qui dit ce qui manque.
function Ouvrir-Gmail {
  $acct = $cfg.GmailAccount
  if ($acct -notmatch '^(\d+|[^@\s/?#]+@[^@\s/?#]+)$') { $acct = "0" }
  [Fen]::keybd_event(0x12, 0, 0, [UIntPtr]::Zero); [Fen]::keybd_event(0x12, 0, 2, [UIntPtr]::Zero)
  Start-Process ("https://mail.google.com/mail/u/" + [Uri]::EscapeDataString($acct) + "/?view=cm&fs=1&tf=1")
  # On reconnaît le nouveau message à son champ « À » qui a le curseur, pas au
  # titre de la fenêtre : celui d'un compte d'école ou d'entreprise ne dit pas
  # « Gmail » (« Compose Mail - ... University Mail »).
  $vu = $false
  for ($i = 0; $i -lt 40; $i++) {
    Wait 500
    $h = [Fen]::GetForegroundWindow()
    # Pendant un changement de fenêtre, Windows n'en annonce parfois aucune.
    if ($h -eq [IntPtr]::Zero) { continue }
    $procId = 0; [Fen]::GetWindowThreadProcessId($h, [ref]$procId) | Out-Null
    $titre = (Get-Process -Id $procId -ErrorAction SilentlyContinue).MainWindowTitle
    if ($titre -match 'Mail|Gmail|Compose|Nouveau message') { $vu = $true }
    [Fen]::ReveillerPage($h)
    $f = try { $AE::FocusedElement } catch { $null }
    if ($f -and $f.Current.ProcessId -eq $procId -and $f.Current.ControlType -eq $CT::ComboBox) {
      $script:targetPid = $procId
      return [pscustomobject]@{ Id = $procId; MainWindowHandle = $h }
    }
  }
  if ($vu) { return "Gmail s'est ouvert, mais pas sur un nouveau message avec le curseur dans « À ». Vérifiez que vous êtes connecté à Gmail dans votre navigateur. Rien n'a été tapé." }
  return "Le navigateur n'a pas affiché Gmail. Ouvrez Gmail une fois dans votre navigateur par défaut, connectez-vous, puis relancez. Rien n'a été tapé."
}

# Toutes les frappes passent par ici, ce qui permet de tester la boucle avec un
# faux Outlook.
function Taper([string]$keys) { [System.Windows.Forms.SendKeys]::SendWait($keys) }

# Lit la réponse d'Outlook à q. Une liste n'est retenue que si elle répond
# bien à q : au moins un contact dont un mot commence par la recherche, ou à
# défaut une liste différente de celle de la recherche d'avant. Sinon c'est
# encore l'ancienne liste qui est affichée, et on attend. Selon la vitesse, on
# attend aussi qu'elle ne bouge plus d'une lecture à l'autre : les résultats
# de l'annuaire arrivent après les contacts récents.
$script:prevSig = ""
function Lire-Reponse($p, [string]$q, $sp) {
  $last = $null; $stable = 0; $empty = 0
  for ($t = 0; $t -lt $sp.Tries; $t++) {
    $list = @(Lire-Suggestions $p)
    if ($list.Count -eq 0) {
      $empty++; $stable = 0; $last = $null
      if ($empty -ge $sp.EmptyTries) { $script:prevSig = ""; return }
    } else {
      $empty = 0
      $sig = ($list | ForEach-Object { $_.Mail }) -join "|"
      $answers = @($list | Where-Object { [Planif]::Matches($_.Name + " " + $_.Mail, $q) }).Count -gt 0
      if ($answers -or $sig -ne $script:prevSig) {
        if ($sig -eq $last) { $stable++ } else { $stable = 1; $last = $sig }
        if ($stable -ge $sp.Stable) { $script:prevSig = $sig; return $list }
      }
    }
    Wait $sp.Poll
  }
  # Jamais stable dans le temps imparti : on garde la dernière liste qui
  # répondait à q, faute de mieux.
  if ($last) { $script:prevSig = $last; return @(Lire-Suggestions $p) }
}

function Plural([int]$n, [string]$one, [string]$many) { if ($n -le 1) { "$n $one" } else { "$n $many" } }
function Csv-Field([string]$v) { if ($v -match '[;"\r\n]') { '"' + $v.Replace('"', '""') + '"' } else { $v } }

$script:files = @()
$script:distinctes = @()
$script:stop = $false

# ------------------------------------------------------ fichiers et reprise
#
# Les fichiers de résultats sont écrits au fil du passage : une adresse trouvée
# est sur le disque aussitôt, même si l'ordinateur s'éteint. Le journal garde
# chaque recherche et ce qu'elle a montré, une ligne par recherche :
#   recherche <TAB> nom <US> adresse <US> gardée <RS> nom <US> ...
# (US et RS sont les caractères de contrôle 31 et 30, absents des noms.) Sa
# première ligne décrit le passage : réglages et nom des fichiers.

$US = [string][char]31; $RS = [string][char]30

function Csv-Line($r) { (Csv-Field $r.Query) + ";" + $r.Rank + ";" + (Csv-Field $r.Name) + ";" + (Csv-Field $r.Mail) }

# Le CSV garde un BOM, sans lequel Excel lit mal les accents. Le texte n'en a
# pas besoin : il ne contient que des adresses.
function New-Writer([string]$path, [bool]$bom = $true) {
  $w = New-Object System.IO.StreamWriter($path, $false, (New-Object System.Text.UTF8Encoding $bom))
  $w.AutoFlush = $true
  $w.NewLine = "`n"
  return $w
}

# Le texte est le format le plus léger : une adresse par ligne, chacune une
# seule fois, sans nom ni recherche. Le CSV garde le détail pour qui le veut.
function Write-Txt([string]$mail) {
  if ($script:txtW -and $script:txtSeen.Add($mail)) { $script:txtW.WriteLine($mail) }
}

# Ouvre les fichiers de résultats en y réécrivant ce qui est déjà trouvé (une
# reprise repart ainsi d'un fichier propre), puis les garde ouverts.
function Open-Outputs([string]$base, $records) {
  $script:files = @(); $script:txtW = $null; $script:csvW = $null
  $script:txtSeen = New-Object System.Collections.Generic.HashSet[string]
  if ($cfg.Format -in "txt", "both") {
    $script:txtW = New-Writer "$base.txt" $false
    foreach ($r in $records) { Write-Txt $r.Mail }
    $script:files += "$base.txt"
  }
  if ($cfg.Format -in "csv", "both") {
    $script:csvW = New-Writer "$base.csv"
    $script:csvW.WriteLine("Recherche;Rang;Nom;Adresse")
    foreach ($r in $records) { $script:csvW.WriteLine((Csv-Line $r)) }
    $script:files += "$base.csv"
  }
}

function Write-Record($r) {
  Write-Txt $r.Mail
  if ($script:csvW) { $script:csvW.WriteLine((Csv-Line $r)) }
}

function Close-Outputs {
  if ($script:txtW) { $script:txtW.Close(); $script:txtW = $null }
  if ($script:csvW) { $script:csvW.Close(); $script:csvW = $null }
}

# Les adresses déjà trouvées, pour Exporter et Copier dès l'ouverture : celles
# du passage interrompu, sinon celles du dernier passage.
function Load-Known {
  $j = Read-Journal
  if ($j -and $j.Entries.Count -gt 0) { return @($j.KeptList | Sort-Object -Unique) }
  $f = [string]$cfg.Dernier
  if (-not $f -or -not (Test-Path $f)) { return @() }
  try {
    $lines = [System.IO.File]::ReadAllLines($f, [System.Text.Encoding]::UTF8)
    $mails = if ($f.EndsWith(".csv")) {
      $lines | Select-Object -Skip 1 | ForEach-Object { ($_ -split ";")[-1].Trim('"') }
    } else { $lines }
    return @($mails | Where-Object { $_ -match '^\S+@\S+$' } | ForEach-Object { $_.ToLower() } | Sort-Object -Unique)
  } catch { return @() }
}

# Exporter et Copier servent dès qu'il y a des adresses, même avant tout passage.
function Refresh-Known {
  $script:distinctes = @(Load-Known)
  $n = $script:distinctes.Count
  $ui.BtnExport.IsEnabled = $n -gt 0; $ui.BtnCopy.IsEnabled = $n -gt 0
  $ui.Distinct.Text = if ($n) { Plural $n "adresse distincte" "adresses distinctes" } else { "" }
}

function Export-Addresses {
  if (-not $script:distinctes.Count) { return }
  $d = New-Object Microsoft.Win32.SaveFileDialog
  $d.Title = "Exporter les adresses"
  $d.Filter = "Texte, une adresse par ligne (*.txt)|*.txt"
  $d.FileName = "adresses-" + (Source-Name).ToLower() + "-" + (Get-Date -Format "yyyy-MM-dd") + ".txt"
  if (Test-Path $cfg.Folder) { $d.InitialDirectory = $cfg.Folder }
  if (-not $d.ShowDialog($win)) { return }
  [System.IO.File]::WriteAllText($d.FileName, (($script:distinctes -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding $false))
  $ko = [Math]::Max(1, [Math]::Round((Get-Item $d.FileName).Length / 1KB))
  Notice "$(Plural $script:distinctes.Count "adresse exportée" "adresses exportées") dans $(Split-Path $d.FileName -Leaf) ($ko Ko), une par ligne." "ok"
}

function Elapsed-Text([double]$sec) {
  $m = [Math]::Floor($sec / 60)
  if ($m -lt 1) { return "moins d'une minute" }
  if ($m -lt 60) { return "$m min" }
  return ("{0} h {1:00}" -f [Math]::Floor($m / 60), ($m % 60))
}

function Journal-Line([string]$q, $list, $keep) {
  $parts = for ($i = 0; $i -lt $list.Count; $i++) { $list[$i].Name + $US + $list[$i].Mail + $US + $(if ($keep[$i]) { "1" } else { "0" }) }
  return $q + "`t" + ($parts -join $RS)
}

# Le journal du passage interrompu, ou rien. Une dernière ligne coupée (arrêt
# brutal pendant l'écriture) est ignorée.
function Read-Journal {
  if (-not (Test-Path $resumePath)) { return $null }
  try {
    $lines = [System.IO.File]::ReadAllLines($resumePath, [System.Text.Encoding]::UTF8)
    if ($lines.Count -lt 1) { return $null }
    $head = $lines[0] | ConvertFrom-Json
    $entries = New-Object System.Collections.Generic.List[object]
    $kept = [ordered]@{}
    for ($i = 1; $i -lt $lines.Count; $i++) {
      $tab = $lines[$i].IndexOf("`t")
      if ($tab -lt 0) { continue }
      $names = @(); $mails = @(); $keeps = @()
      $rest = $lines[$i].Substring($tab + 1)
      $ok = $true
      if ($rest) {
        foreach ($item in $rest.Split($RS)) {
          $f = $item.Split($US)
          if ($f.Count -ne 3) { $ok = $false; break }
          $names += $f[0]; $mails += $f[1]; $keeps += ($f[2] -eq "1")
          if ($f[2] -eq "1") { $kept[$f[1].ToLower()] = 1 }
        }
      }
      if (-not $ok) { break }
      $entries.Add([pscustomobject]@{ Q = $lines[$i].Substring(0, $tab); Names = $names; Mails = $mails; Keep = $keeps })
    }
    return [pscustomobject]@{ Header = $head; Entries = $entries; Kept = $kept.Count; KeptList = @($kept.Keys) }
  } catch { return $null }
}

function Clear-Journal { Remove-Item $resumePath -ErrorAction SilentlyContinue }

# Affiche, ou non, la proposition de reprendre.
function Refresh-Resume {
  $j = Read-Journal
  if ($j -and $j.Entries.Count -gt 0) {
    $when = try { ([datetime]$j.Header.Started).ToString("dd/MM 'à' HH:mm") } catch { "" }
    $file = Split-Path $j.Header.Base -Leaf
    $ui.ResumeText.Text = "Commencé le $when : $(Plural $j.Entries.Count "recherche" "recherches"), $(Plural $j.Kept "adresse" "adresses"), dans $file. Reprendre continue là où il s'était arrêté, dans le même fichier, sans retaper ce qui est fait."
    $ui.ResumeBox.Visibility = "Visible"
    $ui.BtnRun.Content = "Reprendre le passage"
  } else {
    $ui.ResumeBox.Visibility = "Collapsed"
    $ui.BtnRun.Content = if ($script:hasRun) { "Relancer" } else { "Lancer le passage" }
  }
  Refresh-Known
}

function Passage([bool]$resume) {
  Read-Config
  $journal = if ($resume) { Read-Journal } else { $null }
  if ($journal) {
    # Une reprise se fait avec les réglages du passage commencé, sinon les
    # recherches ne seraient plus les mêmes.
    foreach ($k in "Source", "GmailAccount", "Prefix", "Keep", "Exclude", "Dedupe", "Format") {
      if ($null -ne $journal.Header.$k) { $cfg[$k] = $journal.Header.$k }
    }
    Apply-Config; Update-Summary
  }
  Save-Config
  $script:hasRun = $true
  $queries = Build-Queries
  $src = Source-Name
  $gmail = $cfg.Source -eq "gmail"
  if (-not $gmail) {
    $p = Outlook-Process
    if (-not $p) { Notice "Outlook n'est pas ouvert. Ouvrez-le, puis relancez." "warn"; return }
    $script:targetPid = $p.Id
  }

  $sp = $SpeedInfo[$cfg.Speed]
  $script:stop = $false
  Reset-Results; Build-Chips $queries

  # Le planificateur choisit chaque recherche : les 26 lettres, puis les
  # pistes les plus rentables, où qu'elles soient dans l'alphabet.
  $plan = [Planif]::new([string]$cfg.Prefix)
  $prefixLen = $cfg.Prefix.Length
  $chipState = @{}
  $records = New-Object System.Collections.Generic.List[object]
  $seen = @{}

  # Ce qu'il faut garder d'une réponse : rend les nouvelles lignes de résultat.
  $keepFrom = {
    param($q, $list, $keep)
    $out = @()
    for ($i = 0; $i -lt $list.Count; $i++) {
      if (-not $keep[$i]) { continue }
      if ($cfg.Dedupe -and $seen.ContainsKey($list[$i].Mail)) { continue }
      $seen[$list[$i].Mail] = 1
      $r = [pscustomobject]@{ Query = $q; Rank = $i + 1; Name = $list[$i].Name; Mail = $list[$i].Mail }
      $records.Add($r); $out += $r
    }
    return $out
  }
  $markChip = {
    param($q, $count)
    $root = $q.Substring(0, [Math]::Min($q.Length, $prefixLen + 1))
    if ($count -gt 0) { $chipState[$root] = "ok" } elseif (-not $chipState.ContainsKey($root)) { $chipState[$root] = "none" }
    Chip $root $chipState[$root]
  }

  # Reprise : on rejoue le journal dans le planificateur, sans rien taper. Il
  # retrouve exactement l'état où il était, recherche par recherche.
  $replayed = New-Object System.Collections.Generic.List[string]
  if ($journal) {
    $diverged = $false
    foreach ($e in $journal.Entries) {
      $list = @(for ($i = 0; $i -lt $e.Mails.Count; $i++) { [pscustomobject]@{ Name = $e.Names[$i]; Mail = $e.Mails[$i] } })
      $keep = [bool[]]@($e.Keep)
      $q = if ($diverged) { $null } else { $plan.Next() }
      if ($q -ne $e.Q) {
        # Le planificateur a changé depuis : on garde les adresses, pas la
        # suite, et la recherche qu'il proposait reste à faire.
        if ($q) { $plan.Requeue($q) }
        $diverged = $true
        & $keepFrom $e.Q $list $keep | Out-Null
        continue
      }
      $plan.Report($q, [string[]]@($e.Names), [string[]]@($e.Mails), $keep) | Out-Null
      & $keepFrom $q $list $keep | Out-Null
      & $markChip $q $list.Count
      $replayed.Add((Journal-Line $q $list $keep))
    }
    $base = $journal.Header.Base
    $startedText = [string]$journal.Header.Started
    $elapsedBefore = [double]$journal.Header.Seconds
  } else {
    $folder = if (Test-Path $cfg.Folder) { $cfg.Folder } else { [Environment]::GetFolderPath("Desktop") }
    $base = Join-Path $folder ("adresses-" + $src.ToLower() + "-" + (Get-Date -Format "yyyy-MM-dd-HHmm"))
    $startedText = (Get-Date).ToString("o")
    $elapsedBefore = 0
  }
  $foundBefore = $records.Count
  $show = [Math]::Max(0, $records.Count - 300)
  for ($i = $show; $i -lt $records.Count; $i++) { Add-Row $records[$i].Query $records[$i].Mail $records[$i].Name }

  Compact-Mode $true
  $ui.LiveEta.Text = ""

  if ($gmail) {
    $ui.LiveMail.Text = "Ouverture de Gmail"; Pump
    $p = Ouvrir-Gmail
    if ($p -is [string]) { Compact-Mode $false; Notice $p "warn"; return }
  } else {
    Mettre-Devant $p
    Wait 1200
    if (-not (Au-Premier-Plan $p)) { Compact-Mode $false; Notice "Impossible de mettre Outlook au premier plan. Cliquez dans Outlook, puis relancez." "warn"; return }

    Taper("^n")
    $pret = $false
    for ($i = 0; $i -lt 10 -and -not $pret; $i++) { Wait 500; $pret = Focus-Dans-Champ-A }
    if (-not $pret) { Compact-Mode $false; Notice "Le nouveau message ne s'est pas ouvert avec le curseur dans le champ « À ». Rien n'a été tapé." "warn"; return }
  }

  # Tout est prêt : les fichiers et le journal s'ouvrent, et restent à jour à
  # chaque recherche.
  Open-Outputs $base $records
  if (-not (Test-Path $cfgDir)) { New-Item -ItemType Directory -Path $cfgDir | Out-Null }
  $head = [ordered]@{ Started = $startedText; Base = $base; Seconds = [int]$elapsedBefore }
  foreach ($k in "Source", "GmailAccount", "Prefix", "Keep", "Exclude", "Dedupe", "Format") { $head[$k] = $cfg[$k] }
  $jw = New-Writer $resumePath
  $jw.WriteLine(($head | ConvertTo-Json -Compress))
  foreach ($l in $replayed) { $jw.WriteLine($l) }

  $arret = $null
  $prevLen = 0
  $runStart = Get-Date
  $script:prevSig = ""

  try {
    while ($true) {
      if ($script:stop) { $arret = "arrêté à votre demande."; break }
      $q = $plan.Next()
      if (-not $q) { break }
      if (-not (Au-Premier-Plan $p)) { $arret = "$src n'est plus au premier plan (recherche « $q »)."; break }
      if (-not (Focus-Dans-Champ-A)) { $arret = "le curseur a quitté le champ « À » (recherche « $q »)."; break }

      # Pour les recherches plus longues, c'est la pastille de la lettre de
      # départ qui s'allume.
      $root = $q.Substring(0, $prefixLen + 1)
      Chip $root "now"; $ui.LiveQuery.Text = $q.ToUpper(); $ui.LiveMail.Text = "Recherche en cours"; Pump
      Taper("{BACKSPACE " + ($prevLen + 2) + "}")
      Wait 400
      Taper((Escape-Keys $q))
      $prevLen = $q.Length
      Wait $sp.Settle

      $list = @(Lire-Reponse $p $q $sp)
      $keep = [bool[]]@($list | ForEach-Object { Passes-Filters $_.Mail })
      $plan.Report($q, [string[]]@($list | ForEach-Object { $_.Name }), [string[]]@($list | ForEach-Object { $_.Mail }), $keep) | Out-Null
      $jw.WriteLine((Journal-Line $q $list $keep))

      $new = @(& $keepFrom $q $list $keep)
      foreach ($r in $new) { Write-Record $r; Add-Row $r.Query $r.Mail $r.Name; $ui.LiveMail.Text = $r.Mail }
      & $markChip $q $list.Count
      if ($new.Count -eq 0) { $ui.LiveMail.Text = if ($list.Count -eq 0) { "Aucune suggestion" } else { "Rien de nouveau" } }

      # La fin n'est connue qu'à peu près : la barre suit les pistes encore
      # rentables, qui se découvrent en route.
      $done = $plan.Done
      Progress $done ($done + [Math]::Max($plan.Pending, 26 - $done))
      $ui.LiveEta.Text = (Elapsed-Text ($elapsedBefore + ((Get-Date) - $runStart).TotalSeconds)) + " · " + (Plural $records.Count "adresse" "adresses")
    }
    if (Focus-Dans-Champ-A) { Taper("{BACKSPACE " + ($prevLen + 2) + "}") }
  } finally {
    $script:distinctes = @($records | ForEach-Object { $_.Mail } | Sort-Object -Unique)
    Close-Outputs
    $jw.Close()
  }

  Compact-Mode $false
  $win.Activate() | Out-Null

  # La durée écoulée sert à la reprise : on la note en tête du journal.
  if ($arret) {
    try {
      $lines = [System.IO.File]::ReadAllLines($resumePath, [System.Text.Encoding]::UTF8)
      $head.Seconds = [int]($elapsedBefore + ((Get-Date) - $runStart).TotalSeconds)
      $lines[0] = $head | ConvertTo-Json -Compress
      [System.IO.File]::WriteAllLines($resumePath, $lines, (New-Object System.Text.UTF8Encoding $false))
    } catch { }
  } else { Clear-Journal }

  $ui.Distinct.Text = Plural $script:distinctes.Count "adresse distincte" "adresses distinctes"
  if ($records.Count -gt 0) {
    $cfg.Dernier = $script:files[0]; Save-Config
    if ($cfg.CopyAtEnd) { [System.Windows.Clipboard]::SetText(($script:distinctes -join [Environment]::NewLine)) }
    if ($cfg.OpenAtEnd) { Open-Files }
  } else {
    # Rien trouvé : pas de fichier vide qui traîne.
    foreach ($f in $script:files) { Remove-Item $f -ErrorAction SilentlyContinue }
    $script:files = @()
  }

  $names = ($script:files | ForEach-Object { Split-Path $_ -Leaf }) -join " et "
  $gained = $records.Count - $foundBefore
  $copied = if ($cfg.CopyAtEnd -and $records.Count -gt 0) { " Adresses copiées." } else { "" }
  if ($arret) {
    Notice ("Arrêté : $arret Tout ce qui a été trouvé est enregistré" + $(if ($names) { " dans $names" } else { "" }) + "." + $copied + " Vous pouvez reprendre le passage plus tard, là où il s'est arrêté.") "warn"
  } elseif ($records.Count -eq 0) {
    Notice "Terminé, mais aucune adresse ne passe vos filtres. Rien n'a été enregistré." "warn"
  } else {
    $more = $copied + " $($plan.Done) recherches."
    if ($journal) { $more += " Cette reprise a ajouté $(Plural $gained "adresse" "adresses")." }
    $more += switch ($plan.StopReason) {
      "saturation" { " Arrêt automatique : les $($plan.Saturation) dernières recherches n'apportaient plus d'adresse nouvelle." }
      default      { " Toutes les pistes ont été explorées." }
    }
    Notice "Terminé : $(Plural $script:distinctes.Count "adresse" "adresses"), enregistrées dans $names.$more Le brouillon reste ouvert dans $src, champ « À » vide : vous pouvez le fermer$(if ($gmail) { ', et le supprimer des brouillons si Gmail l''y a gardé' })." "ok"
  }
}

function Open-Files {
  foreach ($f in $script:files) {
    if (Test-Path $f) { if ($f.EndsWith(".csv")) { Start-Process $f } else { Start-Process notepad.exe $f } }
  }
}

# ---------------------------------------------------------------- branchements

$ui.Header.Add_MouseLeftButtonDown({ $win.DragMove() })
$ui.BtnClose.Add_Click({ Read-Config; Save-Config; $win.Close() })
$ui.BtnMin.Add_Click({ $win.WindowState = "Minimized" })

$ui.TabRun.Add_Checked({ $ui.PageRun.Visibility = "Visible"; $ui.PageSet.Visibility = "Collapsed" })
$ui.TabSet.Add_Checked({ $ui.PageRun.Visibility = "Collapsed"; $ui.PageSet.Visibility = "Visible" })
$ui.BtnEdit.Add_Click({ $ui.TabSet.IsChecked = $true })

foreach ($n in "SrcOutlook","SrcGmail",
               "SpeedFast","SpeedNormal","SpeedCareful","FmtTxt","FmtCsv","FmtBoth") {
  $ui[$n].Add_Checked({ Update-Summary })
}
foreach ($n in "Dedupe","OpenAtEnd","CopyAtEnd") { $ui[$n].Add_Checked({ Update-Summary }); $ui[$n].Add_Unchecked({ Update-Summary }) }
foreach ($n in "Prefix","Keep","Exclude","GmailAccount") { $ui[$n].Add_TextChanged({ Update-Summary }) }

$ui.BtnFolder.Add_Click({
  $d = New-Object System.Windows.Forms.FolderBrowserDialog
  $d.Description = "Où enregistrer les fichiers"
  $d.SelectedPath = $ui.Folder.Text
  if ($d.ShowDialog() -eq "OK") { $ui.Folder.Text = $d.SelectedPath; Update-Summary }
})
$ui.BtnReset.Add_Click({
  $fresh = Default-Config
  foreach ($k in @($fresh.Keys)) { if ($k -ne "Dernier") { $cfg[$k] = $fresh[$k] } }
  Apply-Config; Update-Summary
})

$ui.BtnRun.Add_Click({
  $ui.TabRun.IsChecked = $true
  $ui.BtnRun.IsEnabled = $false; $ui.BtnExport.IsEnabled = $false; $ui.BtnCopy.IsEnabled = $false
  $resume = $ui.ResumeBox.Visibility -eq "Visible"
  try { Passage $resume } catch { Compact-Mode $false; Notice "Erreur : $($_.Exception.Message)" "warn" }
  $ui.BtnRun.IsEnabled = $true
  Refresh-Resume
})
$ui.BtnFresh.Add_Click({
  Clear-Journal
  Refresh-Resume
  Notice "Le passage interrompu est mis de côté : ses fichiers restent là où ils sont. Lancer le passage repart de A." "info"
})
$ui.BtnStop.Add_Click({ $script:stop = $true })
$ui.BtnExport.Add_Click({ Export-Addresses })
$ui.BtnCopy.Add_Click({
  if ($script:distinctes.Count) {
    [System.Windows.Clipboard]::SetText(($script:distinctes -join [Environment]::NewLine))
    Notice "$(Plural $script:distinctes.Count "adresse copiée" "adresses copiées"), une par ligne." "ok"
  }
})

Apply-Config
Update-Summary
Refresh-Resume
[void]$win.ShowDialog()
