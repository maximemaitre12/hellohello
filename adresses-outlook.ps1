# Adresses Outlook
#
# Ouvre un nouveau message dans le nouvel Outlook, tape les lettres de A à Z
# dans le champ « À » et note toutes les adresses que propose la liste de
# suggestions. Là où la liste est pleine, des adresses restent cachées : la
# recherche est approfondie d'une lettre (a devient aa, ab...), jusqu'à la
# profondeur et au nombre de recherches réglés. Le résultat va dans un fichier
# texte et/ou CSV.
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
using System; using System.Runtime.InteropServices;
public class Fen {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, UIntPtr e);
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
#   non), et retire les adresses déjà notées qui occuperaient la liste.
# - Le passage s'arrête seul quand il n'y a plus de piste rentable, ou quand
#   une longue série de recherches n'a plus rien apporté.
#
# Testé hors Outlook par outils\tests\simulation-adresses.ps1, qui extrait ce
# bloc tel quel.
$PlanifSource = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public class Planif {
  class Node { public string Q; public string Parent; public int Depth; public double Est, EstRaw; public int Kind; }
  class Seen { public int Count; public double Novelty; public double Est; public bool Root, Virtual; public string Last; }

  public int Cap = 0, MaxDepth, Budget, Done = 0, SinceNew = 0, Pending = 0, Saturation;
  public string StopReason = "";

  string rawPrefix, lastPrefix;
  string[] fixedWords;
  Queue<Node> roots = new Queue<Node>();
  List<Node> cands = new List<Node>();
  Dictionary<string, Node> nodes = new Dictionary<string, Node>();
  Dictionary<string, Seen> seen = new Dictionary<string, Seen>();
  HashSet<string> mails = new HashSet<string>();
  Dictionary<string, int> known = new Dictionary<string, int>();
  // Suites de lettres vues : [précédente, suivante], 26 = début de mot.
  double[,] big = new double[27, 26]; double[] bigRow = new double[27];
  double[,] start = new double[26, 26]; double[] startRow = new double[26];
  // Calibrage : taille de liste prévue et réellement affichée, par profondeur
  // et selon que la suite de lettres a déjà été vue ou non.
  double[] predicted = new double[8], obtained = new double[8];

  public Planif(string prefix, int maxDepth, int budget) {
    rawPrefix = prefix ?? "";
    string np = Norm(rawPrefix);
    int sp = np.LastIndexOf(' ');
    lastPrefix = np.Substring(sp + 1);
    fixedWords = sp < 0 ? new string[0] : np.Substring(0, sp).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
    MaxDepth = maxDepth; Budget = budget;
    Saturation = Math.Max(40, budget / 6);
    for (char c = 'a'; c <= 'z'; c++) {
      Node n = new Node { Q = rawPrefix + c, Parent = null, Depth = 1, Est = 0 };
      nodes[n.Q] = n; roots.Enqueue(n);
    }
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

  // Le contact correspond-il à la recherche, à la façon d'Outlook ? Chaque mot
  // de la recherche doit commencer un mot du nom ou de l'adresse.
  public static bool Matches(string text, string query) {
    List<string> toks = Tokens(text);
    foreach (string qw in Tokens(query)) {
      bool ok = false;
      foreach (string t in toks) if (t.StartsWith(qw, StringComparison.Ordinal)) { ok = true; break; }
      if (!ok) return false;
    }
    return true;
  }

  static string LastWord(string q) { string n = Norm(q); return n.Substring(n.LastIndexOf(' ') + 1); }
  int Known(string lw) { int k; return known.TryGetValue(lw, out k) ? k : 0; }

  void Learn(string name, string mail) {
    List<string> toks = Tokens(name + " " + mail);
    HashSet<string> prefixes = new HashSet<string>();
    bool inScope = true;
    foreach (string fw in fixedWords) {
      bool ok = false;
      foreach (string t in toks) if (t.StartsWith(fw, StringComparison.Ordinal)) { ok = true; break; }
      if (!ok) { inScope = false; break; }
    }
    int maxLen = lastPrefix.Length + MaxDepth + 1;
    foreach (string t in toks) {
      int prev = 26;
      for (int i = 0; i < t.Length; i++) {
        int c = t[i] - 'a';
        if (c < 0 || c > 25) { prev = -1; continue; }
        if (prev >= 0) { big[prev, c]++; bigRow[prev]++; }
        if (i == 1 && prev >= 0 && prev < 26) { start[prev, c]++; startRow[prev]++; }
        prev = c;
      }
      if (inScope) for (int l = 1; l <= Math.Min(t.Length, maxLen); l++) prefixes.Add(t.Substring(0, l));
    }
    foreach (string p in prefixes) known[p] = Known(p) + 1;
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
    // une initiale collée au nom (gpoirel, tmaitre), dont la deuxième lettre
    // suit alors la répartition des débuts de noms.
    double s = (start[p, c] + 0.2) / (startRow[p] + 5.2);
    double initial = (big[26, c] + 0.2) / (bigRow[26] + 5.2);
    return 0.45 * s + 0.35 * initial + 0.2 * g;
  }

  // Liste pleine : Outlook cache des contacts sous cette recherche. Un nœud
  // virtuel est plein par construction (voir Next).
  bool Full(Seen s) { return s.Virtual || (s.Count >= 3 && s.Count >= Cap); }

  // Nombre estimé de contacts sous une recherche pleine. On n'en voit que le
  // plafond ; on suppose au moins le double de ce qu'on en connaît déjà.
  double TotalEst(Seen s) {
    if (!Full(s)) return s.Count;
    double floor = s.Root ? Cap * 6.0 : Cap * 2.0;
    return Math.Max(floor, Math.Max(Known(s.Last) * 2.0, s.Est));
  }

  // Adresses inconnues que la liste de n devrait montrer, pondérées par ce que
  // la branche a rapporté jusqu'ici.
  double Score(Node n) {
    Seen p = seen[n.Parent];
    if (!Full(p)) return 0;
    string lw = LastWord(n.Q);
    int k = Known(lw);
    // Part des contacts du parent dont un mot continue par cette lettre : ce
    // qu'on en a vu, complété par la fréquence de cette suite de lettres.
    double share = (k + 3.0 * Continuation(lw)) / (Known(p.Last) + 3.0);
    n.Kind = Math.Min(n.Depth, 4) - 1 + (k == 0 ? 4 : 0);
    n.EstRaw = TotalEst(p) * Math.Min(1.0, share);
    double est = n.EstRaw * Calibration(n.Kind);
    n.Est = est;
    double expNew = Math.Min(Cap, Math.Max(est, k)) - k;
    if (expNew <= 0) return 0;
    return expNew * (0.3 + p.Novelty);
  }

  // Taille de liste réelle sur taille prévue pour ce genre de piste. Le modèle
  // se trompe forcément (chaque carnet a ses habitudes de noms) : il se
  // corrige au fil du passage. Les deux côtés sont plafonnés comme la liste
  // d'Outlook, donc seules les pistes mal estimées font bouger le facteur.
  double Calibration(int kind) { return (obtained[kind] + 3.0) / (predicted[kind] + 3.0); }

  void Expand(string q, int depth) {
    if (depth >= MaxDepth) return;
    for (char c = 'a'; c <= 'z'; c++) {
      string x = q + c;
      if (nodes.ContainsKey(x)) continue;
      Node ch = new Node { Q = x, Parent = q, Depth = depth + 1 };
      nodes[x] = ch; cands.Add(ch);
    }
  }

  public const double MinScore = 0.1;

  // Pour le banc d'essai : calibrage par genre de piste, et meilleures pistes restantes.
  public string Diagnostic() {
    StringBuilder b = new StringBuilder();
    for (int i = 0; i < 8; i++) b.AppendFormat("kind{0} pred={1:0.0} obt={2:0} cal={3:0.00}{4}", i, predicted[i], obtained[i], Calibration(i), Environment.NewLine);
    List<KeyValuePair<double, string>> l = new List<KeyValuePair<double, string>>();
    foreach (Node n in cands) { double sc = Score(n); l.Add(new KeyValuePair<double, string>(sc, n.Q + " k=" + Known(LastWord(n.Q)) + " est=" + n.Est.ToString("0.00") + " raw=" + n.EstRaw.ToString("0.000") + " parentTot=" + TotalEst(seen[n.Parent]).ToString("0") + " parentKnown=" + Known(seen[n.Parent].Last))); }
    l.Sort((x, y) => y.Key.CompareTo(x.Key));
    for (int i = 0; i < Math.Min(15, l.Count); i++) b.AppendFormat("{0:0.000} {1}{2}", l[i].Key, l[i].Value, Environment.NewLine);
    return b.ToString();
  }

  public string Next() {
    if (Done >= Budget) { StopReason = "budget"; return null; }
    if (roots.Count > 0) return roots.Dequeue().Q;
    if (SinceNew >= Saturation) { StopReason = "saturation"; return null; }
    while (true) {
      // Une recherche dont on connaît déjà de quoi remplir la liste ne
      // montrerait que du connu : on ne la tape pas, on passe à ses suites.
      List<Node> virt = null;
      foreach (Node n in cands) {
        if (Cap >= 3 && n.Depth < MaxDepth && Known(LastWord(n.Q)) >= Cap && Full(seen[n.Parent])) {
          if (virt == null) virt = new List<Node>();
          virt.Add(n);
        }
      }
      if (virt != null) {
        foreach (Node n in virt) {
          cands.Remove(n);
          Seen p = seen[n.Parent];
          Score(n);
          seen[n.Q] = new Seen { Virtual = true, Est = n.Est, Novelty = p.Novelty, Last = LastWord(n.Q) };
          Expand(n.Q, n.Depth);
        }
        continue;
      }
      Node best = null; double bs = MinScore; int pending = 0;
      foreach (Node n in cands) {
        double sc = Score(n);
        if (sc >= MinScore) pending++;
        if (sc >= bs) { bs = sc; best = n; }
      }
      if (best == null) { Pending = 0; StopReason = "epuise"; return null; }
      Pending = pending - 1;
      cands.Remove(best);
      return best.Q;
    }
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
    Node node; nodes.TryGetValue(q, out node);
    int depth = node == null ? 1 : node.Depth;
    if (node != null && node.Parent != null) { predicted[node.Kind] += Math.Min(Cap, node.EstRaw); obtained[node.Kind] += mailList.Length; }
    seen[q] = new Seen { Count = mailList.Length, Novelty = (fresh + 0.5) / (mailList.Length + 1.0),
                         Est = node == null ? 0 : node.Est, Root = node == null || node.Parent == null, Last = LastWord(q) };
    if (mailList.Length >= 3) Expand(q, depth);
    return fresh;
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
    SmartDepth = 3; SmartBudget = 300
  }
}
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
                <TextBlock x:Name="Count" Text="0 / 26" Foreground="#5B6A80" HorizontalAlignment="Right"/>
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

          <Grid Margin="0,12,0,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="10"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="BtnRun" Style="{StaticResource Primary}" Content="Lancer le passage"/>
            <Button x:Name="BtnOpen" Grid.Column="2" Style="{StaticResource Secondary}" Content="Ouvrir le fichier" IsEnabled="False"/>
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
                <TextBlock Style="{StaticResource SectionTitle}" Text="1. Jusqu'où chercher"/>
                <StackPanel x:Name="SmartBox">
                  <Border Background="#F1F6FD" CornerRadius="10" Padding="12,10" Margin="0,0,0,12">
                    <TextBlock TextWrapping="Wrap" Foreground="#1E4D8C" FontSize="12"
                               Text="Commence par les 26 lettres. Quand la liste de suggestions est pleine, des adresses restent cachées dessous : la recherche devient une piste à creuser (a devient al, am, an...). Les pistes passent par ordre de rendement, sur tout l'alphabet : d'abord les suites de lettres qui existent vraiment dans vos contacts, et jamais celles qui ne montreraient que des adresses déjà notées. Le passage s'arrête seul quand il n'y a plus rien à gagner."/>
                  </Border>
                  <TextBlock Style="{StaticResource Label}" Text="Profondeur maximale" Margin="0,0,0,8"/>
                  <WrapPanel>
                    <RadioButton x:Name="SmartD2" GroupName="sdepth" Style="{StaticResource Pill}" Content="2 lettres"/>
                    <RadioButton x:Name="SmartD3" GroupName="sdepth" Style="{StaticResource Pill}" Content="3 lettres"/>
                    <RadioButton x:Name="SmartD4" GroupName="sdepth" Style="{StaticResource Pill}" Content="4 lettres"/>
                  </WrapPanel>
                  <TextBlock Style="{StaticResource Help}" Margin="0,0,0,12" Text="Plus c'est profond, plus on trouve d'adresses dans les groupes nombreux, et plus c'est long."/>
                  <TextBlock Style="{StaticResource Label}" Text="Nombre maximum de recherches" Margin="0,0,0,8"/>
                  <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="SmartBudget" Style="{StaticResource Input}" Width="110" MaxLength="6" TextAlignment="Right" InputMethod.IsInputMethodEnabled="False"/>
                    <TextBlock Text="recherches" Foreground="#5B6A80" VerticalAlignment="Center" Margin="10,0,0,0"/>
                  </StackPanel>
                  <TextBlock Style="{StaticResource Help}" Text="Un plafond, pas un objectif : le passage s'arrête souvent avant, dès que les recherches ne rapportent plus. Les adresses trouvées sont toujours gardées."/>
                </StackPanel>

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
                  <RadioButton x:Name="FmtTxt" GroupName="fmt" Style="{StaticResource Pill}" Content="Texte"/>
                  <RadioButton x:Name="FmtCsv" GroupName="fmt" Style="{StaticResource Pill}" Content="CSV pour Excel"/>
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
            <TextBlock x:Name="LiveCount" Text="0 / 26" FontWeight="SemiBold" HorizontalAlignment="Right"/>
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
  switch ([int]$cfg.SmartDepth) { 2 { $ui.SmartD2.IsChecked = $true } 4 { $ui.SmartD4.IsChecked = $true } default { $ui.SmartD3.IsChecked = $true } }
  $ui.SmartBudget.Text = [string][int]$cfg.SmartBudget
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
  $cfg.SmartDepth = if ($ui.SmartD2.IsChecked) { 2 } elseif ($ui.SmartD4.IsChecked) { 4 } else { 3 }
  # Le nombre tapé, tel quel. Champ vide ou zéro : on garde le dernier nombre valable.
  $n = 0
  if ([int]::TryParse(($ui.SmartBudget.Text -replace '\D', ''), [ref]$n) -and $n -ge 1) { $cfg.SmartBudget = $n }
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
  # La durée dépend de ce qu'Outlook contient : on n'annonce que le plafond.
  $max = Duration-Text ([int]$cfg.SmartBudget * $SpeedInfo[$cfg.Speed].Sec)
  $ui.Plan.Text = "$start, puis les pistes les plus rentables jusqu'à $($cfg.SmartDepth) lettres · $($cfg.SmartBudget) recherches au plus$ftxt · $fmt · $max au maximum, souvent bien moins"
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
  $ui.Count.Text = "0 / $($queries.Count)"; $ui.Fill.Width = 0
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

function Progress([int]$done, [int]$total) {
  $ui.Count.Text = "$done / $total"; $ui.LiveCount.Text = "$done / $total"
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
$script:targetPid = 0
function Focus-Dans-Champ-A {
  $f = $AE::FocusedElement
  if (-not $f) { return $false }
  if ($cfg.Source -eq "gmail") {
    return ($f.Current.ControlType -eq $CT::ComboBox -and $f.Current.ProcessId -eq $script:targetPid)
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
  $racine = $AE::FromHandle($p.MainWindowHandle)
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
  $vu = $false
  for ($i = 0; $i -lt 40; $i++) {
    Wait 500
    $h = [Fen]::GetForegroundWindow()
    $procId = 0; [Fen]::GetWindowThreadProcessId($h, [ref]$procId) | Out-Null
    $t = (Get-Process -Id $procId -ErrorAction SilentlyContinue).MainWindowTitle
    $titre = $AE::FromHandle($h).Current.Name
    if ($titre -notmatch 'Gmail' -and $t -notmatch 'Gmail') { continue }
    $vu = $true
    $script:targetPid = $procId
    if (Focus-Dans-Champ-A) { return [pscustomobject]@{ Id = $procId; MainWindowHandle = $h } }
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

function Passage {
  Read-Config; Save-Config
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
  Compact-Mode $true
  $ui.LiveEta.Text = Duration-Text ($queries.Count * $sp.Sec)

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

  $records = New-Object System.Collections.Generic.List[object]
  $seen = @{}
  $arret = $null
  $prevLen = 0
  $started = Get-Date
  $script:prevSig = ""

  # Le planificateur choisit chaque recherche : les 26 lettres, puis les
  # pistes les plus rentables, où qu'elles soient dans l'alphabet.
  $budget = [int]$cfg.SmartBudget
  $plan = New-Object Planif -ArgumentList $cfg.Prefix, ([int]$cfg.SmartDepth), $budget
  $prefixLen = $cfg.Prefix.Length
  $chipState = @{}

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
    $plan.Report($q, [string[]]@($list | ForEach-Object { $_.Name }), [string[]]@($list | ForEach-Object { $_.Mail }),
                 [bool[]]@($list | ForEach-Object { Passes-Filters $_.Mail })) | Out-Null

    $kept = 0
    $rank = 0
    foreach ($s in $list) {
      $rank++
      if (-not (Passes-Filters $s.Mail)) { continue }
      if ($cfg.Dedupe -and $seen.ContainsKey($s.Mail)) { continue }
      $seen[$s.Mail] = 1
      $records.Add([pscustomobject]@{ Query = $q; Rank = $rank; Name = $s.Name; Mail = $s.Mail })
      Add-Row $q $s.Mail $s.Name
      $ui.LiveMail.Text = $s.Mail
      $kept++
    }
    # Jaune seulement si Outlook n'a rien proposé, pas si tout était déjà noté.
    if ($list.Count -gt 0) { $chipState[$root] = "ok" }
    elseif (-not $chipState.ContainsKey($root)) { $chipState[$root] = "none" }
    Chip $root $chipState[$root]
    if ($kept -eq 0) { $ui.LiveMail.Text = if ($list.Count -eq 0) { "Aucune suggestion" } else { "Rien de nouveau après filtres" } }

    # Ce qui reste n'est connu qu'à peu près : les pistes encore rentables,
    # dans la limite du nombre de recherches.
    $done = $plan.Done
    $total = [Math]::Min($budget, $done + [Math]::Max($plan.Pending, 26 - $done))
    Progress $done $total
    $left = ((Get-Date) - $started).TotalSeconds / $done * ($total - $done)
    $ui.LiveEta.Text = if ($total -gt $done) { "reste environ " + (Duration-Text $left) } else { "" }
  }
  if (Focus-Dans-Champ-A) { Taper("{BACKSPACE " + ($prevLen + 2) + "}") }

  Compact-Mode $false
  $win.Activate() | Out-Null

  $script:files = @()
  $script:distinctes = @($records | ForEach-Object { $_.Mail } | Sort-Object -Unique)
  $ui.Distinct.Text = Plural $script:distinctes.Count "adresse distincte" "adresses distinctes"
  if ($records.Count -gt 0) {
    $folder = if (Test-Path $cfg.Folder) { $cfg.Folder } else { [Environment]::GetFolderPath("Desktop") }
    $base = Join-Path $folder ("adresses-" + $src.ToLower() + "-" + (Get-Date -Format "yyyy-MM-dd-HHmm"))
    if ($cfg.Format -in "txt", "both") {
      $lines = $records | ForEach-Object { "$($_.Query) : $($_.Mail)    ($($_.Name))" }
      $contenu = @($lines) + @("", "Adresses distinctes : $($script:distinctes.Count)") + $script:distinctes
      $contenu | Set-Content -Path "$base.txt" -Encoding UTF8
      $script:files += "$base.txt"
    }
    if ($cfg.Format -in "csv", "both") {
      $csv = @("Recherche;Rang;Nom;Adresse") + ($records | ForEach-Object { (Csv-Field $_.Query) + ";" + $_.Rank + ";" + (Csv-Field $_.Name) + ";" + (Csv-Field $_.Mail) })
      $csv | Set-Content -Path "$base.csv" -Encoding UTF8
      $script:files += "$base.csv"
    }
    $ui.BtnOpen.IsEnabled = $true; $ui.BtnCopy.IsEnabled = $true
    if ($cfg.CopyAtEnd) { [System.Windows.Clipboard]::SetText(($script:distinctes -join [Environment]::NewLine)) }
    if ($cfg.OpenAtEnd) { Open-Files }
  }

  $names = ($script:files | ForEach-Object { Split-Path $_ -Leaf }) -join " et "
  if ($arret) {
    Notice ("Arrêté : $arret Ce qui était fait est gardé" + $(if ($names) { " dans $names." } else { "." })) "warn"
  } elseif ($records.Count -eq 0) {
    Notice "Terminé, mais aucune adresse ne passe vos filtres. Rien n'a été enregistré." "warn"
  } else {
    $copied = if ($cfg.CopyAtEnd) { " Adresses copiées." } else { "" }
    $copied += " $($plan.Done) recherches."
    $copied += switch ($plan.StopReason) {
      "budget"     { " La limite de $budget recherches a été atteinte alors qu'il restait des pistes : augmentez-la pour aller plus loin." }
      "saturation" { " Arrêt automatique : les $($plan.Saturation) dernières recherches n'apportaient plus d'adresse nouvelle." }
      default      { " Toutes les pistes rentables ont été explorées." }
    }
    Notice "Terminé : $(Plural $script:distinctes.Count "adresse" "adresses"), enregistrées dans $names.$copied Le brouillon reste ouvert dans $src, champ « À » vide : vous pouvez le fermer$(if ($gmail) { ', et le supprimer des brouillons si Gmail l''y a gardé' })." "ok"
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

foreach ($n in "SmartD2","SmartD3","SmartD4","SrcOutlook","SrcGmail",
               "SpeedFast","SpeedNormal","SpeedCareful","FmtTxt","FmtCsv","FmtBoth") {
  $ui[$n].Add_Checked({ Update-Summary })
}
foreach ($n in "Dedupe","OpenAtEnd","CopyAtEnd") { $ui[$n].Add_Checked({ Update-Summary }); $ui[$n].Add_Unchecked({ Update-Summary }) }
foreach ($n in "Prefix","Keep","Exclude","SmartBudget","GmailAccount") { $ui[$n].Add_TextChanged({ Update-Summary }) }

# Le nombre de recherches n'accepte que des chiffres, au clavier comme au collage.
$ui.SmartBudget.Add_PreviewTextInput({ param($s, $e) if ($e.Text -match '\D') { $e.Handled = $true } })
[System.Windows.DataObject]::AddPastingHandler($ui.SmartBudget, [System.Windows.DataObjectPastingEventHandler]{
  param($s, $e)
  $t = [string]$e.DataObject.GetData([string])
  if ($t -match '\D') { $e.CancelCommand() }
})
# En quittant un champ vide, on réaffiche le nombre retenu.
$ui.SmartBudget.Add_LostFocus({ $ui.SmartBudget.Text = [string][int]$cfg.SmartBudget })

$ui.BtnFolder.Add_Click({
  $d = New-Object System.Windows.Forms.FolderBrowserDialog
  $d.Description = "Où enregistrer les fichiers"
  $d.SelectedPath = $ui.Folder.Text
  if ($d.ShowDialog() -eq "OK") { $ui.Folder.Text = $d.SelectedPath; Update-Summary }
})
$ui.BtnReset.Add_Click({
  $fresh = Default-Config
  foreach ($k in @($fresh.Keys)) { $cfg[$k] = $fresh[$k] }
  Apply-Config; Update-Summary
})

$ui.BtnRun.Add_Click({
  $ui.TabRun.IsChecked = $true
  $ui.BtnRun.IsEnabled = $false; $ui.BtnOpen.IsEnabled = $false; $ui.BtnCopy.IsEnabled = $false
  try { Passage } catch { Compact-Mode $false; Notice "Erreur : $($_.Exception.Message)" "warn" }
  $ui.BtnRun.IsEnabled = $true; $ui.BtnRun.Content = "Relancer"
})
$ui.BtnStop.Add_Click({ $script:stop = $true })
$ui.BtnOpen.Add_Click({ Open-Files })
$ui.BtnCopy.Add_Click({
  if ($script:distinctes.Count) {
    [System.Windows.Clipboard]::SetText(($script:distinctes -join [Environment]::NewLine))
    Notice "$(Plural $script:distinctes.Count "adresse copiée" "adresses copiées"), une par ligne." "ok"
  }
})

Apply-Config
Update-Summary
[void]$win.ShowDialog()
