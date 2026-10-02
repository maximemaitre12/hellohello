// Faux carnet d'adresses réaliste, de quelques centaines à plusieurs dizaines
// de milliers de contacts, qui répond comme le nouvel Outlook : un contact sort
// dès que chaque mot de la recherche commence un mot de son nom ou de son
// adresse (domaine compris), les contacts les plus fréquents en tête, liste
// coupée au plafond.
//
// Les noms mêlent plusieurs origines (français, anglais, espagnol, chinois,
// ukrainien, arabe, indien, allemand, italien) avec une répartition très
// inégale, comme dans un vrai annuaire : beaucoup de Martin, peu de Zhelezniak.
// Au-delà des listes réelles, les noms de famille sont fabriqués à partir de
// syllabes, pour la longue traîne d'un annuaire de 40 000 personnes.

public class FauxOutlook {
  public class Contact { public string Name, Mail; public double W; public List<string> Toks; }
  public List<Contact> All = new List<Contact>();
  public int Cap; public int Queries = 0, Empty = 0;
  Dictionary<string, List<Contact>> index = new Dictionary<string, List<Contact>>();

  static string[] First = (
    "maxime thomas samuel gildas sabrina larisa noe yelena irene youriy vladyslav william xavier mason camille lea manon chloe " +
    "ines jade louise emma alice lucas hugo louis jules arthur paul nathan theo raphael sophie julie claire marie anne pierre " +
    "jean nicolas antoine julien romain kevin quentin yann zoe victor mathilde charlotte pauline margaux clara sarah laura " +
    "elodie aurelie caroline isabelle nathalie valerie sandrine celine stephanie audrey emilie oceane lucie juliette eva " +
    "alexandre benjamin guillaume vincent sebastien francois olivier philippe laurent christophe david eric frederic " +
    "mathieu adrien baptiste clement florian dylan enzo axel leo gabriel adam mohamed karim yasmine nour sami amine rayan " +
    "fatima leila mehdi ines salma youssef omar hamza ali bilal sofia " +
    "james john robert michael daniel matthew joseph christopher andrew joshua ryan brandon tyler jacob ethan noah liam " +
    "olivia ava isabella mia emily abigail madison elizabeth grace hannah lily natalie samantha jessica jennifer ashley " +
    "wei jing xiaoming yuxuan zihan haoran lei chen yu hao jie ying xin yan hui lin min jun tao fang ping qiang qi " +
    "xiaoyu ruoxi siyu yichen zhiwei jiahui mengqi shuang tingting yiwen " +
    "ana lucia pablo javier carmen alejandro diego sergio carlos miguel jose antonio manuel francisco jorge raul alba " +
    "paula marta elena laia nuria cristina rocio " +
    "olena oksana dmytro andriy iryna oleksandr serhiy mykola taras bohdan yulia natalia tetiana kateryna svitlana " +
    "anastasia viktoria daria sofiia maksym artem denys " +
    "aarav vihaan arjun rohan rahul priya ananya diya isha neha pooja amit vikram sanjay " +
    "lukas jonas felix maximilian leon finn paul anna lena hannah laura julia " +
    "giulia francesca chiara alessandro lorenzo matteo andrea marco luca giovanni").Split(' ');

  static string[] Last = (
    "martin bernard dubois thomas robert richard petit durand leroy moreau simon laurent lefebvre michel garcia david " +
    "bertrand roux vincent fournier morel girard andre lefevre mercier dupont lambert bonnet francois martinez legrand " +
    "garnier faure rousseau blanc guerin muller henry roussel nicolas perrin morin mathieu clement gauthier dumont lopez " +
    "fontaine chevalier robin masson sanchez gerard nguyen boyer denis lemaire duval joly gautier roger roche roy noel " +
    "meyer lucas meunier jean perez marchand dufour blanchard marie barbier brun dumas brunet schmitt leroux colin fernandez " +
    "bernier poirel maitre jimenez konovalova neuvelt ouidir tian rodriguez strashnyi graff galezowski " +
    "smith johnson williams brown jones miller davis wilson anderson taylor moore jackson white harris thompson " +
    "wang li zhang liu chen yang huang zhao wu zhou xu sun ma zhu hu guo he gao lin luo zheng liang xie song tang " +
    "han feng deng cao peng zeng xiao tian dong pan yuan cai jiang yu du ye cheng wei su lu ding ren " +
    "gonzalez hernandez gomez diaz alvarez romero navarro torres dominguez vazquez ramos gil ruiz serrano molina " +
    "shevchenko kovalenko bondarenko tkachenko kravchenko boyko kovalchuk oliynyk lysenko melnyk savchenko rudenko " +
    "benali haddad mansour khalil amrani bouzid cherif saidi " +
    "sharma patel singh kumar gupta reddy iyer nair mehta joshi " +
    "schmidt schneider fischer weber wagner becker hoffmann koch richter klein wolf " +
    "rossi russo ferrari esposito bianchi romano colombo ricci marino greco bruno gallo conti").Split(' ');

  static string[] On = "b c d f g h j k l m n p r s t v z br ch cl cr dr fr gr kh pl pr sh st tr zh".Split(' ');
  static string[] Vo = "a e i o u a e i o ou ai ie ia".Split(' ');
  static string[] Co = " n r l s t m k nd rt ll nn ch ck".Split(' ');

  public FauxOutlook(int n, int cap, int seed) {
    Cap = cap;
    Random r = new Random(seed);
    HashSet<string> used = new HashSet<string>();
    string[] services = { "comptabilite", "iael", "bba2en.studies", "scolarite", "admissions", "career.center",
      "library", "it.support", "international", "alumni", "housing", "exams", "student.life", "accueil", "rh", "paie",
      "communication", "direction", "relations.entreprises", "stages" };
    foreach (string s in services) Add(s.Replace('.', ' ').ToUpper(), s + "@em-lyon.com", used, r);

    // Noms de famille : la liste réelle en tête, puis une traîne fabriquée,
    // d'autant plus longue que le carnet est grand.
    List<string> last = new List<string>(Last);
    HashSet<string> lastSet = new HashSet<string>(Last);
    int extra = Math.Max(200, n / 3);
    while (last.Count < Last.Length + extra) {
      int syl = 2 + r.Next(2);
      StringBuilder b = new StringBuilder();
      for (int i = 0; i < syl; i++) b.Append(On[r.Next(On.Length)]).Append(Vo[r.Next(Vo.Length)]).Append(i == syl - 1 ? Co[r.Next(Co.Length)] : "");
      string s = b.ToString().Trim();
      if (s.Length >= 4 && lastSet.Add(s)) last.Add(s);
    }
    string[] domains = { "em-lyon.com", "em-lyon.com", "em-lyon.com", "em-lyon.com", "em-lyon.com", "edu.em-lyon.com",
      "edu.em-lyon.com", "edu.em-lyon.com", "gmail.com", "hotmail.com", "deloitte.fr", "groupeonepoint.com", "uv.es" };

    while (All.Count < n) {
      string f = First[Zipf(r, First.Length, 0.9)], l = last[Zipf(r, last.Count, 0.75)], d = domains[r.Next(domains.Length)];
      int fmt = r.Next(10);
      string local = fmt < 6 ? f + "." + l : fmt < 8 ? f[0] + l : fmt < 9 ? l + "." + f : f + l;
      string mail = local + "@" + d;
      for (int k = 2; used.Contains(mail); k++) mail = local + k + "@" + d;
      string name = r.Next(3) > 0 ? Cap1(f) + " " + l.ToUpper() : l.ToUpper() + " " + Cap1(f);
      Add(name, mail, used, r);
    }
    // Fréquence de contact : quelques centaines de personnes avec qui on
    // échange vraiment (très inégalement), et le reste de l'annuaire, jamais
    // contacté, dans un ordre quelconque mais fixe.
    var order = All.OrderBy(c => r.Next()).ToList();
    for (int i = 0; i < order.Count; i++) order[i].W = i < 400 ? 1.0 / Math.Pow(i + 1, 1.1) : r.NextDouble() * 1e-4;
    foreach (var c in All) {
      HashSet<string> pre = new HashSet<string>();
      foreach (string t in c.Toks) for (int k = 1; k <= Math.Min(3, t.Length); k++) pre.Add(t.Substring(0, k));
      foreach (string p in pre) { List<Contact> li; if (!index.TryGetValue(p, out li)) index[p] = li = new List<Contact>(); li.Add(c); }
    }
  }

  static int Zipf(Random r, int n, double s) {
    // Tirage approché : rang ~ n^u avec une pente réglable, rapide et assez inégal.
    double u = r.NextDouble();
    int k = (int)Math.Floor(Math.Pow(n + 1.0, Math.Pow(u, 1.0 / (1.0 + s)))) - 1;
    return Math.Max(0, Math.Min(n - 1, k));
  }
  static string Cap1(string s) { return char.ToUpper(s[0]) + s.Substring(1); }
  void Add(string name, string mail, HashSet<string> used, Random r) {
    if (!used.Add(mail)) return;
    All.Add(new Contact { Name = name, Mail = mail, Toks = Planif.Tokens(name + " " + mail) });
  }

  public List<Contact> Ask(string q) {
    Queries++;
    List<string> qw = Planif.Tokens(q);
    if (qw.Count == 0) return new List<Contact>();
    // Le mot le plus sélectif sert de point d'entrée dans l'index.
    List<Contact> pool = null;
    foreach (string w in qw) {
      List<Contact> li; string k = w.Length > 3 ? w.Substring(0, 3) : w;
      if (!index.TryGetValue(k, out li)) { pool = new List<Contact>(); break; }
      if (pool == null || li.Count < pool.Count) pool = li;
    }
    var hits = pool.Where(c => qw.All(w => c.Toks.Any(t => t.StartsWith(w, StringComparison.Ordinal))))
                   .OrderByDescending(c => c.W).Take(Cap).ToList();
    if (hits.Count == 0) Empty++;
    return hits;
  }

  // Nombre de contacts qui répondent à q, sans plafond (pour le diagnostic).
  public int CountAll(string q) {
    List<string> qw = Planif.Tokens(q);
    return All.Count(c => qw.All(w => c.Toks.Any(t => t.StartsWith(w, StringComparison.Ordinal))));
  }
}
