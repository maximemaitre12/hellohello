// Planificateur des recherches, repris à l'identique de la classe Planif
// d'adresses-outlook.ps1 (version Windows). Même raisonnement, mêmes nombres,
// même ordre de choix : outils/tests/mac-planif.mjs vérifie que les deux
// versions tapent exactement la même suite de recherches.
//
// Il ne connaît de la messagerie que ce qu'il en voit : la liste de
// suggestions, plafonnée à quelques lignes, où un contact apparaît dès qu'un mot
// de son nom ou de son adresse commence par la recherche.
//
// - Les 26 lettres passent d'abord : elles donnent la taille de la liste (le
//   plafond) et un premier vocabulaire.
// - Une recherche dont la liste est pleine cache des contacts : ses suites (a
//   devient aa, ab...) deviennent des pistes. Une liste moins que pleine est
//   complète : rien à creuser dessous.
// - Les pistes passent par rendement attendu, sur tout l'alphabet à la fois.
//   Le calcul apprend en route ce que chaque genre de piste rapporte vraiment
//   en adresses nouvelles, ce qui écarte vite les recherches vides.
// - Pas de profondeur maximale : on creuse tant que la liste est pleine.
// - Un nom trop courant (martin, wang, jean) n'est pas départagé par une
//   lettre de plus. Quand les lettres ne rapportent plus, un deuxième mot
//   prend le relais (« martin j », « martin s »).
// - Le passage s'arrête seul quand il n'y a plus de piste rentable, ou quand
//   une longue série de recherches n'a plus rien apporté.

(function (root) {
  "use strict";

  var MIN_SCORE = 0.1, COLD = 0.01, MAX = Number.MAX_VALUE;

  function norm(s) {
    if (s == null) return "";
    return String(s).toLowerCase().normalize("NFD").replace(/\p{Mn}/gu, "");
  }

  function tokens(s) {
    var t = [], w = "";
    var n = norm(s) + " ";
    for (var i = 0; i < n.length; i++) {
      var c = n[i];
      if ((c >= "a" && c <= "z") || (c >= "0" && c <= "9")) w += c;
      else if (w.length > 0) { t.push(w); w = ""; }
    }
    return t;
  }

  function matchToks(toks, words) {
    for (var i = 0; i < words.length; i++) {
      var ok = false;
      for (var j = 0; j < toks.length; j++) if (toks[j].startsWith(words[i])) { ok = true; break; }
      if (!ok) return false;
    }
    return true;
  }

  // Le contact correspond-il à la recherche, à la façon de la messagerie ?
  // Chaque mot de la recherche doit commencer un mot du nom ou de l'adresse.
  function matches(text, query) { return matchToks(tokens(text), tokens(query)); }

  // Tas binaire, plus grande clé en tête.
  function Heap() { this.k = []; this.v = []; }
  Heap.prototype.count = function () { return this.k.length; };
  Heap.prototype.topKey = function () { return this.k[0]; };
  Heap.prototype.swap = function (a, b) {
    var t = this.k[a]; this.k[a] = this.k[b]; this.k[b] = t;
    var o = this.v[a]; this.v[a] = this.v[b]; this.v[b] = o;
  };
  Heap.prototype.push = function (key, val) {
    this.k.push(key); this.v.push(val);
    var i = this.k.length - 1;
    while (i > 0) { var p = (i - 1) >> 1; if (this.k[p] >= this.k[i]) break; this.swap(i, p); i = p; }
  };
  Heap.prototype.pop = function () {
    var top = this.v[0], last = this.k.length - 1;
    this.swap(0, last); this.k.pop(); this.v.pop();
    var i = 0, n = this.k.length;
    for (;;) {
      var l = 2 * i + 1, r = l + 1, m = i;
      if (l < n && this.k[l] > this.k[m]) m = l;
      if (r < n && this.k[r] > this.k[m]) m = r;
      if (m === i) break;
      this.swap(i, m); i = m;
    }
    return top;
  };

  function grid(rows, cols) { var g = []; for (var i = 0; i < rows; i++) g.push(new Array(cols).fill(0)); return g; }

  function Planif(prefix, maxDepth, budget) {
    this.Cap = 0;
    this.MaxDepth = maxDepth === undefined ? 12 : maxDepth;
    this.Budget = budget === undefined ? Infinity : budget;
    this.Done = 0; this.SinceNew = 0; this.Pending = 0; this.Saturation = 60;
    this.StopReason = "";

    this.roots = [];
    this.heap = new Heap();
    // Pistes qui ne valent presque rien pour l'instant : revues de loin en loin seulement.
    this.cold = [];
    this.nodes = new Map();
    // Pistes à un seul mot libre, par ce mot : pour savoir lesquelles passent le plafond de connus.
    this.byWord = new Map();
    this.promote = [];
    this.mails = new Set();
    this.contacts = [];
    this.index = new Map();
    this.big = grid(27, 26); this.bigRow = new Array(27).fill(0);
    this.start = grid(26, 26); this.startRow = new Array(26).fill(0);
    this.predicted = new Array(12).fill(0); this.obtained = new Array(12).fill(0);
    this.predNew = new Array(12).fill(0); this.obtNew = new Array(12).fill(0);
    this.lastRebuild = 0; this.lastDeep = 0;
    // Les recherches à deux mots ne viennent qu'une fois les lettres épuisées.
    this.splitPhase = false;
    // Branches qui viennent de rapporter : leurs suites sont à revoir à la hausse.
    this.bump = [];

    var p = prefix || "";
    var fw = tokens(p);
    if (fw.length > 0 && !p.endsWith(" ")) fw.pop();
    this.fixedWords = fw;
    for (var c = 0; c < 26; c++) {
      var n = this.make(p + String.fromCharCode(97 + c), null, false);
      n.Root = true;
      this.roots.push(n);
    }
  }

  Planif.MinScore = MIN_SCORE;
  Planif.norm = norm;
  Planif.tokens = tokens;
  Planif.matches = matches;

  Planif.prototype.make = function (q, parent, split) {
    var n = {
      Kids: null, Q: q, Words: tokens(q), Parent: parent, Depth: parent === null ? 1 : parent.Depth + 1, Split: split, Root: false,
      Est: 0, EstRaw: 0, Exact: 1.0, Stale: 0, ExpNew: 0, Kind: 0,
      Answered: false, Virtual: false, Count: 0, Novelty: 0, KCount: 0, KSeen: 0
    };
    this.nodes.set(q, n);
    if (n.Words.length > 1) { n.KCount = this.knownWords(n.Words); n.KSeen = this.contacts.length; }
    if (!split && (parent === null || !parent.Split)) this.byWord.set(n.Words[n.Words.length - 1], n);
    return n;
  };

  // Contacts déjà connus qui répondent à ces mots.
  Planif.prototype.knownWords = function (words) {
    if (words.length === 0) return this.contacts.length;
    var best = null;
    for (var i = 0; i < words.length; i++) {
      var l = this.index.get(words[i]);
      if (l === undefined) return 0;
      if (best === null || l.length < best.length) best = l;
    }
    if (words.length === 1) return best.length;
    var k = 0;
    for (var j = 0; j < best.length; j++) if (matchToks(this.contacts[best[j]], words)) k++;
    return k;
  };

  // Même chose pour une piste, en ne regardant que les contacts appris depuis
  // la dernière fois (les numéros de contact ne font que croître).
  Planif.prototype.known = function (n) {
    if (n.Words.length === 1) return this.knownWords(n.Words);
    // Si l'index d'un de ses mots est plus court que ce qui reste à lire, on recompte par lui.
    var smallest = Infinity;
    for (var i = 0; i < n.Words.length; i++) {
      var l = this.index.get(n.Words[i]);
      smallest = Math.min(smallest, l === undefined ? 0 : l.length);
    }
    if (smallest < this.contacts.length - n.KSeen) { n.KCount = this.knownWords(n.Words); n.KSeen = this.contacts.length; return n.KCount; }
    for (; n.KSeen < this.contacts.length; n.KSeen++) if (matchToks(this.contacts[n.KSeen], n.Words)) n.KCount++;
    return n.KCount;
  };

  Planif.prototype.learn = function (name, mail) {
    var toks = tokens(name + " " + mail);
    var id = this.contacts.length;
    this.contacts.push(toks);
    var prefixes = new Set();
    for (var a = 0; a < toks.length; a++) {
      var t = toks[a], prev = 26;
      for (var i = 0; i < t.length; i++) {
        var c = t.charCodeAt(i) - 97;
        if (c < 0 || c > 25) { prev = -1; continue; }
        if (prev >= 0) { this.big[prev][c]++; this.bigRow[prev]++; }
        if (i === 1 && prev >= 0 && prev < 26) { this.start[prev][c]++; this.startRow[prev]++; }
        prev = c;
      }
      for (var l = 1; l <= t.length; l++) prefixes.add(t.substring(0, l));
    }
    var inScope = matchToks(toks, this.fixedWords);
    var self = this;
    prefixes.forEach(function (p) {
      var li = self.index.get(p);
      if (li === undefined) { li = []; self.index.set(p, li); }
      li.push(id);
      if (!inScope) return;
      var n = self.byWord.get(p);
      if (n === undefined) return;
      if (self.Cap >= 3 && !n.Answered && n.Parent !== null && self.known(n) === self.Cap) self.promote.push(n);
      // Une branche qui rapporte rend ses suites plus prometteuses.
      if (n.Kids !== null) self.bump.push(n);
    });
  };

  // Probabilité qu'un mot qui commence par lw[..-1] continue par sa dernière lettre.
  Planif.prototype.continuation = function (lw) {
    var c = lw.charCodeAt(lw.length - 1) - 97;
    if (c < 0 || c > 25) return 0.03;
    var p = 26;
    if (lw.length >= 2) { p = lw.charCodeAt(lw.length - 2) - 97; if (p < 0 || p > 25) return 0.03; }
    var g = (this.big[p][c] + 0.2) / (this.bigRow[p] + 5.2);
    if (lw.length !== 2) return g;
    // Deuxième lettre d'un mot : soit la suite d'un prénom ou d'un nom, soit
    // une initiale collée au nom (gpoirel, tmaitre).
    var s = (this.start[p][c] + 0.2) / (this.startRow[p] + 5.2);
    var initial = (this.big[26][c] + 0.2) / (this.bigRow[26] + 5.2);
    return 0.45 * s + 0.35 * initial + 0.2 * g;
  };

  Planif.prototype.full = function (s) { return s.Virtual || (s.Answered && s.Count >= 3 && s.Count >= this.Cap); };

  Planif.prototype.totalEst = function (s) {
    if (!this.full(s)) return s.Count;
    var floor = s.Root ? this.Cap * 6.0 : this.Cap * 2.0;
    return Math.max(floor, Math.max(this.known(s) * 2.0, s.Est));
  };

  Planif.prototype.calibration = function (kind) { return (this.obtained[kind] + 3.0) / (this.predicted[kind] + 3.0); };
  // Rendement réel en adresses nouvelles, sur rendement prévu, par genre de
  // piste. Une liste pleine de contacts déjà connus est longue mais n'apporte
  // rien : c'est ce facteur qui l'apprend.
  Planif.prototype.yieldOf = function (kind) { return (this.obtNew[kind] + 2.0) / (this.predNew[kind] + 2.0); };

  // Adresses inconnues que la liste de n devrait montrer, pondérées par ce que
  // la branche a rapporté jusqu'ici. Une recherche dont on connaît déjà de quoi
  // remplir la liste passe en tête : on ne la tapera pas, on ouvrira ses suites.
  Planif.prototype.score = function (n) {
    var p = n.Parent;
    if (!this.full(p)) return 0;
    var k = this.known(n);
    if (this.Cap >= 3 && k >= this.Cap && n.Depth < this.MaxDepth) return MAX;
    var lw = n.Words[n.Words.length - 1];
    var share = (k + 3.0 * n.Exact * this.continuation(lw)) / (this.known(p) + 3.0);
    n.Kind = n.Split ? 8 + (k === 0 ? 1 : 0) : Math.min(n.Depth, 4) - 1 + (k === 0 ? 4 : 0);
    n.EstRaw = this.totalEst(p) * Math.min(1.0, share);
    n.Est = n.EstRaw * this.calibration(n.Kind);
    var expNew = Math.min(this.Cap, Math.max(n.Est, k)) - k;
    if (expNew <= 0) return 0;
    n.ExpNew = expNew;
    return expNew * (0.3 + p.Novelty) * this.yieldOf(n.Kind);
  };

  // Les suites d'une recherche pleine : une lettre de plus au dernier mot, et,
  // pour ceux qui portent exactement ce mot (martin, wang, jean), qu'aucune
  // lettre de plus ne départage, un deuxième mot.
  Planif.prototype.expand = function (p) {
    if (p.Depth >= this.MaxDepth) return;
    var lw = p.Words[p.Words.length - 1];
    for (var c = 0; c < 26; c++) this.push(p.Q + String.fromCharCode(97 + c), p, false, 1.0);
    var exact = this.exactShare(p.Words, lw);
    if (this.splitPhase && exact > 0) for (var d = 0; d < 26; d++) this.splitWord(p, String.fromCharCode(97 + d), exact);
  };

  // Part des contacts connus de la recherche qui n'ont que ce mot exact.
  Planif.prototype.exactShare = function (words, lw) {
    var l = this.index.get(lw);
    if (lw.length < 2 || l === undefined) return 0;
    var all = 0, ex = 0;
    for (var i = 0; i < l.length; i++) {
      var toks = this.contacts[l[i]];
      if (!matchToks(toks, words)) continue;
      all++;
      var longer = false, same = false;
      for (var j = 0; j < toks.length; j++) {
        if (toks[j] === lw) same = true;
        else if (toks[j].startsWith(lw)) longer = true;
      }
      if (same && !longer) ex++;
    }
    return all === 0 ? 0 : ex / all;
  };

  // Un nouveau mot qui commence un mot déjà dans la recherche ne trie rien
  // (« martin ma » répond comme « martin ») : on l'allonge jusqu'à ce qu'il
  // s'en distingue. Un mot qu'un autre commence déjà est redondant.
  Planif.prototype.splitWord = function (p, s, exact) {
    for (var i = 0; i < p.Words.length; i++) {
      var w = p.Words[i];
      if (w === s || s.startsWith(w)) return;
      if (w.startsWith(s)) {
        for (var c = 0; c < 26; c++) this.splitWord(p, s + String.fromCharCode(97 + c), exact);
        return;
      }
    }
    this.push(p.Q + " " + s, p, true, exact);
  };

  Planif.prototype.push = function (q, parent, split, exact) {
    if (this.nodes.has(q)) return;
    var n = this.make(q, parent, split);
    n.Exact = exact;
    if (parent.Kids === null) parent.Kids = [];
    parent.Kids.push(n);
    n.Stale = this.score(n);
    if (n.Stale < COLD) this.cold.push(n); else this.heap.push(n.Stale, n);
  };

  // Deuxième temps : les lettres ne rapportent plus. Les noms trop courants
  // pour être départagés par une lettre de plus le seront par un deuxième mot.
  Planif.prototype.openSplits = function () {
    this.splitPhase = true;
    var all = Array.from(this.nodes.values());
    for (var i = 0; i < all.length; i++) {
      var p = all[i];
      if (!this.full(p) || p.Depth >= this.MaxDepth) continue;
      var lw = p.Words[p.Words.length - 1];
      var exact = this.exactShare(p.Words, lw);
      if (exact > 0) for (var c = 0; c < 26; c++) this.splitWord(p, String.fromCharCode(97 + c), exact);
    }
    this.lastRebuild = this.Done;
  };

  // Les scores ne font en général que baisser (on connaît de plus en plus de
  // contacts), d'où le tas paresseux. Mais une recherche peut aussi passer le
  // seuil du plafond de connus, ou profiter d'un recalibrage : on recalcule
  // tout de temps en temps.
  Planif.prototype.rebuild = function (deep) {
    var h = new Heap();
    var c2 = deep ? [] : this.cold;
    var pending = 0;
    var all = this.heap.v.slice();
    if (deep) all = all.concat(this.cold);
    var once = new Set();
    for (var i = 0; i < all.length; i++) {
      var n = all[i];
      if (once.has(n)) continue;
      once.add(n);
      // Parent qui a tout montré, ou piste déjà traitée : elle ne servira plus.
      if (n.Answered || (n.Parent.Answered && !n.Parent.Virtual && !this.full(n.Parent) && n.Parent.Count < 3)) continue;
      n.Stale = this.score(n);
      if (n.Stale >= MIN_SCORE) pending++;
      if (n.Stale < COLD) c2.push(n); else h.push(n.Stale, n);
    }
    this.heap = h; this.cold = c2; if (deep) this.lastDeep = this.Done; this.Pending = pending; this.lastRebuild = this.Done;
  };

  Planif.prototype.next = function () {
    if (this.Done >= this.Budget) { this.StopReason = "budget"; return null; }
    if (this.roots.length > 0) return this.roots.shift().Q;
    this.Saturation = Math.max(60, Math.trunc(this.Done / 4));
    if (this.SinceNew >= this.Saturation) { this.StopReason = "saturation"; return null; }
    if (this.Done - this.lastRebuild >= 500) this.rebuild(this.Done - this.lastDeep >= 3000);
    var i, j;
    for (i = 0; i < this.promote.length; i++) if (!this.promote[i].Answered) this.heap.push(MAX, this.promote[i]);
    this.promote = [];
    for (i = 0; i < this.bump.length; i++) {
      var kids = this.bump[i].Kids;
      for (j = 0; j < kids.length; j++) {
        var kid = kids[j];
        if (kid.Answered) continue;
        var sc = this.score(kid);
        if (sc > kid.Stale + 1e-9) { kid.Stale = sc; this.heap.push(sc, kid); }
      }
    }
    this.bump = [];
    var rebuilt = false;
    for (;;) {
      if (this.heap.count() === 0 || this.heap.topKey() < MIN_SCORE) {
        // Avant de conclure, on s'assure que ce n'est pas un score périmé.
        // Un recalcul récent compte : sans cela, chaque fin de file en relancerait un.
        if (rebuilt || this.Done - this.lastRebuild < 50) {
          if (!this.splitPhase) { this.openSplits(); rebuilt = false; continue; }
          this.Pending = 0; this.StopReason = "epuise"; return null;
        }
        this.rebuild(this.Done - this.lastDeep >= 500); rebuilt = true; continue;
      }
      var n = this.heap.pop();
      if (n.Answered) continue;
      var s = this.score(n);
      if (this.heap.count() > 0 && s < this.heap.topKey() - 1e-9) { n.Stale = s; this.heap.push(s, n); continue; }
      if (s < MIN_SCORE) { n.Stale = s; this.heap.push(s, n); continue; }
      // Une recherche dont on connaît déjà de quoi remplir la liste ne
      // montrerait que du connu : on ne la tape pas, on passe à ses suites.
      if (s === MAX) {
        n.Virtual = true; n.Answered = true; n.Novelty = n.Parent.Novelty;
        this.expand(n);
        continue;
      }
      if (this.Pending > 0) this.Pending--;
      return n.Q;
    }
  };

  // Remet en file une recherche proposée mais pas tapée.
  Planif.prototype.requeue = function (q) {
    var n = this.nodes.get(q);
    if (n !== undefined && !n.Answered && n.Parent !== null) this.heap.push(n.Stale, n);
  };

  // Ce que la messagerie a montré pour q. keep[i] dit si l'adresse passe les
  // filtres. Rend le nombre d'adresses nouvelles et gardées.
  Planif.prototype.report = function (q, names, mailList, keep) {
    this.Done++;
    if (mailList.length > this.Cap) this.Cap = mailList.length;
    var fresh = 0, newAll = 0;
    for (var i = 0; i < mailList.length; i++) {
      var m = mailList[i].toLowerCase();
      if (!this.mails.has(m)) { this.mails.add(m); this.learn(names[i], m); newAll++; if (keep[i]) fresh++; }
    }
    this.SinceNew = fresh > 0 ? 0 : this.SinceNew + 1;
    var n = this.nodes.get(q);
    if (n === undefined) { n = this.make(q, null, false); n.Root = true; }
    if (n.Parent !== null) {
      this.predicted[n.Kind] += Math.min(this.Cap, n.EstRaw); this.obtained[n.Kind] += mailList.length;
      this.predNew[n.Kind] += n.ExpNew; this.obtNew[n.Kind] += newAll;
    }
    n.Answered = true; n.Count = mailList.length;
    n.Novelty = (fresh + 0.5) / (mailList.length + 1.0);
    if (mailList.length >= 3) this.expand(n);
    return fresh;
  };

  if (typeof module !== "undefined" && module.exports) module.exports = Planif;
  else root.Planif = Planif;
})(typeof globalThis !== "undefined" ? globalThis : this);
