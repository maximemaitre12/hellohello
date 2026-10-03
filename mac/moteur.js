// Moteur du passage, version Mac. Tout ce qui décide se trouve ici, en
// JavaScript ordinaire : quoi taper, quand lire, quoi garder, quoi écrire.
// L'hôte Mac (hote.js) ne fournit que des gestes simples : mettre une app
// devant, taper, effacer, lire ce qu'affiche l'accessibilité, écrire un
// fichier. Le même moteur tourne dans outils/tests/mac-passage.mjs contre un
// faux Outlook, sans Mac.
//
// Rien n'est jamais envoyé : le moteur ne tape que des lettres et des
// effacements dans le champ « À », jamais Entrée ni Tab.

(function (root) {
  "use strict";

  var Planif = root.Planif || (typeof require === "function" ? require("./planif.js") : null);

  // Secondes par recherche pour l'estimation, et façon de lire la liste.
  var SPEEDS = {
    fast:    { sec: 2.2, settle: 300,  poll: 250, tries: 8,  emptyTries: 3, stable: 1,
               text: "Rapide : lit la liste dès qu'elle répond à la recherche. Peut rater des résultats de l'annuaire, qui arrivent en second." },
    normal:  { sec: 3.3, settle: 450,  poll: 350, tries: 10, emptyTries: 4, stable: 2,
               text: "Normal : attend que la liste soit la même deux fois de suite. Le bon équilibre pour la plupart des recherches." },
    careful: { sec: 5.0, settle: 1200, poll: 450, tries: 14, emptyTries: 6, stable: 3,
               text: "Prudent : attend trois lectures identiques, annuaire compris. Le plus fiable, le plus lent." }
  };

  function defaults(desktop) {
    return {
      Source: "outlook", GmailAccount: "", Prefix: "",
      Keep: "", Exclude: "", Dedupe: true, Speed: "normal",
      Format: "txt", Folder: desktop || "", OpenAtEnd: false, CopyAtEnd: false
    };
  }

  function merge(base, saved) {
    var out = {};
    for (var k in base) out[k] = base[k];
    if (saved && typeof saved === "object") {
      for (var k2 in base) if (saved[k2] !== undefined && saved[k2] !== null && typeof saved[k2] === typeof base[k2]) out[k2] = saved[k2];
    }
    if (!SPEEDS[out.Speed]) out.Speed = "normal";
    if (["txt", "csv", "both"].indexOf(out.Format) < 0) out.Format = "txt";
    if (out.Source !== "gmail") out.Source = "outlook";
    // Une virgule ou un point-virgule ferait valider un destinataire.
    out.Prefix = String(out.Prefix).replace(/[;,]/g, "").trim();
    return out;
  }

  function sourceName(cfg) { return cfg.Source === "gmail" ? "Gmail" : "Outlook"; }

  function durationText(sec) {
    if (sec < 60) return "moins d'une minute";
    var m = Math.round(sec / 60);
    if (m < 60) return "environ " + m + " min";
    return "environ " + Math.floor(m / 60) + " h " + String(m % 60).padStart(2, "0");
  }

  function plural(n, one, many) { return n <= 1 ? n + " " + one : n + " " + many; }

  function elapsedText(sec) {
    var m = Math.floor(sec / 60);
    if (m < 1) return "moins d'une minute";
    if (m < 60) return m + " min";
    return Math.floor(m / 60) + " h " + String(m % 60).padStart(2, "0");
  }

  function planText(cfg) {
    var src = sourceName(cfg);
    var start = cfg.Prefix
      ? "Dans " + src + ", 26 recherches de départ (" + cfg.Prefix + "a à " + cfg.Prefix + "z)"
      : "Dans " + src + ", 26 recherches de départ (A à Z)";
    var fmt = cfg.Format === "csv" ? "CSV" : cfg.Format === "both" ? "texte + CSV" : "texte";
    var filters = [];
    if (cfg.Keep) filters.push("domaines filtrés");
    if (cfg.Exclude) filters.push("exclusions");
    var ftxt = filters.length ? " · " + filters.join(", ") : "";
    // La durée dépend de ce que la messagerie contient : on ne la connaît pas d'avance.
    return start + ", puis les pistes les plus rentables jusqu'à ce qu'il n'y ait plus rien à trouver" + ftxt + " · " + fmt +
      ", enregistré au fur et à mesure · de quelques minutes à plusieurs heures selon la taille de votre carnet, arrêt possible à tout moment";
  }

  // ------------------------------------------------------------ filtres

  function domainMatch(mail, entry) {
    var d = entry.trim().toLowerCase().replace(/^@+/, "");
    if (!d) return false;
    var dom = mail.split("@").pop();
    return dom === d || dom.endsWith("." + d);
  }

  function passesFilters(cfg, mail) {
    if (cfg.Keep) {
      var ok = cfg.Keep.split(",").some(function (e) { return domainMatch(mail, e); });
      if (!ok) return false;
    }
    if (cfg.Exclude) {
      var parts = cfg.Exclude.split(",");
      for (var i = 0; i < parts.length; i++) {
        var x = parts[i].trim().toLowerCase();
        if (!x) continue;
        if (x.indexOf("@") >= 0 && x[0] !== "@") { if (mail === x) return false; }
        else if (domainMatch(mail, x)) return false;
      }
    }
    return true;
  }

  // ---------------------------------------------------- lecture des lignes
  //
  // L'hôte rend les lignes de la liste de suggestions telles que
  // l'accessibilité les montre : pour chaque contact, les textes de sa ligne
  // (souvent le nom, puis l'adresse ; parfois « Nom - adresse » d'un bloc).
  // Gmail ajoute une ligne qui colle toutes les suggestions bout à bout : on
  // la reconnaît à ce qu'elle est exactement la suite des autres, sans espaces.

  var MAIL = /[A-Za-z0-9._%+'\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}/g;

  function mailsIn(s) { return String(s).match(MAIL) || []; }
  function compact(s) { return String(s).replace(/\s+/g, ""); }

  function parseRows(rows) {
    var cands = [], seenCompact = {};
    for (var i = 0; i < rows.length; i++) {
      var texts = rows[i].map(function (t) { return String(t).trim(); }).filter(function (t) { return t; });
      var all = texts.join(" ");
      var found = {};
      mailsIn(all).forEach(function (m) { found[m.toLowerCase()] = 1; });
      var keys = Object.keys(found);
      if (keys.length !== 1) continue;
      var mail = keys[0];
      var c = compact(all);
      if (seenCompact[c]) continue;
      seenCompact[c] = 1;
      var name = "";
      for (var j = 0; j < texts.length && !name; j++) {
        var t = texts[j];
        if (t.toLowerCase().indexOf(mail) >= 0) {
          // « Nom - adresse » ou « Nom adresse » d'un seul bloc.
          var before = t.substring(0, t.toLowerCase().indexOf(mail)).replace(/[\s\-–<(]+$/, "").trim();
          if (before && mailsIn(before).length === 0) name = before;
        } else if (mailsIn(t).length === 0) name = t;
      }
      cands.push({ name: name || mail, mail: mail, compact: c });
    }
    var out = [], seenMail = {};
    for (var k = 0; k < cands.length; k++) {
      if (cands.length > 1) {
        var others = cands.filter(function (_, idx) { return idx !== k; }).map(function (x) { return x.compact; }).join("");
        if (cands[k].compact === others) continue;
      }
      if (seenMail[cands[k].mail]) continue;
      seenMail[cands[k].mail] = 1;
      out.push({ name: cands[k].name, mail: cands[k].mail });
    }
    return out;
  }

  function csvField(v) { v = String(v); return /[;"\r\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; }
  function stamp(d) {
    function p(n) { return String(n).padStart(2, "0"); }
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes());
  }
  function joinPath(dir, name) { return dir.replace(/\/+$/, "") + "/" + name; }

  // -------------------------------------------------------------- moteur
  //
  // natif(op, args) rend une promesse : c'est l'hôte Mac, ou le faux dans les
  // tests. ui reçoit l'avancée du passage pour l'afficher.

  function Moteur(natif, ui, wait, journal) {
    this.natif = natif;
    this.ui = ui;
    this.wait = wait || function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
    this.stop = false;
    this.files = [];
    this.distinct = [];
    this.prevSig = "";
    // Chemin du journal de reprise, donné par l'hôte.
    this.journal = journal || "";
  }

  Moteur.SPEEDS = SPEEDS;
  Moteur.defaults = defaults;
  Moteur.merge = merge;
  Moteur.planText = planText;
  Moteur.durationText = durationText;
  Moteur.parseRows = parseRows;
  Moteur.passesFilters = passesFilters;
  Moteur.plural = plural;
  Moteur.sourceName = sourceName;

  Moteur.prototype.lire = async function () {
    return parseRows(await this.natif("lire", {}));
  };

  // Lit la réponse à q. Une liste n'est retenue que si elle répond bien à q :
  // au moins un contact dont un mot commence par la recherche, ou à défaut une
  // liste différente de celle d'avant. Sinon c'est encore l'ancienne liste qui
  // est affichée, et on attend. Selon le rythme, on attend aussi qu'elle ne
  // bouge plus d'une lecture à l'autre.
  Moteur.prototype.lireReponse = async function (q, sp) {
    var last = null, stable = 0, empty = 0, lastList = [];
    for (var t = 0; t < sp.tries; t++) {
      var list = await this.lire();
      if (list.length === 0) {
        empty++; stable = 0; last = null;
        if (empty >= sp.emptyTries) { this.prevSig = ""; return []; }
      } else {
        empty = 0;
        var sig = list.map(function (x) { return x.mail; }).join("|");
        var answers = list.some(function (x) { return Planif.matches(x.name + " " + x.mail, q); });
        if (answers || sig !== this.prevSig) {
          if (sig === last) stable++; else { stable = 1; last = sig; }
          lastList = list;
          if (stable >= sp.stable) { this.prevSig = sig; return list; }
        }
      }
      await this.wait(sp.poll);
    }
    // Jamais stable dans le temps imparti : on garde la dernière liste qui
    // répondait à q, faute de mieux.
    if (last) { this.prevSig = last; return lastList; }
    return [];
  };

  // Met la messagerie devant, sur un nouveau message, curseur dans « À ».
  // Rend null si c'est prêt, ou le texte qui dit ce qui manque.
  Moteur.prototype.preparer = async function (cfg) {
    var n = this.natif, src = sourceName(cfg);
    if (cfg.Source === "gmail") {
      var acct = cfg.GmailAccount;
      if (!/^(\d+|[^@\s\/?#]+@[^@\s\/?#]+)$/.test(acct)) acct = "0";
      await n("ouvrirUrl", { url: "https://mail.google.com/mail/u/" + encodeURIComponent(acct) + "/?view=cm&fs=1&tf=1" });
      var vu = false, ouvert = false;
      for (var i = 0; i < 40 && !ouvert; i++) {
        await this.wait(500);
        if (!(await n("gmailDevant", {}))) continue;
        vu = true;
        ouvert = (await n("focus", { verrouiller: true })).ok;
      }
      if (!vu) return "Le navigateur n'a pas affiché Gmail. Ouvrez Gmail une fois dans votre navigateur par défaut, connectez-vous, puis relancez. Rien n'a été tapé.";
      if (!ouvert) return "Gmail s'est ouvert, mais pas sur un nouveau message avec le curseur dans « À ». Vérifiez que vous êtes connecté à Gmail dans votre navigateur. Rien n'a été tapé.";
    } else {
      if (!(await n("outlookOuvert", {}))) return "Outlook n'est pas ouvert. Ouvrez-le, puis relancez.";
      await n("activerOutlook", {});
      var devant = false;
      for (var j = 0; j < 10 && !devant; j++) { await this.wait(300); devant = await n("auPremierPlan", {}); }
      if (!devant) return "Impossible de mettre Outlook au premier plan. Cliquez dans Outlook, puis relancez.";
      await n("nouveauMessage", {});
      var pret = false;
      for (var k = 0; k < 12 && !pret; k++) { await this.wait(500); pret = (await n("focus", { verrouiller: true })).ok; }
      if (!pret) return "Le nouveau message ne s'est pas ouvert avec le curseur dans le champ « À ». Rien n'a été tapé. Si le message est bien ouvert, utilisez « Diagnostic » dans les paramètres et envoyez le fichier produit.";
    }
    // Ce qui est déjà à l'écran (votre adresse dans « De », un message
    // ouvert...) n'est pas une suggestion : on le note pour l'ignorer.
    await n("instantane", {});
    return null;
  };

  // ------------------------------------------------ fichiers et reprise
  //
  // Les fichiers de résultats sont écrits au fil du passage : une adresse
  // trouvée est sur le disque aussitôt. Le journal garde chaque recherche et
  // ce qu'elle a montré, une ligne par recherche, au même format que la
  // version Windows :
  //   recherche <TAB> nom <US> adresse <US> gardée <RS> nom <US> ...
  // Sa première ligne décrit le passage : réglages et nom des fichiers.

  var US = String.fromCharCode(31), RS = String.fromCharCode(30);
  var HEAD_KEYS = ["Source", "GmailAccount", "Prefix", "Keep", "Exclude", "Dedupe", "Format"];

  function txtLine(r) { return r.query + " : " + r.mail + "    (" + r.name + ")"; }
  function csvLine(r) { return csvField(r.query) + ";" + r.rank + ";" + csvField(r.name) + ";" + csvField(r.mail); }

  function journalLine(q, list, keep) {
    return q + "\t" + list.map(function (x, i) { return x.name + US + x.mail + US + (keep[i] ? "1" : "0"); }).join(RS);
  }

  // Le journal lu, ou null. Une dernière ligne coupée est ignorée.
  function parseJournal(text) {
    if (!text) return null;
    var lines = String(text).replace(/^﻿/, "").split("\n");
    var head;
    try { head = JSON.parse(lines[0]); } catch (e) { return null; }
    var entries = [], kept = {};
    for (var i = 1; i < lines.length; i++) {
      var line = lines[i].replace(/\r$/, "");
      var tab = line.indexOf("\t");
      if (tab < 0) continue;
      var rest = line.substring(tab + 1), e = { q: line.substring(0, tab), names: [], mails: [], keep: [] }, ok = true;
      if (rest) {
        var items = rest.split(RS);
        for (var k = 0; k < items.length; k++) {
          var f = items[k].split(US);
          if (f.length !== 3) { ok = false; break; }
          e.names.push(f[0]); e.mails.push(f[1]); e.keep.push(f[2] === "1");
          if (f[2] === "1") kept[f[1]] = 1;
        }
      }
      if (!ok) break;
      entries.push(e);
    }
    return { head: head, entries: entries, kept: Object.keys(kept).length };
  }

  Moteur.journalLine = journalLine;
  Moteur.parseJournal = parseJournal;
  Moteur.elapsedText = elapsedText;

  // Le passage interrompu, s'il y en a un : de quoi proposer de le reprendre.
  Moteur.prototype.reprise = async function () {
    if (!this.journal) return null;
    var j = parseJournal(await this.natif("lireTexte", { chemin: this.journal }));
    if (!j || j.entries.length === 0) return null;
    var d = new Date(j.head.Started), when = "";
    if (!isNaN(d)) when = String(d.getDate()).padStart(2, "0") + "/" + String(d.getMonth() + 1).padStart(2, "0") + " à " +
      String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
    return {
      text: "Commencé le " + when + " : " + plural(j.entries.length, "recherche", "recherches") + ", " + plural(j.kept, "adresse", "adresses") +
        ", dans " + String(j.head.Base).split("/").pop() + ". Reprendre continue là où il s'était arrêté, dans le même fichier, sans retaper ce qui est fait.",
      head: j.head
    };
  };

  Moteur.prototype.oublierReprise = async function () {
    if (this.journal) await this.natif("supprimer", { chemin: this.journal });
  };

  Moteur.prototype.passage = async function (cfg, resume) {
    var n = this.natif, ui = this.ui, self = this;
    var journal = resume && this.journal ? parseJournal(await n("lireTexte", { chemin: this.journal })) : null;
    if (journal) {
      // Une reprise se fait avec les réglages du passage commencé, sinon les
      // recherches ne seraient plus les mêmes.
      HEAD_KEYS.forEach(function (k) { if (journal.head[k] !== undefined && journal.head[k] !== null) cfg[k] = journal.head[k]; });
      if (ui.reglages) ui.reglages(cfg);
    }
    var src = sourceName(cfg), gmail = cfg.Source === "gmail";
    var sp = SPEEDS[cfg.Speed];
    this.stop = false;
    this.files = []; this.distinct = []; this.prevSig = "";

    // Le planificateur choisit chaque recherche : les 26 lettres, puis les
    // pistes les plus rentables, où qu'elles soient dans l'alphabet.
    var plan = new Planif(cfg.Prefix);
    var prefixLen = cfg.Prefix.length;
    var chipState = {}, records = [], seen = {};

    function keepFrom(q, list, keep) {
      var out = [];
      for (var i = 0; i < list.length; i++) {
        if (!keep[i]) continue;
        if (cfg.Dedupe && seen[list[i].mail]) continue;
        seen[list[i].mail] = 1;
        var r = { query: q, rank: i + 1, name: list[i].name, mail: list[i].mail };
        records.push(r); out.push(r);
      }
      return out;
    }
    function markChip(q, count) {
      var root = q.substring(0, Math.min(q.length, prefixLen + 1));
      if (count > 0) chipState[root] = "ok"; else if (!chipState[root]) chipState[root] = "none";
      ui.chip(root, chipState[root]);
    }

    ui.debut();

    // Reprise : on rejoue le journal dans le planificateur, sans rien taper.
    // Il retrouve exactement l'état où il était, recherche par recherche.
    var replayed = [], base, startedText, elapsedBefore;
    if (journal) {
      var diverged = false;
      journal.entries.forEach(function (e) {
        var list = e.mails.map(function (m, i) { return { name: e.names[i], mail: m }; });
        var q = diverged ? null : plan.next();
        if (q !== e.q) {
          // Le planificateur a changé depuis : on garde les adresses, pas la
          // suite, et la recherche qu'il proposait reste à faire.
          if (q) plan.requeue(q);
          diverged = true;
          keepFrom(e.q, list, e.keep);
          return;
        }
        plan.report(q, e.names, e.mails, e.keep);
        keepFrom(q, list, e.keep);
        markChip(q, list.length);
        replayed.push(journalLine(q, list, e.keep));
      });
      base = journal.head.Base;
      startedText = journal.head.Started;
      elapsedBefore = Number(journal.head.Seconds) || 0;
    } else {
      var folder = (await n("existe", { chemin: cfg.Folder })) ? cfg.Folder : await n("bureau", {});
      base = joinPath(folder, "adresses-" + src.toLowerCase() + "-" + stamp(new Date()));
      startedText = new Date().toISOString();
      elapsedBefore = 0;
    }
    var foundBefore = records.length;
    records.slice(Math.max(0, records.length - 300)).forEach(function (r) { ui.row(r.query, r.mail, r.name); });

    var pb = await this.preparer(cfg);
    if (pb) { ui.fin(pb, "warn"); return; }

    // Tout est prêt : les fichiers et le journal s'ouvrent, et restent à jour
    // à chaque recherche. Une reprise réécrit les fichiers proprement.
    var wantTxt = cfg.Format === "txt" || cfg.Format === "both";
    var wantCsv = cfg.Format === "csv" || cfg.Format === "both";
    if (wantTxt) {
      await n("ecrire", { chemin: base + ".txt", texte: records.map(function (r) { return txtLine(r) + "\n"; }).join("") });
      this.files.push(base + ".txt");
    }
    if (wantCsv) {
      // Le BOM permet à Excel pour Mac de lire les accents.
      await n("ecrire", { chemin: base + ".csv", texte: "﻿Recherche;Rang;Nom;Adresse\r\n" + records.map(function (r) { return csvLine(r) + "\r\n"; }).join("") });
      this.files.push(base + ".csv");
    }
    var head = { Started: startedText, Base: base, Seconds: Math.round(elapsedBefore) };
    HEAD_KEYS.forEach(function (k) { head[k] = cfg[k]; });
    if (this.journal) await n("ecrire", { chemin: this.journal, texte: [JSON.stringify(head)].concat(replayed).map(function (l) { return l + "\n"; }).join("") });

    // Une erreur en route garde le journal, comme un arrêt : seul un passage
    // allé au bout l'efface.
    var arret = null, completed = false, prevLen = 0, runStart = Date.now();
    var elapsed = function () { return elapsedBefore + (Date.now() - runStart) / 1000; };

    try {
      for (;;) {
        if (this.stop) { arret = "arrêté à votre demande."; break; }
        var q = plan.next();
        if (q === null) break;
        if (!(await n("auPremierPlan", {}))) { arret = src + " n'est plus au premier plan (recherche « " + q + " »)."; break; }
        if (!(await n("focus", {})).ok) { arret = "le curseur a quitté le champ « À » (recherche « " + q + " »)."; break; }

        // Pour les recherches plus longues, c'est la pastille de la lettre de
        // départ qui s'allume.
        var root = q.substring(0, prefixLen + 1);
        ui.chip(root, "now"); ui.live(q.toUpperCase(), "Recherche en cours");
        await n("effacer", { n: prevLen + 2 });
        await this.wait(400);
        await n("taper", { texte: q });
        prevLen = q.length;
        await this.wait(sp.settle);

        var list = await this.lireReponse(q, sp);
        var keep = list.map(function (x) { return passesFilters(cfg, x.mail); });
        plan.report(q, list.map(function (x) { return x.name; }), list.map(function (x) { return x.mail; }), keep);
        if (this.journal) await n("ajouter", { chemin: this.journal, texte: journalLine(q, list, keep) + "\n" });

        var fresh = keepFrom(q, list, keep);
        if (fresh.length) {
          if (wantTxt) await n("ajouter", { chemin: base + ".txt", texte: fresh.map(function (r) { return txtLine(r) + "\n"; }).join("") });
          if (wantCsv) await n("ajouter", { chemin: base + ".csv", texte: fresh.map(function (r) { return csvLine(r) + "\r\n"; }).join("") });
        }
        fresh.forEach(function (r) { ui.row(r.query, r.mail, r.name); ui.live(null, r.mail); });
        markChip(q, list.length);
        if (fresh.length === 0) ui.live(null, list.length === 0 ? "Aucune suggestion" : "Rien de nouveau");

        // La fin n'est connue qu'à peu près : la barre suit les pistes encore
        // rentables, qui se découvrent en route.
        var done = plan.Done;
        ui.progress(done, done + Math.max(plan.Pending, 26 - done), elapsedText(elapsed()) + " · " + plural(records.length, "adresse", "adresses"));
      }
      completed = arret === null;
      if ((await n("focus", {})).ok) await n("effacer", { n: prevLen + 2 });
    } finally {
      var uniq = {};
      records.forEach(function (r) { uniq[r.mail] = 1; });
      this.distinct = Object.keys(uniq).sort();
      if (wantTxt) await n("ajouter", { chemin: base + ".txt", texte: ["", "Adresses distinctes : " + this.distinct.length].concat(this.distinct).join("\n") + "\n" });
      if (this.journal) {
        if (completed) await n("supprimer", { chemin: this.journal });
        else {
          // La durée écoulée sert à la reprise : on la note en tête du journal.
          var text = await n("lireTexte", { chemin: this.journal });
          var lines = String(text || "").split("\n");
          head.Seconds = Math.round(elapsed());
          lines[0] = JSON.stringify(head);
          await n("ecrire", { chemin: this.journal, texte: lines.join("\n") });
        }
      }
    }

    if (records.length > 0) {
      if (cfg.CopyAtEnd) await n("copier", { texte: this.distinct.join("\n") });
      if (cfg.OpenAtEnd) await this.ouvrirFichiers();
    } else {
      // Rien trouvé : pas de fichier vide qui traîne.
      for (var f = 0; f < this.files.length; f++) await n("supprimer", { chemin: this.files[f] });
      this.files = [];
    }

    var names = this.files.map(function (x) { return x.split("/").pop(); }).join(" et ");
    var copied = cfg.CopyAtEnd && records.length > 0 ? " Adresses copiées." : "";
    var msg, kind;
    if (arret) {
      msg = "Arrêté : " + arret + " Tout ce qui a été trouvé est enregistré" + (names ? " dans " + names : "") + "." + copied +
        " Vous pouvez reprendre le passage plus tard, là où il s'est arrêté.";
      kind = "warn";
    } else if (records.length === 0) {
      msg = "Terminé, mais aucune adresse ne passe vos filtres. Rien n'a été enregistré.";
      kind = "warn";
    } else {
      var extra = copied + " " + plan.Done + " recherches.";
      if (journal) extra += " Cette reprise a ajouté " + plural(records.length - foundBefore, "adresse", "adresses") + ".";
      extra += plan.StopReason === "saturation"
        ? " Arrêt automatique : les " + plan.Saturation + " dernières recherches n'apportaient plus d'adresse nouvelle."
        : " Toutes les pistes ont été explorées.";
      msg = "Terminé : " + plural(this.distinct.length, "adresse", "adresses") + ", enregistrées dans " + names + "." + extra +
        " Le brouillon reste ouvert dans " + src + ", champ « À » vide : vous pouvez le fermer" +
        (gmail ? ", et le supprimer des brouillons si Gmail l'y a gardé." : ".");
      kind = "ok";
    }
    ui.fin(msg, kind, { distinct: this.distinct.length, files: this.files.slice(), done: plan.Done });
  };

  Moteur.prototype.ouvrirFichiers = async function () {
    for (var i = 0; i < this.files.length; i++) await this.natif("ouvrir", { chemin: this.files[i] });
  };

  // Diagnostic : ouvre un nouveau message, relève tout ce que l'accessibilité
  // montre avant et après avoir tapé « a », puis efface. Le fichier produit
  // permet d'adapter l'app si la messagerie a changé ses noms.
  Moteur.prototype.diagnostic = async function (cfg) {
    var n = this.natif;
    this.ui.debut();
    var pb = await this.preparer(cfg);
    var avant = await n("diagnostic", { etape: "apres ouverture" + (pb ? " (" + pb + ")" : "") });
    var apres = null, lu = [];
    if (!pb) {
      await n("taper", { texte: "a" });
      await this.wait(2500);
      apres = await n("diagnostic", { etape: "apres avoir tape a" });
      lu = await this.lire();
      await n("effacer", { n: 3 });
    }
    var path = await n("diagnosticFichier", { lu: lu });
    var msg = "Diagnostic enregistré : " + path.split("/").pop() + " (sur le bureau). Envoyez ce fichier pour que l'app soit ajustée. " +
      (pb ? "Problème rencontré : " + pb : plural(lu.length, "suggestion lue", "suggestions lues") + " avec « a ».");
    this.ui.fin(msg, lu.length ? "ok" : "warn", { files: [path] });
    this.files = [path];
    return { avant: avant, apres: apres, lu: lu, path: path };
  };

  if (typeof module !== "undefined" && module.exports) module.exports = Moteur;
  else root.Moteur = Moteur;
})(typeof globalThis !== "undefined" ? globalThis : this);
