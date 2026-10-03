// Fait tourner le vrai moteur de la version Mac (outils/mac/moteur.js) contre
// un faux Outlook et un faux Gmail, sans Mac. Les faux carnets viennent de
// mac-trace.ps1, avec la suite de recherches qu'y a tapée la version Windows.
//
// Les faux montrent ce que l'accessibilité du Mac rend vraiment : des lignes
// de textes, avec l'ancienne liste encore affichée à la première lecture qui
// suit chaque frappe, des lignes parasites, et pour Gmail la ligne qui colle
// toutes les suggestions bout à bout. Le moteur doit taper exactement la même
// suite de recherches que Windows et ne garder que les vraies adresses. Les
// fichiers doivent se remplir pendant le passage, et un passage arrêté puis
// repris doit finir exactement comme un passage d'une traite.
//
//   node outils/tests/mac-passage.mjs <traces.json>

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
globalThis.Planif = require("../mac/planif.js");
const Moteur = require("../mac/moteur.js");

const cases = JSON.parse(readFileSync(process.argv[2], "utf8").replace(/^﻿/, ""));
let failed = 0;
const check = (ok, label) => { console.log((ok ? "OK    " : "ECHEC ") + label); if (!ok) failed++; };

async function run(c, source, opts = {}) {
  const all = c.contacts.map(([name, mail, w]) => ({ name, mail, w: Number(w), toks: Moteur.parseRows ? globalThis.Planif.tokens(name + " " + mail) : [] }));
  const ask = (q) => {
    const qw = globalThis.Planif.tokens(q);
    return all.filter((x) => qw.every((w) => x.toks.some((t) => t.startsWith(w))))
      .map((x, i) => [x, i]).sort((a, b) => b[0].w - a[0].w || a[1] - b[1]).slice(0, c.cap).map((p) => p[0]);
  };

  let field = "", shown = [], stale = false, typedQueries = [], keysOutsideField = 0, focusLocked = false, midLines = -1;
  const written = opts.disk || {};
  const rowsFor = (list) => {
    if (source === "gmail") {
      const rows = list.map((x) => [x.name + " " + x.mail]);
      if (list.length) rows.push([list.map((x) => x.name + x.mail).join("")]);
      return rows;
    }
    // Outlook pour Mac : nom et adresse en deux textes, plus un en-tête.
    return [["Suggestions"], ...list.map((x) => [x.name, x.mail]), ["Rechercher dans l'annuaire"]];
  };
  const natif = async (op, a) => {
    switch (op) {
      case "outlookOuvert": return true;
      case "activerOutlook": return true;
      case "ouvrirUrl": return true;
      case "gmailDevant": return true;
      case "auPremierPlan": return true;
      case "nouveauMessage": return true;
      case "focus": if (a.verrouiller) focusLocked = true; return { ok: focusLocked };
      case "instantane": return true;
      case "effacer": field = field.slice(0, Math.max(0, field.length - a.n)); return true;
      case "taper":
        if (!focusLocked) keysOutsideField++;
        field += a.texte; typedQueries.push(field); stale = true;
        // À mi-course, ce qui est trouvé doit déjà être sur le disque.
        if (typedQueries.length === 60) {
          const csvNow = Object.entries(written).find(([k]) => k.endsWith(".csv"));
          midLines = csvNow ? csvNow[1].split("\r\n").length - 2 : 0;
        }
        if (opts.stopAt && typedQueries.length >= opts.stopAt) m.stop = true;
        return true;
      case "lire": {
        // Première lecture après une frappe : encore l'ancienne liste.
        if (stale) { stale = false; return rowsFor(shown); }
        shown = field ? ask(field) : [];
        // Une ligne parasite à deux adresses, qu'il ne faut jamais garder.
        return [["De : moi@test.com, copie a autre@test.com"], ...rowsFor(shown)];
      }
      case "existe": return true;
      case "bureau": return "/Users/test/Desktop";
      case "ecrire": written[a.chemin] = a.texte; return true;
      case "ajouter": written[a.chemin] = (written[a.chemin] || "") + a.texte; return true;
      case "lireTexte": return written[a.chemin] === undefined ? null : written[a.chemin];
      case "supprimer": delete written[a.chemin]; return true;
      case "copier": return true;
      default: throw new Error("op inconnue " + op);
    }
  };
  let fin = null;
  const ui = { debut() {}, chip() {}, live() {}, row() {}, progress() {}, fin(msg, kind, info) { fin = { msg, kind, info }; } };
  const m = new Moteur(natif, ui, async () => {}, JOURNAL);
  const cfg = Moteur.merge(Moteur.defaults("/Users/test/Desktop"), {
    Source: source, Prefix: c.prefix, Keep: c.keepDomain, Format: "both"
  });
  await m.passage(cfg, !!opts.resume);
  if (opts.raw) return { typed: typedQueries, fin, written, m };

  const label = `${source}, carnet ${c.size}, plafond ${c.cap}, prefixe "${c.prefix}"`;
  const same = typedQueries.length === c.queries.length && typedQueries.every((q, i) => q === c.queries[i]);
  check(same, `${label} : ${typedQueries.length} recherches tapees, identiques a Windows`);
  const real = new Set(c.queries.flatMap((q) => ask(q).map((x) => x.mail)).filter((x) => !c.keepDomain || x.endsWith("@" + c.keepDomain)));
  const got = new Set(m.distinct);
  const bogus = [...got].filter((x) => !real.has(x));
  check(bogus.length === 0 && got.size === real.size, `${label} : ${got.size} adresses gardees sur ${real.size} attendues, ${bogus.length} parasite(s)`);
  const txt = Object.entries(written).find(([k]) => k.endsWith(".txt"));
  const csv = Object.entries(written).find(([k]) => k.endsWith(".csv"));
  // Le texte : rien que les adresses, une par ligne, chacune une fois.
  const tl = txt ? txt[1].split("\n").filter((l) => l) : [];
  check(!!txt && txt[0].startsWith("/Users/test/Desktop/adresses-" + source + "-") && tl.length === got.size &&
    new Set(tl).size === tl.length && tl.every((l) => got.has(l)),
    `${label} : fichier texte ${txt && txt[0].split("/").pop()}, ${tl.length} adresses seules`);
  // Au lancement suivant, Exporter et Copier retrouvent le dernier passage.
  const connues = await m.connues(cfg);
  check(connues.length === got.size, `${label} : ${connues.length} adresses exportables au lancement suivant`);
  check(!!csv && csv[1].startsWith("﻿Recherche;Rang;Nom;Adresse\r\n"), `${label} : CSV avec en-tete et BOM`);
  check(fin && fin.kind === "ok" && field === "", `${label} : fin "${fin && fin.msg.slice(0, 60)}...", champ A vide`);
  check(keysOutsideField === 0, `${label} : aucune frappe avant que le curseur soit dans A`);
  if (typedQueries.length > 60) check(midLines > 0, `${label} : ${midLines} adresses deja sur le disque apres 60 recherches`);
  check(written[JOURNAL] === undefined, `${label} : journal de reprise efface a la fin`);
}

const JOURNAL = "/Users/test/Library/Application Support/AetherOutils/reprise.jsonl";

// Arrêté en route puis repris : même suite de recherches, mêmes fichiers.
async function resume(c, stopAt) {
  const label = `reprise, carnet ${c.size}, plafond ${c.cap}, arret apres ${stopAt}`;
  const straight = await run(c, "outlook", { raw: true });
  const disk = {};
  const a = await run(c, "outlook", { raw: true, disk, stopAt });
  check(a.fin.kind === "warn" && a.fin.msg.startsWith("Arrêté") && disk[JOURNAL] !== undefined, `${label} : arret, journal garde`);
  const offer = await a.m.reprise();
  check(!!offer && offer.text.includes(plural(stopAt)), `${label} : reprise proposee (${offer && offer.text.slice(0, 60)}...)`);
  const b = await run(c, "outlook", { raw: true, disk, resume: true });
  const both = a.typed.concat(b.typed);
  check(both.length === straight.typed.length && both.every((q, i) => q === straight.typed[i]),
    `${label} : ${a.typed.length} + ${b.typed.length} recherches, sans en retaper une, comme d'une traite (${straight.typed.length})`);
  const csv = (w) => Object.entries(w).find(([k]) => k.endsWith(".csv"))[1];
  check(csv(disk) === csv(straight.written), `${label} : CSV final identique au passage d'une traite`);
  check(b.fin.kind === "ok" && b.fin.msg.includes("reprise a ajouté") && disk[JOURNAL] === undefined, `${label} : fin "${b.fin.msg.slice(0, 50)}...", journal efface`);
}
const plural = (n) => n + " recherches";

// Le parseur seul, sur les formes connues.
const P = Moteur.parseRows;
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
check(eq(P([["Jean Dupont", "jean.dupont@ex.fr"]]), [{ name: "Jean Dupont", mail: "jean.dupont@ex.fr" }]), "parse : nom et adresse separes");
check(eq(P([["Jean Dupont - Jean.Dupont@ex.fr"]]), [{ name: "Jean Dupont", mail: "jean.dupont@ex.fr" }]), "parse : « Nom - adresse » d'un bloc");
check(eq(P([["Jean Dupont jean@ex.fr"], ["Jean Dupontjean@ex.fr"]]), [{ name: "Jean Dupont", mail: "jean@ex.fr" }]), "parse : Gmail, une suggestion et sa ligne collee");
check(eq(P([["a@ex.fr a@ex.fr"]]), [{ name: "a@ex.fr", mail: "a@ex.fr" }]), "parse : contact sans nom");
check(eq(P([["Deux", "a@ex.fr", "b@ex.fr"]]), []), "parse : ligne a deux adresses ignoree");
check(eq(P([["Jean", "j@ex.fr"], ["Jean", "J@EX.FR"]]), [{ name: "Jean", mail: "j@ex.fr" }]), "parse : doublon dans la meme liste");

// Les cas sans limite de la trace Windows (les autres servent à mac-planif.mjs).
for (const i of [1, 2, 5, 6]) await run(cases[i], "outlook");
for (const i of [0]) await run(cases[i], "gmail");
await resume(cases[1], 150);
await resume(cases[2], 700);
process.exit(failed ? 1 : 0);
