// Le planificateur de la version Mac (outils/mac/planif.js) doit taper
// exactement la même suite de recherches que celui de la version Windows, sur
// les mêmes faux carnets. Les traces Windows viennent de mac-trace.ps1.
//
//   node outils/tests/mac-planif.mjs <traces.json>

import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const Planif = require("../mac/planif.js");

const cases = JSON.parse(readFileSync(process.argv[2], "utf8").replace(/^﻿/, ""));
let failed = 0;
for (const c of cases) {
  const all = c.contacts.map(([name, mail, w]) => ({ name, mail, w: Number(w), toks: Planif.tokens(name + " " + mail) }));
  const ask = (q) => {
    const qw = Planif.tokens(q);
    // Tri stable par poids décroissant, comme OrderByDescending de .NET.
    return all.filter((x) => qw.every((w) => x.toks.some((t) => t.startsWith(w))))
      .map((x, i) => [x, i]).sort((a, b) => b[0].w - a[0].w || a[1] - b[1]).slice(0, c.cap).map((p) => p[0]);
  };
  const p = new Planif(c.prefix, c.depth, c.budget);
  const got = [];
  for (let q; (q = p.next()) !== null;) {
    got.push(q);
    const l = ask(q);
    p.report(q, l.map((x) => x.name), l.map((x) => x.mail), l.map((x) => !c.keepDomain || x.mail.endsWith("@" + c.keepDomain)));
  }
  let diff = -1;
  for (let i = 0; i < Math.max(got.length, c.queries.length); i++) if (got[i] !== c.queries[i]) { diff = i; break; }
  const label = `carnet ${c.size}, plafond ${c.cap}, ${c.budget} rech., prof. ${c.depth}, prefixe "${c.prefix}"`;
  if (diff < 0 && p.StopReason === c.stop) console.log(`OK    ${label} : ${got.length} recherches identiques, arret ${p.StopReason}`);
  else {
    failed++;
    console.log(`ECART ${label} : recherche ${diff}, Windows "${c.queries[diff]}" / Mac "${got[diff]}", arret ${c.stop} / ${p.StopReason}`);
  }
}
process.exit(failed ? 1 : 0);
