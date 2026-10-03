#!/bin/bash
# Lance « Adresses Outlook.app » sur ce Mac et vérifie qu'elle marche, comme
# un utilisateur la lancerait (open). Pour la construction automatique.
#
#   bash mac/natif/tester.sh dist
#
# 1. L'app seule : fenêtre, interface, pont interface <-> système, fichiers.
# 2. Si l'autorisation Accessibilité a pu être donnée sur ce Mac de test : un
#    passage complet sur la fausse messagerie (FauxOutlook), avec de vraies
#    frappes et une vraie lecture par l'accessibilité, vérifié contre le carnet.
# Les résultats et des captures d'écran vont dans dist/tests.

set -uo pipefail
OUT="$(cd "${1:-dist}" && pwd)"
APP="$OUT/Adresses Outlook.app"
FAUX="$OUT/FauxOutlook.app"
T="$OUT/tests"
rm -rf "$T"; mkdir -p "$T/support" "$T/sortie"
ECHECS=0
echoue() { echo "ECHEC : $*"; ECHECS=$((ECHECS + 1)); }
attendre() {   # attendre <fichier> <secondes>
  for _ in $(seq 1 "$2"); do [ -s "$1" ] && return 0; sleep 1; done
  return 1
}

echo "== Autorisation Accessibilité sur ce Mac de test"
EXE="$APP/Contents/MacOS/Adresses Outlook"
for DB in "/Library/Application Support/com.apple.TCC/TCC.db" "$HOME/Library/Application Support/com.apple.TCC/TCC.db"; do
  for CLIENT in "com.aether.adresses-outlook|0" "$EXE|1"; do
    sudo sqlite3 "$DB" "INSERT OR REPLACE INTO access (service, client, client_type, auth_value, auth_reason, auth_version, flags) VALUES ('kTCCServiceAccessibility', '${CLIENT%|*}', ${CLIENT#*|}, 2, 4, 1, 0);" 2>&1 | sed 's/^/  /' || true
  done
done
sudo pkill -9 tccd 2>/dev/null || true
sleep 2

echo "== 1. L'app seule"
open -n --env AO_SUPPORT="$T/support" --env AO_AUTOTEST="$T/app.json" "$APP"
sleep 4; screencapture -x "$T/1-ouverture.png" || true
if attendre "$T/app.json" 60; then
  cat "$T/app.json"
  python3 - "$T/app.json" <<'PY' || ECHECS=$((ECHECS + 1))
import json, sys
r = json.load(open(sys.argv[1]))
ko = []
if r.get("erreur"): ko.append("erreur : " + r["erreur"])
if r.get("moteur") is not True: ko.append("moteur ou planificateur absent")
if r.get("vue") not in ("app", "autorisation"): ko.append("aucun écran affiché")
if r.get("fichier") != "un\ndeux\n": ko.append("écriture de fichier : %r" % r.get("fichier"))
if r.get("efface") is not True: ko.append("suppression de fichier")
if r.get("inconnu") != "refuse": ko.append("un geste inconnu n'est pas refusé")
if "jusqu'à ce qu'il n'y ait plus rien" not in (r.get("plan") or ""): ko.append("résumé du plan : %r" % r.get("plan"))
print("  écran : %s, accès accessibilité : %s" % (r.get("vue"), r.get("init", {}).get("acces")))
for k in ko: print("ECHEC : " + k)
sys.exit(1 if ko else 0)
PY
else
  echoue "l'app n'a pas répondu en 60 s (elle ne s'ouvre pas ?)"
fi
pkill -f "Adresses Outlook.app" 2>/dev/null || true
sleep 1

ACCES=$(python3 -c "import json;print(json.load(open('$T/app.json')).get('init',{}).get('acces'))" 2>/dev/null || echo False)
if [ "$ACCES" != "True" ]; then
  echo "== 2. Passage complet : non fait, ce Mac de test n'a pas pu donner l'autorisation Accessibilité."
  echo "PASSAGE=non-teste" > "$T/passage.txt"
else
  echo "== 2. Passage complet sur la fausse messagerie"
  open -n --env FO_CONTACTS="$T/contacts.json" "$FAUX"
  sleep 3
  open -n --env AO_SUPPORT="$T/support" --env AO_AUTOTEST="$T/passage.json" --env AO_PASSAGE="$T/sortie" \
       --env AO_CIBLE=com.aether.fauxoutlook "$APP"
  ( for i in 1 2 3 4 5 6; do sleep 20; screencapture -x "$T/2-passage-$i.png" || true; done ) &
  if attendre "$T/passage.json" 1500; then
    cat "$T/passage.json" | head -c 3000; echo
    python3 - "$T/passage.json" "$T/contacts.json" "$T/sortie" "$T/support" <<'PY' || ECHECS=$((ECHECS + 1))
import json, sys, os, glob
r = json.load(open(sys.argv[1])); contacts = json.load(open(sys.argv[2])); sortie, support = sys.argv[3], sys.argv[4]
ko = []
p = r.get("passage")
if r.get("erreur"): ko.append("erreur : " + r["erreur"])
if not p: ko.append("pas de passage")
else:
    vrais = {c["mail"] for c in contacts}
    trouves = set(p["adresses"])
    parasites = trouves - vrais
    part = 100.0 * len(trouves & vrais) / len(vrais)
    print("  %d adresses sur %d dans le carnet (%.0f %%), %d parasite(s), %d s" % (len(trouves & vrais), len(vrais), part, len(parasites), p["secondes"]))
    print("  message : " + p["message"])
    if parasites: ko.append("adresses qui ne sont pas des contacts : %s" % sorted(parasites)[:5])
    if "moi@aether.test" in trouves: ko.append("l'adresse de « De : » a été notée")
    if part < 70: ko.append("seulement %.0f %% du carnet" % part)
    if not p["message"].startswith("Terminé"): ko.append("le passage ne s'est pas terminé normalement")
    txt = glob.glob(os.path.join(sortie, "*.txt")); csv = glob.glob(os.path.join(sortie, "*.csv"))
    if not txt or not csv: ko.append("fichiers texte et CSV absents")
    else:
        lignes = open(csv[0], encoding="utf-8-sig").read().splitlines()
        if lignes[0] != "Recherche;Rang;Nom;Adresse" or len(lignes) - 1 != len(trouves): ko.append("CSV : %d lignes pour %d adresses" % (len(lignes) - 1, len(trouves)))
        if "Adresses distinctes : %d" % len(trouves) not in open(txt[0], encoding="utf-8").read(): ko.append("fin du fichier texte")
    if os.path.exists(os.path.join(support, "reprise.jsonl")): ko.append("journal de reprise non effacé")
for k in ko: print("ECHEC : " + k)
sys.exit(1 if ko else 0)
PY
    echo "PASSAGE=teste" > "$T/passage.txt"
  else
    screencapture -x "$T/2-bloque.png" || true
    echoue "le passage n'a pas fini en 25 min"
  fi
  pkill -f "Adresses Outlook.app" 2>/dev/null || true
  pkill -f "FauxOutlook.app" 2>/dev/null || true
fi

echo "== $ECHECS échec(s)"
exit $((ECHECS > 0 ? 1 : 0))
