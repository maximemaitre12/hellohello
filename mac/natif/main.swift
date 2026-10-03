// Adresses Outlook pour Mac : l'application native.
//
// Elle ne décide rien : elle affiche l'interface (web/interface.html, avec le
// moteur et le planificateur, les mêmes que la version Windows) et exécute les
// gestes que le moteur lui demande, un par message :
//
//   { id, op, args }  ->  window.__natif(id, ok, valeur)
//
// Gestes : mettre Outlook (ou le navigateur de Gmail) devant, ouvrir un
// nouveau message, taper, effacer, lire les suggestions par l'accessibilité,
// lire et écrire des fichiers, copier. Elle ne tape jamais Entrée ni Tab :
// rien n'est jamais envoyé.
//
// Une seule autorisation : Accessibilité (pour lire les suggestions et taper).
// Les frappes passent par le système (CGEvent), pas par « System Events ».
//
// Variables d'environnement, pour les tests seulement (mac/natif/tester.sh) :
//   AO_CIBLE     identifiant de l'app où taper (par défaut Outlook)
//   AO_SUPPORT   dossier des réglages et du journal
//   AO_AUTOTEST  fichier où écrire le résultat de l'autotest, puis quitter
//   AO_PASSAGE   dossier où faire un passage complet pendant l'autotest

import Cocoa
import WebKit
import ApplicationServices

let env = ProcessInfo.processInfo.environment
let OUTLOOK = env["AO_CIBLE"] ?? "com.microsoft.Outlook"
let HOME = NSHomeDirectory()
let SUPPORT = env["AO_SUPPORT"] ?? HOME + "/Library/Application Support/AetherOutils"
let REGLAGES = SUPPORT + "/adresses-outlook.json"
let JOURNAL = SUPPORT + "/reprise.jsonl"

struct Echec: Error, CustomStringConvertible { let description: String }

// ------------------------------------------------------------ accessibilité

func ax(_ e: AXUIElement, _ a: String) -> AnyObject? {
  var v: AnyObject?
  return AXUIElementCopyAttributeValue(e, a as CFString, &v) == .success ? v : nil
}
func axElement(_ v: AnyObject?) -> AXUIElement? {
  guard let v = v, CFGetTypeID(v) == AXUIElementGetTypeID() else { return nil }
  return (v as! AXUIElement)
}
func axTexte(_ e: AXUIElement, _ a: String) -> String { (ax(e, a) as? String) ?? "" }
func enfants(_ e: AXUIElement) -> [AXUIElement] { (ax(e, kAXChildrenAttribute) as? [AXUIElement]) ?? [] }
func role(_ e: AXUIElement) -> String { axTexte(e, kAXRoleAttribute) }
func pidDe(_ e: AXUIElement) -> pid_t { var p: pid_t = 0; return AXUIElementGetPid(e, &p) == .success ? p : 0 }
func textesDe(_ e: AXUIElement) -> [String] {
  var out: [String] = []
  for a in [kAXValueAttribute, kAXTitleAttribute, kAXDescriptionAttribute] {
    let t = axTexte(e, a).trimmingCharacters(in: .whitespacesAndNewlines)
    if !t.isEmpty && !out.contains(t) { out.append(t) }
  }
  return out
}

let MAIL = try! NSRegularExpression(pattern: "[A-Za-z0-9._%+'\\-]+@[A-Za-z0-9\\-]+(?:\\.[A-Za-z0-9\\-]+)*\\.[A-Za-z]{2,}")
func nbMails(_ t: String) -> Int { MAIL.numberOfMatches(in: t, range: NSRange(t.startIndex..., in: t)) }
let LIGNE: Set<String> = ["AXRow", "AXMenuItem", "AXCell", "AXListItem"]

func fenetres(_ pid: pid_t) -> [AXUIElement] {
  let app = AXUIElementCreateApplication(pid)
  AXUIElementSetMessagingTimeout(app, 1.0)
  return (ax(app, kAXWindowsAttribute) as? [AXUIElement]) ?? []
}

// Les lignes de la liste de suggestions, telles que l'accessibilité les
// montre : pour chaque contact, les textes de sa ligne. Une ligne est la plus
// petite partie de l'écran qui contient une seule adresse, dont le parent en
// contient plusieurs (ou une ligne de liste, de menu ou de tableau). Le moteur
// fait le tri ensuite (moteur.js, parseRows).
func lignes(_ pid: pid_t, _ deja: Set<String>, budget: Int = 4000) -> [[String]] {
  var rows: [[String]] = [], vus = 0
  func walk(_ el: AXUIElement, _ prof: Int) -> (textes: [String], mails: Int, emis: Bool) {
    vus += 1
    if vus > budget || prof > 40 { return ([], 0, false) }
    let r = role(el)
    var textes = textesDe(el).filter { !deja.contains($0) }
    var mails = textes.reduce(0) { $0 + nbMails($1) }
    var seuls: [[String]] = [], emis = false
    for c in enfants(el) {
      let w = walk(c, prof + 1)
      textes += w.textes; mails += w.mails
      if w.emis { emis = true } else if w.mails == 1 { seuls.append(w.textes) }
    }
    if LIGNE.contains(r) && mails == 1 && !emis { rows.append(textes); return (textes, mails, true) }
    if mails >= 2 && !seuls.isEmpty { rows += seuls; emis = true }
    return (textes, mails, emis)
  }
  for w in fenetres(pid) { _ = walk(w, 0) }
  return rows
}

func arbre(_ pid: pid_t, budget: Int = 6000) -> [String] {
  var out: [String] = [], vus = 0
  func walk(_ el: AXUIElement, _ prof: Int) {
    vus += 1
    if vus > budget || prof > 40 { return }
    out.append(String(repeating: "  ", count: prof) + role(el) + " " + axTexte(el, kAXSubroleAttribute) + " | " + textesDe(el).joined(separator: " | "))
    for c in enfants(el) { walk(c, prof + 1) }
  }
  for w in fenetres(pid) { walk(w, 0) }
  return out
}

// ------------------------------------------------------------ clavier

let SOURCE = CGEventSource(stateID: .hidSystemState)
func touche(_ code: CGKeyCode, _ flags: CGEventFlags = []) {
  for bas in [true, false] {
    guard let e = CGEvent(keyboardEventSource: SOURCE, virtualKey: code, keyDown: bas) else { continue }
    e.flags = flags
    e.post(tap: .cghidEventTap)
  }
  usleep(12000)
}
// Le texte tel quel, quelle que soit la disposition du clavier.
func frapper(_ texte: String) {
  for u in texte.utf16 {
    var c = u
    for bas in [true, false] {
      guard let e = CGEvent(keyboardEventSource: SOURCE, virtualKey: 0, keyDown: bas) else { continue }
      e.flags = []
      e.keyboardSetUnicodeString(stringLength: 1, unicodeString: &c)
      e.post(tap: .cghidEventTap)
    }
    usleep(12000)
  }
}

// ------------------------------------------------------------ fichiers

func existe(_ p: String) -> Bool { FileManager.default.fileExists(atPath: p) }
func ecrire(_ p: String, _ t: String) throws {
  try FileManager.default.createDirectory(atPath: (p as NSString).deletingLastPathComponent, withIntermediateDirectories: true)
  try t.write(toFile: p, atomically: true, encoding: .utf8)
}
func ajouter(_ p: String, _ t: String) throws {
  if !existe(p) { try ecrire(p, t); return }
  guard let h = FileHandle(forWritingAtPath: p) else { throw Echec(description: "impossible d'écrire dans " + p) }
  h.seekToEndOfFile(); h.write(t.data(using: .utf8)!); h.closeFile()
}
func lireTexte(_ p: String) -> String? { existe(p) ? try? String(contentsOfFile: p, encoding: .utf8) : nil }
// Une valeur absente devient null pour l'interface.
func ouNull(_ x: String?) -> Any { x.map { $0 as Any } ?? NSNull() }

// ------------------------------------------------------------ application

final class Hote: NSObject, NSApplicationDelegate, NSWindowDelegate, WKScriptMessageHandler, WKNavigationDelegate {
  var fen: NSWindow!
  var web: WKWebView!
  var cadreNormal: NSRect?
  var activite: NSObjectProtocol?
  var cible: pid_t = 0          // processus où l'on tape : Outlook, ou le navigateur de Gmail
  var verrou: AXUIElement?      // le champ « À » repéré à l'ouverture du message
  var deja = Set<String>()      // textes déjà à l'écran avant la première frappe
  var diag: [String] = []

  func applicationDidFinishLaunching(_ n: Notification) {
    menus()
    let conf = WKWebViewConfiguration()
    conf.userContentController.add(self, name: "natif")
    fen = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 640, height: 880),
                   styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false)
    fen.title = "Adresses Outlook"
    fen.isReleasedWhenClosed = false
    fen.delegate = self
    fen.minSize = NSSize(width: 480, height: 520)
    web = WKWebView(frame: fen.contentView!.bounds, configuration: conf)
    web.autoresizingMask = [.width, .height]
    web.navigationDelegate = self
    fen.contentView!.addSubview(web)
    let res = Bundle.main.resourceURL!.appendingPathComponent("web")
    web.loadFileURL(res.appendingPathComponent("interface.html"), allowingReadAccessTo: res)
    fen.center()
    fen.makeKeyAndOrderFront(nil)
    NSApp.activate(ignoringOtherApps: true)
  }

  func applicationShouldTerminateAfterLastWindowClosed(_ s: NSApplication) -> Bool { true }

  // Menus : sans menu Édition, ⌘C et ⌘V ne marchent pas dans les champs.
  func menus() {
    let bar = NSMenu()
    func sous(_ titre: String, _ items: [(String, Selector, String)]) {
      let it = NSMenuItem(), m = NSMenu(title: titre)
      for (t, s, k) in items { m.addItem(withTitle: t, action: s, keyEquivalent: k) }
      it.submenu = m; bar.addItem(it)
    }
    sous("Adresses Outlook", [("Masquer Adresses Outlook", #selector(NSApplication.hide(_:)), "h"),
                              ("Quitter Adresses Outlook", #selector(NSApplication.terminate(_:)), "q")])
    sous("Édition", [("Annuler", Selector(("undo:")), "z"), ("Couper", #selector(NSText.cut(_:)), "x"),
                     ("Copier", #selector(NSText.copy(_:)), "c"), ("Coller", #selector(NSText.paste(_:)), "v"),
                     ("Tout sélectionner", #selector(NSText.selectAll(_:)), "a")])
    sous("Fenêtre", [("Réduire", #selector(NSWindow.performMiniaturize(_:)), "m")])
    NSApp.mainMenu = bar
  }

  // -------------------------------------------------------------- pont

  func userContentController(_ c: WKUserContentController, didReceive m: WKScriptMessage) {
    guard let s = m.body as? String, let d = s.data(using: .utf8),
          let o = (try? JSONSerialization.jsonObject(with: d)) as? [String: Any],
          let id = o["id"] as? Int, let op = o["op"] as? String else { return }
    let args = o["args"] as? [String: Any] ?? [:]
    do { repondre(id, true, try faire(op, args)) }
    catch { repondre(id, false, "\(error)") }
  }

  func repondre(_ id: Int, _ ok: Bool, _ v: Any?) {
    var tab: [Any] = [v ?? NSNull()]
    if !JSONSerialization.isValidJSONObject(tab) { tab = [NSNull()] }
    let data = (try? JSONSerialization.data(withJSONObject: tab)) ?? Data("[null]".utf8)
    let json = String(data: data, encoding: .utf8)!
    web.evaluateJavaScript("window.__natif(\(id),\(ok),\(json)[0])", completionHandler: nil)
  }

  func s(_ a: [String: Any], _ k: String) -> String { (a[k] as? String) ?? "" }

  func faire(_ op: String, _ a: [String: Any]) throws -> Any? {
    switch op {
    case "init":
      try? FileManager.default.createDirectory(atPath: SUPPORT, withIntermediateDirectories: true)
      return ["acces": AXIsProcessTrusted(), "bureau": HOME + "/Desktop", "reglages": ouNull(lireTexte(REGLAGES)), "journal": JOURNAL]
    case "acces":
      return AXIsProcessTrusted()
    case "demanderAcces":
      let opt = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary
      _ = AXIsProcessTrustedWithOptions(opt)
      NSWorkspace.shared.open(URL(string: "x-apple.systempreferences:com.apple.preference.security?Privacy_Accessibility")!)
      return true
    case "relancer":
      // macOS ne prend parfois en compte l'autorisation qu'au lancement suivant.
      let p = Process()
      p.executableURL = URL(fileURLWithPath: "/bin/sh")
      p.arguments = ["-c", "sleep 1; /usr/bin/open -n \"$0\"", Bundle.main.bundlePath]
      try p.run()
      DispatchQueue.main.asyncAfter(deadline: .now() + 0.2) { NSApp.terminate(nil) }
      return true
    case "sauverReglages": try ecrire(REGLAGES, s(a, "json")); return true
    case "choisirDossier":
      let p = NSOpenPanel()
      p.canChooseDirectories = true; p.canChooseFiles = false; p.allowsMultipleSelection = false; p.prompt = "Choisir"
      if existe(s(a, "actuel")) { p.directoryURL = URL(fileURLWithPath: s(a, "actuel")) }
      NSApp.activate(ignoringOtherApps: true)
      return ouNull(p.runModal() == .OK ? p.url?.path : nil)
    // Exporter : la fenêtre « Enregistrer sous » du Mac, puis le fichier.
    case "enregistrerSous":
      let p = NSSavePanel()
      p.nameFieldStringValue = s(a, "nom"); p.allowedFileTypes = ["txt"]; p.canCreateDirectories = true
      if existe(s(a, "dossier")) { p.directoryURL = URL(fileURLWithPath: s(a, "dossier")) }
      NSApp.activate(ignoringOtherApps: true)
      guard p.runModal() == .OK, let u = p.url else { return NSNull() }
      try ecrire(u.path, s(a, "texte"))
      return u.path
    case "compact": compact(a["on"] as? Bool ?? false); return true
    case "existe": return !s(a, "chemin").isEmpty && existe(s(a, "chemin"))
    case "bureau": return HOME + "/Desktop"
    case "ecrire": try ecrire(s(a, "chemin"), s(a, "texte")); return true
    case "ajouter": try ajouter(s(a, "chemin"), s(a, "texte")); return true
    case "lireTexte": return ouNull(lireTexte(s(a, "chemin")))
    case "supprimer": if existe(s(a, "chemin")) { try FileManager.default.removeItem(atPath: s(a, "chemin")) }; return true
    case "copier":
      NSPasteboard.general.clearContents(); NSPasteboard.general.setString(s(a, "texte"), forType: .string); return true
    case "ouvrir": return NSWorkspace.shared.open(URL(fileURLWithPath: s(a, "chemin")))
    case "ouvrirUrl":
      verrou = nil; deja = []
      return URL(string: s(a, "url")).map { NSWorkspace.shared.open($0) } ?? false

    case "outlookOuvert": return outlook() != nil
    case "activerOutlook":
      guard let o = outlook() else { return false }
      cible = o.processIdentifier; verrou = nil; deja = []
      o.unhide()
      return o.activate(options: [.activateIgnoringOtherApps])
    // Le navigateur est devant, sur Gmail : il devient la cible, et on lui
    // demande d'exposer sa page à l'accessibilité (Chrome ne le fait pas seul).
    case "gmailDevant":
      guard let f = NSWorkspace.shared.frontmostApplication, f.processIdentifier != getpid() else { return false }
      let app = AXUIElementCreateApplication(f.processIdentifier)
      guard let w = axElement(ax(app, kAXFocusedWindowAttribute)), axTexte(w, kAXTitleAttribute).contains("Gmail") else { return false }
      AXUIElementSetAttributeValue(app, "AXManualAccessibility" as CFString, kCFBooleanTrue)
      AXUIElementSetAttributeValue(app, "AXEnhancedUserInterface" as CFString, kCFBooleanTrue)
      cible = f.processIdentifier
      return true
    case "auPremierPlan": return cible != 0 && NSWorkspace.shared.frontmostApplication?.processIdentifier == cible
    case "nouveauMessage": touche(45, .maskCommand); return true   // ⌘N

    // Le curseur est-il dans « À » ? Avec verrouiller, le champ qui a le
    // curseur devient celui où l'on tape ; ensuite, on vérifie que c'est
    // toujours lui avant chaque frappe.
    case "focus":
      guard NSWorkspace.shared.frontmostApplication?.processIdentifier == cible, let f = champActif() else { return ["ok": false] }
      if a["verrouiller"] as? Bool == true { verrou = f; return ["ok": true] }
      return ["ok": verrou.map { memeChamp(f, $0) } ?? false]
    // Ce qui porte déjà une adresse dans le message ouvert (votre adresse dans
    // « De », un destinataire...) n'est pas une suggestion : on le note pour
    // l'ignorer. Seulement dans ce message, pour ne pas masquer un contact
    // affiché ailleurs dans la messagerie.
    case "instantane":
      deja = []
      let app = AXUIElementCreateApplication(cible)
      if let w = axElement(ax(app, kAXFocusedWindowAttribute)) {
        var vus = 0
        func walk(_ el: AXUIElement, _ prof: Int) {
          vus += 1
          if vus > 4000 || prof > 40 { return }
          for t in textesDe(el) where nbMails(t) > 0 { deja.insert(t) }
          for c in enfants(el) { walk(c, prof + 1) }
        }
        walk(w, 0)
      }
      return Array(deja)
    case "effacer":
      guard verrou != nil else { return false }
      for _ in 0..<max(0, a["n"] as? Int ?? 0) { touche(51) }
      return true
    case "taper":
      guard verrou != nil else { return false }
      // Rien qui puisse valider un destinataire ou envoyer.
      frapper(s(a, "texte").filter { !"\r\n\t;,".contains($0) })
      return true
    case "lire": return cible != 0 ? lignes(cible, deja) : []

    case "diagnostic":
      diag.append("=== \(s(a, "etape")) (\(Date()))")
      let f = axElement(ax(AXUIElementCreateSystemWide(), kAXFocusedUIElementAttribute))
      diag.append("curseur : " + (f.map { role($0) + " | " + textesDe($0).joined(separator: " | ") + " | processus \(pidDe($0))" } ?? "aucun")
                  + " ; cible \(cible) ; accès \(AXIsProcessTrusted())")
      if cible != 0 { diag += arbre(cible) }
      return true
    case "diagnosticFichier":
      diag.append("=== lignes lues par l'app")
      for x in (a["lu"] as? [[String: Any]]) ?? [] { diag.append("\(x["name"] ?? "") <\(x["mail"] ?? "")>") }
      let df = DateFormatter(); df.dateFormat = "yyyy-MM-dd-HHmm"
      let p = HOME + "/Desktop/diagnostic-adresses-" + df.string(from: Date()) + ".txt"
      try ecrire(p, diag.joined(separator: "\n") + "\n")
      diag = []
      return p

    case "autotestFin":
      if let out = env["AO_AUTOTEST"] {
        let d = (try? JSONSerialization.data(withJSONObject: a, options: [.prettyPrinted])) ?? Data()
        FileManager.default.createFile(atPath: out, contents: d)
      }
      DispatchQueue.main.asyncAfter(deadline: .now() + 0.3) { NSApp.terminate(nil) }
      return true
    default:
      throw Echec(description: "geste inconnu : " + op)
    }
  }

  func outlook() -> NSRunningApplication? { NSRunningApplication.runningApplications(withBundleIdentifier: OUTLOOK).first }

  // Le champ qui a le curseur, s'il est dans la messagerie et qu'on peut y taper.
  func champActif() -> AXUIElement? {
    guard let f = axElement(ax(AXUIElementCreateSystemWide(), kAXFocusedUIElementAttribute)) else { return nil }
    guard ["AXTextField", "AXTextArea", "AXComboBox", "AXSearchField"].contains(role(f)) else { return nil }
    return cible != 0 && pidDe(f) != cible ? nil : f
  }

  func memeChamp(_ a: AXUIElement, _ b: AXUIElement) -> Bool {
    if CFEqual(a, b) { return true }
    // Certaines messageries recréent le champ après chaque frappe : on accepte
    // le même rôle sous le même parent.
    guard let pa = axElement(ax(a, kAXParentAttribute)), let pb = axElement(ax(b, kAXParentAttribute)) else { return false }
    return CFEqual(pa, pb) && role(a) == role(b)
  }

  // Pendant le passage : un petit panneau en bas à droite, au-dessus des autres
  // fenêtres, qui ne prend pas le clavier. macOS ne doit pas endormir l'app.
  func compact(_ on: Bool) {
    if on {
      cadreNormal = fen.frame
      let v = (fen.screen ?? NSScreen.main!).visibleFrame
      fen.setFrame(NSRect(x: v.maxX - 428, y: v.minY + 8, width: 420, height: 196), display: true, animate: true)
      fen.level = .floating
      if activite == nil {
        activite = ProcessInfo.processInfo.beginActivity(options: [.userInitiated, .idleSystemSleepDisabled, .latencyCritical], reason: "Passage Adresses Outlook")
      }
    } else {
      if let c = cadreNormal { fen.setFrame(c, display: true, animate: true) }
      fen.level = .normal
      if let a = activite { ProcessInfo.processInfo.endActivity(a); activite = nil }
      NSApp.activate(ignoringOtherApps: true)
    }
  }

  // -------------------------------------------------------------- autotest
  //
  // Pour la construction automatique : vérifie l'interface, le pont et les
  // fichiers, et, si AO_PASSAGE est donné, fait un passage complet sur la
  // fausse messagerie (FauxOutlook). Écrit le résultat dans AO_AUTOTEST.

  func webView(_ w: WKWebView, didFinish n: WKNavigation!) {
    guard env["AO_AUTOTEST"] != nil else { return }
    let passage = env["AO_PASSAGE"].map { "\"\($0)\"" } ?? "null"
    let js = """
    (async function () {
      var r = {};
      try {
        await new Promise(function (ok) { setTimeout(ok, 1500); });
        r.moteur = typeof Moteur === "function" && typeof Planif === "function";
        r.init = await natif("init", {});
        r.vue = document.getElementById("full").hidden ? (document.getElementById("gate").hidden ? "aucune" : "autorisation") : "app";
        r.plan = document.getElementById("plan").textContent;
        var p = r.init.journal + ".test";
        await natif("ecrire", { chemin: p, texte: "un\\n" });
        await natif("ajouter", { chemin: p, texte: "deux\\n" });
        r.fichier = await natif("lireTexte", { chemin: p });
        await natif("supprimer", { chemin: p });
        r.efface = (await natif("lireTexte", { chemin: p })) === null;
        r.inconnu = await natif("pasUnGeste", {}).then(function () { return "accepte"; }, function (e) { return "refuse"; });
        var dossier = \(passage);
        if (dossier && r.init.acces) {
          cfg.Folder = dossier; cfg.Format = "both"; cfg.Speed = "fast"; cfg.Prefix = ""; cfg.Keep = ""; cfg.Exclude = "";
          var t0 = Date.now();
          await moteur.passage(cfg, false);
          r.passage = { secondes: Math.round((Date.now() - t0) / 1000), distinct: moteur.distinct.length, adresses: moteur.distinct,
                        fichiers: moteur.files, message: document.getElementById("notice").textContent };
        }
      } catch (e) { r.erreur = String(e && e.stack || e); }
      window.webkit.messageHandlers.natif.postMessage(JSON.stringify({ id: 0, op: "autotestFin", args: r }));
    })();
    """
    w.evaluateJavaScript(js, completionHandler: nil)
  }
}

AXUIElementSetMessagingTimeout(AXUIElementCreateSystemWide(), 1.0)
let app = NSApplication.shared
let hote = Hote()
app.delegate = hote
app.setActivationPolicy(.regular)
app.run()
