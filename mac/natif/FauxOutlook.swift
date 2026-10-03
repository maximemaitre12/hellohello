// Fausse messagerie pour tester Adresses Outlook sur un Mac sans Outlook (la
// construction automatique de GitHub). ⌘N ouvre un nouveau message, curseur
// dans « À » ; chaque frappe filtre un carnet de contacts comme Outlook le
// fait (chaque mot tapé commence un mot du nom ou de l'adresse, les plus
// fréquents d'abord, 5 au plus), dans un tableau que l'accessibilité expose
// en lignes. Le message montre aussi « De : » avec une adresse, que l'app ne
// doit jamais noter.
//
// FO_CONTACTS : fichier où écrire le carnet (nom, adresse, poids), pour que le
// test sache ce qu'il y avait à trouver.

import Cocoa

struct Contact { let nom: String, mail: String, poids: Double, mots: [String] }

func mots(_ s: String) -> [String] {
  let n = s.folding(options: [.diacriticInsensitive, .caseInsensitive], locale: nil).lowercased()
  var out: [String] = [], w = ""
  for c in n + " " {
    if ("a"..."z").contains(c) || ("0"..."9").contains(c) { w.append(c) } else if !w.isEmpty { out.append(w); w = "" }
  }
  return out
}

// Un carnet fixe : les mêmes contacts à chaque construction.
func carnet() -> [Contact] {
  let prenoms = "jean marie pierre sophie lucas emma louis chloe hugo lea paul julie thomas camille nicolas sarah antoine laura maxime ines wei jing olena dmytro carlos lucia".split(separator: " ").map(String.init)
  let noms = "martin bernard dubois durand leroy moreau simon laurent michel garcia roux fournier morel girard mercier dupont lambert bonnet wang zhang shevchenko lopez rossi muller".split(separator: " ").map(String.init)
  let domaines = ["em-lyon.com", "em-lyon.com", "edu.em-lyon.com", "gmail.com", "deloitte.fr"]
  var graine: UInt64 = 7
  func hasard(_ n: Int) -> Int { graine = graine &* 6364136223846793005 &+ 1442695040888963407; return Int((graine >> 33) % UInt64(n)) }
  var out: [Contact] = [], vus = Set<String>()
  while out.count < 260 {
    let p = prenoms[hasard(prenoms.count)], n = noms[hasard(noms.count)], d = domaines[hasard(domaines.count)]
    var local = hasard(3) == 0 ? String(p.first!) + n : p + "." + n
    if vus.contains(local + "@" + d) { local += String(hasard(90) + 2) }
    let mail = local + "@" + d
    if !vus.insert(mail).inserted { continue }
    let nom = p.prefix(1).uppercased() + p.dropFirst() + " " + n.uppercased()
    out.append(Contact(nom: nom, mail: mail, poids: 1.0 / Double(out.count + 1), mots: mots(nom + " " + mail)))
  }
  return out
}

final class Fausse: NSObject, NSApplicationDelegate, NSTextFieldDelegate, NSTableViewDataSource, NSTableViewDelegate {
  let contacts = carnet()
  var montres: [Contact] = []
  var fen: NSWindow?
  var champ: NSTextField!
  var table: NSTableView!

  func applicationDidFinishLaunching(_ n: Notification) {
    let bar = NSMenu()
    let it = NSMenuItem(), m = NSMenu(title: "FauxOutlook")
    m.addItem(withTitle: "Quitter", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")
    it.submenu = m; bar.addItem(it)
    let it2 = NSMenuItem(), m2 = NSMenu(title: "Fichier")
    m2.addItem(withTitle: "Nouveau message", action: #selector(nouveau), keyEquivalent: "n").target = self
    it2.submenu = m2; bar.addItem(it2)
    NSApp.mainMenu = bar
    if let p = ProcessInfo.processInfo.environment["FO_CONTACTS"] {
      let l = contacts.map { ["nom": $0.nom, "mail": $0.mail, "poids": $0.poids] as [String: Any] }
      try? JSONSerialization.data(withJSONObject: l).write(to: URL(fileURLWithPath: p))
    }
    // Une fenêtre principale, comme la boîte de réception.
    let w = NSWindow(contentRect: NSRect(x: 80, y: 200, width: 500, height: 300), styleMask: [.titled], backing: .buffered, defer: false)
    w.title = "Boîte de réception"
    w.makeKeyAndOrderFront(nil)
    NSApp.activate(ignoringOtherApps: true)
  }

  @objc func nouveau() {
    let w = NSWindow(contentRect: NSRect(x: 120, y: 160, width: 560, height: 360), styleMask: [.titled, .closable], backing: .buffered, defer: false)
    w.title = "Nouveau message"
    let v = w.contentView!
    let de = NSTextField(labelWithString: "De : moi@aether.test")
    de.frame = NSRect(x: 20, y: 320, width: 500, height: 20); v.addSubview(de)
    let a = NSTextField(labelWithString: "À")
    a.frame = NSRect(x: 20, y: 288, width: 30, height: 22); v.addSubview(a)
    champ = NSTextField(frame: NSRect(x: 50, y: 286, width: 480, height: 24))
    champ.delegate = self
    champ.setAccessibilityLabel("À")
    v.addSubview(champ)
    table = NSTableView()
    for (id, larg) in [("nom", 220.0), ("mail", 250.0)] {
      let c = NSTableColumn(identifier: NSUserInterfaceItemIdentifier(id)); c.width = CGFloat(larg); table.addTableColumn(c)
    }
    table.headerView = nil
    table.dataSource = self; table.delegate = self
    let sv = NSScrollView(frame: NSRect(x: 50, y: 120, width: 480, height: 160))
    sv.documentView = table
    v.addSubview(sv)
    fen = w
    w.makeKeyAndOrderFront(nil)
    w.makeFirstResponder(champ)
  }

  // Comme Outlook : chaque mot tapé commence un mot du nom ou de l'adresse ;
  // les plus fréquents d'abord ; 5 au plus. La liste arrive un peu après la
  // frappe, comme une vraie recherche.
  func controlTextDidChange(_ n: Notification) {
    let q = mots(champ.stringValue)
    DispatchQueue.main.asyncAfter(deadline: .now() + 0.15) {
      self.montres = q.isEmpty ? [] : Array(self.contacts.filter { c in q.allSatisfy { w in c.mots.contains { $0.hasPrefix(w) } } }
        .sorted { $0.poids > $1.poids }.prefix(5))
      self.table.reloadData()
    }
  }

  func numberOfRows(in t: NSTableView) -> Int { montres.count }
  func tableView(_ t: NSTableView, viewFor col: NSTableColumn?, row: Int) -> NSView? {
    let c = montres[row]
    let f = NSTextField(labelWithString: col?.identifier.rawValue == "nom" ? c.nom : c.mail)
    return f
  }
}

let app = NSApplication.shared
let fausse = Fausse()
app.delegate = fausse
app.setActivationPolicy(.regular)
app.run()
