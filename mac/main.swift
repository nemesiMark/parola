// Word per macOS — finestra nativa che contiene l'editor (cartella "web" nelle risorse dell'app).
import Cocoa
import WebKit
import UniformTypeIdentifiers

final class AppDelegate: NSObject, NSApplicationDelegate, NSWindowDelegate,
                         WKNavigationDelegate, WKUIDelegate, WKScriptMessageHandler {
    var window: NSWindow!
    var webView: WKWebView!
    var pageLoaded = false
    var pendingFiles: [URL] = []

    // Colore della barra in alto dell'editor (chiaro/scuro), così la barra del titolo si fonde con la pagina
    let chromeColor = NSColor(name: nil) { appearance in
        appearance.bestMatch(from: [.darkAqua, .aqua]) == .darkAqua
            ? NSColor(srgbRed: 0x15 / 255.0, green: 0x18 / 255.0, blue: 0x1e / 255.0, alpha: 1)
            : NSColor(srgbRed: 0xf5 / 255.0, green: 0xf6 / 255.0, blue: 0xfa / 255.0, alpha: 1)
    }

    // MARK: Avvio

    func applicationDidFinishLaunching(_ note: Notification) {
        buildMenu()

        let config = WKWebViewConfiguration()
        config.websiteDataStore = .default()
        if #available(macOS 12.3, *) { config.preferences.isElementFullscreenEnabled = true }
        config.preferences.javaScriptCanOpenWindowsAutomatically = false
        config.userContentController.add(self, name: "word")

        webView = WKWebView(frame: .zero, configuration: config)
        webView.navigationDelegate = self
        webView.uiDelegate = self
        webView.allowsMagnification = false
        webView.setValue(false, forKey: "drawsBackground")

        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 1400, height: 900),
                          styleMask: [.titled, .closable, .miniaturizable, .resizable],
                          backing: .buffered, defer: false)
        window.title = "Word"
        window.titlebarAppearsTransparent = true
        window.backgroundColor = chromeColor
        window.minSize = NSSize(width: 760, height: 520)
        window.collectionBehavior = [.fullScreenPrimary]
        window.contentView = webView
        window.delegate = self
        window.setFrameAutosaveName("WordFinestra")
        if !window.setFrameUsingName("WordFinestra") { window.center() }
        window.makeKeyAndOrderFront(nil)

        let web = Bundle.main.resourceURL!.appendingPathComponent("web")
        webView.loadFileURL(web.appendingPathComponent("index.html"), allowingReadAccessTo: web)
        NSApp.activate(ignoringOtherApps: true)
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { true }

    // Documenti aperti dal Finder (doppio clic o «Apri con»)
    func application(_ application: NSApplication, open urls: [URL]) {
        if pageLoaded { urls.forEach(sendFileToPage) } else { pendingFiles.append(contentsOf: urls) }
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        pageLoaded = true
        // la pagina carica le librerie con "defer": aspetto un attimo prima di passarle i file
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.6) {
            self.pendingFiles.forEach(self.sendFileToPage)
            self.pendingFiles.removeAll()
        }
    }

    func sendFileToPage(_ url: URL) {
        guard let data = try? Data(contentsOf: url) else { return }
        let name = jsString(url.lastPathComponent)
        webView.evaluateJavaScript("window.word && window.word.openBase64(\(name), \"\(data.base64EncodedString())\")")
        NSDocumentController.shared.noteNewRecentDocumentURL(url)
    }

    // MARK: Messaggi dalla pagina

    func userContentController(_ uc: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let body = message.body as? [String: Any], let type = body["type"] as? String else { return }
        switch type {
        case "save":
            guard let name = body["name"] as? String, let b64 = body["data"] as? String,
                  let data = Data(base64Encoded: b64) else { return }
            saveFile(data: data, suggestedName: name)
        case "print":
            printDocument(title: body["title"] as? String ?? "Documento")
        case "title":
            let t = (body["title"] as? String ?? "").trimmingCharacters(in: .whitespaces)
            window.title = t.isEmpty ? "Word" : t
        case "theme":
            switch body["theme"] as? String {
            case "dark": window.appearance = NSAppearance(named: .darkAqua)
            case "light": window.appearance = NSAppearance(named: .aqua)
            default: window.appearance = nil
            }
        default: break
        }
    }

    func saveFile(data: Data, suggestedName: String) {
        let panel = NSSavePanel()
        panel.nameFieldStringValue = suggestedName
        let ext = (suggestedName as NSString).pathExtension
        if !ext.isEmpty, let type = UTType(filenameExtension: ext) { panel.allowedContentTypes = [type] }
        panel.canCreateDirectories = true
        panel.isExtensionHidden = false
        panel.beginSheetModal(for: window) { resp in
            guard resp == .OK, let url = panel.url else { return }
            do {
                try data.write(to: url, options: .atomic)
                NSDocumentController.shared.noteNewRecentDocumentURL(url)
                self.webView.evaluateJavaScript("window.word && window.word.saved(\(self.jsString(url.lastPathComponent)))")
            } catch {
                self.toast("Errore durante il salvataggio: \(error.localizedDescription)")
            }
        }
    }

    func printDocument(title: String) {
        let info = NSPrintInfo.shared.copy() as! NSPrintInfo
        info.jobDisposition = .spool
        info.horizontalPagination = .fit
        info.verticalPagination = .automatic
        info.topMargin = 0; info.bottomMargin = 0; info.leftMargin = 0; info.rightMargin = 0
        let op = webView.printOperation(with: info)
        op.jobTitle = title
        op.showsPrintPanel = true
        op.showsProgressPanel = true
        op.view?.frame = webView.bounds
        op.runModal(for: window, delegate: nil, didRun: nil, contextInfo: nil)
    }

    func toast(_ text: String) {
        webView.evaluateJavaScript("window.word && window.word.toast(\(jsString(text)))")
    }

    func jsString(_ s: String) -> String {
        let data = try! JSONSerialization.data(withJSONObject: [s])
        let arr = String(data: data, encoding: .utf8)!
        return String(arr.dropFirst().dropLast())
    }

    // MARK: Finestre di dialogo richieste dalla pagina

    func webView(_ webView: WKWebView, runOpenPanelWith parameters: WKOpenPanelParameters,
                 initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping ([URL]?) -> Void) {
        let panel = NSOpenPanel()
        panel.allowsMultipleSelection = parameters.allowsMultipleSelection
        panel.canChooseDirectories = false
        panel.beginSheetModal(for: window) { resp in
            if resp == .OK { panel.urls.forEach { NSDocumentController.shared.noteNewRecentDocumentURL($0) } }
            completionHandler(resp == .OK ? panel.urls : nil)
        }
    }

    func webView(_ webView: WKWebView, runJavaScriptAlertPanelWithMessage message: String,
                 initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping () -> Void) {
        let a = NSAlert(); a.messageText = message; a.addButton(withTitle: "OK")
        a.beginSheetModal(for: window) { _ in completionHandler() }
    }

    func webView(_ webView: WKWebView, runJavaScriptConfirmPanelWithMessage message: String,
                 initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping (Bool) -> Void) {
        let a = NSAlert(); a.messageText = message
        a.addButton(withTitle: "OK"); a.addButton(withTitle: "Annulla")
        a.beginSheetModal(for: window) { r in completionHandler(r == .alertFirstButtonReturn) }
    }

    func webView(_ webView: WKWebView, runJavaScriptTextInputPanelWithPrompt prompt: String, defaultText: String?,
                 initiatedByFrame frame: WKFrameInfo, completionHandler: @escaping (String?) -> Void) {
        let a = NSAlert(); a.messageText = prompt
        a.addButton(withTitle: "OK"); a.addButton(withTitle: "Annulla")
        let field = NSTextField(frame: NSRect(x: 0, y: 0, width: 320, height: 24))
        field.stringValue = defaultText ?? ""
        a.accessoryView = field
        a.window.initialFirstResponder = field
        a.beginSheetModal(for: window) { r in completionHandler(r == .alertFirstButtonReturn ? field.stringValue : nil) }
    }

    // I collegamenti esterni si aprono nel browser
    func webView(_ webView: WKWebView, decidePolicyFor action: WKNavigationAction,
                 decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        if let url = action.request.url, let scheme = url.scheme?.lowercased(),
           ["http", "https", "mailto", "tel"].contains(scheme) {
            NSWorkspace.shared.open(url)
            decisionHandler(.cancel); return
        }
        decisionHandler(.allow)
    }
    func webView(_ webView: WKWebView, createWebViewWith configuration: WKWebViewConfiguration,
                 for action: WKNavigationAction, windowFeatures: WKWindowFeatures) -> WKWebView? {
        if let url = action.request.url { NSWorkspace.shared.open(url) }
        return nil
    }

    // MARK: Menu

    @objc func menuCommand(_ sender: NSMenuItem) {
        guard let c = sender.representedObject as? String else { return }
        webView.evaluateJavaScript("window.word && window.word.command('\(c)')")
    }

    func buildMenu() {
        let main = NSMenu()
        func item(_ title: String, _ action: Selector?, _ key: String = "", _ mods: NSEvent.ModifierFlags = .command) -> NSMenuItem {
            let i = NSMenuItem(title: title, action: action, keyEquivalent: key)
            i.keyEquivalentModifierMask = mods
            return i
        }
        // voce che passa un comando alla pagina
        func cmd(_ title: String, _ command: String, _ key: String = "", _ mods: NSEvent.ModifierFlags = .command) -> NSMenuItem {
            let i = item(title, #selector(menuCommand(_:)), key, mods)
            i.target = self
            i.representedObject = command
            return i
        }
        func submenu(_ title: String, _ items: [NSMenuItem], into parent: NSMenu? = nil) -> NSMenuItem {
            let top = NSMenuItem(title: title, action: nil, keyEquivalent: ""); let m = NSMenu(title: title)
            items.forEach { m.addItem($0) }
            top.submenu = m
            (parent ?? main).addItem(top)
            return top
        }

        _ = submenu("Word", [
            item("Informazioni su Word", #selector(NSApplication.orderFrontStandardAboutPanel(_:))),
            .separator(),
            item("Nascondi Word", #selector(NSApplication.hide(_:)), "h"),
            item("Nascondi altre", #selector(NSApplication.hideOtherApplications(_:)), "h", [.command, .option]),
            item("Mostra tutte", #selector(NSApplication.unhideAllApplications(_:))),
            .separator(),
            item("Esci da Word", #selector(NSApplication.terminate(_:)), "q"),
        ])
        let exportMenu = NSMenu(title: "Esporta")
        [cmd("Pagina web (.html)…", "html"), cmd("Markdown (.md)…", "md"), cmd("Testo semplice (.txt)…", "txt")].forEach { exportMenu.addItem($0) }
        let exportItem = NSMenuItem(title: "Esporta", action: nil, keyEquivalent: "")
        exportItem.submenu = exportMenu
        _ = submenu("File", [
            cmd("Nuovo", "new", "n"),
            cmd("Apri…", "open", "o"),
            .separator(),
            cmd("Salva .docx…", "save", "s"),
            exportItem,
            cmd("Stampa / PDF…", "print", "p"),
            .separator(),
            item("Chiudi finestra", #selector(NSWindow.performClose(_:)), "w"),
        ])
        _ = submenu("Modifica", [
            item("Annulla", Selector(("undo:")), "z"),
            item("Ripeti", Selector(("redo:")), "z", [.command, .shift]),
            .separator(),
            item("Taglia", #selector(NSText.cut(_:)), "x"),
            item("Copia", #selector(NSText.copy(_:)), "c"),
            item("Incolla", #selector(NSText.paste(_:)), "v"),
            item("Incolla e adatta stile", #selector(NSTextView.pasteAsPlainText(_:)), "v", [.command, .option, .shift]),
            item("Seleziona tutto", #selector(NSText.selectAll(_:)), "a"),
            .separator(),
            cmd("Trova…", "find", "f"),
            cmd("Sostituisci…", "replace", "h", [.command, .option]),
            .separator(),
            cmd("Cerca un comando…", "search", "j"),
        ])
        _ = submenu("Vista", [
            cmd("Struttura del documento", "nav"),
            cmd("Modalità concentrazione", "focus"),
            .separator(),
            cmd("Ingrandisci", "zoomIn", "+"),
            cmd("Riduci", "zoomOut", "-"),
            cmd("Dimensioni reali", "zoomReset", "0"),
            .separator(),
            item("Schermo intero", #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control]),
        ])
        _ = submenu("Strumenti", [
            cmd("Conteggio parole", "count", "g", [.command, .shift]),
        ])
        let win = submenu("Finestra", [
            item("Contrai", #selector(NSWindow.performMiniaturize(_:)), "m"),
            item("Zoom", #selector(NSWindow.performZoom(_:))),
        ])
        let help = submenu("Aiuto", [
            cmd("Scorciatoie da tastiera", "keys", "/"),
        ])
        NSApp.windowsMenu = win.submenu
        NSApp.helpMenu = help.submenu
        NSApp.mainMenu = main
    }
}

let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.setActivationPolicy(.regular)
app.run()
