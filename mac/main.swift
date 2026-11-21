// Parola per macOS — finestra nativa che contiene l'editor (cartella "web" nelle risorse dell'app).
import Cocoa
import WebKit
import UniformTypeIdentifiers

final class AppDelegate: NSObject, NSApplicationDelegate, NSWindowDelegate,
                         WKNavigationDelegate, WKUIDelegate, WKScriptMessageHandler {
    var window: NSWindow!
    var webView: WKWebView!
    var pageLoaded = false
    var pendingFiles: [URL] = []

    // MARK: Avvio

    func applicationDidFinishLaunching(_ note: Notification) {
        buildMenu()

        let config = WKWebViewConfiguration()
        config.websiteDataStore = .default()
        if #available(macOS 12.3, *) { config.preferences.isElementFullscreenEnabled = true }
        config.preferences.javaScriptCanOpenWindowsAutomatically = false
        config.userContentController.add(self, name: "parola")

        webView = WKWebView(frame: .zero, configuration: config)
        webView.navigationDelegate = self
        webView.uiDelegate = self
        webView.allowsMagnification = false
        webView.setValue(false, forKey: "drawsBackground")

        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 1360, height: 880),
                          styleMask: [.titled, .closable, .miniaturizable, .resizable],
                          backing: .buffered, defer: false)
        window.title = "Parola"
        window.minSize = NSSize(width: 720, height: 500)
        window.collectionBehavior = [.fullScreenPrimary]
        window.contentView = webView
        window.delegate = self
        window.setFrameAutosaveName("ParolaFinestra")
        if !window.setFrameUsingName("ParolaFinestra") { window.center() }
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
        webView.evaluateJavaScript("window.parola && window.parola.openBase64(\(name), \"\(data.base64EncodedString())\")")
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
        default: break
        }
    }

    func saveFile(data: Data, suggestedName: String) {
        let panel = NSSavePanel()
        panel.nameFieldStringValue = suggestedName
        if let docx = UTType(filenameExtension: "docx") { panel.allowedContentTypes = [docx] }
        panel.canCreateDirectories = true
        panel.beginSheetModal(for: window) { resp in
            guard resp == .OK, let url = panel.url else { return }
            do {
                try data.write(to: url, options: .atomic)
                self.toast("Salvato in «\(url.lastPathComponent)»")
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
        webView.evaluateJavaScript("window.parola && window.parola.toast(\(jsString(text)))")
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
        panel.beginSheetModal(for: window) { resp in completionHandler(resp == .OK ? panel.urls : nil) }
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

    @objc func cmdNew() { run("new") }
    @objc func cmdOpen() { run("open") }
    @objc func cmdSave() { run("save") }
    @objc func cmdPrint() { run("print") }
    @objc func cmdFind() { run("find") }
    @objc func cmdReplace() { run("replace") }
    func run(_ c: String) { webView.evaluateJavaScript("window.parola && window.parola.command('\(c)')") }

    func buildMenu() {
        let main = NSMenu()
        func item(_ title: String, _ action: Selector?, _ key: String = "", _ mods: NSEvent.ModifierFlags = .command, target: AnyObject? = nil) -> NSMenuItem {
            let i = NSMenuItem(title: title, action: action, keyEquivalent: key)
            i.keyEquivalentModifierMask = mods
            if let t = target { i.target = t }
            return i
        }
        func submenu(_ title: String, _ items: [NSMenuItem]) -> NSMenuItem {
            let top = NSMenuItem(); let m = NSMenu(title: title)
            items.forEach { m.addItem($0) }
            top.submenu = m; main.addItem(top); return top
        }

        _ = submenu("Parola", [
            item("Informazioni su Parola", #selector(NSApplication.orderFrontStandardAboutPanel(_:))),
            .separator(),
            item("Nascondi Parola", #selector(NSApplication.hide(_:)), "h"),
            item("Nascondi altre", #selector(NSApplication.hideOtherApplications(_:)), "h", [.command, .option]),
            item("Mostra tutte", #selector(NSApplication.unhideAllApplications(_:))),
            .separator(),
            item("Esci da Parola", #selector(NSApplication.terminate(_:)), "q"),
        ])
        _ = submenu("File", [
            item("Nuovo", #selector(cmdNew), "n", target: self),
            item("Apri…", #selector(cmdOpen), "o", target: self),
            .separator(),
            item("Salva .docx…", #selector(cmdSave), "s", target: self),
            item("Stampa / PDF…", #selector(cmdPrint), "p", target: self),
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
            item("Trova…", #selector(cmdFind), "f", target: self),
            item("Sostituisci…", #selector(cmdReplace), "h", [.command, .option], target: self),
        ])
        _ = submenu("Vista", [
            item("Schermo intero", #selector(NSWindow.toggleFullScreen(_:)), "f", [.command, .control]),
        ])
        let win = submenu("Finestra", [
            item("Contrai", #selector(NSWindow.performMiniaturize(_:)), "m"),
            item("Zoom", #selector(NSWindow.performZoom(_:))),
        ])
        NSApp.windowsMenu = win.submenu
        NSApp.mainMenu = main
    }
}

let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.setActivationPolicy(.regular)
app.run()
