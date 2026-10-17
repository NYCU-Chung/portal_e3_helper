//
//  ViewController.swift
//  NYCU E3 Helper
//
//  Created by 蔡侑呈 on 2026/10/1.
//

import Cocoa
import SafariServices
import WebKit

let extensionBundleIdentifier = "com.yoyo1112.nycu-e3-helper.Extension"

class ViewController: NSViewController, WKNavigationDelegate, WKScriptMessageHandler {

    @IBOutlet var webView: WKWebView!

    override func viewDidLoad() {
        super.viewDidLoad()

        self.webView.navigationDelegate = self

        self.webView.configuration.userContentController.add(self, name: "controller")

        self.webView.loadFileURL(Bundle.main.url(forResource: "Main", withExtension: "html")!, allowingReadAccessTo: Bundle.main.resourceURL!)
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        SFSafariExtensionManager.getStateOfSafariExtension(withIdentifier: extensionBundleIdentifier) { (state, error) in
            guard let state = state, error == nil else {
                // Insert code to inform the user that something went wrong.
                return
            }

            DispatchQueue.main.async {
                webView.evaluateJavaScript("show(\(state.isEnabled))")
            }
        }
    }

    func userContentController(_ userContentController: WKUserContentController, didReceive message: WKScriptMessage) {
        guard let action = message.body as? String else { return }
        if action == "open-e3", let url = URL(string: "https://e3p.nycu.edu.tw/") {
            NSWorkspace.shared.open(url)
            return
        }
        guard action == "open-preferences" else { return }
        SFSafariApplication.showPreferencesForExtension(withIdentifier: extensionBundleIdentifier) { error in
            guard error != nil else { return }
            DispatchQueue.main.async {
                let alert = NSAlert()
                alert.messageText = NSLocalizedString("Unable to open Safari Settings", comment: "")
                alert.informativeText = error?.localizedDescription ?? ""
                alert.runModal()
            }
        }
    }
}
