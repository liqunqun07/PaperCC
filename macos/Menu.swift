import AppKit

final class MenuController: NSObject, NSApplicationDelegate {
    let item = NSStatusBar.system.statusItem(withLength: NSStatusItem.variableLength)
    let port = ProcessInfo.processInfo.environment["PAPERCC_PORT"] ?? "8765"

    func applicationDidFinishLaunching(_ notification: Notification) {
        item.button?.title = "PaperCC"
        let menu = NSMenu()
        let open = menu.addItem(withTitle: "打开论文审阅", action: #selector(openReview), keyEquivalent: "")
        open.target = self
        menu.addItem(.separator())
        menu.addItem(withTitle: "退出菜单栏入口", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")
        item.menu = menu
    }

    @objc func openReview() {
        guard let n = UInt16(port), n > 0, let url = URL(string: "http://127.0.0.1:\(n)/review.html") else { return }
        NSWorkspace.shared.open(url)
    }
}

let app = NSApplication.shared
let controller = MenuController()
app.setActivationPolicy(.accessory)
app.delegate = controller
app.run()
