import AppKit

final class Provider: NSObject {
    @objc func openWorkbench(_ pasteboard: NSPasteboard, userData: String?, error: AutoreleasingUnsafeMutablePointer<NSString?>) {
        guard let host = userData, ["word", "excel", "powerpoint"].contains(host) else {
            error.pointee = "请在 Word、Excel 或 PowerPoint 中打开三桥 AI。"; return
        }
        let program = FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent("Applications/TrioBridge AI.app/Contents/MacOS/TrioBridge AI")
        let task = Process()
        task.executableURL = program
        task.arguments = ["--office=" + host]
        do { try task.run() }
        catch let failure { error.pointee = "无法打开三桥 AI，请重新运行 Install.command。\n\(failure.localizedDescription)" as NSString }
    }
}
let application = NSApplication.shared
let provider = Provider()
application.setActivationPolicy(.accessory)
application.servicesProvider = provider
NSUpdateDynamicServices()
application.run()
