import SafariServices
import UserNotifications

// Native messages are limited to notification operations; they do not execute commands.
final class SafariWebExtensionHandler: NSObject, NSExtensionRequestHandling {
    func beginRequest(with context: NSExtensionContext) {
        func reply(_ message: [String: Any]) {
            let response = NSExtensionItem()
            response.userInfo = [SFExtensionMessageKey: message]
            context.completeRequest(returningItems: [response], completionHandler: nil)
        }
        guard let item = context.inputItems.first as? NSExtensionItem,
              let message = item.userInfo?[SFExtensionMessageKey] as? [String: Any],
              let action = message["action"] as? String else {
            reply(["success": false, "error": "Invalid native message"])
            return
        }
        let center = UNUserNotificationCenter.current()
        switch action {
        case "notificationStatus":
            center.getNotificationSettings { settings in
                let granted = settings.authorizationStatus == .authorized || settings.authorizationStatus == .provisional
                reply(["success": true, "permission": granted ? "granted" : "denied"])
            }
        case "authorizeNotifications":
            center.requestAuthorization(options: [.alert, .sound]) { granted, error in
                if error != nil {
                    reply(["success": false, "error": "Unable to request macOS notification permission"])
                } else {
                    reply(["success": true, "permission": granted ? "granted" : "denied"])
                }
            }
        case "showNotification":
            guard let id = message["id"] as? String, !id.isEmpty, id.count <= 512,
                  let title = message["title"] as? String, !title.isEmpty, title.count <= 256,
                  let body = message["message"] as? String, body.count <= 4000 else {
                reply(["success": false, "error": "Invalid notification content"])
                return
            }
            let urlString = message["url"] as? String ?? ""
            if !urlString.isEmpty {
                guard let url = URL(string: urlString), url.scheme == "https",
                      ["e3.nycu.edu.tw", "e3p.nycu.edu.tw"].contains(url.host ?? ""),
                      url.user == nil, url.password == nil else {
                    reply(["success": false, "error": "Invalid notification URL"])
                    return
                }
            }
            center.getNotificationSettings { settings in
                guard settings.authorizationStatus == .authorized || settings.authorizationStatus == .provisional else {
                    reply(["success": false, "error": "Allow NYCU E3 Helper notifications in macOS System Settings"])
                    return
                }
                let content = UNMutableNotificationContent()
                content.title = title
                content.body = body
                content.sound = .default
                content.userInfo = ["url": urlString]
                let request = UNNotificationRequest(identifier: id, content: content, trigger: nil)
                center.add(request) { error in
                    reply(error == nil ? ["success": true] : ["success": false, "error": "Unable to schedule macOS notification"])
                }
            }
        default:
            reply(["success": false, "error": "Unsupported native operation"])
        }
    }
}
