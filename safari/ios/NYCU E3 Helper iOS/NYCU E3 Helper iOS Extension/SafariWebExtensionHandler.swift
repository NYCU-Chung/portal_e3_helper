import SafariServices

// iPad has no desktop notification bridge. Reject native operations explicitly.
final class SafariWebExtensionHandler: NSObject, NSExtensionRequestHandling {
    func beginRequest(with context: NSExtensionContext) {
        let response = NSExtensionItem()
        response.userInfo = [SFExtensionMessageKey: [
            "success": false,
            "error": "Native operations are unavailable on iPad. Use sidebar notifications."
        ]]
        context.completeRequest(returningItems: [response], completionHandler: nil)
    }
}
