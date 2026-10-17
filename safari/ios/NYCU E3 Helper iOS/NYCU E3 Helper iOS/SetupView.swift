import SwiftUI

struct SetupView: View {
    var body: some View {
        NavigationStack {
            List {
                Section {
                    Label("E3 Helper for iPad", systemImage: "books.vertical")
                        .font(.title2)
                        .bold()
                    Text("把作業、課程、公告與教材集中在 Safari 側欄。")
                }
                Section("首次啟用") {
                    Text("1. 開啟 iPad「設定」→「App」→「Safari」→「延伸功能」。")
                    Text("2. 啟用 NYCU E3 Helper，允許存取 e3p.nycu.edu.tw 與 e3.nycu.edu.tw。")
                    Text("3. 用 Safari 登入 E3，重新整理頁面，點右側 E3 Helper，再點「同步」。")
                    if let url = URL(string: "https://e3p.nycu.edu.tw/") {
                        Link("開啟 E3 平台", destination: url)
                            .frame(minHeight: 44)
                    }
                }
                Section("使用提醒") {
                    Text("請使用 Safari，並保持已登入的 E3 分頁開啟。公告與信件可在「公告」分頁重新載入。")
                    Text("側欄入口與寬度可用手指拖曳。分開下載會交由 Safari 開啟或下載；大型 ZIP 請分批處理。")
                    Text("iPad 版提供側欄通知，沒有系統推播。Safari 關閉或進入背景後，無法保證同步或提醒時間。")
                }
                Section("隱私與選用功能") {
                    Text("本工具使用 Safari 既有的 E3 登入狀態，不要求校務密碼。課程資料與設定儲存在擴充功能本機儲存空間。")
                    Text("翻譯會將文字傳送至 Google Translate；AI 摘要會將內容傳送至 OpenAI，需自行設定 API Key。啟用這些功能時，需允許對應服務的網站存取。")
                    Text("這是非官方輔助工具；作業要求與截止時間請以 E3 原文為準。")
                    if let url = URL(string: "https://github.com/Yoyo1112/portal_e3_helper/issues") {
                        Link("使用支援與問題回報", destination: url)
                            .frame(minHeight: 44)
                    }
                }
            }
            .navigationTitle("NYCU E3 Helper")
            .tint(.orange)
        }
    }
}

#Preview {
    SetupView()
}
