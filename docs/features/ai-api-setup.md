# Gemini／OpenAI API Key 設定教學

[English](#english) · [AI 摘要與今日總覽](openai-digest.md)

最後查核官方文件：2026-10-23。資格、模型、額度與介面可能變更，請以自己的供應商控制台與下方官方文件為準。本教學適用於 Chrome、Safari 與 iPad 版的 AI 摘要設定。

API Key 是連接服務的私密金鑰；token 是計算輸入與輸出用量的單位。建立金鑰不代表取得免費用量。翻譯使用 Google Translate，不需要 AI 金鑰；單篇摘要與今日總覽才使用所選的 AI 供應商。

## Gemini：使用免費方案

1. 使用自己的 Google 帳號登入 [Google AI Studio API Keys](https://aistudio.google.com/apikey)，確認符合 [使用資格與地區限制](https://ai.google.dev/gemini-api/docs/available-regions)；Gemini API 要求使用者年滿 18 歲。
2. 接受服務條款後，選擇建立 API Key，依畫面指示建立或選取專案，再複製金鑰。詳細操作見 [官方金鑰教學](https://ai.google.dev/gemini-api/docs/api-key)。
3. 在 AI Studio 確認專案方案與 [目前用量限制](https://ai.google.dev/gemini-api/docs/rate-limits)。若只想使用免費方案，先確認專案未升級付費，再從 [價格表](https://ai.google.dev/gemini-api/docs/pricing) 選擇有免費用量的模型。
4. 依下方步驟貼入 E3 Helper。模型清單包含可生成內容的模型，但不代表每個模型都免費或帳號都有權限使用；請先核對價格表。

免費方案受模型與專案額度限制。未付費服務的內容可能用於改善 Google 產品，並可能由人工審查；部分地區或帳戶適用不同條款。請閱讀 [Gemini API 條款](https://ai.google.dev/gemini-api/terms)，不要傳送敏感或機密內容。

## OpenAI：建立金鑰與確認費用

1. 登入 [OpenAI Platform](https://platform.openai.com/)，選擇或建立供 E3 Helper 使用的專案。
2. 在 [API Keys](https://platform.openai.com/api-keys) 建立該專案的 secret key，立即複製並妥善保存。
3. 檢查 [API 帳務](https://platform.openai.com/settings/organization/billing/overview) 與 [用量](https://platform.openai.com/usage)。ChatGPT 訂閱不包含 API 用量；請另外確認 API 餘額與 [模型價格](https://developers.openai.com/api/docs/pricing)。
4. 依下方步驟在 E3 Helper 選擇 OpenAI、貼上金鑰及選擇模型。

### 選用：分享輸入與輸出換取免費 token

組織擁有者可在 [Data controls → Sharing](https://platform.openai.com/settings/organization/data-controls/sharing) 查看免費用量資格。若有優惠且願意分享，啟用 **Share inputs and outputs with OpenAI**，選取 E3 Helper 金鑰所屬專案，確認已加入優惠。分享將用於模型改善；不要分享個資或機密課程內容。

仍須有正餘額；僅涵蓋指定模型與專案。每日額度依帳戶與模型群組而異，於 UTC 00:00（台灣 08:00）重設；讓總量超額的整筆請求會計費。到 Usage 核對優惠；沒有優惠提示就不要假設免費。可隨時關閉分享。完整資格、額度及排除項目見 [官方說明](https://help.openai.com/en/articles/10306912-sharing-feedback-evaluation-and-fine-tuning-data-and-api-inputs-and-outputs-with-openai)。

## 在 E3 Helper 啟用

1. 開啟側欄 **更多 → 設定**，找到 **AI 摘要與今日總覽**，勾選 **啟用 AI 摘要**。
2. 選擇 **Gemini** 或 **OpenAI**，貼上該供應商的 API Key。
3. Gemini：等待模型清單載入，必要時按 **重新載入模型**；也可手動輸入模型 ID。OpenAI：從摘要模型選單選擇自己的帳戶可用模型。
4. 按 **測試連接** 檢查，再按 **儲存設定**。測試會呼叫供應商 API，也可能消耗用量。
5. 在公告或信件預覽按 **AI摘要**；或在公告頁按 **產生總覽**。兩者共用上述設定。

單篇摘要會傳送內文；今日總覽會傳送當日項目的標題、課程、寄件者及時間。金鑰儲存在本機瀏覽器，切換供應商會保留兩組金鑰。不要把金鑰放進截圖、Git 或問題回報；若外洩，立即至供應商撤銷並建立新金鑰。

OpenAI 請求使用 `store: false`，但這不代表零保留，也不會取消你在供應商端選擇的資料分享；一般 API 濫用監控仍可能保留資料。見 [OpenAI 資料控制](https://developers.openai.com/api/docs/guides/your-data)。

## 常見問題

| 狀況 | 處理方式 |
| --- | --- |
| 金鑰無效／401 | 確認供應商正確、金鑰完整且未撤銷；重新貼上並儲存。 |
| 429／額度不足 | 查看供應商用量、速率限制與 API 餘額；等待額度重設或減少請求。不要只因有金鑰或 ChatGPT 訂閱就假設有額度。 |
| 模型不存在或沒有權限 | 選擇帳戶可用的模型；Gemini 可重新載入清單，並核對免費方案是否支援該模型。 |
| 測試成功但摘要失敗 | 查看錯誤訊息；長篇內容可能受 token 或速率限制，測試成功不保證之後請求免費。 |
| Safari 無法連線 | 檢查網路與 Safari 的供應商網站存取權限，再重試。 |

## English

Official guidance checked on **2026-10-23**. Eligibility, quotas and models can change. Use your provider dashboard and the linked official documentation to confirm current terms. These steps apply to Chrome, Safari and iPad.

An **API key** is a private credential; **tokens** measure input and output usage. Creating a key does not grant free usage. Translation needs no AI key; summaries and daily digests use your selected provider.

### Gemini free tier

1. Sign in to [Google AI Studio](https://aistudio.google.com/apikey). Check [age and region eligibility](https://ai.google.dev/gemini-api/docs/available-regions); Gemini API users must be 18 or older.
2. Accept the terms, create an API key in a new or existing project, and copy it. See the [official key guide](https://ai.google.dev/gemini-api/docs/api-key).
3. Confirm the project's plan and [limits](https://ai.google.dev/gemini-api/docs/rate-limits). For free usage, verify that the project has not been upgraded to paid billing and choose a model with a free tier in the [pricing table](https://ai.google.dev/gemini-api/docs/pricing). E3 Helper's model list does not indicate price or account access.

Unpaid-service content may be used for product improvement and human review; regional and account exceptions apply. Read the [Gemini terms](https://ai.google.dev/gemini-api/terms) before sending course content.

### OpenAI key and optional complimentary tokens

Sign in to [OpenAI Platform](https://platform.openai.com/), create or select a project, and create a secret key under [API Keys](https://platform.openai.com/api-keys). Copy it securely. Check [billing](https://platform.openai.com/settings/organization/billing/overview), [usage](https://platform.openai.com/usage) and [pricing](https://developers.openai.com/api/docs/pricing); ChatGPT subscriptions do not cover API usage.

Organization owners can check [Sharing](https://platform.openai.com/settings/organization/data-controls/sharing) for eligibility. If offered and acceptable, enable **Share inputs and outputs with OpenAI** for the key's project and confirm enrollment. Shared content supports model improvement; exclude sensitive data.

A positive balance is required. Only eligible models and shared projects qualify. Daily group quotas vary; reset is 00:00 UTC (08:00 Taiwan). An entire request crossing the quota is billed. Verify the incentive in Usage; absent an offer, assume no complimentary tokens. Opt out anytime. Check [official eligibility and exclusions](https://help.openai.com/en/articles/10306912-sharing-feedback-evaluation-and-fine-tuning-data-and-api-inputs-and-outputs-with-openai).

### Connect E3 Helper

1. Open **More → Settings → AI summaries and daily digest**, and enable AI summaries.
2. Choose the provider and paste its key. For Gemini, load or refresh the model list, or enter a model ID manually. For OpenAI, select an available summary model.
3. Click **Test connection**, then **Save settings**. Testing calls the provider and can consume quota.
4. Use **AI summary** in an announcement/message preview or **Generate digest** in Announcements.

Individual summaries send the body; digests send today's titles, courses, senders and timestamps. Keys remain in local browser storage and persist when switching providers. Keep them out of screenshots, Git and issue reports; revoke leaked keys immediately.

OpenAI requests set `store: false`; this does not guarantee zero retention or override your data-sharing choice. See [OpenAI data controls](https://developers.openai.com/api/docs/guides/your-data).

For invalid-key/401 errors, check the provider and replace the key. For 429/quota errors, check usage, rate limits and balance. For unavailable models, choose one your account supports; refresh Gemini's list if needed. If a short connection test works but a summary fails, check the error and content length. In Safari, also check provider website permissions and connectivity.
