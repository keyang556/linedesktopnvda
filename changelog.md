# Changelog / 更新紀錄 / 変更履歴 / บันทึกการเปลี่ยนแปลง

## 1.2.5-beta7

### English

- Fixed the y/n/p/a/d keys being left unusable in LINE after a recall or photo-consent prompt.
- Moved incoming-call handling, OCR reading, dialog detection and voice-message playback off NVDA's main thread, so NVDA no longer freezes for several seconds.
- The copy-to-read feature no longer clears the clipboard; file and image clipboard contents are now preserved.
- Hardened focus handling and the tray / PIN-code / menu virtual windows against stale UI objects and missing braille.
- Navigation no longer keeps hijacking the arrow keys after a popup closes, and rapid arrow presses no longer send stray Tab keys.
- Raised the minimum supported NVDA version to 2022.1 (the add-on relies on the controlTypes Role/State enums).

### 繁體中文

- 修正在 LINE 中使用收回或照片同意提示後，y/n/p/a/d 等按鍵被永久攔截、無法輸入的問題。
- 將來電處理、OCR 朗讀、對話框偵測與語音訊息播放移至背景執行緒，NVDA 不再凍結數秒。
- 複製朗讀功能不再清空剪貼簿；檔案與圖片等剪貼簿內容會被保留。
- 強化焦點處理與系統匣／認證碼／選單虛擬視窗，避免陳舊的介面物件或點字未初始化造成錯誤。
- 彈出選單關閉後不再持續攔截方向鍵，快速連按方向鍵也不會送出多餘的 Tab。
- 最低支援的 NVDA 版本提高為 2022.1（本附加元件使用 controlTypes 的 Role／State 列舉）。

### 日本語

- 取り消しや写真同意のプロンプトの後に、LINE で y/n/p/a/d キーが使えなくなる問題を修正しました。
- 着信処理、OCR 読み上げ、ダイアログ検出、ボイスメッセージ再生を NVDA のメインスレッドの外に移し、NVDA が数秒間フリーズしないようにしました。
- コピー読み上げ機能がクリップボードを消去しないようにしました。ファイルや画像のクリップボード内容も保持されます。
- フォーカス処理と、トレイ／PIN コード／メニューの仮想ウィンドウを、無効になった UI オブジェクトや点字未初期化に対して強化しました。
- ポップアップが閉じた後も矢印キーを奪い続けないようにし、矢印キーの連打で余分な Tab が送られないようにしました。
- サポートする NVDA の最低バージョンを 2022.1 に引き上げました（controlTypes の Role/State 列挙体を使用するため）。

### ภาษาไทย

- แก้ปัญหาที่ปุ่ม y/n/p/a/d ใช้งานไม่ได้ใน LINE หลังจากกล่องยืนยันการเรียกคืนข้อความหรือการยินยอมรูปภาพ
- ย้ายการรับสาย การอ่านด้วย OCR การตรวจจับกล่องโต้ตอบ และการเล่นข้อความเสียงออกจากเธรดหลักของ NVDA เพื่อไม่ให้ NVDA ค้างหลายวินาที
- ฟีเจอร์คัดลอกเพื่ออ่านจะไม่ล้างคลิปบอร์ดอีกต่อไป เนื้อหาคลิปบอร์ดที่เป็นไฟล์และรูปภาพจะถูกเก็บไว้
- เพิ่มความทนทานของการจัดการโฟกัสและหน้าต่างเสมือนของถาด / รหัส PIN / เมนู ต่ออ็อบเจ็กต์ UI ที่หมดอายุและกรณีที่ยังไม่ได้เริ่มต้นอักษรเบรลล์
- การนำทางจะไม่ยึดปุ่มลูกศรค้างไว้หลังจากปิดป๊อปอัป และการกดลูกศรอย่างรวดเร็วจะไม่ส่งปุ่ม Tab เกินมา
- ปรับเวอร์ชัน NVDA ขั้นต่ำที่รองรับเป็น 2022.1 (ส่วนเสริมใช้ Role/State enum ของ controlTypes)

## 1.2.5-beta6

### English

- Added an AI image description dialog with a read-only transcript and follow-up questions for the same image.
- Added settings for image-description service, API key, model, and custom prompt; blank fields use the configured defaults.
- Completed runtime translations, manifest text, and README documentation for English, Traditional Chinese, Japanese, and Thai.

### 繁體中文

- 新增 AI 圖片描述對話視窗，支援唯讀對話紀錄，並可針對同一張圖片繼續追問。
- 新增圖片描述服務、API Key、模型與自訂提示詞設定；欄位留空時使用預設設定。
- 補齊英文、繁體中文、日文與泰文的程式翻譯、manifest 文字與 README 說明文件。

### 日本語

- 同じ画像について追加質問できる、読み取り専用の会話履歴付き AI 画像説明ダイアログを追加しました。
- 画像説明サービス、API Key、モデル、カスタムプロンプトの設定を追加しました。空欄の場合は既定の設定を使用します。
- 英語、繁体字中国語、日本語、タイ語の実行時翻訳、manifest テキスト、README ドキュメントを補完しました。

### ภาษาไทย

- เพิ่มหน้าต่างอธิบายรูปภาพด้วย AI พร้อมประวัติการสนทนาแบบอ่านอย่างเดียว และสามารถถามต่อเกี่ยวกับรูปเดิมได้
- เพิ่มการตั้งค่าบริการอธิบายรูปภาพ API Key โมเดล และพรอมต์แบบกำหนดเอง หากเว้นว่างไว้จะใช้ค่าเริ่มต้น
- เติมคำแปลขณะทำงาน ข้อความ manifest และเอกสาร README สำหรับภาษาอังกฤษ จีนตัวเต็ม ญี่ปุ่น และไทยให้ครบถ้วน
