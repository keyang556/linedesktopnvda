# Changelog / 更新紀錄 / 変更履歴 / บันทึกการเปลี่ยนแปลง

## 1.3.0-beta1

### English

- The incoming-call shortcuts (NVDA+Windows+A/D/S/F) and their Tools-menu equivalents now run on a worker thread on the actual gesture path, so answering or rejecting a call no longer freezes NVDA for several seconds.
- AI image description now asks for one-time consent before uploading a screenshot to a cloud AI service with the add-on's bundled shared key. Users who configured their own API key are never asked.
- All of the add-on's gestures now appear under a single translated "LINE Desktop" category in NVDA's Input Gestures dialog, and the navigation scripts (Tab/arrows, Enter, Ctrl+1/2/3, mic/camera toggles) now have translatable descriptions in input help.
- Braille messages now follow NVDA's braille message timeout settings (set the timeout to "show indefinitely" to keep the old persistent behavior).
- NVDA no longer reacts to the add-on's own synthetic mouse clicks, and the mouse pointer returns to its previous position after such clicks.
- The file-dialog watcher and window-foregrounding waits were moved off NVDA's main thread and now react as soon as the state changes instead of on fixed timers.
- Fixed a leak that could keep OCR pixel buffers alive after LINE exited.
- The minimum supported NVDA version is now 2024.1; last tested with NVDA 2026.2.

### 繁體中文

- 來電快速鍵（NVDA+Windows+A/D/S/F）與工具功能表的對應項目現在會在實際按鍵路徑上於背景執行緒執行，接聽或拒絕來電不再讓 NVDA 凍結數秒。
- AI 圖片描述在使用附加元件內建的共用金鑰上傳螢幕截圖到雲端 AI 服務前，會先徵求一次性同意；已設定自己 API 金鑰的使用者不會被詢問。
- 附加元件的所有手勢現在都顯示在 NVDA 輸入手勢對話方塊中單一的「LINE Desktop」翻譯分類下，導覽腳本（Tab／方向鍵、Enter、Ctrl+1/2/3、麥克風／鏡頭切換）在輸入說明中也有可翻譯的描述。
- 點字訊息現在遵循 NVDA 的點字訊息逾時設定（將逾時設定為「持續顯示」可保留舊有的永久顯示行為）。
- NVDA 不再對附加元件自身合成的滑鼠點擊做出反應，點擊後滑鼠游標會回到原本的位置。
- 檔案對話方塊偵測與視窗前景等待已移出 NVDA 主執行緒，並在狀態改變時立即反應，而非依固定計時器。
- 修正 LINE 結束後 OCR 像素緩衝區可能持續佔用記憶體的問題。
- 最低支援的 NVDA 版本提高為 2024.1；最後測試版本為 NVDA 2026.2。

### 日本語

- 着信ショートカット（NVDA+Windows+A/D/S/F）とツールメニューの対応項目が、実際のジェスチャ経路でもワーカースレッドで実行されるようになり、着信の応答や拒否で NVDA が数秒間フリーズしなくなりました。
- AI 画像説明は、アドオン内蔵の共有キーでスクリーンショットをクラウド AI サービスにアップロードする前に、一度だけ同意を求めるようになりました。自分の API キーを設定しているユーザーには確認しません。
- アドオンのすべてのジェスチャが NVDA の入力ジェスチャダイアログで翻訳された単一の「LINE Desktop」カテゴリに表示されるようになり、ナビゲーション スクリプト（Tab／矢印キー、Enter、Ctrl+1/2/3、マイク／カメラ切り替え）にも入力ヘルプで翻訳可能な説明が付きました。
- 点字メッセージが NVDA の点字メッセージ表示時間の設定に従うようになりました（従来の常時表示にするには表示時間を「無期限に表示」に設定してください）。
- アドオン自身が合成したマウスクリックに NVDA が反応しなくなり、クリック後にマウスポインターが元の位置に戻るようになりました。
- ファイルダイアログの監視とウィンドウの前面化待ちを NVDA のメインスレッドの外に移し、固定タイマーではなく状態の変化に即座に反応するようにしました。
- LINE の終了後に OCR のピクセルバッファーが解放されないことがある問題を修正しました。
- サポートする NVDA の最低バージョンは 2024.1 になりました。最終テストは NVDA 2026.2 で行っています。

### ภาษาไทย

- ปุ่มลัดสายเรียกเข้า (NVDA+Windows+A/D/S/F) และรายการที่ตรงกันในเมนูเครื่องมือ ตอนนี้ทำงานบนเธรดพื้นหลังในเส้นทางการกดปุ่มจริง การรับหรือปฏิเสธสายจึงไม่ทำให้ NVDA ค้างหลายวินาทีอีกต่อไป
- การอธิบายรูปภาพด้วย AI จะขอความยินยอมครั้งเดียวก่อนอัปโหลดภาพหน้าจอไปยังบริการ AI คลาวด์ด้วยคีย์ส่วนกลางที่มากับส่วนเสริม ผู้ใช้ที่ตั้งค่า API คีย์ของตนเองจะไม่ถูกถาม
- ท่าทางทั้งหมดของส่วนเสริมแสดงอยู่ใต้หมวดหมู่ "LINE Desktop" ที่แปลแล้วเพียงหมวดเดียวในกล่องโต้ตอบท่าทางการป้อนข้อมูลของ NVDA และสคริปต์การนำทาง (Tab/ลูกศร, Enter, Ctrl+1/2/3, สลับไมค์/กล้อง) มีคำอธิบายที่แปลได้ในวิธีใช้การป้อนข้อมูลแล้ว
- ข้อความอักษรเบรลล์เป็นไปตามการตั้งค่าเวลาแสดงข้อความเบรลล์ของ NVDA แล้ว (ตั้งค่าเป็น "แสดงตลอดไป" เพื่อคงพฤติกรรมแสดงถาวรแบบเดิม)
- NVDA จะไม่ตอบสนองต่อการคลิกเมาส์สังเคราะห์ของส่วนเสริมเอง และตัวชี้เมาส์จะกลับสู่ตำแหน่งเดิมหลังการคลิก
- ย้ายการตรวจจับกล่องโต้ตอบไฟล์และการรอหน้าต่างขึ้นเป็นพื้นหน้าออกจากเธรดหลักของ NVDA และตอบสนองทันทีเมื่อสถานะเปลี่ยน แทนการใช้ตัวจับเวลาแบบตายตัว
- แก้ไขการรั่วไหลที่อาจทำให้บัฟเฟอร์พิกเซลของ OCR ค้างอยู่ในหน่วยความจำหลังจาก LINE ปิดไปแล้ว
- เวอร์ชัน NVDA ขั้นต่ำที่รองรับคือ 2024.1 และทดสอบล่าสุดกับ NVDA 2026.2

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
