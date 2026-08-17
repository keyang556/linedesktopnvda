## 1.3.0-beta4

### English

- Message recall and LINE's first-run "Convert to text" photo upload notice are now confirmed in a standard NVDA message dialog instead of a temporary single-key mode. The same actions are available as buttons: Recall (`Y`), Cancel (`N`), Stealth recall (`P`, Premium required, shown only when LINE offers it), Agree (`A`) and Decline (`D`).
- Because those letters are now button accelerators inside a dialog rather than global bindings, they are no longer swallowed while you are typing in LINE, and a failed confirmation can no longer leave them stuck.
- The confirmation no longer times out after ten seconds; the dialog waits for your answer. Pressing escape still cancels the recall, or declines the photo upload, exactly as the timeout used to.
- All of the add-on's dialogs, including the AI image-description consent dialog and the settings-panel error messages, now use NVDA's message dialog API (`gui.message.MessageDialog`) instead of the deprecated `gui.messageBox`.
- The minimum supported NVDA version is now 2025.1, which introduced that dialog API; last tested with NVDA 2026.2.

### 繁體中文

- 收回訊息與 LINE 第一次使用「轉為文字」時的照片上傳提示，現在改用標準的 NVDA 訊息對話方塊確認，不再使用臨時單鍵模式。原本的動作都以按鈕呈現：收回 (`Y`)、取消 (`N`)、無痕收回 (`P`，需 Premium，只在 LINE 有提供時顯示)、同意 (`A`)、不同意 (`D`)。
- 這些字母現在是對話方塊內的按鈕快捷鍵，而非全域綁定，因此在 LINE 中輸入時不會再被攔截，確認失敗也不會讓這些按鍵卡住。
- 確認不再於 10 秒後逾時，對話方塊會等待您回答。按 `Escape` 仍等同取消收回或不同意照片上傳，行為與原本的逾時相同。
- 附加元件所有的對話方塊，包含 AI 圖片描述同意對話方塊與設定面板的錯誤訊息，現在都改用 NVDA 的訊息對話方塊 API (`gui.message.MessageDialog`)，不再使用已棄用的 `gui.messageBox`。
- 最低支援的 NVDA 版本提高為 2025.1（該版本開始提供此對話方塊 API）；最後測試版本為 NVDA 2026.2。

### 日本語

- メッセージの取り消しと、LINE の初回「テキストに変換」時の写真アップロード通知が、一時的な単独キー モードではなく標準の NVDA メッセージ ダイアログで確認されるようになりました。同じ操作がボタンとして用意されています: 取り消し (`Y`)、キャンセル (`N`)、ステルス取り消し (`P`、Premium が必要、LINE が提供している場合のみ表示)、同意 (`A`)、同意しない (`D`)。
- これらの文字はグローバルな割り当てではなくダイアログ内のボタンのアクセスキーになったため、LINE で入力中に横取りされることがなくなり、確認に失敗してもキーが押せなくなることがなくなりました。
- 確認は 10 秒でタイムアウトしなくなり、ダイアログは応答を待ちます。`Escape` は従来のタイムアウトと同じく、取り消しのキャンセル、または写真アップロードの拒否として機能します。
- AI 画像説明の同意ダイアログや設定パネルのエラー メッセージを含め、アドオンのすべてのダイアログが、非推奨の `gui.messageBox` ではなく NVDA のメッセージ ダイアログ API (`gui.message.MessageDialog`) を使用するようになりました。
- サポートする NVDA の最低バージョンは、このダイアログ API が導入された 2025.1 になりました。最終テストは NVDA 2026.2 で行っています。

### ภาษาไทย

- การเรียกคืนข้อความ และประกาศอัปโหลดรูปภาพเมื่อใช้ "แปลงเป็นข้อความ" ครั้งแรกของ LINE ตอนนี้ยืนยันผ่านกล่องโต้ตอบข้อความมาตรฐานของ NVDA แทนโหมดคีย์เดี่ยวชั่วคราว คำสั่งเดิมทั้งหมดมีให้เป็นปุ่ม ได้แก่ เรียกคืน (`Y`), ยกเลิก (`N`), เรียกคืนแบบไร้ร่องรอย (`P`, ต้องใช้ Premium และแสดงเฉพาะเมื่อ LINE มีตัวเลือกนั้น), ยอมรับ (`A`) และไม่ยอมรับ (`D`)
- เนื่องจากตัวอักษรเหล่านี้กลายเป็นคีย์เข้าถึงปุ่มภายในกล่องโต้ตอบแทนการผูกคีย์ส่วนกลาง จึงไม่ถูกดักจับขณะพิมพ์ใน LINE อีกต่อไป และการยืนยันที่ล้มเหลวก็ไม่ทำให้คีย์เหล่านั้นค้างอีก
- การยืนยันไม่หมดเวลาใน 10 วินาทีอีกต่อไป กล่องโต้ตอบจะรอคำตอบของคุณ การกด `Escape` ยังคงเท่ากับยกเลิกการเรียกคืนหรือปฏิเสธการอัปโหลดรูปภาพ เช่นเดียวกับการหมดเวลาแบบเดิม
- กล่องโต้ตอบทั้งหมดของส่วนเสริม รวมถึงกล่องยินยอมสำหรับคำอธิบายรูปภาพด้วย AI และข้อความแจ้งข้อผิดพลาดในแผงการตั้งค่า ตอนนี้ใช้ API กล่องโต้ตอบข้อความของ NVDA (`gui.message.MessageDialog`) แทน `gui.messageBox` ที่เลิกใช้แล้ว
- เวอร์ชัน NVDA ขั้นต่ำที่รองรับคือ 2025.1 ซึ่งเป็นเวอร์ชันที่เริ่มมี API ดังกล่าว ทดสอบล่าสุดกับ NVDA 2026.2
