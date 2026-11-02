## 1.4.0-beta1

### English

- LINE for Windows 26.6.0.4004 or later is now recommended. Starting with that version, LINE exposes the chat list, messages (including reply quotes, times, and read counts), tabs, buttons, and edit fields to NVDA, so NVDA reads them itself as you press Tab or the arrow keys. On older LINE versions the chat list and messages may not be read at all, so please update LINE first.
- The add-on no longer intercepts Tab, Shift+Tab, or the arrow keys, and no longer polls the focus, right-clicks messages to copy them, touches the clipboard, or runs OCR to read messages and chat names. This removes double speech, the synthetic Tab presses after each arrow key in the chat list, and the `ocr.wav` sound.
- The background chat cache (`NVDA+Windows+U`) has been removed. It only existed to correct OCR-read messages, which LINE now provides directly.
- `Control+C` in the message list no longer falls back to OCR when LINE's "Copy" menu item cannot be found; it reports the problem instead of copying guessed text.
- Pressing `Enter` in a search box whose name mentions messages (for example "Search chats and messages") no longer plays the message-sent sound.
- The readmes, the website, and the add-on description now recommend LINE 26.6.0.4004 or later instead of LINE MR. If you also use LINE MR, do not enable both add-ons at the same time.

### 繁體中文

- 現在建議使用 Windows 版 LINE 26.6.0.4004 以上版本。從這個版本開始，LINE 會直接向 NVDA 提供聊天清單、訊息（包含回覆引用的原始訊息、時間與已讀人數）、分頁、按鈕與編輯欄位，按 Tab 或方向鍵時由 NVDA 直接讀出。較舊的 LINE 版本可能完全無法讀出聊天清單與訊息，請先更新 LINE。
- 附加元件不再攔截 Tab、Shift+Tab 與方向鍵，也不再輪詢焦點、對訊息按右鍵複製、動用剪貼簿，或用 OCR 讀取訊息與聊天室名稱。因此不會再重複報讀、不會在聊天清單中每按一次方向鍵就自動送出 Tab，也不會再播放 `ocr.wav` 音效。
- 移除背景聊天快取 (`NVDA+Windows+U`)。這個功能只是用來校正 OCR 讀出的訊息，現在 LINE 已直接提供訊息內容。
- 在訊息列表中按 `Control+C` 時，若找不到 LINE 的「複製」選項，不再改用 OCR，而是直接告知，不會把猜測的文字放進剪貼簿。
- 在名稱提到訊息的搜尋框（例如英文介面的 "Search chats and messages"）中按 `Enter`，不再播放訊息送出音效。
- 說明文件、網站與附加元件描述改為建議使用 LINE 26.6.0.4004 以上版本，不再推薦 LINE MR。若同時安裝了 LINE MR，請勿同時啟用兩個附加元件。

### 日本語

- Windows 版 LINE 26.6.0.4004 以降を推奨するようになりました。このバージョンから、LINE がチャットリスト、メッセージ（返信の引用元、時刻、既読数を含む）、タブ、ボタン、入力欄を NVDA に提供するため、Tab キーや矢印キーを押すと NVDA が直接読み上げます。古いバージョンの LINE ではチャットリストやメッセージがまったく読み上げられない場合があるため、先に LINE を更新してください。
- アドオンは Tab、Shift+Tab、矢印キーを横取りしなくなり、フォーカスのポーリング、メッセージを右クリックしてのコピー、クリップボードの使用、OCR によるメッセージやチャット名の読み取りを行わなくなりました。これにより、二重読み上げ、チャットリストで矢印キーを押すたびに送られていた Tab、`ocr.wav` の音がなくなりました。
- バックグラウンドのチャットキャッシュ (`NVDA+Windows+U`) を削除しました。OCR で読み取ったメッセージを補正するためだけの機能で、現在は LINE がメッセージを直接提供します。
- メッセージリストで `Control+C` を押したときに LINE の「コピー」項目が見つからない場合、OCR に切り替えず、推測した文字をコピーせずにその旨をお知らせするようになりました。
- 名前にメッセージを含む検索ボックス（例: "Search chats and messages"）で `Enter` を押しても、送信音が鳴らなくなりました。
- ドキュメント、Web サイト、アドオンの説明は、LINE MR ではなく LINE 26.6.0.4004 以降を推奨するようになりました。LINE MR も使用している場合は、両アドオンを同時に有効化しないでください。

### ภาษาไทย

- ตอนนี้แนะนำให้ใช้ LINE สำหรับ Windows เวอร์ชัน 26.6.0.4004 หรือใหม่กว่า ตั้งแต่เวอร์ชันนี้ LINE ส่งข้อมูลรายการแชท ข้อความ (รวมถึงข้อความต้นฉบับที่ถูกตอบกลับ เวลา และจำนวนผู้อ่าน) แท็บ ปุ่ม และช่องกรอกข้อความให้ NVDA เอง เมื่อกด Tab หรือปุ่มลูกศร NVDA จะอ่านได้โดยตรง ใน LINE เวอร์ชันเก่าอาจอ่านรายการแชทและข้อความไม่ได้เลย โปรดอัปเดต LINE ก่อน
- ส่วนเสริมไม่ดักจับปุ่ม Tab, Shift+Tab และปุ่มลูกศรอีกต่อไป และไม่ตรวจจับโฟกัสซ้ำ ๆ ไม่คลิกขวาที่ข้อความเพื่อคัดลอก ไม่แตะต้องคลิปบอร์ด และไม่ใช้ OCR อ่านข้อความหรือชื่อห้องแชท จึงไม่มีการอ่านซ้ำ ไม่มีการกด Tab อัตโนมัติหลังกดปุ่มลูกศรในรายการแชท และไม่มีเสียง `ocr.wav` อีก
- นำแคชแชทเบื้องหลัง (`NVDA+Windows+U`) ออกแล้ว ฟังก์ชันนี้มีไว้เพื่อแก้ไขข้อความที่อ่านด้วย OCR เท่านั้น ซึ่งตอนนี้ LINE ส่งข้อความให้โดยตรงแล้ว
- เมื่อกด `Control+C` ในรายการข้อความแล้วไม่พบรายการ "คัดลอก" ของ LINE ส่วนเสริมจะไม่สลับไปใช้ OCR อีก แต่จะแจ้งให้ทราบแทนการคัดลอกข้อความที่เดาเอา
- การกด `Enter` ในช่องค้นหาที่มีชื่อเกี่ยวกับข้อความ (เช่น "Search chats and messages") จะไม่เล่นเสียงส่งข้อความอีก
- เอกสาร เว็บไซต์ และคำอธิบายส่วนเสริมแนะนำ LINE 26.6.0.4004 หรือใหม่กว่าแทน LINE MR หากคุณใช้ LINE MR ด้วย อย่าเปิดใช้ทั้งสองส่วนเสริมพร้อมกัน
