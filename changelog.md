## 1.4.0-beta1

### English

- LINE for Windows 26.6.0.4004 or later is now recommended. Starting with that version, LINE exposes the chat list, messages (including reply quotes, times, and read counts), tabs, buttons, and edit fields to NVDA, so NVDA reads them itself as you press Tab or the arrow keys. On older LINE versions the chat list and messages may not be read at all, so please update LINE first.
- The add-on no longer intercepts Tab, Shift+Tab, or the arrow keys, and no longer polls the focus, right-clicks messages to copy them, touches the clipboard, or runs OCR to read messages and chat names. This removes double speech, the synthetic Tab presses after each arrow key in the chat list, and the `ocr.wav` sound.
- The background chat cache (`NVDA+Windows+U`) has been removed. It only existed to correct OCR-read messages, which LINE now provides directly.
- `Control+C` in the message list no longer falls back to OCR when LINE's "Copy" menu item cannot be found; it reports the problem instead of copying guessed text.
- Pressing `Enter` in a search box whose name mentions messages (for example "Search chats and messages") no longer plays the message-sent sound.
- `Shift+F10` and the `Applications` key are now left to LINE, which opens its own context menu for the focused message (reply, copy, share, recall, translate, "Convert to text", and more) that NVDA reads directly. The add-on's OCR-based virtual message menu has been removed.
- After you choose recall, LINE's own confirmation dialog is used. NVDA reads it, so the add-on no longer opens a second dialog with the `Y`/`N`/`P` keys or runs OCR on LINE's dialog to click its buttons. The same applies to the first-run photo upload notice for "Convert to text" (`A`/`D`). `NVDA+Windows+Delete` still starts a recall.
- `Control+1`, `Control+2`, and `Control+3` go straight to LINE. NVDA already announces where the focus lands, so the add-on no longer speaks a fixed tab name on top of it.
- The add-on no longer guesses names for the login fields, the search box, the message input, or chat list items, because LINE now names them itself. Only buttons that LINE still leaves unnamed get a fallback name.
- Removed debug-only event handlers that read an object's name on every selection, state, or name change in LINE, each costing a cross-process call.
- `NVDA+Windows+R` (reply), `NVDA+Windows+K` (save as), `NVDA+Windows+Delete` (recall), and `Control+C` (copy) now open LINE's own context menu with `Shift+F10` and choose the item by its name, instead of right-clicking several spots on the message and reading the popup with OCR. They wait until you release the shortcut's keys first. If LINE's menu has no such item for that message (for example recall on someone else's message), the add-on closes the menu and says so. `Control+C` now also passes through normally on buttons and chat list items.
- `NVDA+Windows+O` now opens LINE's own "More options" menu, which NVDA reads directly, instead of an OCR-based virtual menu. The message reader (`NVDA+Windows+J`) chooses "Save chat" from that same menu by name.
- The virtual windows for LINE's tray menu and the PC verification code screen have been removed. NVDA reads LINE's tray menu directly. The tray virtual window also took over the chat's "More options" menu, saying "Line menu" and swallowing the arrow keys and `Enter`; that no longer happens. The verification code shown on first login is no longer read with OCR and is left to LINE and NVDA.
- The readmes, the website, and the add-on description now recommend LINE 26.6.0.4004 or later instead of LINE MR. If you also use LINE MR, do not enable both add-ons at the same time.

### 繁體中文

- 現在建議使用 Windows 版 LINE 26.6.0.4004 以上版本。從這個版本開始，LINE 會直接向 NVDA 提供聊天清單、訊息（包含回覆引用的原始訊息、時間與已讀人數）、分頁、按鈕與編輯欄位，按 Tab 或方向鍵時由 NVDA 直接讀出。較舊的 LINE 版本可能完全無法讀出聊天清單與訊息，請先更新 LINE。
- 附加元件不再攔截 Tab、Shift+Tab 與方向鍵，也不再輪詢焦點、對訊息按右鍵複製、動用剪貼簿，或用 OCR 讀取訊息與聊天室名稱。因此不會再重複報讀、不會在聊天清單中每按一次方向鍵就自動送出 Tab，也不會再播放 `ocr.wav` 音效。
- 移除背景聊天快取 (`NVDA+Windows+U`)。這個功能只是用來校正 OCR 讀出的訊息，現在 LINE 已直接提供訊息內容。
- 在訊息列表中按 `Control+C` 時，若找不到 LINE 的「複製」選項，不再改用 OCR，而是直接告知，不會把猜測的文字放進剪貼簿。
- 在名稱提到訊息的搜尋框（例如英文介面的 "Search chats and messages"）中按 `Enter`，不再播放訊息送出音效。
- `Shift+F10` 與 `Applications` 鍵現在直接交給 LINE，開啟 LINE 本身的訊息右鍵選單（回覆、複製、分享、收回、翻譯、轉為文字等），由 NVDA 直接讀出。附加元件以 OCR 建立的虛擬訊息選單已移除。
- 選擇收回後改用 LINE 本身的確認對話框，由 NVDA 直接讀出；附加元件不再另外開啟有 `Y`/`N`/`P` 按鍵的對話框，也不再用 OCR 辨識 LINE 的對話框來點按鈕。第一次使用「轉為文字」時的照片上傳提示 (`A`/`D`) 也一樣。`NVDA+Windows+Delete` 仍可開始收回。
- `Control+1`、`Control+2`、`Control+3` 直接交給 LINE。NVDA 本來就會報讀焦點移到哪裡，附加元件不再額外念出固定的分頁名稱。
- LINE 已自行為登入欄位、搜尋框、訊息輸入框與聊天清單項目命名，附加元件不再替它們猜名稱；只有 LINE 仍未命名的按鈕才會補上名稱。
- 移除僅供除錯的事件處理。它們會在 LINE 每次選取、狀態或名稱變更時讀取物件名稱，每個事件都多一次跨程序呼叫。
- `NVDA+Windows+R`（回覆）、`NVDA+Windows+K`（另存新檔）、`NVDA+Windows+Delete`（收回）與 `Control+C`（複製）改為用 `Shift+F10` 開啟 LINE 本身的右鍵選單，再依名稱選擇項目，不再對訊息的多個位置按右鍵並用 OCR 辨識彈出選單。執行前會先等您放開快速鍵的按鍵。若該訊息的選單沒有這個項目（例如收回別人的訊息），附加元件會關閉選單並告知。在按鈕與聊天清單項目上按 `Control+C` 也會照常傳給 LINE。
- `NVDA+Windows+O` 改為開啟 LINE 本身的「其他功能」選單，由 NVDA 直接讀出，不再使用 OCR 虛擬選單。訊息閱讀器 (`NVDA+Windows+J`) 也從同一個選單依名稱選擇「儲存聊天」。
- 移除 LINE 系統匣選單與電腦版認證畫面的虛擬視窗。LINE 的系統匣選單由 NVDA 直接讀出。原本系統匣虛擬視窗也會攔下聊天室的「其他功能」選單，念出「Line 功能表」並吞掉方向鍵與 `Enter`，現在不會再發生。首次登入時顯示的認證碼不再用 OCR 讀取，改由 LINE 與 NVDA 處理。
- 說明文件、網站與附加元件描述改為建議使用 LINE 26.6.0.4004 以上版本，不再推薦 LINE MR。若同時安裝了 LINE MR，請勿同時啟用兩個附加元件。

### 日本語

- Windows 版 LINE 26.6.0.4004 以降を推奨するようになりました。このバージョンから、LINE がチャットリスト、メッセージ（返信の引用元、時刻、既読数を含む）、タブ、ボタン、入力欄を NVDA に提供するため、Tab キーや矢印キーを押すと NVDA が直接読み上げます。古いバージョンの LINE ではチャットリストやメッセージがまったく読み上げられない場合があるため、先に LINE を更新してください。
- アドオンは Tab、Shift+Tab、矢印キーを横取りしなくなり、フォーカスのポーリング、メッセージを右クリックしてのコピー、クリップボードの使用、OCR によるメッセージやチャット名の読み取りを行わなくなりました。これにより、二重読み上げ、チャットリストで矢印キーを押すたびに送られていた Tab、`ocr.wav` の音がなくなりました。
- バックグラウンドのチャットキャッシュ (`NVDA+Windows+U`) を削除しました。OCR で読み取ったメッセージを補正するためだけの機能で、現在は LINE がメッセージを直接提供します。
- メッセージリストで `Control+C` を押したときに LINE の「コピー」項目が見つからない場合、OCR に切り替えず、推測した文字をコピーせずにその旨をお知らせするようになりました。
- 名前にメッセージを含む検索ボックス（例: "Search chats and messages"）で `Enter` を押しても、送信音が鳴らなくなりました。
- `Shift+F10` と `Applications` キーは LINE にそのまま渡されるようになり、LINE 自体のメッセージのコンテキストメニュー（返信、コピー、共有、取り消し、翻訳、「テキストに変換」など）が開いて NVDA が直接読み上げます。アドオンの OCR による仮想メッセージメニューは削除しました。
- 取り消しを選んだ後は LINE 自体の確認ダイアログを使います。NVDA が直接読み上げるため、アドオンは `Y`/`N`/`P` キーの付いた独自のダイアログを開いたり、OCR で LINE のダイアログのボタンを押したりしなくなりました。初回の「テキストに変換」の写真アップロード通知 (`A`/`D`) も同様です。`NVDA+Windows+Delete` で引き続き取り消しを開始できます。
- `Control+1`、`Control+2`、`Control+3` は LINE にそのまま渡されます。フォーカスの移動先は NVDA が読み上げるため、アドオンが固定のタブ名を重ねて読み上げることはなくなりました。
- ログイン欄、検索ボックス、メッセージ入力欄、チャットリストの項目は LINE 自体が名前を付けるようになったため、アドオンは名前を推測しなくなりました。LINE がまだ名前を付けていないボタンにだけ代わりの名前を付けます。
- LINE で選択・状態・名前が変わるたびにオブジェクト名を読み取っていた、デバッグ専用のイベント処理を削除しました。イベントごとにプロセス間呼び出しが発生していました。
- `NVDA+Windows+R`（返信）、`NVDA+Windows+K`（名前を付けて保存）、`NVDA+Windows+Delete`（取り消し）、`Control+C`（コピー）は、`Shift+F10` で LINE 自体のコンテキストメニューを開き、名前で項目を選ぶようになりました。メッセージの複数の位置を右クリックしてポップアップを OCR で読み取ることはなくなりました。実行前にショートカットのキーが離されるのを待ちます。そのメッセージのメニューに該当する項目がない場合（他人のメッセージの取り消しなど）、アドオンはメニューを閉じてその旨をお知らせします。ボタンやチャットリストの項目では `Control+C` が通常どおり LINE に渡されます。
- `NVDA+Windows+O` は OCR による仮想メニューではなく、NVDA が直接読み上げる LINE 自体の「その他」メニューを開くようになりました。メッセージリーダー (`NVDA+Windows+J`) も同じメニューから名前で「トーク保存」を選びます。
- LINE のシステムトレイメニューと PC 版認証画面の仮想ウィンドウを削除しました。LINE のシステムトレイメニューは NVDA が直接読み上げます。トレイの仮想ウィンドウはトークの「その他」メニューも横取りし、「Line メニュー」と読み上げて矢印キーと `Enter` を奪っていましたが、これはなくなりました。初回ログイン時に表示される認証コードは OCR で読み取らず、LINE と NVDA に任せます。
- ドキュメント、Web サイト、アドオンの説明は、LINE MR ではなく LINE 26.6.0.4004 以降を推奨するようになりました。LINE MR も使用している場合は、両アドオンを同時に有効化しないでください。

### ภาษาไทย

- ตอนนี้แนะนำให้ใช้ LINE สำหรับ Windows เวอร์ชัน 26.6.0.4004 หรือใหม่กว่า ตั้งแต่เวอร์ชันนี้ LINE ส่งข้อมูลรายการแชท ข้อความ (รวมถึงข้อความต้นฉบับที่ถูกตอบกลับ เวลา และจำนวนผู้อ่าน) แท็บ ปุ่ม และช่องกรอกข้อความให้ NVDA เอง เมื่อกด Tab หรือปุ่มลูกศร NVDA จะอ่านได้โดยตรง ใน LINE เวอร์ชันเก่าอาจอ่านรายการแชทและข้อความไม่ได้เลย โปรดอัปเดต LINE ก่อน
- ส่วนเสริมไม่ดักจับปุ่ม Tab, Shift+Tab และปุ่มลูกศรอีกต่อไป และไม่ตรวจจับโฟกัสซ้ำ ๆ ไม่คลิกขวาที่ข้อความเพื่อคัดลอก ไม่แตะต้องคลิปบอร์ด และไม่ใช้ OCR อ่านข้อความหรือชื่อห้องแชท จึงไม่มีการอ่านซ้ำ ไม่มีการกด Tab อัตโนมัติหลังกดปุ่มลูกศรในรายการแชท และไม่มีเสียง `ocr.wav` อีก
- นำแคชแชทเบื้องหลัง (`NVDA+Windows+U`) ออกแล้ว ฟังก์ชันนี้มีไว้เพื่อแก้ไขข้อความที่อ่านด้วย OCR เท่านั้น ซึ่งตอนนี้ LINE ส่งข้อความให้โดยตรงแล้ว
- เมื่อกด `Control+C` ในรายการข้อความแล้วไม่พบรายการ "คัดลอก" ของ LINE ส่วนเสริมจะไม่สลับไปใช้ OCR อีก แต่จะแจ้งให้ทราบแทนการคัดลอกข้อความที่เดาเอา
- การกด `Enter` ในช่องค้นหาที่มีชื่อเกี่ยวกับข้อความ (เช่น "Search chats and messages") จะไม่เล่นเสียงส่งข้อความอีก
- ปุ่ม `Shift+F10` และ `Applications` ส่งตรงไปยัง LINE แล้ว ซึ่งจะเปิดเมนูคลิกขวาของ LINE เองสำหรับข้อความที่โฟกัสอยู่ (ตอบกลับ คัดลอก แชร์ เรียกคืน แปล "แปลงเป็นข้อความ" และอื่น ๆ) และ NVDA อ่านได้โดยตรง นำเมนูข้อความเสมือนที่ใช้ OCR ของส่วนเสริมออกแล้ว
- หลังเลือกเรียกคืน จะใช้กล่องโต้ตอบยืนยันของ LINE เองซึ่ง NVDA อ่านได้โดยตรง ส่วนเสริมจึงไม่เปิดกล่องโต้ตอบของตัวเองที่มีปุ่ม `Y`/`N`/`P` และไม่ใช้ OCR คลิกปุ่มในกล่องโต้ตอบของ LINE อีก เช่นเดียวกับประกาศอัปโหลดรูปภาพครั้งแรกของ "แปลงเป็นข้อความ" (`A`/`D`) ยังใช้ `NVDA+Windows+Delete` เพื่อเริ่มเรียกคืนได้
- `Control+1`, `Control+2` และ `Control+3` ส่งตรงไปยัง LINE เนื่องจาก NVDA ประกาศตำแหน่งโฟกัสอยู่แล้ว ส่วนเสริมจึงไม่พูดชื่อแท็บแบบตายตัวซ้ำอีก
- LINE ตั้งชื่อช่องเข้าสู่ระบบ ช่องค้นหา ช่องพิมพ์ข้อความ และรายการในรายการแชทเองแล้ว ส่วนเสริมจึงไม่เดาชื่อให้อีก และจะตั้งชื่อสำรองให้เฉพาะปุ่มที่ LINE ยังไม่ได้ตั้งชื่อเท่านั้น
- นำตัวจัดการเหตุการณ์สำหรับดีบักออก ซึ่งอ่านชื่อวัตถุทุกครั้งที่มีการเลือก สถานะ หรือชื่อเปลี่ยนใน LINE ทำให้เกิดการเรียกข้ามโปรเซสทุกเหตุการณ์
- `NVDA+Windows+R` (ตอบกลับ), `NVDA+Windows+K` (บันทึกเป็น), `NVDA+Windows+Delete` (เรียกคืน) และ `Control+C` (คัดลอก) จะเปิดเมนูคลิกขวาของ LINE เองด้วย `Shift+F10` แล้วเลือกรายการตามชื่อ แทนการคลิกขวาหลายตำแหน่งบนข้อความและอ่านเมนูด้วย OCR โดยจะรอให้คุณปล่อยปุ่มของคีย์ลัดก่อน หากเมนูของข้อความนั้นไม่มีรายการดังกล่าว (เช่น เรียกคืนข้อความของผู้อื่น) ส่วนเสริมจะปิดเมนูและแจ้งให้ทราบ นอกจากนี้ `Control+C` บนปุ่มและรายการในรายการแชทจะส่งต่อให้ LINE ตามปกติ
- `NVDA+Windows+O` จะเปิดเมนู "ตัวเลือกเพิ่มเติม" ของ LINE เองซึ่ง NVDA อ่านได้โดยตรง แทนเมนูเสมือนที่ใช้ OCR และโปรแกรมอ่านข้อความ (`NVDA+Windows+J`) ก็เลือก "บันทึกแชท" จากเมนูเดียวกันตามชื่อ
- ลบหน้าต่างเสมือนของเมนูถาดระบบ LINE และหน้าจอยืนยันตัวตนสำหรับเวอร์ชันคอมพิวเตอร์ออกแล้ว NVDA อ่านเมนูถาดระบบของ LINE ได้โดยตรง หน้าต่างเสมือนของถาดระบบเคยแทรกเข้าไปในเมนู "ตัวเลือกเพิ่มเติม" ของแชทด้วย โดยพูดว่า "เมนู Line" และกินปุ่มลูกศรกับ `Enter` ซึ่งจะไม่เกิดขึ้นอีก รหัสยืนยันที่แสดงตอนเข้าสู่ระบบครั้งแรกไม่อ่านด้วย OCR อีกต่อไป แต่ปล่อยให้ LINE และ NVDA จัดการ
- เอกสาร เว็บไซต์ และคำอธิบายส่วนเสริมแนะนำ LINE 26.6.0.4004 หรือใหม่กว่าแทน LINE MR หากคุณใช้ LINE MR ด้วย อย่าเปิดใช้ทั้งสองส่วนเสริมพร้อมกัน
