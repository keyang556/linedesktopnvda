# LINE Desktop Accessibility Add-on for NVDA

> **Recommended LINE version / 建議安裝版本 / 推奨バージョン / เวอร์ชันที่แนะนำ: LINE for Windows 26.6.0.4004+**
>
> **English:** LINE 26.6.0.4004 and later expose the chat list, messages, tabs, buttons, and edit fields to NVDA directly, so this add-on no longer reads messages with OCR or the clipboard. Please update LINE first; on older versions the chat list and messages may not be read. If you also use [LINE MR for Windows](https://github.com/Peetchy/line-mr-for-win), do not enable both add-ons together.
>
> **繁體中文：** LINE 26.6.0.4004 以上版本會直接向 NVDA 提供聊天清單、訊息、分頁、按鈕與編輯欄位，因此本附加元件不再用 OCR 或剪貼簿閱讀訊息。請先更新 LINE；在較舊版本中，聊天清單與訊息可能無法讀出。若同時安裝了 [LINE MR for Windows](https://github.com/Peetchy/line-mr-for-win)，請勿同時啟用兩個附加元件。
>
> **日本語：** LINE 26.6.0.4004 以降は、チャットリスト、メッセージ、タブ、ボタン、入力欄を NVDA に直接提供するため、本アドオンは OCR やクリップボードでメッセージを読み取らなくなりました。先に LINE を更新してください。古いバージョンではチャットリストやメッセージが読み上げられない場合があります。[LINE MR for Windows](https://github.com/Peetchy/line-mr-for-win) も使用している場合は、両アドオンを同時に有効化しないでください。
>
> **ไทย:** LINE 26.6.0.4004 หรือใหม่กว่าส่งข้อมูลรายการแชท ข้อความ แท็บ ปุ่ม และช่องกรอกข้อความให้ NVDA โดยตรง ส่วนเสริมนี้จึงไม่ใช้ OCR หรือคลิปบอร์ดในการอ่านข้อความอีกต่อไป โปรดอัปเดต LINE ก่อน ในเวอร์ชันเก่าอาจอ่านรายการแชทและข้อความไม่ได้ หากคุณใช้ [LINE MR for Windows](https://github.com/Peetchy/line-mr-for-win) ด้วย อย่าเปิดใช้ทั้งสองส่วนเสริมพร้อมกัน

---

This project provides an NVDA screen reader add-on designed to improve the accessibility of the LINE Desktop application on Windows.

## Documentation / 文件 / ドキュメント

Detailed documentation is available in the following languages:

- [English (English)](./addon/doc/en/readme.md)
- [繁體中文 (Traditional Chinese)](./addon/doc/zh_TW/readme.md)
- [日本語 (Japanese)](./addon/doc/ja/readme.md)
- [ภาษาไทย (Thai)](./addon/doc/th/readme.md)

---

## 快速摘要 (Quick Summary)

### Traditional Chinese (繁體中文)
此附加元件增強了 LINE 電腦版的無障礙支援，包含：
- 支援語音與視訊通話快速鍵 (`NVDA+Windows+C/V`)。
- 支援來電接聽與拒絕 (`NVDA+Windows+A/D`)。
- 跳到通話視窗 (`NVDA+Windows+F`)。
- 支援訊息閱讀器，可匯出聊天室並在視窗中逐則閱讀 (`NVDA+Windows+J`)。
- 支援 AI 圖片描述，可描述圖片或貼圖訊息並在對話視窗追問 (`NVDA+Windows+I`)。
- 可在 LINE Desktop 設定面板選擇圖片描述服務、API Key、模型與提示詞。
- 通話中麥克風與相機狀態播報。
- 聊天清單、訊息、`Shift+F10` 右鍵選單與收回確認對話框都由 NVDA 直接讀取（需 LINE 26.6.0.4004 以上）；訊息快速鍵也改用 LINE 的選單；OCR 只用於通話視窗與聊天室名稱。
- **注意**：OCR 並非 100% 準確。

### English
This add-on enhances LINE Desktop accessibility, featuring:
- Shortcuts for voice and video calls (`NVDA+Windows+C/V`).
- Answer/reject incoming calls (`NVDA+Windows+A/D`).
- Focus call window (`NVDA+Windows+F`).
- Message Reader for exporting and browsing chats (`NVDA+Windows+J`).
- AI Image Description for images and stickers, with follow-up questions in a dialog (`NVDA+Windows+I`).
- LINE Desktop settings for image-description service, API key, model, and prompt.
- Microphone and camera status announcements during calls.
- The chat list, messages, the `Shift+F10` context menu, and the recall confirmation are read by NVDA directly (LINE 26.6.0.4004 or later required); the message shortcuts use LINE's own menus too; OCR is only used for call windows and the chat room name.
- **Note**: OCR is not perfectly accurate.

### Japanese (日本語)
このアドオンは LINE デスクトップ版のアクセシビリティを向上させます：
- 音声・ビデオ通話のショートカット (`NVDA+Windows+C/V`)。
- 着信への応答・拒否 (`NVDA+Windows+A/D`)。
- 通話ウィンドウにフォーカス (`NVDA+Windows+F`)。
- チャットを書き出して閲覧できるメッセージリーダー (`NVDA+Windows+J`)。
- 画像やスタンプを説明し、ダイアログで追加質問できる AI 画像説明 (`NVDA+Windows+I`)。
- LINE Desktop 設定パネルで画像説明サービス、API Key、モデル、プロンプトを設定。
- 通話中のマイクとカメラの状態読み上げ。
- チャットリスト、メッセージ、`Shift+F10` のコンテキストメニュー、取り消しの確認ダイアログは NVDA が直接読み上げます（LINE 26.6.0.4004 以降が必要）。メッセージ用ショートカットも LINE のメニューを使い、OCR は通話ウィンドウとチャットルーム名にのみ使用します。
- **注意**: OCR は完全に正確ではありません。

### Thai (ภาษาไทย)
ส่วนเสริมนี้ช่วยเพิ่มประสิทธิภาพการเข้าถึง LINE Desktop:
- คีย์ลัดสำหรับการโทรด้วยเสียงและวิดีโอ (`NVDA+Windows+C/V`)
- รับสาย/ปฏิเสธสายเรียกเข้า (`NVDA+Windows+A/D`)
- ไปยังหน้าต่างการโทร (`NVDA+Windows+F`)
- มีโปรแกรมอ่านข้อความสำหรับส่งออกและเปิดอ่านแชท (`NVDA+Windows+J`)
- อธิบายรูปภาพหรือสติกเกอร์ด้วย AI และถามต่อในหน้าต่างโต้ตอบ (`NVDA+Windows+I`)
- ตั้งค่าบริการอธิบายรูปภาพ API Key โมเดล และพรอมต์ได้ในแผง LINE Desktop
- แจ้งสถานะไมโครโฟนและกล้องระหว่างการโทร
- NVDA อ่านรายการแชท ข้อความ เมนูคลิกขวา `Shift+F10` และกล่องโต้ตอบยืนยันการเรียกคืนได้โดยตรง (ต้องใช้ LINE 26.6.0.4004 หรือใหม่กว่า) คีย์ลัดสำหรับข้อความก็ใช้เมนูของ LINE เช่นกัน ใช้ OCR เฉพาะหน้าต่างการโทรและชื่อห้องแชท
- **หมายเหตุ**: OCR ไม่ถูกต้อง 100%

---

## Community & Support

* **LINE User Group**: [Join the Group](https://line.me/R/ti/g/BKQ2dZtTjx)
  Join our LINE group for feature suggestions, troubleshooting, and discussions with the development team.
* **Source Code & Issue Tracker**: [GitHub Repository](https://github.com/keyang556/linedesktopnvda)
  Feel free to open an Issue for feature requests or bug reports. Pull Requests are always welcome for code contributions.
* **Contact Developer**: [Contact via LINE](https://line.me/ti/p/3GigC88lAt)

---

## Development

Copyright (C) 2026 LINE Desktop NVDA Add-on Contributors.
Distributed under the terms of the GNU General Public License, version 2 or later.
