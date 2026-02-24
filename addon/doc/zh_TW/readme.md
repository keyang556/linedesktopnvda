# LINE 電腦版 NVDA 輔助工具

## 概述

此附加元件增強了 NVDA 螢幕閱讀器對 Windows 上 LINE 電腦版（Qt6 版本）的支援。它透過 OCR（文字辨識）和座標自動化技術，改善了聊天列表導覽、訊息閱讀以及撥打電話的無障礙體驗。

> [!IMPORTANT]
> 此附加元件部分功能使用 OCR（文字辨識），因此辨識結果不會完全準確。

## 功能特色

* 改善聊天列表與訊息輸入欄位的導覽體驗。
* **語音與視訊通話**：直接從聊天視窗發起通話。
* **來電處理**：接聽、拒絕來電，以及查看來電者資訊。
* **OCR 支援**：自動嘗試讀取無法透過標準無障礙 API 存取的文字。
* **除錯工具**：提供快速鍵以檢查 UI 結構，便於排除故障。

## 使用技巧與提醒

* **建立紀錄**：在使用此版本傳送訊息給某人之前，建議先使用手機或 Chrome 網頁版傳送一次訊息。有聊天紀錄後，附加元件更容易定位 UI 元素。
* **傳送訊息**：
    1. 搜尋好友名稱。盡可能輸入準確的關鍵字，使搜尋結果只有一個好友，以避免錯誤。
    2. 在訊息清單/側邊欄，按 `Shift+Tab` 移動至編輯區。
    3. 輸入訊息後按 `Enter` 傳送。
* **事前檢查**：傳送訊息或撥打電話之前，請務必先檢查聊天紀錄，確認對象是否正確。

## 鍵盤快速鍵

| 快速鍵 | 類別 | 操作 |
|---|---|---|
| **NVDA+Shift+C** | LINE Desktop | 撥打語音通話 |
| **NVDA+Shift+V** | LINE Desktop | 撥打視訊通話 |
| **NVDA+Windows+A** | LINE Desktop | 接聽來電 |
| **NVDA+Windows+D** | LINE Desktop | 拒絕來電 |
| **NVDA+Windows+S** | LINE Desktop | 查看來電者 |
| **Control+O** | LINE Desktop | 開啟附加檔案（暫停附加元件直到檔案選擇完成） |
| **NVDA+Shift+K** | LINE Desktop | 除錯：檢查 UIA 與 OCR (內容將複製到剪貼簿) |
| **NVDA+Shift+J** | 全域 | 回報目前焦點所在的應用程式與程序資訊 |

## 社群與支援

* **LINE 使用者交流群組**：[加入群組](https://line.me/R/ti/g/BKQ2dZtTjx)
  歡迎加入群組以提出功能建議、回報使用問題，或與開發團隊進行討論。
* **原始碼與問題追蹤**：[GitHub 專案](https://github.com/keyang556/linedesktopnvda)
  若有功能建議或發現錯誤，歡迎開啟 Issue；若有意願貢獻程式碼，非常歡迎提交 Pull Request。
* **聯繫開發者**：[聯絡張可揚 (LINE)](https://line.me/ti/p/3GigC88lAt)

## 支援版本

* LINE 電腦版 for Windows (標準安裝包或 Microsoft Store 版本)。
* NVDA 2022.1 或更高版本。
