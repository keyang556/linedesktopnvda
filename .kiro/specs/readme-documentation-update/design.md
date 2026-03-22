# 設計文檔：README 文檔更新

## Overview

本設計說明如何更新 LINE Desktop NVDA 附加元件的四個語言版本 README 文檔（英文、日文、泰文、繁體中文），以整合新增的快速鍵功能、虛擬視窗說明，並優化文檔結構。

### 設計目標

1. 在快速鍵表格中新增兩個新快速鍵（檔案儲存、播放語音訊息）
2. 新增獨立的虛擬視窗說明章節
3. 將音效提示內容整合至使用技巧章節
4. 新增輸入手勢調整說明
5. 確保四個語言版本的結構和內容保持一致

### 設計原則

- **一致性**：所有語言版本保持相同的章節結構和內容順序
- **可讀性**：使用清晰的 Markdown 格式，遵循現有文檔風格
- **完整性**：確保所有新增內容在四個語言版本中都有對應的翻譯
- **可維護性**：保持文檔結構清晰，便於未來更新

## Architecture

### 文檔結構層次

README 文檔採用以下章節結構：

```
1. 標題與概述 (Overview)
2. 功能特色 (Features)
3. 使用技巧與提醒 (Usage Tips)
   3.1 音效提示 (Sound Effects) - 新整合位置
4. 訊息閱讀與複製 (Message Reading)
5. 虛擬視窗 (Virtual Windows) - 新增章節
6. 鍵盤快速鍵 (Keyboard Shortcuts)
   6.1 輸入手勢調整說明 - 新增內容
   6.2 各類快速鍵表格
7. 社群與支援 (Community)
8. 支援版本 (Supported Versions)
```

### 更新策略

採用「模板驅動」的更新方式：
1. 定義每個更新點的精確位置
2. 為每個語言版本準備對應的內容
3. 使用 `strReplace` 工具進行精確替換
4. 確保 Markdown 格式的一致性

## Components and Interfaces

### 更新組件

#### 1. 快速鍵表格更新組件

**位置 A - 訊息操作表格**：「鍵盤快速鍵」章節內的「訊息操作」表格

**新增內容**：
- `NVDA+Windows+K` - 開啟另存新檔對話框（附加元件類別）

**插入位置**：
- 在「訊息操作」表格的適當位置插入新行（建議在 `NVDA+Windows+T` 之後）

**表格格式**：
```markdown
| 快速鍵 | 類別 | 操作 | 備註 |
|---|---|---|---|
| **NVDA+Windows+K** | 附加元件 | 開啟另存新檔對話框 | 檔案上傳後只能下載 7 天 |
```

**位置 B - 媒體播放器表格**：「鍵盤快速鍵」章節內的「媒體播放器」表格

**新增內容**：
- `NVDA+Windows+K` - 儲存語音訊息（附加元件類別）
- `NVDA+Windows+P` - 播放語音訊息（附加元件類別）

**插入位置**：
- 在「媒體播放器」表格的第一行（在 `Ctrl+S` 之前）插入這兩個快速鍵

**表格格式**：
```markdown
| 快速鍵 | 類別 | 操作 |
|---|---|---|
| **NVDA+Windows+K** | 附加元件 | 儲存語音訊息 |
| **NVDA+Windows+P** | 附加元件 | 播放語音訊息 |
```

#### 2. 輸入手勢調整說明組件

**位置**：「鍵盤快速鍵」章節開頭，在 NOTE 區塊之後

**內容結構**：
```markdown
> [!TIP]
> 您可以透過「NVDA 功能表 → 偏好 → 輸入手勢」來自訂這些快速鍵。
```

**各語言版本內容**：
- 繁體中文：「您可以透過「NVDA 功能表 → 偏好 → 輸入手勢」來自訂這些快速鍵。」
- 英文：「You can customize these shortcuts through "NVDA Menu → Preferences → Input Gestures".」
- 日文：「これらのショートカットは「NVDA メニュー → 設定 → 入力ジェスチャー」でカスタマイズできます。」
- 泰文：「คุณสามารถปรับแต่งคีย์ลัดเหล่านี้ได้ผ่าน "เมนู NVDA → การตั้งค่า → ท่าทางการป้อนข้อมูล"」

#### 3. 虛擬視窗章節組件

**位置**：在「訊息閱讀與複製」章節之後，「鍵盤快速鍵」章節之前

**章節結構**：
```markdown
## 虛擬視窗 (Virtual Windows)

### 什麼是虛擬視窗？

虛擬視窗是此附加元件提供的一種特殊介面，用於改善某些難以透過標準無障礙 API 存取的 LINE 功能。當您使用特定快速鍵時，附加元件會建立一個虛擬的選單或對話框，讓您能夠更方便地操作這些功能。

### 虛擬視窗的類型

此附加元件目前提供以下虛擬視窗：

1. **聊天室更多選項選單**：按 `NVDA+Windows+O` 開啟
2. **訊息右鍵選單**：用於複製、回覆、收回訊息等操作
3. **PIN 碼輸入視窗**：用於鎖定模式的 PIN 碼輸入
4. **系統匣選單**：用於存取系統匣中的 LINE 選項

### 如何使用虛擬視窗

1. **開啟虛擬視窗**：使用對應的快速鍵（如 `NVDA+Windows+O`）
2. **導覽選項**：使用上下方向鍵在選項之間移動
3. **選擇選項**：按 `Enter` 鍵執行選中的選項
4. **關閉視窗**：按 `Escape` 鍵關閉虛擬視窗

### 注意事項

- 虛擬視窗是模擬的介面，實際操作仍由附加元件透過座標自動化執行
- 某些虛擬視窗選項可能因 LINE 版本或介面狀態而有所不同
- 如果虛擬視窗無法正常運作，請確認您的 LINE 視窗處於正常狀態且未被其他視窗遮擋
```

**各語言版本標題**：
- 繁體中文：「虛擬視窗」
- 英文：「Virtual Windows」
- 日文：「仮想ウィンドウ」
- 泰文：「หน้าต่างเสมือน (Virtual Windows)」

#### 4. 音效提示整合組件

**原位置**：獨立的「音效提示」章節（在「訊息閱讀與複製」章節內）

**新位置**：整合到「使用技巧與提醒」章節的末尾

**整合方式**：
- 移除「### 音效提示」的獨立標題
- 將音效提示內容作為「使用技巧與提醒」章節的最後一個項目
- 保持原有的內容和格式不變

## Data Models

### 文檔內容模型

```typescript
interface ReadmeDocument {
  language: 'zh_TW' | 'en' | 'ja' | 'th';
  sections: Section[];
}

interface Section {
  title: string;
  content: string;
  subsections?: Section[];
  order: number;
}

interface KeyboardShortcut {
  key: string;           // 例如：'NVDA+Windows+K'
  category: string;      // '附加元件' 或 'LINE'
  action: string;        // 操作描述
  note?: string;         // 可選的備註
  tableGroup: string;    // 所屬表格組別
}

interface VirtualWindow {
  name: string;          // 虛擬視窗名稱
  trigger: string;       // 觸發快速鍵
  description: string;   // 功能描述
}
```

### 更新位置映射

```typescript
interface UpdateLocation {
  section: string;       // 章節名稱
  subsection?: string;   // 子章節名稱（可選）
  position: 'before' | 'after' | 'replace' | 'append';
  anchor: string;        // 定位錨點（用於精確定位）
}

const updateLocations = {
  saveFileShortcutMessage: {
    section: '鍵盤快速鍵',
    subsection: '訊息操作',
    position: 'after',
    anchor: '| **NVDA+Windows+T** | 附加元件 | 讀出目前聊天室名稱 |'
  },
  saveVoiceShortcut: {
    section: '鍵盤快速鍵',
    subsection: '媒體播放器',
    position: 'before',
    anchor: '| **Ctrl+S** | LINE | 儲存檔案 |'
  },
  playVoiceShortcut: {
    section: '鍵盤快速鍵',
    subsection: '媒體播放器',
    position: 'before',
    anchor: '| **Ctrl+S** | LINE | 儲存檔案 |'
  },
  inputGestureNote: {
    section: '鍵盤快速鍵',
    position: 'after',
    anchor: '> [!NOTE]'
  },
  virtualWindowSection: {
    section: '虛擬視窗',
    position: 'after',
    anchor: '## 訊息閱讀與複製'
  },
  soundEffectsIntegration: {
    section: '使用技巧與提醒',
    position: 'append',
    anchor: '* **事前檢查**'
  }
};
```

### 語言內容映射

```typescript
interface LanguageContent {
  [key: string]: {
    zh_TW: string;
    en: string;
    ja: string;
    th: string;
  };
}

const contentMap: LanguageContent = {
  inputGestureTip: {
    zh_TW: '您可以透過「NVDA 功能表 → 偏好 → 輸入手勢」來自訂這些快速鍵。',
    en: 'You can customize these shortcuts through "NVDA Menu → Preferences → Input Gestures".',
    ja: 'これらのショートカットは「NVDA メニュー → 設定 → 入力ジェスチャー」でカスタマイズできます。',
    th: 'คุณสามารถปรับแต่งคีย์ลัดเหล่านี้ได้ผ่าน "เมนู NVDA → การตั้งค่า → ท่าทางการป้อนข้อมูล"'
  },
  virtualWindowTitle: {
    zh_TW: '虛擬視窗',
    en: 'Virtual Windows',
    ja: '仮想ウィンドウ',
    th: 'หน้าต่างเสมือน (Virtual Windows)'
  },
  saveFileActionMessage: {
    zh_TW: '開啟另存新檔對話框',
    en: 'Open Save As dialog',
    ja: '名前を付けて保存ダイアログを開く',
    th: 'เปิดกล่องโต้ตอบบันทึกเป็น'
  },
  saveFileNote: {
    zh_TW: '檔案上傳後只能下載 7 天',
    en: 'Files can only be downloaded for 7 days after upload',
    ja: 'ファイルはアップロード後7日間のみダウンロード可能',
    th: 'ไฟล์สามารถดาวน์โหลดได้เพียง 7 วันหลังอัปโหลด'
  },
  saveVoiceAction: {
    zh_TW: '儲存語音訊息',
    en: 'Save voice message',
    ja: 'ボイスメッセージを保存',
    th: 'บันทึกข้อความเสียง'
  },
  playVoiceAction: {
    zh_TW: '播放語音訊息',
    en: 'Play voice message',
    ja: 'ボイスメッセージを再生',
    th: 'เล่นข้อความเสียง'
  }
};
```


## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Multi-language Structural Consistency

*For any* pair of language versions (zh_TW, en, ja, th), the chapter structure and order should be identical, with all new content present in both versions.

**Validates: Requirements 6.1, 6.2**

### Property 2: Keyboard Shortcut Key Consistency

*For any* keyboard shortcut key combination (e.g., `NVDA+Windows+K`), the key string should be identical across all four language versions.

**Validates: Requirements 6.4**

### Property 3: Markdown Format Consistency

*For any* newly added content, the Markdown formatting (table structure, heading levels, special markers) should conform to the existing document's formatting conventions.

**Validates: Requirements 7.1, 7.2, 7.3**


## Error Handling

### 文件不存在錯誤

**情境**：嘗試更新不存在的 README 文件

**處理方式**：
- 在執行更新前驗證所有四個語言版本的文件是否存在
- 如果任何文件不存在，報告錯誤並列出缺失的文件
- 不執行任何更新操作，保持現有文件不變

### 定位錨點找不到錯誤

**情境**：用於定位插入位置的錨點文字在文件中不存在

**處理方式**：
- 使用 `strReplace` 工具時，如果 `oldStr` 無法匹配，工具會自動報錯
- 提供清晰的錯誤訊息，說明哪個錨點在哪個文件中找不到
- 建議檢查文件內容是否已被修改或版本不符

### Markdown 格式錯誤

**情境**：更新後的內容導致 Markdown 格式錯誤

**處理方式**：
- 在實施更新前，驗證新增內容的 Markdown 語法
- 確保表格的欄位數量與現有表格一致
- 確保標題層級符合文檔結構
- 使用 Markdown 解析器驗證更新後的文件

### 內容不完整錯誤

**情境**：某個語言版本缺少必要的翻譯內容

**處理方式**：
- 在執行更新前，驗證所有語言版本的內容映射是否完整
- 如果發現缺少翻譯，報告錯誤並列出缺失的內容
- 提供英文版本作為臨時替代方案的建議

### 編碼問題

**情境**：文件編碼不一致導致特殊字元顯示錯誤

**處理方式**：
- 確保所有文件使用 UTF-8 編碼
- 在讀取和寫入文件時明確指定 UTF-8 編碼
- 驗證特殊字元（如泰文、日文字元）正確顯示

## Testing Strategy

### 測試方法

本功能採用雙重測試策略：

1. **單元測試 (Unit Tests)**：驗證特定的文檔更新操作
2. **屬性測試 (Property-Based Tests)**：驗證跨語言版本的一致性屬性

### 單元測試範圍

單元測試專注於驗證特定的更新操作和邊界情況：

#### 1. 快速鍵新增測試

**測試案例**：
- 驗證「訊息操作」表格中 `NVDA+Windows+K` 快速鍵已新增
- 驗證快速鍵的類別標示為「附加元件」
- 驗證操作描述為「開啟另存新檔對話框」或「儲存檔案」（或對應語言的翻譯）
- 驗證備註欄位包含 7 天下載限制的提醒
- 驗證「媒體播放器」表格中 `NVDA+Windows+K` 快速鍵已新增（儲存語音訊息）
- 驗證「媒體播放器」表格中 `NVDA+Windows+P` 快速鍵已新增（播放語音訊息）

**測試方法**：
```python
def test_save_file_shortcut_added_message_actions_zh_tw():
    content = read_file('addon/doc/zh_TW/readme.md')
    # 在訊息操作章節中查找
    message_actions_section = extract_section(content, '訊息操作')
    assert 'NVDA+Windows+K' in message_actions_section
    assert '另存新檔' in message_actions_section or '儲存檔案' in message_actions_section
    assert '7 天' in message_actions_section or '7天' in message_actions_section

def test_save_voice_shortcut_added_media_player_zh_tw():
    content = read_file('addon/doc/zh_TW/readme.md')
    # 在媒體播放器章節中查找
    media_player_section = extract_section(content, '媒體播放器')
    assert 'NVDA+Windows+K' in media_player_section
    assert '儲存語音訊息' in media_player_section or '語音訊息' in media_player_section

def test_play_voice_shortcut_added_en():
    content = read_file('addon/doc/en/readme.md')
    media_player_section = extract_section(content, 'Media Player')
    assert 'NVDA+Windows+P' in media_player_section
    assert 'Play voice message' in media_player_section
```

#### 2. 輸入手勢說明測試

**測試案例**：
- 驗證每個語言版本都包含輸入手勢調整說明
- 驗證說明文字包含正確的路徑資訊
- 驗證說明使用 TIP 標記格式

**測試方法**：
```python
def test_input_gesture_tip_present():
    languages = ['zh_TW', 'en', 'ja', 'th']
    for lang in languages:
        content = read_file(f'addon/doc/{lang}/readme.md')
        assert '[!TIP]' in content
        # 驗證包含「輸入手勢」或對應翻譯
```

#### 3. 虛擬視窗章節測試

**測試案例**：
- 驗證虛擬視窗章節存在
- 驗證章節包含必要的子章節（什麼是虛擬視窗、類型、使用方法、注意事項）
- 驗證章節位置在「訊息閱讀與複製」之後
- 驗證章節位置在「鍵盤快速鍵」之前

**測試方法**：
```python
def test_virtual_window_section_structure():
    content = read_file('addon/doc/zh_TW/readme.md')
    sections = extract_sections(content)
    
    # 驗證虛擬視窗章節存在
    assert '虛擬視窗' in [s.title for s in sections]
    
    # 驗證章節順序
    section_titles = [s.title for s in sections]
    msg_idx = section_titles.index('訊息閱讀與複製')
    vw_idx = section_titles.index('虛擬視窗')
    kb_idx = section_titles.index('鍵盤快速鍵')
    
    assert msg_idx < vw_idx < kb_idx
```

#### 4. 音效提示整合測試

**測試案例**：
- 驗證「使用技巧與提醒」章節包含音效提示內容
- 驗證不存在獨立的「音效提示」章節
- 驗證音效提示內容完整（包含 ocr.wav 說明）

**測試方法**：
```python
def test_sound_effects_integrated():
    content = read_file('addon/doc/zh_TW/readme.md')
    sections = extract_sections(content)
    
    # 驗證沒有獨立的音效提示章節
    section_titles = [s.title for s in sections]
    assert '音效提示' not in section_titles
    
    # 驗證使用技巧章節包含音效提示內容
    usage_section = next(s for s in sections if '使用技巧' in s.title)
    assert 'ocr.wav' in usage_section.content
```

### 屬性測試範圍

屬性測試使用 property-based testing 框架（如 Python 的 Hypothesis）來驗證通用屬性：

#### Property Test 1: Multi-language Structural Consistency

**測試配置**：
- 最小迭代次數：100
- 測試標籤：**Feature: readme-documentation-update, Property 1: For any pair of language versions, the chapter structure and order should be identical**

**測試邏輯**：
```python
from hypothesis import given, strategies as st

@given(
    lang1=st.sampled_from(['zh_TW', 'en', 'ja', 'th']),
    lang2=st.sampled_from(['zh_TW', 'en', 'ja', 'th'])
)
def test_structural_consistency(lang1, lang2):
    """
    Feature: readme-documentation-update
    Property 1: For any pair of language versions, the chapter structure 
    and order should be identical
    """
    if lang1 == lang2:
        return  # Skip same language comparison
    
    content1 = read_file(f'addon/doc/{lang1}/readme.md')
    content2 = read_file(f'addon/doc/{lang2}/readme.md')
    
    sections1 = extract_section_structure(content1)
    sections2 = extract_section_structure(content2)
    
    # 驗證章節數量相同
    assert len(sections1) == len(sections2)
    
    # 驗證章節層級結構相同
    for s1, s2 in zip(sections1, sections2):
        assert s1.level == s2.level
        assert s1.has_subsections == s2.has_subsections
```

#### Property Test 2: Keyboard Shortcut Key Consistency

**測試配置**：
- 最小迭代次數：100
- 測試標籤：**Feature: readme-documentation-update, Property 2: For any keyboard shortcut key combination, the key string should be identical across all versions**

**測試邏輯**：
```python
@given(shortcut=st.sampled_from([
    'NVDA+Windows+K',
    'NVDA+Windows+P',
    'NVDA+Windows+C',
    'NVDA+Windows+V',
    # ... 其他快速鍵
]))
def test_shortcut_key_consistency(shortcut):
    """
    Feature: readme-documentation-update
    Property 2: For any keyboard shortcut key combination, the key string 
    should be identical across all versions
    """
    languages = ['zh_TW', 'en', 'ja', 'th']
    
    for lang in languages:
        content = read_file(f'addon/doc/{lang}/readme.md')
        shortcuts = extract_shortcuts(content)
        
        # 如果該快速鍵存在，驗證鍵組合字串完全相同
        if shortcut in shortcuts:
            assert shortcuts[shortcut].key == shortcut
```

#### Property Test 3: Markdown Format Consistency

**測試配置**：
- 最小迭代次數：100
- 測試標籤：**Feature: readme-documentation-update, Property 3: For any newly added content, the Markdown formatting should conform to existing conventions**

**測試邏輯**：
```python
@given(lang=st.sampled_from(['zh_TW', 'en', 'ja', 'th']))
def test_markdown_format_consistency(lang):
    """
    Feature: readme-documentation-update
    Property 3: For any newly added content, the Markdown formatting 
    should conform to existing conventions
    """
    content = read_file(f'addon/doc/{lang}/readme.md')
    
    # 驗證表格格式一致
    tables = extract_tables(content)
    for table in tables:
        # 所有表格應該有相同的欄位數
        column_counts = [len(row) for row in table.rows]
        assert len(set(column_counts)) == 1
        
        # 表格分隔行應該使用正確格式
        assert table.separator_row.count('|') == column_counts[0] + 1
    
    # 驗證標題層級正確
    headings = extract_headings(content)
    for i in range(len(headings) - 1):
        # 標題層級不應該跳級（如從 ## 直接到 ####）
        level_diff = headings[i+1].level - headings[i].level
        assert level_diff <= 1
```

### 測試工具

**Markdown 解析器**：使用 `markdown-it-py` 或類似工具解析 Markdown 文檔

**測試框架**：
- 單元測試：`pytest`
- 屬性測試：`hypothesis`

**輔助函數**：
```python
def extract_sections(content: str) -> List[Section]:
    """提取文檔的章節結構"""
    pass

def extract_shortcuts(content: str) -> Dict[str, Shortcut]:
    """提取快速鍵表格內容"""
    pass

def extract_tables(content: str) -> List[Table]:
    """提取所有表格"""
    pass

def extract_headings(content: str) -> List[Heading]:
    """提取所有標題"""
    pass
```

### 測試執行順序

1. **格式驗證**：先執行 Markdown 格式檢查
2. **內容驗證**：執行單元測試驗證特定內容
3. **一致性驗證**：執行屬性測試驗證跨語言一致性
4. **整合驗證**：手動檢查文檔的可讀性和完整性

### 測試覆蓋率目標

- 單元測試：覆蓋所有驗收標準中的 example 類型測試
- 屬性測試：覆蓋所有驗收標準中的 property 類型測試
- 目標覆蓋率：100% 的驗收標準
