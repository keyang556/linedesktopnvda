# LINE Desktop Accessibility Add-on for NVDA

## Overview

This add-on enhances the NVDA screen reader's support for the LINE Desktop application on Windows (Qt6 version). It uses OCR (Optical Character Recognition) and coordinate-based automation to improve the accessibility of chat list navigation, message reading, and making phone calls.

> [!IMPORTANT]
> Some features in this add-on rely on OCR (Optical Character Recognition), so the recognition results may not be 100% accurate.

## Features

* Improved navigation experience for chat lists and the message input field.
* **Voice and Video Calls**: Initiate calls directly from the friends tab.
* **Incoming Call Handling**: Answer and reject incoming calls, and check caller information.
* **OCR Support**: Automatically attempts to read text that cannot be accessed through standard accessibility APIs.
* **Debug Tools**: Provides shortcuts to inspect the UI structure, making troubleshooting easier.

## Usage Tips and Reminders

* **Establish Connection**: Before using this version to send a message to someone, it is recommended to send a message first using a mobile phone or the Chrome web version. Having chat history makes it easier for the add-on to locate UI elements.
* **Sending Messages**:
    1. Search for a friend's name. Enter specific keywords so the search result yields only one friend to avoid errors.
    2. In the message list or sidebar, press `Shift+Tab` to move to the edit field.
    3. Type your message and press `Enter` to send.
* **Making Calls**:
    Currently, you cannot initiate calls directly from the chat tab using this add-on. Please follow these steps:
    1. Press `Control+1` to switch to the friends tab.
    2. Search for the friend's name, select them, and open the chat room.
    3. Press `NVDA+Windows+C` to make a voice call, or `NVDA+Windows+V` to make a video call.
* **Pre-check**: Before sending a message or making a call, always check the chat history to ensure you are contacting the right person.

## Message Reading and Copying

* **Reading Messages**: When navigating through the message list, the add-on uses a "copy-first" approach to read messages. It automatically right-clicks the message → selects "Copy" → reads the clipboard content. The original clipboard content is restored after reading.
* **Copying Messages**: Press `Control+C` in the message list, and the add-on will copy the message text to the clipboard via the right-click context menu.
* **OCR Fallback**: If copying via the context menu fails (e.g., the menu or menu item cannot be found), the add-on will automatically fall back to OCR (Optical Character Recognition) to read the message content.

### Sound Effects

* **Default (Copy Success)**: If the message is read via the context menu copy method, no special sound effect is played.
* **OCR Fallback Sound**: If the message is read via OCR, the `ocr.wav` sound effect will be played to indicate that the content may not be fully accurate.

> [!WARNING]
> OCR (Optical Character Recognition) results are **not 100% accurate**. If you hear the `ocr.wav` sound effect, be aware that the text may differ from the actual message content. To verify the exact content, it is recommended to check the message on your phone or another platform.

## Keyboard Shortcuts

> [!NOTE]
> In the "Category" column, "Add-on" indicates shortcuts provided by this add-on, and "LINE" indicates built-in LINE Desktop shortcuts.

### Calls & Incoming Calls

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Windows+C** | Add-on | Make a voice call |
| **NVDA+Windows+V** | Add-on | Make a video call |
| **NVDA+Windows+A** | Add-on | Answer an incoming call |
| **NVDA+Windows+D** | Add-on | Reject an incoming call |
| **NVDA+Windows+S** | Add-on | Check caller information |
| **NVDA+Windows+F** | Add-on | Focus the call window |
| **Ctrl+Shift+A** | LINE | Toggle microphone on/off |
| **Ctrl+Shift+V** | LINE | Toggle camera on/off |

### Message Actions

| Shortcut | Category | Action |
|---|---|---|
| **Control+C** | Add-on | Copy the current message (in message list; normal copy in edit fields) |
| **NVDA+Windows+R** | Add-on | Reply to the current message |
| **NVDA+Windows+Delete** | Add-on | Recall (unsend) the current message |
| **NVDA+Windows+T** | Add-on | Read current chat room name |

### Basic Shortcuts

| Shortcut | Category | Action |
|---|---|---|
| **Ctrl+W** | LINE | Close window |
| **Ctrl+L** | LINE | Enable lock mode |
| **Ctrl+1** | LINE | Go to Friends tab |
| **Ctrl+2** | LINE | Go to Chats tab |
| **Ctrl+3** | LINE | Go to Add Friends tab |
| **Ctrl+Shift+P** | LINE | Capture screen |
| **Ctrl+Shift+K** | LINE | Open Keep Memo |
| **F1** | LINE | Show shortcut key tips (Note: screen readers cannot read this content; please refer to this document instead) |

### Chat Room Tab Navigation

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Windows+1** | Add-on | Navigate to All Chats tab |
| **NVDA+Windows+2** | Add-on | Navigate to Friends tab |
| **NVDA+Windows+3** | Add-on | Navigate to Groups tab |
| **NVDA+Windows+4** | Add-on | Navigate to Communities tab |
| **NVDA+Windows+5** | Add-on | Navigate to Official Accounts tab |

### Friends & Chat List

| Shortcut | Category | Action |
|---|---|---|
| **Ctrl+Shift+F** | LINE | Search |
| **Ctrl+Tab** | LINE | Move to previous chat room |

### Chat Room

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Windows+O** | Add-on | Click more options button |
| **Shift+Enter** | LINE | New line |
| **Ctrl+F** | LINE | Search |
| **Ctrl+N** | LINE | Open notes |
| **Ctrl+E** | LINE | Open sticker window |
| **Ctrl+O** | LINE | Send file |
| **Ctrl+K** | LINE | Open Keep Memo |
| **Ctrl+P** | LINE | Capture screen |

### Text Formatting

| Shortcut | Category | Action |
|---|---|---|
| **Ctrl+B** | LINE | Bold |
| **Ctrl+I** | LINE | Italic |
| **Ctrl+Shift+X** | LINE | Strikethrough |
| **Ctrl+Shift+C** | LINE | Red text box |
| **Ctrl+Shift+D** | LINE | Paragraph box |

### Media Player

| Shortcut | Category | Action |
|---|---|---|
| **Ctrl+S** | LINE | Save file |
| **Ctrl+C** | LINE | Copy file |
| **Enter** | LINE | Full screen |
| **Space** | LINE | Play/Pause |
| **Ctrl+Left / Right** | LINE | Fast forward/rewind video 5 seconds |
| **Ctrl+Up / Down** | LINE | Increase/decrease video volume |
| **Ctrl+Shift+Plus / Ctrl+Minus** | LINE | Zoom in/out image |

### Screen Capture

| Shortcut | Category | Action |
|---|---|---|
| **Space** | LINE | Reset capture area |
| **Ctrl+Z** | LINE | Undo |
| **Ctrl+Shift+Z / Ctrl+Y** | LINE | Redo |
| **Ctrl+C** | LINE | Copy |
| **Ctrl+S** | LINE | Save |
| **Shift** | LINE | Draw straight lines, rectangles, and circles |

### Chat Room Categories

| Shortcut | Category | Action |
|---|---|---|
| **Alt+Left / Right** | LINE | Move category |

### Debug Tools

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Shift+K** | Add-on | Debug: Check UIA and OCR (contents dumped to clipboard) |
| **NVDA+Shift+J** | Global | Report information about the currently focused app and process |

## Community and Support

* **LINE User Group**: [Join Group](https://line.me/R/ti/g/BKQ2dZtTjx)
  Feel free to join the group to suggest features, report issues, or discuss with the development team.
* **Source Code & Issue Tracker**: [GitHub Repository](https://github.com/keyang556/linedesktopnvda)
  If you have feature suggestions or find bugs, feel free to open an Issue; if you are willing to contribute code, Pull Requests are entirely welcome.
* **Contact Developer**: [Contact Ken Chang (LINE)](https://line.me/ti/p/3GigC88lAt)

## Supported Versions

* LINE Desktop for Windows (Standard installer or Microsoft Store version).
* NVDA 2022.1 or higher.
