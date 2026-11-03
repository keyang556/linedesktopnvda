# LINE Desktop Accessibility Add-on for NVDA

> **Recommended LINE version: LINE for Windows 26.6.0.4004 or later**
>
> Starting with 26.6.0.4004, LINE itself exposes the chat list, messages (including reply quotes, times, and read counts), tabs, buttons, and edit fields to NVDA. NVDA reads them directly as you press Tab or the arrow keys, so this add-on no longer reads messages with OCR or the clipboard. Please update LINE before using this version of the add-on: on older LINE versions the chat list and messages may not be read at all.
>
> If you also use [LINE MR for Windows](https://github.com/Peetchy/line-mr-for-win), do not enable it together with this add-on; both provide an NVDA application module for LINE.

## Overview

This add-on enhances the NVDA screen reader's support for the LINE Desktop application on Windows (Qt6 version). NVDA reads LINE's interface, messages, menus, and dialogs natively; the add-on adds calls, message shortcuts, a chat-history reader, and optional AI image description on top, using OCR (Optical Character Recognition) and coordinate-based automation where LINE does not expose the information yet.

> [!IMPORTANT]
> Some actions (call windows and reading the chat room name) still rely on OCR, so their recognition may not be 100% accurate. Reading messages and the message shortcuts do not use OCR.

## Features

* **Voice and Video Calls**: Initiate calls directly from the friends tab.
* **Incoming Call Handling**: Answer and reject incoming calls, and check caller information.
* **OCR Assistance**: Used only where LINE does not expose text yet, such as call windows and the chat room name.
* **Debug Tools**: Provides shortcuts to inspect the UI structure, making troubleshooting easier.
* **Accessible Message Reader**: Export the current chat and browse it in a simple reader dialog.
* **AI Image Description**: Describe the focused image or sticker message with a configurable AI service, then ask follow-up questions in an accessible dialog.
* **Status Feedback**: Announces microphone/camera state during calls, and plays a short sound after a message is actually sent.

## Usage Tips and Reminders

* **Establish Connection**: Before using this version to send a message to someone, it is recommended to send a message first using a mobile phone or the Chrome web version. Having chat history makes it easier for the add-on to locate UI elements.
* **Sending Messages**:
    1. Search for a friend's name. Enter specific keywords so the search result yields only one friend to avoid errors.
    2. Press `Enter` on the chat to open it; focus moves to the message list. Press `Tab` to move to the message input.
    3. Type your message and press `Enter` to send.
* **Making Calls**:
    Currently, you cannot initiate calls directly from the chat tab using this add-on. Please follow these steps:
    1. Press `Control+1` to switch to the friends tab.
    2. Search for the friend's name, select them, and open the chat room.
    3. Press `NVDA+Windows+C` to make a voice call, or `NVDA+Windows+V` to make a video call.
* **Describing Images**:
    1. Move focus to an image or sticker message.
    2. Press `NVDA+Windows+I`.
    3. Read the result in the Image Description dialog, or type a follow-up question and press `Ctrl+Enter`.
* **Pre-check**: Before sending a message or making a call, always check the chat history to ensure you are contacting the right person.

### Sound Effects

* **Send Sound**: When you press `Enter` in the message input and LINE really sends the message, the add-on plays a short sound. If `Enter` only confirms IME composition, no sound is played.

## Message Reading and Copying

* **Reading Messages**: With LINE 26.6.0.4004 or later, NVDA reads each message itself as you move through the message list with the arrow keys, including the sender, the quoted original of a reply, the time, and the read count. The add-on no longer right-clicks messages or runs OCR to read them, and your clipboard is left untouched.
* **Copying Messages**: Press `Control+C` in the message list, and the add-on chooses "Copy" from LINE's own context menu for the focused message. If that menu has no "Copy" item, the add-on says so.

## Message Reader and Context Menus

* **Message Reader (`NVDA+Windows+J`)**: The add-on opens the chat room's "More options" menu, chooses "Save chat", parses the exported text, and opens an accessible reader dialog.
* **Reader Navigation**: Use `UpArrow` / `DownArrow` to move one entry, `PageUp` / `PageDown` to jump ten entries, `Ctrl+Home` / `Ctrl+End` to reach the first / last entry, and `Escape` to close the window. Date separators are preserved, and the current position is announced.
* **Message Context Menu (`Applications` / `Shift+F10`)**: In the message list, these keys open LINE's own context menu for the focused message (reply, copy, share, recall, translate, "Convert to text", and more), and NVDA reads it directly: use `UpArrow` / `DownArrow` to move, `Enter` to choose, and `Escape` to close. The add-on no longer replaces it with an OCR-based virtual menu.
* **Message Shortcuts**: `NVDA+Windows+R` (reply), `NVDA+Windows+K` (save as), `NVDA+Windows+Delete` (recall), and `Control+C` (copy) open that same menu and choose the item for you, so they work only where LINE's menu offers the item.
* **Recall Confirmation**: After you choose recall (from that menu or with `NVDA+Windows+Delete`), LINE shows its own confirmation dialog and NVDA reads it: press `Tab` to hear the question and reach the **Recall** button, then press `Space` or `Enter`. The add-on no longer opens a dialog of its own for this.
* **Photo-to-Text Notice**: The first-run photo upload notice for "Convert to text" is likewise left to LINE's own dialog.

## AI Image Description

* **Shortcut (`NVDA+Windows+I`)**: Captures the focused image or sticker message and sends it to the configured image-description service.
* **Dialog**: The first description opens in a resizable dialog with a read-only transcript. Use the edit field to ask follow-up questions about the same image; press `Ctrl+Enter` to send.
* **Settings**: Open "NVDA Menu -> Preferences -> Settings -> LINE Desktop" to choose the image-description service, API key, model, and prompt. Leaving the API key or prompt blank uses the configured default.
* **Privacy Reminder**: Image description sends the captured message image to the selected cloud AI service. The first time you use the feature with a given service, a consent dialog naming that service appears; nothing is uploaded unless you agree. This applies whether the request uses the add-on's bundled shared key or your own API key, because the question is which company receives your screenshot. Consent is remembered per service, so switching to a different service asks again. Use a provider and prompt appropriate for the content you are describing.

## Virtual Windows

### What are Virtual Windows?

Virtual Windows are special interfaces provided by this add-on to improve access to certain LINE features that are difficult to access through standard accessibility APIs. When one of those LINE screens appears, the add-on creates a virtual menu or dialog box, allowing you to operate these features more conveniently.

### Types of Virtual Windows

This add-on currently provides the following virtual windows:

1. **Verification Code Window**: Reads out the verification code when logging into LINE for the first time; enter this code on your mobile device
2. **System Tray Menu**: Used to access LINE options in the system tray

### How to Use Virtual Windows

1. **Open Virtual Window**: It opens by itself when the verification code screen or LINE's tray menu appears
2. **Navigate Options**: Use the up and down arrow keys to move between options
3. **Select Option**: Press `Enter` to execute the selected option
4. **Close Window**: Press `Escape` to close the virtual window

### Notes

- Virtual windows are simulated interfaces; actual operations are still performed by the add-on through coordinate automation
- Some virtual window options may vary depending on the LINE version or interface state
- If a virtual window does not work properly, ensure your LINE window is in a normal state and not obscured by other windows

## Keyboard Shortcuts

> [!NOTE]
> In the "Category" column, "Add-on" indicates shortcuts provided by this add-on, and "LINE" indicates built-in LINE Desktop shortcuts.

> [!TIP]
> You can customize these shortcuts through "NVDA Menu → Preferences → Input Gestures".

### Calls & Incoming Calls

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Windows+C** | Add-on | Make a voice call |
| **NVDA+Windows+V** | Add-on | Make a video call |
| **NVDA+Windows+A** | Add-on | Answer an incoming call |
| **NVDA+Windows+D** | Add-on | Reject an incoming call |
| **NVDA+Windows+S** | Add-on | Check caller information |
| **NVDA+Windows+F** | Add-on | Focus the call window |
| **Ctrl+Shift+A** | LINE | Toggle microphone on/off; during calls the add-on announces the new state |
| **Ctrl+Shift+V** | LINE | Toggle camera on/off; during calls the add-on announces the new state |

### Message Actions

| Shortcut | Category | Action |
|---|---|---|
| **Control+C** | Add-on | Copy the current message (in message list; normal copy in edit fields) |
| **NVDA+Windows+R** | Add-on | Reply to the current message |
| **NVDA+Windows+Delete** | Add-on | Recall (unsend) the current message; LINE then shows its own confirmation dialog |
| **NVDA+Windows+T** | Add-on | Read current chat room name |
| **NVDA+Windows+K** | Add-on | Open Save As dialog (Files can only be downloaded for 7 days after upload) |
| **NVDA+Windows+I** | Add-on | Describe the current image or sticker message with the configured AI service |
| **Applications / Shift+F10** | LINE | Open LINE's context menu for the current message |

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
| **Enter** | LINE | Send message; the add-on plays a sound only when sending succeeds |
| **NVDA+Windows+J** | Add-on | Open the message reader |
| **NVDA+Windows+O** | Add-on | Open the chat room's More options menu (LINE's own menu) |
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
| **NVDA+Windows+K** | Add-on | Save voice message |
| **NVDA+Windows+P** | Add-on | Play voice message |
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
| **NVDA+Shift+V** | Add-on | Report LINE version, language, window type, and Qt accessibility status |
| **NVDA+Shift+J** | Global | Report information about the currently focused app and process |

## Community and Support

* **LINE User Group**: [Join Group](https://line.me/R/ti/g/BKQ2dZtTjx)
  Feel free to join the group to suggest features, report issues, or discuss with the development team.
* **Source Code & Issue Tracker**: [GitHub Repository](https://github.com/keyang556/linedesktopnvda)
  If you have feature suggestions or find bugs, feel free to open an Issue; if you are willing to contribute code, Pull Requests are entirely welcome.
* **Contact Developer**: [Contact Ken Chang (LINE)](https://line.me/ti/p/3GigC88lAt)

## Supported Versions

* LINE Desktop for Windows 26.6.0.4004 or later (standard installer or Microsoft Store version). Older LINE versions do not expose messages to NVDA, so the chat list and messages may not be read.
* NVDA 2025.1 or higher.
