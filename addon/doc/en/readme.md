# LINE Desktop Accessibility Add-on for NVDA

## Overview

This add-on enhances NVDA screen reader support for the LINE desktop application on Windows (Qt6 version). It provides accessibility enhancements for navigating chats, reading messages, and making calls using OCR and coordinate-based automation.

> [!IMPORTANT]
> This add-on uses OCR (Optical Character Recognition) for some features. It may not be 100% accurate.

## Features

* Improved navigation for chat lists and message input fields.
* **Voice and Video Calls**: Initiate calls directly from a chat window.
* **OCR Support**: Automatically attempts to read text that isn't exposed via standard accessibility APIs.
* **Debug Tools**: Shortcuts to inspect the UI structure for troubleshooting.

## Usage Tips

* **Establishing Connection**: Before using this add-on to message someone for the first time, send them at least one message via your phone or the Chrome extension. Having a chat history makes it much easier for the add-on to locate elements correctly.
* **Sending Messages**: 
    1. Search for the friend's name. Try to use a search term that returns only one result to avoid mistakes.
    2. In the message list/sidebar, use `Shift+Tab` to reach the edit field.
    3. Type your message and press `Enter`.
* **Verification**: Always verify the chat history and the recipient's name before sending messages or making calls.
* **Limitations**: Currently, you cannot answer incoming calls via this add-on.

## Keyboard Shortcuts

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Shift+C** | LINE Desktop | Start a Voice Call |
| **NVDA+Shift+V** | LINE Desktop | Start a Video Call |
| **NVDA+Shift+T** | LINE Desktop | Open Attachment File Picker |
| **NVDA+Shift+K** | LINE Desktop | Debug: Inspect UIA and OCR (Copy to clipboard) |
| **NVDA+Shift+J** | Global | Report focused app and process info |

## Supported Versions

* LINE desktop for Windows (Standard or Microsoft Store version).
* NVDA 2022.1 or later.
