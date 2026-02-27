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

## Keyboard Shortcuts

| Shortcut | Category | Action |
|---|---|---|
| **NVDA+Windows+C** | LINE Desktop | Make a voice call |
| **NVDA+Windows+V** | LINE Desktop | Make a video call |
| **NVDA+Windows+A** | LINE Desktop | Answer an incoming call |
| **NVDA+Windows+D** | LINE Desktop | Reject an incoming call |
| **NVDA+Windows+S** | LINE Desktop | Check caller information |
| **NVDA+Windows+T** | LINE Desktop | Read current chat room name |
| **Control+1** | LINE Desktop | Switch to friends tab |
| **Control+2** | LINE Desktop | Switch to chats tab |
| **Control+3** | LINE Desktop | Switch to add friends tab |
| **Control+O** | LINE Desktop | Open file attachment (pauses add-on until file selection completes) |
| **NVDA+Shift+K** | LINE Desktop | Debug: Check UIA and OCR (contents dumped to clipboard) |
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
