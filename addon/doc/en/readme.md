# LINE Desktop Accessibility Add-on for NVDA

## Overview

This add-on enhances NVDA screen reader support for the LINE desktop application on Windows. It provides improved accessibility for navigating chats, reading messages, managing contacts, and composing messages.

## Features

* Improved announcement of chat list items (chat name, last message, unread count)
* Better reading of individual chat messages (sender, content, timestamp)
* Automatic labeling of unlabeled buttons and controls
* Quick navigation keyboard shortcuts
* Support for both standard installer and Microsoft Store versions of LINE

## Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| NVDA+Shift+L | Read the last received message in the current chat |
| NVDA+Shift+M | Move focus to the message input field |
| NVDA+Shift+T | Report the current chat/conversation name |
| NVDA+Shift+D | Log debug info about the focused control (for development) |

## Supported Versions

* LINE desktop for Windows (standard installer or Microsoft Store version)
* NVDA 2024.1 or later

## Known Limitations

* This is an initial version. The overlay classes may need refinement based on the actual UIA tree structure of your LINE desktop version.
* Some UI elements may not be perfectly labeled depending on the LINE desktop version and its accessibility implementation.
* The debug shortcut (NVDA+Shift+D) can be used to inspect control properties and help improve the add-on.

## Troubleshooting

If the add-on does not seem to work:

1. Make sure LINE desktop is running and focused.
2. Use NVDA+Shift+D to log debug information about the current control.
3. Check the NVDA log (NVDA menu > Tools > View Log) for debug information starting with "LINE Debug UIA Info".
4. The log information will help identify the correct UIA properties for refining the overlay classes.
