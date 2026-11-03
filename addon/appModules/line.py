# LINE Desktop App Module for NVDA
# Provides accessibility enhancements for the LINE desktop application.
# LINE desktop uses Qt6 framework, which exposes UI via UIA on Windows.

import appModuleHandler
from scriptHandler import script
import controlTypes
import api
import ui
import os
import nvwave
import textInfos
import speech
import core
from logHandler import log
from NVDAObjects.UIA import UIA
import UIAHandler
import ctypes
import ctypes.wintypes
import comtypes
import re
import threading
import time
import configparser
import winreg
from ._utils import pickOcrLanguage, setPreferredOcrLanguage
from . import _lineLabels
import addonHandler

addonHandler.initTranslation()

# Translators: The single category under which all of this add-on's scripts
# appear in NVDA's Input Gestures dialog.
SCRIPT_CATEGORY = _("LINE Desktop")


# ---------------------------------------------------------------------------
# LINE installation info — version detection and window type classification
# ---------------------------------------------------------------------------


def _getLineDataDir():
	"""Return the LINE data directory path, or None if not found."""
	localAppData = os.environ.get("LOCALAPPDATA", "")
	if not localAppData:
		return None
	lineDataDir = os.path.join(localAppData, "LINE", "Data")
	if os.path.isdir(lineDataDir):
		return lineDataDir
	return None


def _readLineVersion():
	"""Read the installed LINE version from LINE.ini.

	Returns the version string (e.g. '26.1.0.3865') or None.
	"""
	dataDir = _getLineDataDir()
	if not dataDir:
		return None
	iniPath = os.path.join(dataDir, "LINE.ini")
	if not os.path.isfile(iniPath):
		return None
	try:
		parser = configparser.ConfigParser()
		parser.read(iniPath, encoding="utf-8")
		return parser.get("global", "last_updated_version", fallback=None)
	except Exception:
		return None


def _readLineLanguage():
	"""Read the LINE UI language from installLang.ini.

	Returns the language code (e.g. 'zh-TW') or None.
	"""
	dataDir = _getLineDataDir()
	if not dataDir:
		return None
	iniPath = os.path.join(dataDir, "installLang.ini")
	if not os.path.isfile(iniPath):
		return None
	try:
		parser = configparser.ConfigParser()
		parser.read(iniPath, encoding="utf-8")
		return parser.get("General", "installLang", fallback=None)
	except Exception:
		return None


# ---------------------------------------------------------------------------
# Qt accessibility environment variable management
# ---------------------------------------------------------------------------

_QT_ACCESSIBILITY_ENV_NAME = "QT_ACCESSIBILITY"
_HWND_BROADCAST = 0xFFFF
_WM_SETTINGCHANGE = 0x001A
_SMTO_ABORTIFHUNG = 0x0002


def _isQtAccessibleSet():
	"""Check if QT_ACCESSIBILITY=1 is set in user environment variables.

	Returns True if the variable is set to '1', False otherwise.
	"""
	try:
		with winreg.OpenKey(
			winreg.HKEY_CURRENT_USER,
			"Environment",
			0,
			winreg.KEY_READ,
		) as key:
			value, _ = winreg.QueryValueEx(key, _QT_ACCESSIBILITY_ENV_NAME)
			return str(value) == "1"
	except FileNotFoundError:
		return False
	except Exception:
		log.debugWarning("Failed to read QT_ACCESSIBILITY from registry", exc_info=True)
		return False


def _setQtAccessible(enable=True):
	"""Set or remove QT_ACCESSIBILITY in user environment variables.

	Writes to HKCU\\Environment so the setting persists across reboots.
	Broadcasts WM_SETTINGCHANGE so new processes pick up the change.
	Returns True on success, False on failure.
	"""
	try:
		with winreg.OpenKey(
			winreg.HKEY_CURRENT_USER,
			"Environment",
			0,
			winreg.KEY_SET_VALUE | winreg.KEY_READ,
		) as key:
			if enable:
				winreg.SetValueEx(
					key,
					_QT_ACCESSIBILITY_ENV_NAME,
					0,
					winreg.REG_SZ,
					"1",
				)
				log.info("QT_ACCESSIBILITY=1 set in user environment")
			else:
				try:
					winreg.DeleteValue(key, _QT_ACCESSIBILITY_ENV_NAME)
					log.info("QT_ACCESSIBILITY removed from user environment")
				except FileNotFoundError:
					pass
		# Broadcast environment change to all windows
		ctypes.windll.user32.SendMessageTimeoutW(
			_HWND_BROADCAST,
			_WM_SETTINGCHANGE,
			0,
			"Environment",
			_SMTO_ABORTIFHUNG,
			5000,
			None,
		)
		return True
	except Exception:
		log.warning("Failed to set QT_ACCESSIBILITY in registry", exc_info=True)
		return False


# Window type classification —
# LINE has two main window modes:
#   "AllInOneWindow" — sidebar (chat list) + chat area in one window
#   "ChatWindow" — standalone chat window (no sidebar)
# In ChatWindow mode, all list items are message items (no sidebar heuristic needed).

# Cache to avoid repeated window classification within the same focus cycle.
_windowTypeCache = {
	"hwnd": None,
	"type": None,  # "allinone", "chat", or "unknown"
	"expiresAt": 0.0,
}
_WINDOW_TYPE_CACHE_TTL = 2.0  # seconds


def _classifyLineWindow(hwnd=None):
	"""Classify the current LINE window as 'allinone', 'chat', or 'unknown'.

	Uses window dimensions as heuristic:
	  - AllInOneWindow is wider (has sidebar ~250-350px + chat area)
	  - ChatWindow is narrower (chat area only, typically < 500px wide)
	Also checks UIA tree for sidebar list presence as a secondary signal.
	"""
	global _windowTypeCache

	if hwnd is None:
		hwnd = ctypes.windll.user32.GetForegroundWindow()
	if not hwnd:
		return "unknown"

	now = time.monotonic()
	cache = _windowTypeCache
	if cache["hwnd"] == int(hwnd) and cache["expiresAt"] > now:
		return cache["type"]

	windowType = "unknown"
	try:
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winWidth = rect.right - rect.left

		# AllInOneWindow typically has sidebar (~250-350px) + chat area (~400px+)
		# so total width >= ~600px. ChatWindow is standalone chat, narrower.
		# However, users can resize. Use a soft threshold + child window count.
		if winWidth >= 600:
			# Likely AllInOneWindow, but could be a wide ChatWindow.
			# Check for the presence of multiple pane children (sidebar + chat)
			# via quick UIA check.
			try:
				handler = UIAHandler.handler
				if handler:
					_walker = handler.clientObject.RawViewWalker
					rootCond = handler.clientObject.CreatePropertyCondition(
						30003,
						50033,  # ControlType == Pane
					)
					rootEl = handler.clientObject.ElementFromHandle(hwnd)
					if rootEl:
						panes = rootEl.FindAll(2, rootCond)  # TreeScope.Children=2
						paneCount = panes.Length if panes else 0
						if paneCount >= 2:
							windowType = "allinone"
						else:
							# Single pane or no panes — could be ChatWindow
							# at larger size
							windowType = "allinone" if winWidth >= 700 else "chat"
			except Exception:
				# Fallback: wide window is likely AllInOneWindow
				windowType = "allinone"
		else:
			windowType = "chat"

	except Exception:
		log.debug("_classifyLineWindow failed", exc_info=True)

	cache["hwnd"] = int(hwnd)
	cache["type"] = windowType
	cache["expiresAt"] = now + _WINDOW_TYPE_CACHE_TTL
	log.debug(f"LINE: window type classified as '{windowType}' (hwnd={hwnd})")
	return windowType


# Sound file to play after a message is successfully sent
_SEND_SOUND_PATH = os.path.join(
	os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
	"sounds",
	"sent.wav",
)


# Regex pattern to remove spurious spaces between CJK characters.
# Windows OCR inserts spaces between every CJK character.
# Covers: CJK Unified (\u4E00-\u9FFF), CJK Radicals (\u2E80-\u2FFF),
#         CJK Compatibility (\u3200-\u33FF), CJK Ext-A (\u3400-\u4DBF),
#         Hiragana (\u3040-\u309F), Katakana (\u30A0-\u30FF),
#         Fullwidth (\uFF00-\uFFEF), CJK Compatibility Ideographs (\uF900-\uFAFF),
#         CJK Compatibility Forms (\uFE30-\uFE4F), CJK Symbols (\u3000-\u303F),
#         Bopomofo (\u3100-\u312F, \u31A0-\u31BF)
_CJK_CHAR = (
	"["
	"\u2e80-\u2fff"  # CJK Radicals
	"\u3000-\u303f"  # CJK Symbols and Punctuation
	"\u3040-\u309f"  # Hiragana
	"\u30a0-\u30ff"  # Katakana
	"\u3100-\u312f"  # Bopomofo
	"\u31a0-\u31bf"  # Bopomofo Extended
	"\u3200-\u33ff"  # CJK Compatibility
	"\u3400-\u4dbf"  # CJK Unified Ext A
	"\u4e00-\u9fff"  # CJK Unified Ideographs
	"\uf900-\ufaff"  # CJK Compatibility Ideographs
	"\ufe30-\ufe4f"  # CJK Compatibility Forms
	"\uff00-\uffef"  # Fullwidth Forms
	"]"
)
_CJK_SPACE_RE = re.compile(
	r"(?<=" + _CJK_CHAR + r") (?=" + _CJK_CHAR + r")",
)


def _removeCJKSpaces(text):
	"""Remove spaces between CJK characters inserted by Windows OCR.

	'可 能 因 為' → '可能因為'
	Spaces between Latin characters are preserved.
	"""
	if not text:
		return text
	return _CJK_SPACE_RE.sub("", text)


_VOICE_CALL_CONFIRMATION_START_LABELS = (
	"開始",
	"开始",
	"開始通話",
	"开始通话",
	"開始語音通話",
	"开始语音通话",
	"start",
	"startcall",
	"startvoicecall",
)
_VOICE_CALL_CONFIRMATION_JOIN_LABELS = (
	"加入",
	"參加",
	"参加",
	"加入通話",
	"加入通话",
	"加入語音通話",
	"加入语音通话",
	"參加通話",
	"参加通话",
	"join",
	"joincall",
	"joinvoicecall",
)
_GROUP_VOICE_CALL_CONFIRMATION_HINTS = (
	"群組",
	"群组",
	"群聊",
	"多人",
	"成員",
	"成员",
	"已加入",
	"已參加",
	"已参加",
	"參加中",
	"参加中",
	"正在通話",
	"通話中",
	"group",
)


def _normalizeVoiceCallConfirmationLine(text):
	"""Normalize call confirmation OCR/UIA text for action matching."""
	normalized = _removeCJKSpaces(str(text or "").strip())
	normalized = re.sub(r"\s+", "", normalized)
	return normalized.strip(".,:;!?|/\\()[]{}<>。！？、．：；「」『』（）")


def _matchVoiceCallConfirmationActionLabel(text):
	"""Return the primary voice-call confirmation action represented by text."""
	normalized = _normalizeVoiceCallConfirmationLine(text)
	if not normalized:
		return None

	lower = normalized.lower()
	if normalized in _VOICE_CALL_CONFIRMATION_JOIN_LABELS or lower in _VOICE_CALL_CONFIRMATION_JOIN_LABELS:
		return "join"
	if normalized in _VOICE_CALL_CONFIRMATION_START_LABELS or lower in _VOICE_CALL_CONFIRMATION_START_LABELS:
		return "start"
	return None


def _extractVoiceCallConfirmationActionLabels(text):
	"""Extract start/join action labels from voice-call confirmation OCR text."""
	labels = []
	for rawLine in str(text or "").splitlines():
		label = _matchVoiceCallConfirmationActionLabel(rawLine)
		if label and (not labels or labels[-1] != label):
			labels.append(label)
	return labels


def _isGroupVoiceCallConfirmationText(text, actionLabels=()):
	"""Return True when OCR text describes a group-call confirmation or join prompt."""
	compact = _normalizeVoiceCallConfirmationLine(text)
	if not compact:
		return False

	lower = compact.lower()
	if any(keyword in compact or keyword in lower for keyword in _GROUP_VOICE_CALL_CONFIRMATION_HINTS):
		return True
	if re.search(r"\d+(?:人|位|名).{0,8}(?:加入|參加|参加|通話中|正在通話)", compact):
		return True
	if "join" in set(actionLabels or ()) and ("通話" in compact or "call" in lower):
		return True
	return False


def _getVoiceCallConfirmationState(text):
	"""Return the action/layout state for a LINE voice-call confirmation prompt."""
	actionLabels = _extractVoiceCallConfirmationActionLabels(text)
	isGroup = _isGroupVoiceCallConfirmationText(text, actionLabels)
	action = "join" if "join" in actionLabels else "start"

	if action != "join":
		compact = _normalizeVoiceCallConfirmationLine(text)
		lower = compact.lower()
		hasJoinHint = (
			any(keyword in compact for keyword in ("加入", "參加", "参加", "已加入", "已參加", "已参加"))
			or "join" in lower
		)
		if hasJoinHint and isGroup:
			action = "join"

	return {
		"action": action,
		"isGroup": bool(isGroup),
	}


def _extractVoiceCallConfirmationActionClickPoints(ocrLines, dialogRect):
	"""Map voice-call confirmation action labels to OCR-derived click points."""
	if not ocrLines or not dialogRect:
		return {}

	dialogLeft, _dialogTop, dialogRight, _dialogBottom = dialogRect
	dialogCenterX = (dialogLeft + dialogRight) / 2.0
	matched = {}

	for index, line in enumerate(ocrLines):
		label = _matchVoiceCallConfirmationActionLabel(line.get("text", ""))
		rect = line.get("rect")
		if not label or not rect or not _rectsIntersect(rect, dialogRect):
			continue

		rectLeft, rectTop, rectRight, rectBottom = rect
		if rectRight <= rectLeft or rectBottom <= rectTop:
			continue

		centerX = (rectLeft + rectRight) / 2.0
		centerY = (rectTop + rectBottom) / 2.0
		score = (
			centerY,
			(rectRight - rectLeft) * (rectBottom - rectTop),
			-abs(centerX - dialogCenterX),
			-index,
		)
		current = matched.get(label)
		if current is None or score > current["score"]:
			matched[label] = {
				"clickPoint": (int(centerX), int(centerY)),
				"rect": rect,
				"score": score,
			}

	return {
		label: {
			"clickPoint": data["clickPoint"],
			"rect": data["rect"],
		}
		for label, data in matched.items()
	}


def _selectVoiceCallConfirmationActionTarget(ocrActionTargets, action):
	"""Return the safest OCR target for the primary voice-call dialog action."""
	if not ocrActionTargets:
		return {}

	normalizedAction = "join" if action == "join" else "start"
	preferredLabels = ("start", "join") if normalizedAction == "join" else ("start",)
	for label in preferredLabels:
		target = ocrActionTargets.get(label)
		if target:
			return target
	return {}


def _extractOcrRectLike(obj):
	"""Extract a screen-space rectangle from a UWP OCR line/word object."""

	def _getValue(source, *names):
		if source is None:
			return None
		if isinstance(source, dict):
			for name in names:
				value = source.get(name)
				if value is not None:
					return value
		for name in names:
			value = getattr(source, name, None)
			if value is not None:
				return value
		return None

	def _coerceNumber(value):
		if value is None:
			return None
		for attr in ("value", "Value"):
			nested = getattr(value, attr, None)
			if nested is not None and nested is not value:
				value = nested
				break
		try:
			return float(value)
		except Exception:
			return None

	def _coerceIntTuple(*values):
		coerced = []
		for value in values:
			number = _coerceNumber(value)
			if number is None:
				return None
			coerced.append(int(round(number)))
		return tuple(coerced)

	def _rectFromLeftTopRightBottom(source):
		rect = _coerceIntTuple(
			_getValue(source, "left", "Left", "minX", "MinX", "x1", "X1"),
			_getValue(source, "top", "Top", "minY", "MinY", "y1", "Y1"),
			_getValue(source, "right", "Right", "maxX", "MaxX", "x2", "X2"),
			_getValue(source, "bottom", "Bottom", "maxY", "MaxY", "y2", "Y2"),
		)
		if rect and rect[2] > rect[0] and rect[3] > rect[1]:
			return rect
		return None

	def _rectFromXYWH(source):
		rect = _coerceIntTuple(
			_getValue(source, "x", "X"),
			_getValue(source, "y", "Y"),
			_getValue(source, "width", "Width", "w", "W"),
			_getValue(source, "height", "Height", "h", "H"),
		)
		if rect and rect[2] > 0 and rect[3] > 0:
			left, top, width, height = rect
			return (left, top, left + width, top + height)
		return None

	def _pointFrom(source):
		if source is None:
			return None
		if isinstance(source, (list, tuple)) and len(source) >= 2:
			point = _coerceIntTuple(source[0], source[1])
			if point:
				return point
		point = _coerceIntTuple(
			_getValue(source, "x", "X", "left", "Left"),
			_getValue(source, "y", "Y", "top", "Top"),
		)
		if point:
			return point
		return None

	def _rectFromPoints(points):
		extracted = []
		if isinstance(points, (list, tuple)) and len(points) >= 8:
			flatNumbers = [_coerceNumber(value) for value in points]
			if len(flatNumbers) % 2 == 0 and all(value is not None for value in flatNumbers):
				extracted.extend(
					(
						int(round(flatNumbers[index])),
						int(round(flatNumbers[index + 1])),
					)
					for index in range(0, len(flatNumbers), 2)
				)
		if not extracted:
			for point in points or ():
				parsed = _pointFrom(point)
				if parsed:
					extracted.append(parsed)
		if not extracted:
			return None
		left = min(x for x, _y in extracted)
		top = min(y for _x, y in extracted)
		right = max(x for x, _y in extracted)
		bottom = max(y for _x, y in extracted)
		if right <= left or bottom <= top:
			return None
		return (left, top, right, bottom)

	def _rectFromSequence(source, preferXYWH=False):
		if not isinstance(source, (list, tuple)) or len(source) < 4:
			return None
		pointRect = _rectFromPoints(source)
		if pointRect:
			return pointRect
		flat = _coerceIntTuple(*source[:4])
		if not flat:
			return None
		left, top, third, fourth = flat
		edgeRect = None
		if third > left and fourth > top:
			edgeRect = (left, top, third, fourth)
		sizeRect = None
		if third > 0 and fourth > 0:
			sizeRect = (left, top, left + third, top + fourth)
		for rect in (sizeRect, edgeRect) if preferXYWH else (edgeRect, sizeRect):
			if rect and rect[2] > rect[0] and rect[3] > rect[1]:
				return rect
		return None

	def _rectFromSource(source, preferXYWH=False):
		if source is None:
			return None
		for rect in (
			_rectFromLeftTopRightBottom(source),
			_rectFromXYWH(source),
		):
			if rect:
				return rect
		origin = _getValue(source, "origin", "Origin")
		size = _getValue(source, "size", "Size")
		if origin is not None and size is not None:
			rect = _coerceIntTuple(
				_getValue(origin, "x", "X"),
				_getValue(origin, "y", "Y"),
				_getValue(size, "width", "Width"),
				_getValue(size, "height", "Height"),
			)
			if rect and rect[2] > 0 and rect[3] > 0:
				left, top, width, height = rect
				return (left, top, left + width, top + height)
		for attr in (
			"points",
			"Points",
			"vertices",
			"Vertices",
			"corners",
			"Corners",
			"polygon",
			"Polygon",
			"coordinates",
			"Coordinates",
			"boundingPoints",
			"BoundingPoints",
		):
			rect = _rectFromPoints(_getValue(source, attr))
			if rect:
				return rect
		return _rectFromSequence(source, preferXYWH=preferXYWH)

	for attr in (
		"boundingRect",
		"boundingRectangle",
		"bounding_rect",
		"bounding_rectangle",
		"rect",
		"location",
		"bounds",
		"box",
	):
		rect = _rectFromSource(_getValue(obj, attr), preferXYWH=True)
		if rect:
			return rect

	rect = _rectFromSource(obj)
	if rect:
		return rect

	words = _getValue(obj, "words", "Words")
	try:
		wordIterator = iter(words or ())
	except Exception:
		wordIterator = ()
	wordRects = []
	for word in wordIterator:
		rect = _extractOcrRectLike(word)
		if rect:
			wordRects.append(rect)
	if wordRects:
		return (
			min(rect[0] for rect in wordRects),
			min(rect[1] for rect in wordRects),
			max(rect[2] for rect in wordRects),
			max(rect[3] for rect in wordRects),
		)

	return None


def _extractOcrLines(result):
	"""Return OCR line entries with text and optional screen rectangles."""
	rawLines = getattr(result, "lines", None) or []
	extracted = []
	for rawLine in rawLines:
		text = (getattr(rawLine, "text", "") or "").strip()
		if not text:
			continue
		extracted.append(
			{
				"text": text,
				"rect": _extractOcrRectLike(rawLine),
			},
		)

	if extracted:
		return extracted

	rawWords = getattr(result, "words", None) or []
	lineEnds = []
	for rawLine in rawLines:
		try:
			lineEnd = int(rawLine)
		except Exception:
			continue
		if lineEnd >= 0:
			lineEnds.append(lineEnd)

	wordEntries = []
	for rawWord in rawWords:
		try:
			offset = int(getattr(rawWord, "offset"))
			left = int(getattr(rawWord, "left"))
			top = int(getattr(rawWord, "top"))
			width = int(getattr(rawWord, "width"))
			height = int(getattr(rawWord, "height"))
		except Exception:
			continue
		if width <= 0 or height <= 0:
			continue
		wordEntries.append(
			{
				"offset": offset,
				"rect": (left, top, left + width, top + height),
			},
		)

	text = getattr(result, "text", "") or ""
	if lineEnds and wordEntries and text:
		rebuilt = []
		lineStart = 0
		wordIndex = 0
		wordCount = len(wordEntries)
		textLen = len(text)
		for rawLineEnd in lineEnds:
			lineEnd = max(lineStart, min(int(rawLineEnd), textLen))
			lineText = text[lineStart:lineEnd].rstrip("\r\n")
			lineRects = []
			while wordIndex < wordCount and wordEntries[wordIndex]["offset"] < rawLineEnd:
				if wordEntries[wordIndex]["offset"] >= lineStart:
					lineRects.append(wordEntries[wordIndex]["rect"])
				wordIndex += 1
			if lineText.strip():
				rect = None
				if lineRects:
					rect = (
						min(item[0] for item in lineRects),
						min(item[1] for item in lineRects),
						max(item[2] for item in lineRects),
						max(item[3] for item in lineRects),
					)
				rebuilt.append(
					{
						"text": lineText.strip(),
						"rect": rect,
					},
				)
			lineStart = lineEnd
		if rebuilt:
			return rebuilt

	text = _removeCJKSpaces((getattr(result, "text", "") or "").strip())
	return [{"text": line.strip(), "rect": None} for line in text.splitlines() if line.strip()]


def _invokeUIAInvokePattern(pattern):
	"""Invoke a UIA InvokePattern without relying on generated comtypes stubs."""
	if not pattern:
		return False

	invoke = getattr(pattern, "Invoke", None)
	if callable(invoke):
		invoke()
		return True

	queryInterface = getattr(pattern, "QueryInterface", None)
	if not callable(queryInterface):
		return False

	comMethod = getattr(comtypes, "COMMETHOD", None)
	if comMethod is None:
		from comtypes import COMMETHOD as comMethod

	hresult = getattr(ctypes, "HRESULT", ctypes.c_long)

	class _IUIAutomationInvokePattern(comtypes.IUnknown):
		_iid_ = comtypes.GUID("{FB377FBE-8EA6-46D5-9C73-6499642D3059}")
		_methods_ = [comMethod([], hresult, "Invoke")]

	queryInterface(_IUIAutomationInvokePattern).Invoke()
	return True


def _tryInvokeUIAElement(element):
	"""Return True when a UIA element successfully activates InvokePattern."""
	if not element:
		return False

	pattern = element.GetCurrentPattern(10000)  # InvokePattern
	if not pattern:
		return False
	return _invokeUIAInvokePattern(pattern)


# ---------------------------------------------------------------------------
# LINE's own popup menus
# ---------------------------------------------------------------------------
#
# LINE 26.6.0.4004+ opens the message menu (Shift+F10) and the chat's "More
# options" menu as popups of named buttons that take the keyboard focus. The
# helpers below pick an item from such a menu by its name. They do raw UIA
# work and wait for LINE, so they run on worker threads (_runOnWorkerThread).

_UIA_CONTROL_TYPE_PROPERTY_ID = 30003
_UIA_NAME_PROPERTY_ID = 30005
_UIA_BUTTON_CONTROL_TYPE_ID = 50000
_UIA_EDIT_CONTROL_TYPE_ID = 50004

# _chooseNativeMenuItem results
_MENU_ITEM_CLICKED = "clicked"
_MENU_ITEM_NOT_FOUND = "notFound"  # the expected menu opened without that item
_MENU_NOT_OPENED = "noMenu"
_MENU_ABANDONED = "abandoned"  # LINE lost the foreground; nothing was clicked


def _waitForModifierRelease(timeoutSec=2.0):
	"""Worker thread: wait until Shift, Control, Alt and Windows are released.

	Keys the add-on sends while the user still holds the shortcut's modifiers
	combine with them: Shift+F10 sent during Control+C reaches LINE as
	Control+Shift+F10, and during NVDA+Windows+R as Windows+Shift+F10.
	"""
	getAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState
	modifiers = (0x10, 0x11, 0x12, 0x5B, 0x5C)  # Shift, Control, Alt, left/right Windows
	deadline = time.monotonic() + timeoutSec
	while time.monotonic() < deadline:
		if not any(getAsyncKeyState(vk) & 0x8000 for vk in modifiers):
			return True
		time.sleep(0.05)
	return False


def _isLineForeground(processId):
	"""True when the foreground window belongs to the LINE process ``processId``."""
	hwnd = ctypes.windll.user32.GetForegroundWindow()
	if not hwnd:
		return False
	foregroundProcessId = ctypes.wintypes.DWORD()
	ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(foregroundProcessId))
	return foregroundProcessId.value == processId


def _clickElementCenter(element):
	"""Click the middle of a UIA element, or use its InvokePattern if it has no rect.

	A real click comes first because LINE's Qt buttons are known to act on it,
	while an advertised InvokePattern may have no handler behind it.
	"""
	try:
		rect = element.CurrentBoundingRectangle
		if rect.right > rect.left and rect.bottom > rect.top:
			return _sendMouseClick(
				int((rect.left + rect.right) / 2),
				int((rect.top + rect.bottom) / 2),
			)
	except Exception:
		log.debug("LINE: reading a UIA element's rect failed", exc_info=True)
	try:
		return _tryInvokeUIAElement(element)
	except Exception:
		log.debug("LINE: invoking a UIA element failed", exc_info=True)
		return False


def _findNamedButton(handler, hwnd, names):
	"""Return the first button in window ``hwnd`` whose UIA name is one of ``names``."""
	client = handler.clientObject
	root = client.ElementFromHandle(hwnd)
	if not root:
		return None
	for name in names:
		condition = client.CreateAndCondition(
			client.CreatePropertyCondition(_UIA_CONTROL_TYPE_PROPERTY_ID, _UIA_BUTTON_CONTROL_TYPE_ID),
			client.CreatePropertyCondition(_UIA_NAME_PROPERTY_ID, name),
		)
		element = root.FindFirst(UIAHandler.TreeScope_Descendants, condition)
		if element:
			return element
	return None


def _waitForFocusedMenuItem(handler, processId, previousRuntimeId, timeoutSec=1.5):
	"""Return the LINE button that takes the focus once a popup menu opens, or None.

	Only buttons of the LINE process ``processId`` count, so a button focused
	in another application the user switched to is never taken for the menu.
	"""
	deadline = time.monotonic() + timeoutSec
	while time.monotonic() < deadline:
		try:
			element = handler.clientObject.GetFocusedElement()
			if (
				element
				and element.CurrentProcessId == processId
				and element.CurrentControlType == _UIA_BUTTON_CONTROL_TYPE_ID
				and _getElementRuntimeId(element) != previousRuntimeId
			):
				return element
		except Exception:
			log.debug("LINE: polling the focus for a popup menu failed", exc_info=True)
		time.sleep(0.05)
	return None


def _findMenuItem(handler, focusedItem, table, key):
	"""Return ``(item, isKnownMenu)`` for the menu that holds ``focusedItem``.

	LINE opens each menu in a window of its own, so the search climbs only to
	the nearest ancestor with a window handle (at most a few levels) and looks
	at the buttons below it. ``isKnownMenu`` says whether any of them is a
	label from ``table``, i.e. whether this is the menu the caller expected.
	"""
	walker = handler.clientObject.RawViewWalker
	root = focusedItem
	for _level in range(4):
		if root.CurrentNativeWindowHandle:
			break
		parent = walker.GetParentElement(root)
		if not parent:
			break
		root = parent
	buttons = root.FindAll(
		UIAHandler.TreeScope_Descendants,
		handler.clientObject.CreatePropertyCondition(
			_UIA_CONTROL_TYPE_PROPERTY_ID,
			_UIA_BUTTON_CONTROL_TYPE_ID,
		),
	)
	isKnownMenu = False
	for index in range(buttons.Length if buttons else 0):
		button = buttons.GetElement(index)
		labelKey = _lineLabels.labelKeyOf(button.CurrentName, table)
		if labelKey == key:
			return button, True
		if labelKey is not None:
			isKnownMenu = True
	return None, isKnownMenu


def _chooseNativeMenuItem(handler, processId, previousRuntimeId, table, key):
	"""Worker thread: click ``key``'s item in the LINE popup menu that is opening.

	``processId`` is LINE's; nothing outside that process is searched or
	clicked. Returns ``(status, itemName)``: _MENU_ITEM_CLICKED with the name
	LINE gives the item, _MENU_ITEM_NOT_FOUND when the expected menu opened
	without that item (it is still open), _MENU_NOT_OPENED when no LINE button
	took the focus in time or the menu is not one ``table`` describes (left open
	for the user), or _MENU_ABANDONED when LINE lost the foreground, e.g. the
	user switched applications meanwhile.
	"""
	focusedItem = _waitForFocusedMenuItem(handler, processId, previousRuntimeId)
	if focusedItem is None:
		return (_MENU_NOT_OPENED if _isLineForeground(processId) else _MENU_ABANDONED), ""
	item, isKnownMenu = _findMenuItem(handler, focusedItem, table, key)
	if item is None:
		return (_MENU_ITEM_NOT_FOUND if isKnownMenu else _MENU_NOT_OPENED), ""
	itemName = item.CurrentName or key
	# The search took a moment; make sure the menu is still LINE's active
	# popup before moving the mouse there.
	try:
		stillInLine = (
			_isLineForeground(processId)
			and item.CurrentProcessId == processId
			and handler.clientObject.GetFocusedElement().CurrentProcessId == processId
		)
	except Exception:
		stillInLine = False
	if not stillInLine:
		return _MENU_ABANDONED, ""
	if not _clickElementCenter(item):
		return _MENU_ITEM_NOT_FOUND, ""
	return _MENU_ITEM_CLICKED, itemName


def _normalizeRuntimeId(runtimeId):
	if runtimeId is None:
		return None
	try:
		return tuple(int(part) for part in runtimeId)
	except Exception:
		try:
			return tuple(runtimeId)
		except Exception:
			return None


def _getElementRuntimeId(element):
	if element is None:
		return None
	try:
		return _normalizeRuntimeId(element.GetRuntimeId())
	except Exception:
		return None


def _queueUiMessage(text):
	"""Speak a message safely from any thread.

	ui.message must not be called from worker threads; this queues it onto
	NVDA's main event queue instead. Also safe to call from the main thread.
	"""
	try:
		import queueHandler

		queueHandler.queueFunction(queueHandler.eventQueue, ui.message, text)
	except Exception:
		try:
			core.callLater(0, ui.message, text)
		except Exception:
			pass


def _coInitializeWorker():
	"""Best-effort COM init for worker threads that touch UIA/OCR."""
	try:
		ctypes.windll.ole32.CoInitializeEx(None, 0)  # COINIT_MULTITHREADED
	except Exception:
		pass


def _waitForForegroundWindow(hwnd, timeoutSec=1.0):
	"""Wait until ``hwnd`` is the foreground window (worker threads only).

	Returns as soon as the window actually reaches the foreground instead of
	sleeping a fixed interval; gives up after ``timeoutSec``. Blocking, so
	never call from NVDA's main thread.
	"""
	try:
		from utils.blockUntilConditionMet import blockUntilConditionMet

		blockUntilConditionMet(
			getValue=lambda: ctypes.windll.user32.GetForegroundWindow() == hwnd,
			giveUpAfterSeconds=timeoutSec,
			intervalBetweenSeconds=0.05,
		)
	except Exception:
		time.sleep(0.3)


# Pointer-sized integer for the SendInput dwExtraInfo field.
if ctypes.sizeof(ctypes.c_void_p) == 8:
	_ULONG_PTR = ctypes.c_ulonglong
else:
	_ULONG_PTR = ctypes.c_ulong


class _MOUSEINPUT(ctypes.Structure):
	_fields_ = (
		("dx", ctypes.wintypes.LONG),
		("dy", ctypes.wintypes.LONG),
		("mouseData", ctypes.wintypes.DWORD),
		("dwFlags", ctypes.wintypes.DWORD),
		("time", ctypes.wintypes.DWORD),
		("dwExtraInfo", _ULONG_PTR),
	)


class _INPUT(ctypes.Structure):
	class _INPUTunion(ctypes.Union):
		_fields_ = (("mi", _MOUSEINPUT),)

	_anonymous_ = ("u",)
	_fields_ = (
		("type", ctypes.wintypes.DWORD),
		("u", _INPUTunion),
	)


_INPUT_MOUSE = 0
_MOUSEEVENTF_LEFTDOWN = 0x0002
_MOUSEEVENTF_LEFTUP = 0x0004
_MOUSEEVENTF_RIGHTDOWN = 0x0008
_MOUSEEVENTF_RIGHTUP = 0x0010


def _sendMouseClick(x, y, button="left"):
	"""Move the cursor to (x, y) and click atomically via SendInput.

	SendInput delivers the button-down and button-up events in a single
	call, so no inter-event ``time.sleep`` is needed. This keeps clicks
	reliable without ever blocking the calling (often main) thread, which
	the previous SetCursorPos + mouse_event + sleep pattern did. Safe on
	any thread.

	The whole sequence runs under mouseHandler.ignoreInjection() so NVDA's
	mouse hook does not react to the synthetic move/click (e.g. speaking the
	object under the mouse), and the cursor is restored afterwards. The
	injected events are queued
	in order, so restoring immediately after SendInput cannot retarget the
	click.
	"""
	try:
		import mouseHandler
		import winUser

		originalPos = None
		try:
			originalPos = winUser.getCursorPos()
		except Exception:
			pass
		with mouseHandler.ignoreInjection():
			try:
				ctypes.windll.user32.SetCursorPos(int(x), int(y))
				if button == "right":
					downFlag, upFlag = _MOUSEEVENTF_RIGHTDOWN, _MOUSEEVENTF_RIGHTUP
				else:
					downFlag, upFlag = _MOUSEEVENTF_LEFTDOWN, _MOUSEEVENTF_LEFTUP
				events = (_INPUT * 2)()
				for i, flag in enumerate((downFlag, upFlag)):
					events[i].type = _INPUT_MOUSE
					events[i].mi = _MOUSEINPUT(0, 0, 0, flag, 0, 0)
				ctypes.windll.user32.SendInput(2, ctypes.byref(events), ctypes.sizeof(_INPUT))
			finally:
				if originalPos is not None:
					try:
						winUser.setCursorPos(*originalPos)
					except Exception:
						pass
		return True
	except Exception:
		log.debugWarning("LINE: _sendMouseClick failed", exc_info=True)
		return False


def _runOnWorkerThread(work, onDone=None, taskName="task"):
	"""Run ``work()`` on a daemon thread, then ``onDone(result)`` on the main thread.

	``onDone`` receives None when ``work`` raised. Used to keep screen OCR and
	raw-COM UIA scans (which block for seconds) off NVDA's main thread.
	"""

	def _worker():
		_coInitializeWorker()
		try:
			result = work()
		except Exception:
			log.debugWarning(f"LINE: background {taskName} failed", exc_info=True)
			result = None
		if onDone is not None:
			core.callLater(0, onDone, result)

	threading.Thread(target=_worker, daemon=True, name=f"lineAddon-{taskName}").start()


# Flag to suppress addon while a file dialog is open
_suppressAddon = False

# Image description backend providers exposed in the settings panel.
_IMAGE_DESCRIPTION_PROVIDER_GOOGLE = "google"
_IMAGE_DESCRIPTION_PROVIDER_OLLAMA = "ollama"
_IMAGE_DESCRIPTION_PROVIDER_NVIDIA = "nvidia"
_IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS = "pollinations"
_IMAGE_DESCRIPTION_PROVIDER_OPENAI = "openai"
_IMAGE_DESCRIPTION_PROVIDER_MISTRAL = "mistral"
_IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS = (
	_IMAGE_DESCRIPTION_PROVIDER_GOOGLE,
	_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
	_IMAGE_DESCRIPTION_PROVIDER_NVIDIA,
	_IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS,
	_IMAGE_DESCRIPTION_PROVIDER_OPENAI,
	_IMAGE_DESCRIPTION_PROVIDER_MISTRAL,
)
_IMAGE_DESCRIPTION_DEFAULT_PROVIDER = _IMAGE_DESCRIPTION_PROVIDER_GOOGLE
_IMAGE_DESCRIPTION_PROVIDER_LABELS = {
	_IMAGE_DESCRIPTION_PROVIDER_GOOGLE: "Google AI",
	_IMAGE_DESCRIPTION_PROVIDER_OLLAMA: "Ollama Cloud",
	_IMAGE_DESCRIPTION_PROVIDER_NVIDIA: "NVIDIA NIM",
	_IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS: "Pollinations.AI",
	_IMAGE_DESCRIPTION_PROVIDER_OPENAI: "OpenAI",
	_IMAGE_DESCRIPTION_PROVIDER_MISTRAL: "Mistral AI",
}
_IMAGE_DESCRIPTION_USER_PROVIDER_FILENAME = "line_desktop_image_provider.txt"

# API keys used by NVDA+Windows+I image description.
# The bundled convenience keys are encrypted with a PBKDF2-HMAC-SHA256 derived
# keystream plus HMAC integrity tag, using a per-blob random salt, so the
# ciphertext is opaque to naive string / pattern extraction. They are NOT
# cryptographic secrets: this is an open-source addon and a determined reader
# can still recover them by running the decryption code. The purpose is to
# raise the bar for casual extraction and to spare end-users from having to
# obtain their own key. Users can override either key at any time through the
# settings panel; their key is persisted (also encrypted) under NVDA's user
# config directory.
_IMAGE_DESCRIPTION_DEFAULT_KEY_BLOB = (
	"g+Ku1l+8YmbpO4/JPwy+ZyMlZw4Nfm9gO5bbn8K/vPkz7VFoMRpiFsx2hgfKpUqbxmWqQGo2h8Ph7YjZljEEFkmc3+HlxuE="
)
_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_KEY_BLOB = "oLtfHW4hFhTMLQ0mKKcEqd70nU8Z9EsEjjSKfueLCUCnJD9oee2CTd3GtTi0LyS6ZJMuee/jIxFiXDH9kzdyFOMVRfIjMRxJzqZxnccS+p5B1gjbWxYyMLY="
_IMAGE_DESCRIPTION_NVIDIA_DEFAULT_KEY_BLOB = "z4L9t8i44kXzffkss7ukUYezYDRvTRxxgKURaLeyOl2Ea8kLN4dNweWybeuf4F8SD6I8ArEWqbvry5BB2o4MxerrJgd0OZD8pQNm9Nw6P7RE2mDI1xktlM4bj6XpwU9G/0wFnMMn"
_IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_KEY_BLOB = (
	"9djyxTd4B3//wSRROFV/x5bpuusesK71/tOYdrtn4jf6YeRcniZVExF4//O0M+0G61hO0fkmB99ToQl9b3awbnwgPw=="
)
_IMAGE_DESCRIPTION_OPENAI_DEFAULT_KEY_BLOB = (
	"lAm3ikqneka+uk2yk1QsqH08/47ADNiQ72/WbiqYSBpFL5WP4DlJbj4iir03qcxpcA0s1Bq92JPom2+5LrbpAqNSYngmW9aIrPQ"
	"svswUoTj/3/DGp8+xOWf/vEIHPyTdQgafiJDlqlnf4/PI9oPSmmT0f4qRQn99TZzrdl+VbyQlXRoc+7N9cikjINQyT/N6j0sxA0"
	"9N1noxqL6EBcouFFUAe6xtfy9lWzamsggrc/BGwQeCMiojshNKZQY6xfBg1bnF/Q=="
)
_IMAGE_DESCRIPTION_MISTRAL_DEFAULT_KEY_BLOB = (
	"gv024MC6257LD65vvtsLcSv06bdTsV1dCeRDwhaaob/PHZFHRusKQixxW1RHSFIDfJRfF937g9Z7Df1VLkazZg=="
)
_IMAGE_DESCRIPTION_USER_KEY_FILENAME = "line_desktop_image_api_key.dat"
_IMAGE_DESCRIPTION_USER_OLLAMA_KEY_FILENAME = "line_desktop_ollama_api_key.dat"
_IMAGE_DESCRIPTION_USER_NVIDIA_KEY_FILENAME = "line_desktop_nvidia_api_key.dat"
_IMAGE_DESCRIPTION_USER_POLLINATIONS_KEY_FILENAME = "line_desktop_pollinations_api_key.dat"
_IMAGE_DESCRIPTION_USER_OPENAI_KEY_FILENAME = "line_desktop_openai_api_key.dat"
_IMAGE_DESCRIPTION_USER_MISTRAL_KEY_FILENAME = "line_desktop_mistral_api_key.dat"
_IMAGE_DESCRIPTION_USER_MODEL_FILENAME = "line_desktop_image_model.txt"
_IMAGE_DESCRIPTION_USER_OLLAMA_MODEL_FILENAME = "line_desktop_ollama_model.txt"
_IMAGE_DESCRIPTION_USER_NVIDIA_MODEL_FILENAME = "line_desktop_nvidia_model.txt"
_IMAGE_DESCRIPTION_USER_POLLINATIONS_MODEL_FILENAME = "line_desktop_pollinations_model.txt"
_IMAGE_DESCRIPTION_USER_OPENAI_MODEL_FILENAME = "line_desktop_openai_model.txt"
_IMAGE_DESCRIPTION_USER_MISTRAL_MODEL_FILENAME = "line_desktop_mistral_model.txt"
_IMAGE_DESCRIPTION_USER_PROMPT_FILENAME = "line_desktop_image_prompt.txt"
_IMAGE_DESCRIPTION_USER_MAX_TOKENS_FILENAME = "line_desktop_image_max_tokens.txt"
# Default Google model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_DEFAULT_MODEL = "gemini-flash-lite-latest"
# Default Ollama Cloud model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL = "gemma4:31b-cloud"
# Default NVIDIA NIM model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_NVIDIA_DEFAULT_MODEL = "google/diffusiongemma-26b-a4b-it"
# Models exposed in the settings panel dropdown. The order here is the order
# the user sees. Strings are the actual model IDs sent to the API endpoint.
# VERIFY all IDs against: GET https://generativelanguage.googleapis.com/v1beta/models
_IMAGE_DESCRIPTION_AVAILABLE_MODELS = (
	"gemini-3-flash-preview",
	"gemini-3.1-flash-lite",
	"gemini-3.5-flash",
	"gemini-3.5-flash-lite",
	"gemini-3.6-flash",
	"gemini-3.7-flash",
	"gemini-3.8-flash",
	"gemini-flash-latest",
	"gemini-flash-lite-latest",
	"gemini-2.5-flash",
	"gemini-2.5-flash-lite",
	"gemma-4-26b-a4b-it",
	"gemma-4-31b-it",
)
# Vision-capable models commonly available on Ollama Cloud. Users can pick
# whichever they have access to; unavailable IDs surface as an HTTP error.
_IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS = ("gemma4:31b-cloud",)
# Vision-capable models exposed for NVIDIA NIM. Users can pick whichever they have
# access to; unavailable IDs surface as an HTTP error.
_IMAGE_DESCRIPTION_NVIDIA_AVAILABLE_MODELS = (
	"deepseek-ai/deepseek-v4.1-flash",
	"google/gemma-4-31b-it",
	"google/diffusiongemma-26b-a4b-it",
	"meta/muse-glimmer-30b",
	"minimaxai/minimax-m3.1-flash",
	"nvidia/ising-calibration-1.5-31b",
	"nvidia/nemotron-3-nano-omni-30b-a3b-reasoning",
	"nvidia/nemotron-3.5-super-vl-120b-a12b",
	"stepfun-ai/step-5-preview",
	"z-ai/glm-5.3-flash",
)
# Default Pollinations.AI model used when the user has not picked one.
_IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_MODEL = "gpt-5.6-luna"
# Vision-capable Pollinations.AI models. Strings are the actual model IDs sent
# to the API; the user-facing labels live in _IMAGE_DESCRIPTION_POLLINATIONS_MODEL_LABELS.
_IMAGE_DESCRIPTION_POLLINATIONS_AVAILABLE_MODELS = (
	"deepseek/deepseek-v4.1-flash",
	"openai/gpt-5.3-codex",
	"gpt-5.4",
	"gpt-5.4-mini",
	"openai",
	"openai-large",
	"gpt-5.6-sol",
	"gpt-5.6-terra",
	"gpt-5.6-luna",
	"openai/gpt-6-astra",
	"openai/gpt-6-sol",
	"openai/gpt-6-luna",
	"openai/gpt-6.1-sol",
	"grok",
	"grok-large",
	"grok-4.6",
	"minimax",
	"mistral-large",
	"kimi",
	"kimi-code",
	"kimi-k3",
	"z-ai/glm-5.3-flash",
)
# Display labels for Pollinations.AI models, surfaced in the settings dropdown.
_IMAGE_DESCRIPTION_POLLINATIONS_MODEL_LABELS = {
	"deepseek/deepseek-v4.1-flash": "DeepSeek V4.1 Flash",
	"openai/gpt-5.3-codex": "GPT-5.3 Codex",
	"gpt-5.4": "gpt-5.4",
	"gpt-5.4-mini": "GPT-5.4 Mini",
	"openai": "GPT-5.4 Nano",
	"openai-large": "GPT-5.5",
	"gpt-5.6-sol": "GPT-5.6 Sol",
	"gpt-5.6-terra": "GPT-5.6 Terra",
	"gpt-5.6-luna": "GPT-5.6 Luna",
	"openai/gpt-6-astra": "GPT-6 Astra",
	"openai/gpt-6-sol": "GPT-6 Sol",
	"openai/gpt-6-luna": "GPT-6 Luna",
	"openai/gpt-6.1-sol": "GPT-6.1 Sol",
	"grok": "Grok 4.20 Non-Reasoning",
	"grok-large": "Grok 4.3",
	"grok-4.6": "Grok 4.6",
	"minimax": "MiniMax M3",
	"mistral-large": "Mistral Large 3",
	"kimi": "Moonshot Kimi K2.6",
	"kimi-code": "Moonshot Kimi K2.7 Code",
	"kimi-k3": "Moonshot Kimi K3",
	"z-ai/glm-5.3-flash": "Z.ai GLM-5.3 Flash",
}
# Default OpenAI model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_OPENAI_DEFAULT_MODEL = "gpt-5.6-terra"
# Vision-capable OpenAI models exposed in the settings panel.
_IMAGE_DESCRIPTION_OPENAI_AVAILABLE_MODELS = (
	"gpt-6-astra",
	"gpt-6-sol",
	"gpt-6-luna",
	"gpt-5.6-sol",
	"gpt-5.6-terra",
	"gpt-5.6-luna",
	"gpt-5.5",
	"gpt-5.4",
	"gpt-5.4-mini",
	"gpt-5.4-nano",
)
# Default Mistral AI model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_MISTRAL_DEFAULT_MODEL = "mistral-medium-latest"
# Vision-capable Mistral AI models exposed in the settings panel.
_IMAGE_DESCRIPTION_MISTRAL_AVAILABLE_MODELS = (
	"mistral-large-latest",
	"mistral-medium-latest",
	"mistral-small-latest",
)
_IMAGE_DESCRIPTION_ENDPOINT = (
	"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
)
_IMAGE_DESCRIPTION_OLLAMA_ENDPOINT = "https://ollama.com/api/chat"
_IMAGE_DESCRIPTION_NVIDIA_ENDPOINT = "https://integrate.api.nvidia.com/v1/chat/completions"
_IMAGE_DESCRIPTION_POLLINATIONS_ENDPOINT = "https://gen.pollinations.ai/v1/chat/completions"
_IMAGE_DESCRIPTION_OPENAI_ENDPOINT = "https://api.openai.com/v1/chat/completions"
_IMAGE_DESCRIPTION_MISTRAL_ENDPOINT = "https://api.mistral.ai/v1/chat/completions"
_IMAGE_DESCRIPTION_DEFAULT_PROMPT = (
	"Describe this image succinctly, but in as much detail as possible. "
	"If there is text, ensure it is included in your response exactly as shown."
)
# Maximum length accepted from user-supplied prompts, to avoid pathological
# payloads being sent to the API.
_IMAGE_DESCRIPTION_PROMPT_MAX_LEN = 2000
# User-configurable max output tokens used by every backend that accepts the
# parameter. Bounded so the user cannot accidentally send pathological values.
_IMAGE_DESCRIPTION_DEFAULT_MAX_TOKENS = 700
_IMAGE_DESCRIPTION_MIN_MAX_TOKENS = 50
_IMAGE_DESCRIPTION_MAX_MAX_TOKENS = 4000

_IMAGE_API_KEY_SALT_LEN = 16
_IMAGE_API_KEY_MAC_LEN = 16
_IMAGE_API_KEY_PBKDF2_ITERS = 100000

# Sentinel indicating the effective API key has not been computed yet.
_NOT_COMPUTED = object()
# Populated once at AppModule.__init__ via _initEffectiveImageApiKey().
# After that, _getEffectiveImageApiKey() is a pure memory read (no PBKDF2).
_cachedEffectiveImageApiKey = _NOT_COMPUTED
_cachedEffectiveOllamaApiKey = _NOT_COMPUTED
_cachedEffectiveNvidiaApiKey = _NOT_COMPUTED
_cachedEffectivePollinationsApiKey = _NOT_COMPUTED
_cachedEffectiveOpenaiApiKey = _NOT_COMPUTED
_cachedEffectiveMistralApiKey = _NOT_COMPUTED
_cachedEffectiveImageProvider = _NOT_COMPUTED
_cachedEffectiveOllamaModel = _NOT_COMPUTED
_cachedEffectiveNvidiaModel = _NOT_COMPUTED
_cachedEffectivePollinationsModel = _NOT_COMPUTED
_cachedEffectiveOpenaiModel = _NOT_COMPUTED
_cachedEffectiveMistralModel = _NOT_COMPUTED


def _deriveImageApiKeyMaterial(salt, length):
	import hashlib
	import hmac
	import struct

	# Passphrase is assembled at call time rather than stored as a single
	# module-level literal, so a plain string dump of the .pyc is less
	# revealing about what the blob decrypts to.
	_p = (b"nvda", b"-line-", b"desktop-", b"image-", b"api-", b"2026-", b"v2")
	passphrase = b"".join(_p)
	master = hashlib.pbkdf2_hmac(
		"sha256",
		passphrase,
		salt,
		_IMAGE_API_KEY_PBKDF2_ITERS,
		dklen=32,
	)
	stream = bytearray()
	counter = 0
	while len(stream) < length:
		counter += 1
		stream.extend(
			hmac.new(master, struct.pack(">I", counter), hashlib.sha256).digest(),
		)
	return bytes(stream[:length]), master


def _obfuscateImageApiKey(plain):
	import base64
	import hashlib
	import hmac

	data = plain.encode("utf-8")
	salt = os.urandom(_IMAGE_API_KEY_SALT_LEN)
	stream, master = _deriveImageApiKeyMaterial(salt, len(data))
	cipher = bytes(a ^ b for a, b in zip(data, stream))
	mac = hmac.new(master, salt + cipher, hashlib.sha256).digest()[:_IMAGE_API_KEY_MAC_LEN]
	return base64.b64encode(salt + mac + cipher).decode("ascii")


def _deobfuscateImageApiKey(blob):
	import base64
	import hashlib
	import hmac

	if not blob:
		return None
	try:
		raw = base64.b64decode(blob.encode("ascii"))
		if len(raw) < _IMAGE_API_KEY_SALT_LEN + _IMAGE_API_KEY_MAC_LEN:
			return None
		salt = raw[:_IMAGE_API_KEY_SALT_LEN]
		mac = raw[_IMAGE_API_KEY_SALT_LEN : _IMAGE_API_KEY_SALT_LEN + _IMAGE_API_KEY_MAC_LEN]
		cipher = raw[_IMAGE_API_KEY_SALT_LEN + _IMAGE_API_KEY_MAC_LEN :]
		stream, master = _deriveImageApiKeyMaterial(salt, len(cipher))
		expected = hmac.new(master, salt + cipher, hashlib.sha256).digest()[:_IMAGE_API_KEY_MAC_LEN]
		if not hmac.compare_digest(mac, expected):
			return None
		plain = bytes(a ^ b for a, b in zip(cipher, stream)).decode("utf-8")
	except Exception:
		return None
	return plain or None


def _getImageApiKeyStorePath():
	"""Return the filesystem path for the user-supplied API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_KEY_FILENAME)


def getUserImageApiKey():
	"""Return the plain API key previously set by the user, or None."""
	path = _getImageApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserImageApiKey(plain):
	"""Persist a user-supplied API key (obfuscated). Empty/None clears it."""
	global _cachedEffectiveImageApiKey
	path = _getImageApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveImageApiKey = _deobfuscateImageApiKey(_IMAGE_DESCRIPTION_DEFAULT_KEY_BLOB)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectiveImageApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user API key: {e}", exc_info=True)
		return False


def _initEffectiveImageApiKey():
	"""Decrypt and cache the effective Google API key. Called once at addon startup."""
	global _cachedEffectiveImageApiKey
	userKey = getUserImageApiKey()
	_cachedEffectiveImageApiKey = userKey or _deobfuscateImageApiKey(_IMAGE_DESCRIPTION_DEFAULT_KEY_BLOB)


def _getEffectiveImageApiKey():
	"""Return the cached Google API key; falls back to lazy init if not yet computed."""
	if _cachedEffectiveImageApiKey is _NOT_COMPUTED:
		_initEffectiveImageApiKey()
	return _cachedEffectiveImageApiKey


def _getOllamaApiKeyStorePath():
	"""Return the filesystem path for the user-supplied Ollama API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_OLLAMA_KEY_FILENAME)


def getUserOllamaApiKey():
	"""Return the plain Ollama API key previously set by the user, or None."""
	path = _getOllamaApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Ollama API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserOllamaApiKey(plain):
	"""Persist a user-supplied Ollama API key (obfuscated). Empty/None clears it."""
	global _cachedEffectiveOllamaApiKey
	path = _getOllamaApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveOllamaApiKey = _deobfuscateImageApiKey(_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_KEY_BLOB)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectiveOllamaApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Ollama API key: {e}", exc_info=True)
		return False


def _initEffectiveOllamaApiKey():
	"""Decrypt and cache the effective Ollama API key."""
	global _cachedEffectiveOllamaApiKey
	userKey = getUserOllamaApiKey()
	_cachedEffectiveOllamaApiKey = userKey or _deobfuscateImageApiKey(
		_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_KEY_BLOB,
	)


def _getEffectiveOllamaApiKey():
	"""Return the cached Ollama API key; falls back to lazy init if not yet computed."""
	if _cachedEffectiveOllamaApiKey is _NOT_COMPUTED:
		_initEffectiveOllamaApiKey()
	return _cachedEffectiveOllamaApiKey


def _getNvidiaApiKeyStorePath():
	"""Return the filesystem path for the user-supplied NVIDIA NIM API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_NVIDIA_KEY_FILENAME)


def getUserNvidiaApiKey():
	"""Return the plain NVIDIA NIM API key previously set by the user, or None."""
	path = _getNvidiaApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user NVIDIA API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserNvidiaApiKey(plain):
	"""Persist a user-supplied NVIDIA NIM API key (obfuscated). Empty/None clears it."""
	global _cachedEffectiveNvidiaApiKey
	path = _getNvidiaApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveNvidiaApiKey = _deobfuscateImageApiKey(_IMAGE_DESCRIPTION_NVIDIA_DEFAULT_KEY_BLOB)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectiveNvidiaApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user NVIDIA API key: {e}", exc_info=True)
		return False


def _initEffectiveNvidiaApiKey():
	"""Decrypt and cache the effective NVIDIA NIM API key."""
	global _cachedEffectiveNvidiaApiKey
	userKey = getUserNvidiaApiKey()
	_cachedEffectiveNvidiaApiKey = userKey or _deobfuscateImageApiKey(
		_IMAGE_DESCRIPTION_NVIDIA_DEFAULT_KEY_BLOB,
	)


def _getEffectiveNvidiaApiKey():
	"""Return the cached NVIDIA NIM API key; falls back to lazy init if not yet computed."""
	if _cachedEffectiveNvidiaApiKey is _NOT_COMPUTED:
		_initEffectiveNvidiaApiKey()
	return _cachedEffectiveNvidiaApiKey


def _getPollinationsApiKeyStorePath():
	"""Return the filesystem path for the user-supplied Pollinations.AI API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_POLLINATIONS_KEY_FILENAME)


def getUserPollinationsApiKey():
	"""Return the plain Pollinations.AI API key previously set by the user, or None."""
	path = _getPollinationsApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Pollinations API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserPollinationsApiKey(plain):
	"""Persist a user-supplied Pollinations.AI API key (obfuscated). Empty/None clears it."""
	global _cachedEffectivePollinationsApiKey
	path = _getPollinationsApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectivePollinationsApiKey = _deobfuscateImageApiKey(
				_IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_KEY_BLOB,
			)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectivePollinationsApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Pollinations API key: {e}", exc_info=True)
		return False


def _initEffectivePollinationsApiKey():
	"""Decrypt and cache the effective Pollinations.AI API key."""
	global _cachedEffectivePollinationsApiKey
	userKey = getUserPollinationsApiKey()
	_cachedEffectivePollinationsApiKey = userKey or _deobfuscateImageApiKey(
		_IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_KEY_BLOB,
	)


def _getEffectivePollinationsApiKey():
	"""Return the cached Pollinations.AI API key; falls back to lazy init if not yet computed."""
	if _cachedEffectivePollinationsApiKey is _NOT_COMPUTED:
		_initEffectivePollinationsApiKey()
	return _cachedEffectivePollinationsApiKey


def _getOpenaiApiKeyStorePath():
	"""Return the filesystem path for the user-supplied OpenAI API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_OPENAI_KEY_FILENAME)


def getUserOpenaiApiKey():
	"""Return the plain OpenAI API key previously set by the user, or None."""
	path = _getOpenaiApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user OpenAI API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserOpenaiApiKey(plain):
	"""Persist a user-supplied OpenAI API key (obfuscated). Empty/None clears it."""
	global _cachedEffectiveOpenaiApiKey
	path = _getOpenaiApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveOpenaiApiKey = _deobfuscateImageApiKey(_IMAGE_DESCRIPTION_OPENAI_DEFAULT_KEY_BLOB)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectiveOpenaiApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user OpenAI API key: {e}", exc_info=True)
		return False


def _initEffectiveOpenaiApiKey():
	"""Decrypt and cache the effective OpenAI API key."""
	global _cachedEffectiveOpenaiApiKey
	userKey = getUserOpenaiApiKey()
	_cachedEffectiveOpenaiApiKey = userKey or _deobfuscateImageApiKey(
		_IMAGE_DESCRIPTION_OPENAI_DEFAULT_KEY_BLOB,
	)


def _getEffectiveOpenaiApiKey():
	"""Return the cached OpenAI API key; falls back to lazy init if not yet computed."""
	if _cachedEffectiveOpenaiApiKey is _NOT_COMPUTED:
		_initEffectiveOpenaiApiKey()
	return _cachedEffectiveOpenaiApiKey


def _getMistralApiKeyStorePath():
	"""Return the filesystem path for the user-supplied Mistral AI API key file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_MISTRAL_KEY_FILENAME)


def getUserMistralApiKey():
	"""Return the plain Mistral AI API key previously set by the user, or None."""
	path = _getMistralApiKeyStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			blob = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Mistral API key: {e}", exc_info=True)
		return None
	return _deobfuscateImageApiKey(blob)


def setUserMistralApiKey(plain):
	"""Persist a user-supplied Mistral AI API key (obfuscated). Empty/None clears it."""
	global _cachedEffectiveMistralApiKey
	path = _getMistralApiKeyStorePath()
	if not path:
		return False
	try:
		if not plain:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveMistralApiKey = _deobfuscateImageApiKey(
				_IMAGE_DESCRIPTION_MISTRAL_DEFAULT_KEY_BLOB,
			)
		else:
			blob = _obfuscateImageApiKey(plain)
			with open(path, "w", encoding="utf-8") as f:
				f.write(blob)
			_cachedEffectiveMistralApiKey = plain
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Mistral API key: {e}", exc_info=True)
		return False


def _initEffectiveMistralApiKey():
	"""Decrypt and cache the effective Mistral AI API key."""
	global _cachedEffectiveMistralApiKey
	userKey = getUserMistralApiKey()
	_cachedEffectiveMistralApiKey = userKey or _deobfuscateImageApiKey(
		_IMAGE_DESCRIPTION_MISTRAL_DEFAULT_KEY_BLOB,
	)


def _getEffectiveMistralApiKey():
	"""Return the cached Mistral AI API key; falls back to lazy init if not yet computed."""
	if _cachedEffectiveMistralApiKey is _NOT_COMPUTED:
		_initEffectiveMistralApiKey()
	return _cachedEffectiveMistralApiKey


def _getImageProviderStorePath():
	"""Return the filesystem path for the user-selected provider preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_PROVIDER_FILENAME)


def getUserImageProvider():
	"""Return the provider ID previously chosen by the user, or None."""
	path = _getImageProviderStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user image provider: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS:
		log.debug(f"LINE: stored image provider {value!r} is not in the allowed list")
		return None
	return value


def setUserImageProvider(provider):
	"""Persist the user-selected provider. Empty/None or default clears the file."""
	global _cachedEffectiveImageProvider
	path = _getImageProviderStorePath()
	if not path:
		return False
	try:
		if not provider or provider == _IMAGE_DESCRIPTION_DEFAULT_PROVIDER:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveImageProvider = _IMAGE_DESCRIPTION_DEFAULT_PROVIDER
			return True
		if provider not in _IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS:
			log.warning(f"LINE: refusing to save unknown image provider {provider!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(provider)
		_cachedEffectiveImageProvider = provider
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user image provider: {e}", exc_info=True)
		return False


def _getEffectiveImageProvider():
	"""Return the cached provider ID; lazily resolved from disk on first call."""
	global _cachedEffectiveImageProvider
	if _cachedEffectiveImageProvider is _NOT_COMPUTED:
		_cachedEffectiveImageProvider = getUserImageProvider() or _IMAGE_DESCRIPTION_DEFAULT_PROVIDER
	return _cachedEffectiveImageProvider


# One-time consent before the image-description feature uploads screenshots to
# a cloud AI provider. Asked on first use of each provider regardless of whose
# API key is used: the privacy question is "which company receives my
# screenshot", which a user-supplied key does not answer. Stored as a loose
# file like the other image-description settings, one provider ID per line.
_AI_CONSENT_FILENAME = "line_desktop_ai_consent.txt"


def _getAiConsentStorePath():
	"""Return the filesystem path for the AI-upload consent marker file."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _AI_CONSENT_FILENAME)


# Providers consented to during this NVDA session. Also serves as the fallback
# when _recordAiConsent can't persist to disk (e.g. a read-only config
# directory), so a granted "Yes" stays honored for the rest of the session
# instead of every subsequent use failing with "not consented" right after the
# user agreed.
_aiConsentGrantedInSession = set()


def _readConsentedAiProviders():
	"""Return the set of provider IDs the user has consented to, from disk."""
	path = _getAiConsentStorePath()
	if not path or not os.path.isfile(path):
		return set()
	try:
		with open(path, "r", encoding="utf-8") as f:
			# Older versions stored a bare "1" with no provider ID; such a
			# marker matches no provider, so consent is asked once more per
			# provider — the safe direction.
			return {line.strip() for line in f if line.strip()}
	except Exception:
		return set()


def _hasAiConsent(provider=None):
	"""True when the user has consented to uploads for ``provider``.

	Consent is per provider: agreeing to send screenshots to one cloud service
	must never imply agreement for a different company the user was never
	shown by name.
	"""
	if provider is None:
		provider = _getEffectiveImageProvider()
	if provider in _aiConsentGrantedInSession:
		return True
	return provider in _readConsentedAiProviders()


def _recordAiConsent(provider=None):
	"""Persist the user's acceptance of AI uploads for ``provider``."""
	if provider is None:
		provider = _getEffectiveImageProvider()
	_aiConsentGrantedInSession.add(provider)
	path = _getAiConsentStorePath()
	if not path:
		return False
	try:
		providers = _readConsentedAiProviders() | {provider}
		with open(path, "w", encoding="utf-8") as f:
			f.write("\n".join(sorted(providers)))
		return True
	except Exception as e:
		log.warning(
			f"LINE: failed to persist AI consent to disk; will re-ask next session: {e}",
			exc_info=True,
		)
		return False


def _ensureImageDescriptionConsent():
	"""Ask once per cloud AI provider before sending it a screenshot.

	Returns True when sending is allowed, i.e. the user already accepted for
	the active provider. Otherwise shows a blocking Yes/No dialog naming that
	provider (main thread only). Asked on first use of each provider whether
	the request uses the add-on's bundled key or the user's own, since the
	disclosure is about which company receives the screenshot. Declining is
	not persisted, so the question is asked again on next use.
	"""
	provider = _getEffectiveImageProvider()
	if _hasAiConsent(provider):
		return True
	providerLabel = _IMAGE_DESCRIPTION_PROVIDER_LABELS.get(provider, provider)
	try:
		import gui
		from gui.message import DefaultButtonSet, DialogType, MessageDialog, ReturnCode

		dlg = MessageDialog(
			gui.mainFrame,
			# Translators: Message of the consent dialog shown before the
			# image-description feature uploads a screenshot to a cloud AI
			# service. {provider} is the name of that service.
			_(
				"圖片描述功能會將目前訊息的螢幕截圖上傳至雲端 AI 服務 {provider} 進行辨識。\n"
				"是否同意？選「是」之後，使用這項服務時不會再詢問。\n"
				"可在「NVDA 功能表 → 偏好設定 → 設定 → LINE Desktop」更換服務或設定自己的 API 金鑰；"
				"更換服務後會再次詢問。",
			).format(provider=providerLabel),
			# Translators: Title of the one-time AI image description consent dialog
			_("LINE 圖片描述 - 隱私確認"),
			DialogType.WARNING,
			buttons=DefaultButtonSet.YES_NO,
		)
		# YES_NO alone has no fallback, so escape would not close the dialog;
		# declining is the safe answer to a privacy question.
		dlg.setFallbackAction(ReturnCode.NO)
		if dlg.ShowModal() == ReturnCode.YES:
			_recordAiConsent(provider)
			return True
		return False
	except Exception:
		log.warning("LINE: AI consent dialog failed", exc_info=True)
		return False


def _getImageModelStorePath():
	"""Return the filesystem path for the user-selected model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_MODEL_FILENAME)


def getUserImageModel():
	"""Return the model ID previously chosen by the user, or None."""
	path = _getImageModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user image model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_AVAILABLE_MODELS:
		log.debug(f"LINE: stored image model {value!r} is not in the allowed list")
		return None
	return value


_cachedEffectiveImageModel = _NOT_COMPUTED


def setUserImageModel(name):
	"""Persist a user-selected model. Empty/None or the default clears the file."""
	global _cachedEffectiveImageModel
	path = _getImageModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveImageModel = _IMAGE_DESCRIPTION_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown image model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectiveImageModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user image model: {e}", exc_info=True)
		return False


def _getEffectiveImageModel():
	"""Return the cached model ID; lazily resolved from disk on first call."""
	global _cachedEffectiveImageModel
	if _cachedEffectiveImageModel is _NOT_COMPUTED:
		_cachedEffectiveImageModel = getUserImageModel() or _IMAGE_DESCRIPTION_DEFAULT_MODEL
	return _cachedEffectiveImageModel


def _getOllamaModelStorePath():
	"""Return the filesystem path for the user-selected Ollama model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_OLLAMA_MODEL_FILENAME)


def getUserOllamaModel():
	"""Return the Ollama model ID previously chosen by the user, or None."""
	path = _getOllamaModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Ollama model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS:
		log.debug(f"LINE: stored Ollama model {value!r} is not in the allowed list")
		return None
	return value


def setUserOllamaModel(name):
	"""Persist a user-selected Ollama model. Empty/None or default clears the file."""
	global _cachedEffectiveOllamaModel
	path = _getOllamaModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveOllamaModel = _IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown Ollama model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectiveOllamaModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Ollama model: {e}", exc_info=True)
		return False


def _getEffectiveOllamaModel():
	"""Return the cached Ollama model ID; lazily resolved from disk on first call."""
	global _cachedEffectiveOllamaModel
	if _cachedEffectiveOllamaModel is _NOT_COMPUTED:
		_cachedEffectiveOllamaModel = getUserOllamaModel() or _IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL
	return _cachedEffectiveOllamaModel


def _getNvidiaModelStorePath():
	"""Return the filesystem path for the user-selected NVIDIA NIM model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_NVIDIA_MODEL_FILENAME)


def getUserNvidiaModel():
	"""Return the NVIDIA NIM model ID previously chosen by the user, or None."""
	path = _getNvidiaModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user NVIDIA model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_NVIDIA_AVAILABLE_MODELS:
		log.debug(f"LINE: stored NVIDIA model {value!r} is not in the allowed list")
		return None
	return value


def setUserNvidiaModel(name):
	"""Persist a user-selected NVIDIA NIM model. Empty/None or default clears the file."""
	global _cachedEffectiveNvidiaModel
	path = _getNvidiaModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_NVIDIA_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveNvidiaModel = _IMAGE_DESCRIPTION_NVIDIA_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_NVIDIA_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown NVIDIA model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectiveNvidiaModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user NVIDIA model: {e}", exc_info=True)
		return False


def _getEffectiveNvidiaModel():
	"""Return the cached NVIDIA model ID; lazily resolved from disk on first call."""
	global _cachedEffectiveNvidiaModel
	if _cachedEffectiveNvidiaModel is _NOT_COMPUTED:
		_cachedEffectiveNvidiaModel = getUserNvidiaModel() or _IMAGE_DESCRIPTION_NVIDIA_DEFAULT_MODEL
	return _cachedEffectiveNvidiaModel


def _getPollinationsModelStorePath():
	"""Return the filesystem path for the user-selected Pollinations.AI model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_POLLINATIONS_MODEL_FILENAME)


def getUserPollinationsModel():
	"""Return the Pollinations.AI model ID previously chosen by the user, or None."""
	path = _getPollinationsModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Pollinations model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_POLLINATIONS_AVAILABLE_MODELS:
		log.debug(f"LINE: stored Pollinations model {value!r} is not in the allowed list")
		return None
	return value


def setUserPollinationsModel(name):
	"""Persist a user-selected Pollinations.AI model. Empty/None or default clears the file."""
	global _cachedEffectivePollinationsModel
	path = _getPollinationsModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectivePollinationsModel = _IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_POLLINATIONS_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown Pollinations model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectivePollinationsModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Pollinations model: {e}", exc_info=True)
		return False


def _getEffectivePollinationsModel():
	"""Return the cached Pollinations.AI model ID; lazily resolved from disk on first call."""
	global _cachedEffectivePollinationsModel
	if _cachedEffectivePollinationsModel is _NOT_COMPUTED:
		_cachedEffectivePollinationsModel = (
			getUserPollinationsModel() or _IMAGE_DESCRIPTION_POLLINATIONS_DEFAULT_MODEL
		)
	return _cachedEffectivePollinationsModel


def _getOpenaiModelStorePath():
	"""Return the filesystem path for the user-selected OpenAI model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_OPENAI_MODEL_FILENAME)


def getUserOpenaiModel():
	"""Return the OpenAI model ID previously chosen by the user, or None."""
	path = _getOpenaiModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user OpenAI model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_OPENAI_AVAILABLE_MODELS:
		log.debug(f"LINE: stored OpenAI model {value!r} is not in the allowed list")
		return None
	return value


def setUserOpenaiModel(name):
	"""Persist a user-selected OpenAI model. Empty/None or default clears the file."""
	global _cachedEffectiveOpenaiModel
	path = _getOpenaiModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_OPENAI_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveOpenaiModel = _IMAGE_DESCRIPTION_OPENAI_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_OPENAI_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown OpenAI model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectiveOpenaiModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user OpenAI model: {e}", exc_info=True)
		return False


def _getEffectiveOpenaiModel():
	"""Return the cached OpenAI model ID; lazily resolved from disk on first call."""
	global _cachedEffectiveOpenaiModel
	if _cachedEffectiveOpenaiModel is _NOT_COMPUTED:
		_cachedEffectiveOpenaiModel = getUserOpenaiModel() or _IMAGE_DESCRIPTION_OPENAI_DEFAULT_MODEL
	return _cachedEffectiveOpenaiModel


def _getMistralModelStorePath():
	"""Return the filesystem path for the user-selected Mistral AI model preference."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_MISTRAL_MODEL_FILENAME)


def getUserMistralModel():
	"""Return the Mistral AI model ID previously chosen by the user, or None."""
	path = _getMistralModelStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user Mistral model: {e}", exc_info=True)
		return None
	if not value:
		return None
	if value not in _IMAGE_DESCRIPTION_MISTRAL_AVAILABLE_MODELS:
		log.debug(f"LINE: stored Mistral model {value!r} is not in the allowed list")
		return None
	return value


def setUserMistralModel(name):
	"""Persist a user-selected Mistral AI model. Empty/None or default clears the file."""
	global _cachedEffectiveMistralModel
	path = _getMistralModelStorePath()
	if not path:
		return False
	try:
		if not name or name == _IMAGE_DESCRIPTION_MISTRAL_DEFAULT_MODEL:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveMistralModel = _IMAGE_DESCRIPTION_MISTRAL_DEFAULT_MODEL
			return True
		if name not in _IMAGE_DESCRIPTION_MISTRAL_AVAILABLE_MODELS:
			log.warning(f"LINE: refusing to save unknown Mistral model {name!r}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(name)
		_cachedEffectiveMistralModel = name
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user Mistral model: {e}", exc_info=True)
		return False


def _getEffectiveMistralModel():
	"""Return the cached Mistral AI model ID; lazily resolved from disk on first call."""
	global _cachedEffectiveMistralModel
	if _cachedEffectiveMistralModel is _NOT_COMPUTED:
		_cachedEffectiveMistralModel = getUserMistralModel() or _IMAGE_DESCRIPTION_MISTRAL_DEFAULT_MODEL
	return _cachedEffectiveMistralModel


def _getImagePromptStorePath():
	"""Return the filesystem path for the user-supplied description prompt."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_PROMPT_FILENAME)


def getUserImagePrompt():
	"""Return the prompt previously set by the user, or None."""
	path = _getImagePromptStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read()
	except Exception as e:
		log.debug(f"LINE: failed to read user image prompt: {e}", exc_info=True)
		return None
	value = value.strip()
	if not value:
		return None
	return value


_cachedEffectiveImagePrompt = _NOT_COMPUTED


def setUserImagePrompt(text):
	"""Persist a user-supplied prompt. Empty/None or the default clears the file."""
	global _cachedEffectiveImagePrompt
	path = _getImagePromptStorePath()
	if not path:
		return False
	try:
		cleaned = (text or "").strip()
		if cleaned:
			cleaned = cleaned[:_IMAGE_DESCRIPTION_PROMPT_MAX_LEN]
		if not cleaned or cleaned == _IMAGE_DESCRIPTION_DEFAULT_PROMPT:
			if os.path.isfile(path):
				os.remove(path)
			_cachedEffectiveImagePrompt = _IMAGE_DESCRIPTION_DEFAULT_PROMPT
			return True
		with open(path, "w", encoding="utf-8") as f:
			f.write(cleaned)
		_cachedEffectiveImagePrompt = cleaned
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user image prompt: {e}", exc_info=True)
		return False


def _getEffectiveImagePrompt():
	"""Return the cached prompt; lazily resolved from disk on first call."""
	global _cachedEffectiveImagePrompt
	if _cachedEffectiveImagePrompt is _NOT_COMPUTED:
		_cachedEffectiveImagePrompt = getUserImagePrompt() or _IMAGE_DESCRIPTION_DEFAULT_PROMPT
	return _cachedEffectiveImagePrompt


def _nvdaLanguageHint():
	"""Map NVDA's UI language to an English language name for the AI prompt.

	NVDA's languageHandler returns codes like ``zh_TW`` or ``ja``; we translate
	those into a phrase the AI service can recognise and use to choose its
	output language. Returns None when the language is English or unknown so
	the caller can skip appending a hint.
	"""
	try:
		import languageHandler

		lang = (languageHandler.getLanguage() or "").strip()
	except Exception:
		return None
	if not lang:
		return None
	normalised = lang.replace("-", "_").split(".", 1)[0].lower()
	# Exact match takes precedence so traditional/simplified Chinese variants
	# do not collapse onto a single label.
	exactMap = {
		"zh_tw": "Chinese Traditional",
		"zh_hk": "Chinese Traditional",
		"zh_mo": "Chinese Traditional",
		"zh_cn": "Chinese Simplified",
		"zh_sg": "Chinese Simplified",
		"zh_hans": "Chinese Simplified",
		"zh_hant": "Chinese Traditional",
		"pt_br": "Brazilian Portuguese",
	}
	if normalised in exactMap:
		return exactMap[normalised]
	base = normalised.split("_", 1)[0]
	baseMap = {
		"en": None,
		"zh": "Chinese Simplified",
		"ja": "Japanese",
		"ko": "Korean",
		"th": "Thai",
		"vi": "Vietnamese",
		"id": "Indonesian",
		"ms": "Malay",
		"es": "Spanish",
		"fr": "French",
		"de": "German",
		"it": "Italian",
		"pt": "Portuguese",
		"ru": "Russian",
		"ar": "Arabic",
		"tr": "Turkish",
		"nl": "Dutch",
		"pl": "Polish",
		"sv": "Swedish",
		"fi": "Finnish",
		"da": "Danish",
		"no": "Norwegian",
		"cs": "Czech",
		"hu": "Hungarian",
		"el": "Greek",
		"he": "Hebrew",
		"hi": "Hindi",
	}
	return baseMap.get(base)


def _getRuntimeImagePrompt():
	"""Return the prompt actually sent to the API, with a language hint when applicable."""
	base = _getEffectiveImagePrompt()
	# Language hint is appended only for the default prompt; custom prompts are
	# left untouched so the user can pick their own output language.
	if base != _IMAGE_DESCRIPTION_DEFAULT_PROMPT:
		return base
	hint = _nvdaLanguageHint()
	if not hint:
		return base
	return f"{base} Respond in {hint}."


def _getImageMaxTokensStorePath():
	"""Return the filesystem path for the user-configured max-tokens value."""
	try:
		import globalVars

		configPath = globalVars.appArgs.configPath
	except Exception:
		return None
	if not configPath:
		return None
	return os.path.join(configPath, _IMAGE_DESCRIPTION_USER_MAX_TOKENS_FILENAME)


def getUserImageMaxTokens():
	"""Return the user-configured max output tokens, or None when unset."""
	path = _getImageMaxTokensStorePath()
	if not path or not os.path.isfile(path):
		return None
	try:
		with open(path, "r", encoding="utf-8") as f:
			value = f.read().strip()
	except Exception as e:
		log.debug(f"LINE: failed to read user image max tokens: {e}", exc_info=True)
		return None
	if not value:
		return None
	try:
		n = int(value)
	except ValueError:
		log.debug(f"LINE: stored image max tokens {value!r} is not an integer")
		return None
	if n < _IMAGE_DESCRIPTION_MIN_MAX_TOKENS or n > _IMAGE_DESCRIPTION_MAX_MAX_TOKENS:
		log.debug(f"LINE: stored image max tokens {n} is out of range")
		return None
	return n


def setUserImageMaxTokens(value):
	"""Persist a user-configured max-tokens value. ``None`` clears the file,
	leaving the backend unconstrained (no ``max_tokens`` sent to the API)."""
	path = _getImageMaxTokensStorePath()
	if not path:
		return False
	try:
		if value is None:
			if os.path.isfile(path):
				os.remove(path)
			return True
		try:
			n = int(value)
		except (TypeError, ValueError):
			log.warning(f"LINE: refusing to save non-integer max tokens {value!r}")
			return False
		if n < _IMAGE_DESCRIPTION_MIN_MAX_TOKENS or n > _IMAGE_DESCRIPTION_MAX_MAX_TOKENS:
			log.warning(f"LINE: refusing to save out-of-range max tokens {n}")
			return False
		with open(path, "w", encoding="utf-8") as f:
			f.write(str(n))
		return True
	except Exception as e:
		log.warning(f"LINE: failed to save user image max tokens: {e}", exc_info=True)
		return False


def _captureRegionAsPng(left, top, width, height):
	"""Capture the given screen rect and return PNG bytes, or None on failure.

	Uses NVDA's ``screenBitmap`` to grab pixels and encodes them to PNG
	with only the standard library (``struct`` + ``zlib``).  Keeping the
	encoder self-contained avoids taking a hard dependency on Pillow,
	which is not bundled with NVDA.
	"""
	if width <= 0 or height <= 0:
		return None

	try:
		import ctypes as _ctypes
		import screenBitmap

		sb = screenBitmap.ScreenBitmap(width, height)
		pixels = sb.captureImage(left, top, width, height)
		# pixels is a ctypes array of RGBQUAD (BGRA, 4 bytes per pixel).
		bgra = bytearray(_ctypes.string_at(pixels, width * height * 4))
	except Exception as e:
		log.debug(
			f"LINE: _captureRegionAsPng capture failed: {e}",
			exc_info=True,
		)
		return None

	try:
		import struct
		import zlib

		# Swap B <-> R channels in place to turn BGRA into RGBA.
		# Tuple assignment evaluates the RHS fully before assigning.
		bgra[0::4], bgra[2::4] = bgra[2::4], bgra[0::4]
		rgba = bgra

		stride = width * 4
		raw = bytearray(height * (stride + 1))
		# Prepend a 0x00 filter byte to every scanline (filter type: None).
		for y in range(height):
			srcStart = y * stride
			dstStart = y * (stride + 1)
			raw[dstStart] = 0
			raw[dstStart + 1 : dstStart + 1 + stride] = rgba[srcStart : srcStart + stride]

		def _chunk(tag, data):
			return (
				struct.pack(">I", len(data))
				+ tag
				+ data
				+ struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
			)

		signature = b"\x89PNG\r\n\x1a\n"
		ihdr = struct.pack(
			">IIBBBBB",
			width,
			height,
			8,  # bit depth
			6,  # color type: RGBA
			0,
			0,
			0,  # compression, filter, interlace
		)
		idat = zlib.compress(bytes(raw), 6)
		return signature + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b"")
	except Exception as e:
		log.debug(
			f"LINE: _captureRegionAsPng encode failed: {e}",
			exc_info=True,
		)
		return None


def _buildInitialImageContents(pngBytes, prompt=None):
	"""Build the initial multi-turn ``contents`` payload for image description.

	The returned list contains a single user turn bearing the configured
	prompt text and the inline PNG bytes, matching the Gemini API schema.
	"""
	import base64

	return [
		{
			"role": "user",
			"parts": [
				{"text": prompt if prompt is not None else _getRuntimeImagePrompt()},
				{
					"inline_data": {
						"mime_type": "image/png",
						"data": base64.b64encode(pngBytes).decode("ascii"),
					},
				},
			],
		},
	]


def _callImageDescriptionApi(contents, timeout=None):
	"""Dispatch a pre-built ``contents`` list to the active provider.

	Returns (text, None) on success or (None, error_msg) on failure. The
	provider is resolved at call time, so changing the setting mid-conversation
	switches subsequent turns to the new backend. When ``timeout`` is None,
	each backend uses its own appropriate default (30 s for Google, 60 s for
	Ollama Cloud which typically needs more time).
	"""
	provider = _getEffectiveImageProvider()
	if not _hasAiConsent(provider):
		# Defense in depth: script_describeImage asks for consent on the main
		# thread before this runs; never show UI here (worker thread). Checked
		# against the provider resolved above, so switching provider
		# mid-conversation cannot inherit the previous provider's consent.
		# Translators: Error when image description is attempted without AI consent
		return None, _("尚未同意將圖片傳送給第三方 AI 服務")
	if provider == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
		return _callOllamaImageDescriptionApi(
			contents,
			timeout=timeout if timeout is not None else 60.0,
		)
	if provider == _IMAGE_DESCRIPTION_PROVIDER_NVIDIA:
		return _callNvidiaImageDescriptionApi(
			contents,
			timeout=timeout if timeout is not None else 60.0,
		)
	if provider == _IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS:
		return _callPollinationsImageDescriptionApi(
			contents,
			timeout=timeout if timeout is not None else 60.0,
		)
	if provider == _IMAGE_DESCRIPTION_PROVIDER_OPENAI:
		return _callOpenaiImageDescriptionApi(
			contents,
			timeout=timeout if timeout is not None else 60.0,
		)
	if provider == _IMAGE_DESCRIPTION_PROVIDER_MISTRAL:
		return _callMistralImageDescriptionApi(
			contents,
			timeout=timeout if timeout is not None else 60.0,
		)
	return _callGoogleImageDescriptionApi(
		contents,
		timeout=timeout if timeout is not None else 30.0,
	)


def _callGoogleImageDescriptionApi(contents, timeout=30.0):
	"""Send a pre-built ``contents`` list to Google AI and return (text, error_msg).

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectiveImageApiKey()
		if not apiKey:
			log.warning("LINE: no image description API key available")
			return None, _("未設定 API Key")
		url = _IMAGE_DESCRIPTION_ENDPOINT.format(
			model=_getEffectiveImageModel(),
			key=apiKey,
		)
		body = {"contents": contents}
		userMaxTokens = getUserImageMaxTokens()
		if userMaxTokens is not None:
			body["generationConfig"] = {"maxOutputTokens": userMaxTokens}
		req = urllib.request.Request(
			url,
			data=json.dumps(body).encode("utf-8"),
			headers={"Content-Type": "application/json"},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(f"LINE: image description network error: {e}", exc_info=True)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		candidates = data.get("candidates") or []
		if not candidates:
			log.info(
				f"LINE: image description returned no candidates: {data!r}",
			)
			return None, _("圖片描述失敗 (無回應)")
		parts = candidates[0].get("content", {}).get("parts", []) if isinstance(candidates[0], dict) else []
		for part in parts:
			if isinstance(part, dict):
				text = part.get("text")
				if text:
					return text.strip(), None
	except Exception as e:
		log.warning(
			f"LINE: image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗")


def _geminiContentsToOllamaMessages(contents):
	"""Convert the canonical Gemini-shaped contents into Ollama chat messages.

	The dialog stores history as Gemini turns ({"role": "user"|"model",
	"parts": [{"text": ...}, {"inline_data": {...}}]}). Ollama's /api/chat
	expects {"role": "user"|"assistant", "content": str, "images": [b64]}.
	"""
	messages = []
	for turn in contents or []:
		role = turn.get("role") if isinstance(turn, dict) else None
		ollamaRole = "assistant" if role == "model" else "user"
		texts = []
		images = []
		for part in (turn.get("parts") or []) if isinstance(turn, dict) else []:
			if not isinstance(part, dict):
				continue
			if "text" in part and part["text"]:
				texts.append(part["text"])
			elif "inline_data" in part:
				inline = part.get("inline_data") or {}
				data = inline.get("data")
				if data:
					images.append(data)
		message = {"role": ollamaRole, "content": "\n".join(texts)}
		if images:
			message["images"] = images
		messages.append(message)
	return messages


def _callOllamaImageDescriptionApi(contents, timeout=60.0):
	"""Send the canonical contents to Ollama Cloud's /api/chat endpoint.

	Returns (text, None) on success or (None, error_msg) on failure. Uses a
	longer default timeout because cloud-hosted vision models can take noticeably
	longer than Google's flash variants to produce a first response.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectiveOllamaApiKey()
		if not apiKey:
			log.warning("LINE: no Ollama Cloud API key available")
			return None, _("未設定 API Key")
		messages = _geminiContentsToOllamaMessages(contents)
		body = {
			"model": _getEffectiveOllamaModel(),
			"messages": messages,
			"stream": False,
		}
		req = urllib.request.Request(
			_IMAGE_DESCRIPTION_OLLAMA_ENDPOINT,
			data=json.dumps(body).encode("utf-8"),
			headers={
				"Content-Type": "application/json",
				"Authorization": f"Bearer {apiKey}",
				"Accept": "application/json",
			},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: Ollama image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(f"LINE: Ollama image description network error: {e}", exc_info=True)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: Ollama image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		message = data.get("message") if isinstance(data, dict) else None
		if isinstance(message, dict):
			text = message.get("content")
			if text:
				return text.strip(), None
		log.info(f"LINE: Ollama image description returned no content: {data!r}")
	except Exception as e:
		log.warning(
			f"LINE: Ollama image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗 (無回應)")


def _geminiContentsToNvidiaMessages(contents):
	"""Convert the canonical Gemini-shaped contents into NVIDIA NIM chat messages.

	NVIDIA NIM exposes an OpenAI-compatible chat completions API. Each message has
	{"role": "user"|"assistant", "content": [...]}, where content parts are either
	{"type": "text", "text": ...} or {"type": "image_url",
	"image_url": {"url": "data:image/png;base64,..."}}.
	"""
	messages = []
	for turn in contents or []:
		role = turn.get("role") if isinstance(turn, dict) else None
		nvRole = "assistant" if role == "model" else "user"
		parts = []
		for part in (turn.get("parts") or []) if isinstance(turn, dict) else []:
			if not isinstance(part, dict):
				continue
			if "text" in part and part["text"]:
				parts.append({"type": "text", "text": part["text"]})
			elif "inline_data" in part:
				inline = part.get("inline_data") or {}
				data = inline.get("data")
				mime = inline.get("mime_type") or "image/png"
				if data:
					parts.append(
						{
							"type": "image_url",
							"image_url": {"url": f"data:{mime};base64,{data}"},
						},
					)
		if not parts:
			log.debug(f"LINE: skipping empty NVIDIA turn with role {nvRole!r}")
			continue
		messages.append({"role": nvRole, "content": parts})
	return messages


def _callNvidiaImageDescriptionApi(contents, timeout=60.0):
	"""Send the canonical contents to NVIDIA NIM's OpenAI-compatible chat endpoint.

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectiveNvidiaApiKey()
		if not apiKey:
			log.warning("LINE: no NVIDIA NIM API key available")
			return None, _("未設定 API Key")
		messages = _geminiContentsToNvidiaMessages(contents)
		body = {
			"model": _getEffectiveNvidiaModel(),
			"messages": messages,
			"stream": False,
		}
		userMaxTokens = getUserImageMaxTokens()
		if userMaxTokens is not None:
			body["max_tokens"] = userMaxTokens
		req = urllib.request.Request(
			_IMAGE_DESCRIPTION_NVIDIA_ENDPOINT,
			data=json.dumps(body).encode("utf-8"),
			headers={
				"Content-Type": "application/json",
				"Authorization": f"Bearer {apiKey}",
				"Accept": "application/json",
			},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: NVIDIA image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(f"LINE: NVIDIA image description network error: {e}", exc_info=True)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: NVIDIA image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		choices = data.get("choices") if isinstance(data, dict) else None
		if isinstance(choices, list) and choices:
			message = choices[0].get("message") if isinstance(choices[0], dict) else None
			if isinstance(message, dict):
				text = message.get("content")
				# Some NIM responses return content as a list of parts.
				if isinstance(text, list):
					collected = []
					for part in text:
						if isinstance(part, dict):
							inner = part.get("text")
							if inner:
								collected.append(inner)
					text = "".join(collected)
				if text:
					return text.strip(), None
		log.info(f"LINE: NVIDIA image description returned no content: {data!r}")
	except Exception as e:
		log.warning(
			f"LINE: NVIDIA image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗 (無回應)")


def _geminiContentsToPollinationsMessages(contents):
	"""Convert canonical Gemini-shaped contents into OpenAI-compatible chat messages.

	Pollinations.AI's text endpoint mirrors OpenAI's /chat/completions schema, so
	the conversion is identical to the NVIDIA NIM helper above. Kept as a
	separate function to leave the existing NVIDIA helper untouched and to make
	provider-specific tweaks (e.g. content/parts shape) easy in the future.
	"""
	messages = []
	for turn in contents or []:
		role = turn.get("role") if isinstance(turn, dict) else None
		oaiRole = "assistant" if role == "model" else "user"
		parts = []
		for part in (turn.get("parts") or []) if isinstance(turn, dict) else []:
			if not isinstance(part, dict):
				continue
			if "text" in part and part["text"]:
				parts.append({"type": "text", "text": part["text"]})
			elif "inline_data" in part:
				inline = part.get("inline_data") or {}
				data = inline.get("data")
				mime = inline.get("mime_type") or "image/png"
				if data:
					parts.append(
						{
							"type": "image_url",
							"image_url": {"url": f"data:{mime};base64,{data}"},
						},
					)
		if not parts:
			log.debug(f"LINE: skipping empty Pollinations turn with role {oaiRole!r}")
			continue
		messages.append({"role": oaiRole, "content": parts})
	return messages


def _callPollinationsImageDescriptionApi(contents, timeout=60.0):
	"""Send the canonical contents to Pollinations.AI's OpenAI-compatible chat endpoint.

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectivePollinationsApiKey()
		messages = _geminiContentsToPollinationsMessages(contents)
		body = {
			"model": _getEffectivePollinationsModel(),
			"messages": messages,
			"stream": False,
		}
		userMaxTokens = getUserImageMaxTokens()
		if userMaxTokens is not None:
			body["max_tokens"] = userMaxTokens
		headers = {
			"Content-Type": "application/json",
			"Accept": "application/json",
			# Pollinations.AI's edge layer rejects requests carrying urllib's
			# default ``Python-urllib/x.y`` User-Agent with HTTP 403; matching
			# the AI Content Describer add-on's workaround clears the block.
			"User-Agent": "curl/8.4.0",
		}
		# Pollinations.AI accepts unauthenticated traffic with rate limits, but
		# supplying the bundled token unlocks the higher tier.
		if apiKey:
			headers["Authorization"] = f"Bearer {apiKey}"
		req = urllib.request.Request(
			_IMAGE_DESCRIPTION_POLLINATIONS_ENDPOINT,
			data=json.dumps(body).encode("utf-8"),
			headers=headers,
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: Pollinations image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(
				f"LINE: Pollinations image description network error: {e}",
				exc_info=True,
			)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: Pollinations image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		choices = data.get("choices") if isinstance(data, dict) else None
		if isinstance(choices, list) and choices:
			message = choices[0].get("message") if isinstance(choices[0], dict) else None
			if isinstance(message, dict):
				text = message.get("content")
				if isinstance(text, list):
					collected = []
					for part in text:
						if isinstance(part, dict):
							inner = part.get("text")
							if inner:
								collected.append(inner)
					text = "".join(collected)
				if text:
					return text.strip(), None
		log.info(f"LINE: Pollinations image description returned no content: {data!r}")
	except Exception as e:
		log.warning(
			f"LINE: Pollinations image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗 (無回應)")


def _geminiContentsToOpenaiMessages(contents):
	"""Convert canonical Gemini-shaped contents into OpenAI chat messages.

	OpenAI's /v1/chat/completions schema is the same OpenAI-compatible shape used
	by the NVIDIA NIM and Pollinations.AI helpers above. Kept as a separate
	function so provider-specific tweaks remain easy without touching the other
	backends.
	"""
	messages = []
	for turn in contents or []:
		role = turn.get("role") if isinstance(turn, dict) else None
		oaiRole = "assistant" if role == "model" else "user"
		parts = []
		for part in (turn.get("parts") or []) if isinstance(turn, dict) else []:
			if not isinstance(part, dict):
				continue
			if "text" in part and part["text"]:
				parts.append({"type": "text", "text": part["text"]})
			elif "inline_data" in part:
				inline = part.get("inline_data") or {}
				data = inline.get("data")
				mime = inline.get("mime_type") or "image/png"
				if data:
					parts.append(
						{
							"type": "image_url",
							"image_url": {"url": f"data:{mime};base64,{data}"},
						},
					)
		if not parts:
			log.debug(f"LINE: skipping empty OpenAI turn with role {oaiRole!r}")
			continue
		messages.append({"role": oaiRole, "content": parts})
	return messages


def _callOpenaiImageDescriptionApi(contents, timeout=60.0):
	"""Send the canonical contents to OpenAI's chat completions endpoint.

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectiveOpenaiApiKey()
		if not apiKey:
			log.warning("LINE: no OpenAI API key available")
			return None, _("未設定 API Key")
		messages = _geminiContentsToOpenaiMessages(contents)
		body = {
			"model": _getEffectiveOpenaiModel(),
			"messages": messages,
			"stream": False,
		}
		userMaxTokens = getUserImageMaxTokens()
		if userMaxTokens is not None:
			body["max_tokens"] = userMaxTokens
		req = urllib.request.Request(
			_IMAGE_DESCRIPTION_OPENAI_ENDPOINT,
			data=json.dumps(body).encode("utf-8"),
			headers={
				"Content-Type": "application/json",
				"Authorization": f"Bearer {apiKey}",
				"Accept": "application/json",
			},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: OpenAI image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(f"LINE: OpenAI image description network error: {e}", exc_info=True)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: OpenAI image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		choices = data.get("choices") if isinstance(data, dict) else None
		if isinstance(choices, list) and choices:
			message = choices[0].get("message") if isinstance(choices[0], dict) else None
			if isinstance(message, dict):
				text = message.get("content")
				if isinstance(text, list):
					collected = []
					for part in text:
						if isinstance(part, dict):
							inner = part.get("text")
							if inner:
								collected.append(inner)
					text = "".join(collected)
				if text:
					return text.strip(), None
		log.info(f"LINE: OpenAI image description returned no content: {data!r}")
	except Exception as e:
		log.warning(
			f"LINE: OpenAI image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗 (無回應)")


def _geminiContentsToMistralMessages(contents):
	"""Convert canonical Gemini-shaped contents into Mistral AI chat messages.

	Mistral's /v1/chat/completions schema is OpenAI-compatible except that the
	"image_url" content part is a plain string (the URL or data URI) rather than
	an {"url": ...} object as used by OpenAI/NVIDIA/Pollinations.
	"""
	messages = []
	for turn in contents or []:
		role = turn.get("role") if isinstance(turn, dict) else None
		mistralRole = "assistant" if role == "model" else "user"
		parts = []
		for part in (turn.get("parts") or []) if isinstance(turn, dict) else []:
			if not isinstance(part, dict):
				continue
			if "text" in part and part["text"]:
				parts.append({"type": "text", "text": part["text"]})
			elif "inline_data" in part:
				inline = part.get("inline_data") or {}
				data = inline.get("data")
				mime = inline.get("mime_type") or "image/png"
				if data:
					parts.append(
						{
							"type": "image_url",
							"image_url": f"data:{mime};base64,{data}",
						},
					)
		if not parts:
			log.debug(f"LINE: skipping empty Mistral turn with role {mistralRole!r}")
			continue
		messages.append({"role": mistralRole, "content": parts})
	return messages


def _callMistralImageDescriptionApi(contents, timeout=60.0):
	"""Send the canonical contents to Mistral AI's chat completions endpoint.

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	try:
		import json
		import urllib.request
		import urllib.error

		apiKey = _getEffectiveMistralApiKey()
		if not apiKey:
			log.warning("LINE: no Mistral API key available")
			return None, _("未設定 API Key")
		messages = _geminiContentsToMistralMessages(contents)
		body = {
			"model": _getEffectiveMistralModel(),
			"messages": messages,
			"stream": False,
		}
		userMaxTokens = getUserImageMaxTokens()
		if userMaxTokens is not None:
			body["max_tokens"] = userMaxTokens
		req = urllib.request.Request(
			_IMAGE_DESCRIPTION_MISTRAL_ENDPOINT,
			data=json.dumps(body).encode("utf-8"),
			headers={
				"Content-Type": "application/json",
				"Authorization": f"Bearer {apiKey}",
				"Accept": "application/json",
			},
			method="POST",
		)
		try:
			with urllib.request.urlopen(req, timeout=timeout) as resp:
				raw = resp.read()
		except urllib.error.HTTPError as e:
			errBody = e.read().decode("utf-8", errors="replace")
			log.warning(
				f"LINE: Mistral image description HTTP {e.code} {e.reason}: {errBody[:500]}",
			)
			return None, _("圖片描述失敗 (HTTP {code})").format(code=e.code)
		except Exception as e:
			log.warning(f"LINE: Mistral image description network error: {e}", exc_info=True)
			return None, _("圖片描述失敗 (網路錯誤)")
		data = json.loads(raw.decode("utf-8", errors="replace"))
	except Exception as e:
		log.warning(
			f"LINE: Mistral image description request failed: {e}",
			exc_info=True,
		)
		return None, _("圖片描述失敗")

	try:
		choices = data.get("choices") if isinstance(data, dict) else None
		if isinstance(choices, list) and choices:
			message = choices[0].get("message") if isinstance(choices[0], dict) else None
			if isinstance(message, dict):
				text = message.get("content")
				if isinstance(text, list):
					collected = []
					for part in text:
						if isinstance(part, dict):
							inner = part.get("text")
							if inner:
								collected.append(inner)
					text = "".join(collected)
				if text:
					return text.strip(), None
		log.info(f"LINE: Mistral image description returned no content: {data!r}")
	except Exception as e:
		log.warning(
			f"LINE: Mistral image description response parse failed: {e}",
			exc_info=True,
		)
	return None, _("圖片描述失敗 (無回應)")


def _getDpiScale(hwnd=None):
	"""Get DPI scale factor for the given window (or foreground window).

	Uses GetDpiForWindow (Win10 1607+), falls back to GetDpiForSystem.
	Returns float: 1.0 = 100%, 1.25 = 125%, 1.5 = 150%, 2.0 = 200%, etc.
	"""
	import ctypes

	if hwnd is None:
		hwnd = ctypes.windll.user32.GetForegroundWindow()
	dpi = 96
	try:
		# GetDpiForWindow is available on Windows 10 1607+
		dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
	except Exception:
		try:
			dpi = ctypes.windll.user32.GetDpiForSystem()
		except Exception:
			dpi = 96
	if dpi <= 0:
		dpi = 96
	scale = dpi / 96.0
	log.debug(f"LINE: DPI={dpi}, scale={scale:.2f}")
	return scale


_CHAT_HEADER_ICON_Y_BASE = 55
_CHAT_HEADER_ICON_SPACING_BASE = 27
_CHAT_HEADER_FIRST_ICON_OFFSET_BASE = 15


def _scaleLineUiPixels(value, scale):
	try:
		scale = float(scale)
	except Exception:
		scale = 1.0
	if scale <= 0:
		scale = 1.0
	return int(value * scale)


def _rectTupleFromWinRect(rect):
	try:
		left = int(rect.left)
		top = int(rect.top)
		right = int(rect.right)
		bottom = int(rect.bottom)
	except Exception:
		return None
	if right <= left or bottom <= top:
		return None
	return (left, top, right, bottom)


def _getWindowScreenRect(hwnd, ctypesModule=None):
	"""Return a top-level window rectangle in screen coordinates.

	Prefer DWM extended frame bounds when available because they are not affected
	by the invisible resize border around Qt windows.
	"""
	if not hwnd:
		return None
	if ctypesModule is None:
		ctypesModule = ctypes

	try:
		rect = ctypesModule.wintypes.RECT()
		result = ctypesModule.windll.user32.GetWindowRect(hwnd, ctypesModule.byref(rect))
		if result == 0:
			return None
		windowRect = _rectTupleFromWinRect(rect)
	except Exception:
		windowRect = None

	try:
		dwmapi = getattr(ctypesModule.windll, "dwmapi", None)
		if dwmapi:
			DWMWA_EXTENDED_FRAME_BOUNDS = 9
			dwmRect = ctypesModule.wintypes.RECT()
			hr = dwmapi.DwmGetWindowAttribute(
				hwnd,
				DWMWA_EXTENDED_FRAME_BOUNDS,
				ctypesModule.byref(dwmRect),
				ctypesModule.sizeof(dwmRect),
			)
			extendedRect = _rectTupleFromWinRect(dwmRect)
			if hr == 0 and extendedRect:
				return extendedRect
	except Exception:
		pass

	return windowRect


def _getWindowClientScreenRect(hwnd, ctypesModule=None):
	"""Return a window client rectangle converted to screen coordinates."""
	if not hwnd:
		return None
	if ctypesModule is None:
		ctypesModule = ctypes

	try:
		clientRect = ctypesModule.wintypes.RECT()
		result = ctypesModule.windll.user32.GetClientRect(hwnd, ctypesModule.byref(clientRect))
		if result == 0:
			return None
		width = int(clientRect.right - clientRect.left)
		height = int(clientRect.bottom - clientRect.top)
		if width <= 0 or height <= 0:
			return None

		topLeft = ctypesModule.wintypes.POINT()
		topLeft.x = 0
		topLeft.y = 0
		bottomRight = ctypesModule.wintypes.POINT()
		bottomRight.x = width
		bottomRight.y = height
		if ctypesModule.windll.user32.ClientToScreen(hwnd, ctypesModule.byref(topLeft)) == 0:
			return None
		if ctypesModule.windll.user32.ClientToScreen(hwnd, ctypesModule.byref(bottomRight)) == 0:
			return None
		left = int(topLeft.x)
		top = int(topLeft.y)
		right = int(bottomRight.x)
		bottom = int(bottomRight.y)
		if right <= left or bottom <= top:
			return None
		return (left, top, right, bottom)
	except Exception:
		return None


def _getChatHeaderIconPointFromRect(screenRect, scale, iconIndex=0):
	try:
		left, top, right, bottom = [int(value) for value in screenRect]
		iconIndex = max(0, int(iconIndex))
	except Exception:
		return None
	if right <= left or bottom <= top:
		return None

	iconY = top + _scaleLineUiPixels(_CHAT_HEADER_ICON_Y_BASE, scale)
	iconSpacing = _scaleLineUiPixels(_CHAT_HEADER_ICON_SPACING_BASE, scale)
	firstIconOffset = _scaleLineUiPixels(_CHAT_HEADER_FIRST_ICON_OFFSET_BASE, scale)
	iconX = right - firstIconOffset - (iconIndex * iconSpacing)

	if iconX < left or iconX > right or iconY < top or iconY > bottom:
		return None
	return (iconX, iconY)


def _getChatHeaderIconPointInfo(hwnd, iconIndex=0):
	"""Return a LINE chat header icon point and source rect in screen coordinates.

	Use the client area before the outer window frame so hidden resize borders
	and DWM frame differences do not shift header clicks across DPI settings.
	"""
	scale = _getDpiScale(hwnd)
	contentRect = _getWindowClientScreenRect(hwnd)
	rectSource = "client"
	if not contentRect:
		contentRect = _getWindowScreenRect(hwnd)
		rectSource = "window"
	if not contentRect:
		log.debug("LINE: no window rectangle for chat header icon")
		return None

	point = _getChatHeaderIconPointFromRect(contentRect, scale, iconIndex=iconIndex)
	if point:
		log.info(
			f"LINE: chat header icon index={iconIndex} at {point}, "
			f"rect={contentRect}, source={rectSource}, dpiScale={scale:.2f}",
		)
	else:
		log.warning(
			f"LINE: chat header icon index={iconIndex} outside rect={contentRect}, "
			f"source={rectSource}, dpiScale={scale:.2f}",
		)
	if not point:
		return None
	return (point, contentRect)


def _getChatHeaderIconPoint(hwnd, iconIndex=0):
	"""Return a LINE chat header icon point in screen coordinates."""
	iconInfo = _getChatHeaderIconPointInfo(hwnd, iconIndex=iconIndex)
	if not iconInfo:
		return None
	return iconInfo[0]


def _rectsIntersect(rectA, rectB):
	if not rectA or not rectB:
		return False
	return not (rectA[2] <= rectB[0] or rectA[0] >= rectB[2] or rectA[3] <= rectB[1] or rectA[1] >= rectB[3])


def _getEditPlaceholder(element):
	"""Extract placeholder or current text hints from an edit control."""
	placeholder = ""
	try:
		name = element.CurrentName
		if name and name.strip() and name.strip().lower() not in ("line",):
			placeholder = name.strip().lower()
	except Exception:
		pass
	if not placeholder:
		try:
			val = element.GetCurrentPropertyValue(30045)
			if val and isinstance(val, str) and val.strip():
				placeholder = val.strip().lower()
		except Exception:
			pass
	if not placeholder:
		try:
			desc = element.GetCurrentPropertyValue(30159)
			if desc and isinstance(desc, str) and desc.strip():
				placeholder = desc.strip().lower()
		except Exception:
			pass
	return placeholder


def _getTextViaUIAFindAll(obj, maxElements=30):
	"""Use raw UIA FindAll to get text from descendants.

	Qt6 elements often report childCount=0 to NVDA but DO have
	UIA descendants accessible via FindAll. This method bypasses
	NVDA's child enumeration and queries UIA directly.
	"""
	texts = []
	if not hasattr(obj, "UIAElement") or obj.UIAElement is None:
		return texts
	try:
		element = obj.UIAElement
		handler = UIAHandler.handler
		if handler is None:
			return texts
		# Create a condition that matches all elements
		condition = handler.clientObject.CreateTrueCondition()
		# Find all descendants
		elements = element.FindAll(
			UIAHandler.TreeScope_Descendants,
			condition,
		)
		if elements:
			count = min(elements.Length, maxElements)
			for i in range(count):
				try:
					child = elements.GetElement(i)
					name = child.CurrentName
					if name and name.strip():
						text = name.strip()
						if text not in texts:
							texts.append(text)
				except Exception:
					continue
	except Exception:
		log.debug("_getTextViaUIAFindAll failed", exc_info=True)
	return texts


def _getTextFromDisplay(obj):
	"""Read text from the screen area of the object using display model.

	This is the ultimate fallback when UIA provides no text at all.
	Works because Qt6 renders text visually even if UIA tree is empty.
	"""
	try:
		# Guard: check that the appModule has a valid binding handle
		# DisplayModel requires this, and it's not available during NVDA startup
		appMod = obj.appModule
		if appMod is None:
			return ""
		try:
			# _getBindingHandle will raise if not available
			if not hasattr(appMod, "_getBindingHandle"):
				return ""
			appMod._getBindingHandle()
		except Exception:
			return ""

		if not obj.location:
			return ""
		left, top, width, height = obj.location
		if width <= 0 or height <= 0:
			return ""

		import displayModel

		info = displayModel.DisplayModelTextInfo(obj, textInfos.POSITION_ALL)
		text = info.text
		if text and text.strip():
			return text.strip()
	except Exception:
		pass
	return ""


def _extractTextFromUIAElement(element):
	"""Extract text content from a raw UIA COM element using safe property queries.

	Returns a list of text strings found, or empty list.
	Qt6 elements in LINE typically have empty Name, so we try multiple
	UIA properties via GetCurrentPropertyValue (safe, no comtypes casts).

	UIA Property IDs used:
	  30005 = NameProperty
	  30045 = ValueValue (from ValuePattern)
	  30092 = LegacyIAccessible.Name
	  30093 = LegacyIAccessible.Value
	  30094 = LegacyIAccessible.Description
	  30159 = FullDescription
	"""
	texts = []

	# Strategy 1: Element Name
	try:
		name = element.CurrentName
		if name and name.strip():
			texts.append(name.strip())
			return texts
	except Exception:
		pass

	# Strategy 2: UIA property values via GetCurrentPropertyValue (SAFE)
	propertyIds = [
		(30045, "ValueValue"),
		(30092, "LegacyName"),
		(30093, "LegacyValue"),
		(30094, "LegacyDescription"),
		(30159, "FullDescription"),
	]
	for propId, propLabel in propertyIds:
		try:
			val = element.GetCurrentPropertyValue(propId)
			if val and isinstance(val, str) and val.strip():
				t = val.strip()
				if t not in texts:
					texts.append(t)
					log.debug(f"LINE UIA property {propLabel}({propId}): '{t}'")
		except Exception:
			pass

	if texts:
		return texts

	# Strategy 3: Raw UIA FindAll on descendants
	try:
		handler = UIAHandler.handler
		if handler:
			condition = handler.clientObject.CreateTrueCondition()
			children = element.FindAll(UIAHandler.TreeScope_Descendants, condition)
			if children:
				count = min(children.Length, 20)
				for i in range(count):
					try:
						child = children.GetElement(i)
						childName = child.CurrentName
						if childName and childName.strip():
							t = childName.strip()
							if t not in texts:
								texts.append(t)
						else:
							# Also try ValueValue on descendants
							try:
								childVal = child.GetCurrentPropertyValue(30045)
								if childVal and isinstance(childVal, str) and childVal.strip():
									t = childVal.strip()
									if t not in texts:
										texts.append(t)
							except Exception:
								pass
					except Exception:
						continue
	except Exception:
		pass

	if texts:
		return texts

	# Strategy 4: Walk UIA tree using TreeWalker for direct children
	try:
		handler = UIAHandler.handler
		if handler:
			walker = handler.clientObject.RawViewWalker
			child = walker.GetFirstChildElement(element)
			childCount = 0
			while child and childCount < 20:
				try:
					childName = child.CurrentName
					if childName and childName.strip():
						t = childName.strip()
						if t not in texts:
							texts.append(t)
				except Exception:
					pass
				try:
					child = walker.GetNextSiblingElement(child)
				except Exception:
					break
				childCount += 1
	except Exception:
		pass

	return texts


def _isMessageInputElement(element, handler):
	"""Return True when a raw UIA edit element is LINE's chat message input.

	The login window has two sibling edit fields (email + password), so those
	never count. LINE 26.6.0.4004+ names its edit fields, so the name decides
	first; a search hint wins over a message hint because the sidebar search
	box can be named like "Search chats and messages". Unnamed fields fall back
	to position: the message input sits right of the sidebar or near the bottom.
	"""
	try:
		walker = handler.clientObject.RawViewWalker
		parentEl = walker.GetParentElement(element)
		if not parentEl:
			return False

		editCondition = handler.clientObject.CreatePropertyCondition(
			30003,
			50004,
		)
		editElements = parentEl.FindAll(
			UIAHandler.TreeScope_Children,
			editCondition,
		)
		if editElements and editElements.Length >= 2:
			return False

		placeholder = _getEditPlaceholder(element)
		searchKeywords = ("搜尋", "search", "検索", "ค้นหา", "찾기")
		messageKeywords = ("輸入訊息", "message", "メッセージ", "ข้อความ", "입력")
		if placeholder and any(kw in placeholder for kw in searchKeywords):
			return False
		if placeholder and any(kw in placeholder for kw in messageKeywords):
			return True

		myRect = element.CurrentBoundingRectangle
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		wndRect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(wndRect))
		wndWidth = wndRect.right - wndRect.left
		wndHeight = wndRect.bottom - wndRect.top
		if wndWidth <= 0 or wndHeight <= 0:
			return False
		relativeX = (myRect.left - wndRect.left) / wndWidth
		relativeBottom = (myRect.bottom - wndRect.top) / wndHeight
		return relativeX >= 0.5 or relativeBottom > 0.7
	except Exception:
		log.debug("_isMessageInputElement failed", exc_info=True)
		return False


class LineToolbarButton(UIA):
	"""Overlay class for toolbar/sidebar buttons that lack labels."""

	def _get_name(self):
		name = super().name
		if name and name.strip():
			return name
		# Try tooltip / help text
		try:
			helpText = self.helpText
			if helpText and helpText.strip():
				return helpText.strip()
		except Exception:
			pass
		# Try UIA FindAll for nested text
		texts = _getTextViaUIAFindAll(self, maxElements=5)
		if texts:
			return " ".join(texts)
		# Try automation ID as fallback label
		try:
			automationId = self.UIAAutomationId
			if automationId:
				return automationId.replace("_", " ").replace("-", " ")
		except Exception:
			pass
		# Try display model
		displayText = _getTextFromDisplay(self)
		if displayText:
			return displayText
		return ""


class AppModule(appModuleHandler.AppModule):
	"""NVDA App Module for LINE Desktop.

	Provides accessibility enhancements for LINE desktop application,
	which uses Qt6 framework with incomplete UIA exposure.
	"""

	disableBrowseModeByDefault: bool = True
	sleepMode = None

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		# Read and cache LINE installation info
		self._lineVersion = _readLineVersion()
		self._lineLanguage = _readLineLanguage()
		# Let OCR follow LINE's UI language (falls back to zh-Hant first).
		setPreferredOcrLanguage(self._lineLanguage)
		self._qtAccessibleSet = _isQtAccessibleSet()
		# Decrypt the image API keys once at startup so PBKDF2 never runs
		# during an actual image-description request.
		_initEffectiveImageApiKey()
		_initEffectiveOllamaApiKey()
		_initEffectiveNvidiaApiKey()
		_initEffectivePollinationsApiKey()
		_initEffectiveOpenaiApiKey()
		_initEffectiveMistralApiKey()
		log.info(
			f"LINE AppModule loaded for process: {self.processID}, "
			f"exe: {self.appName}, "
			f"lineVersion: {self._lineVersion}, "
			f"lineLanguage: {self._lineLanguage}, "
			f"qtAccessible: {self._qtAccessibleSet}",
		)

	def terminate(self):
		"""Drop cached COM references and invalidate in-flight async chains.

		Module-level globals would otherwise keep stale UIA pointers alive
		after LINE exits, and pending callLater chains would fire against a
		dead process. All of these caches are best-effort and rebuild
		themselves, so clearing them is safe even when another LINE process
		(e.g. LineCall) is still running.
		"""
		global _suppressAddon
		try:
			_suppressAddon = False
			# Invalidate any in-flight file-dialog watcher so its worker thread
			# stops at the next poll instead of outliving this AppModule (it
			# polls by process ID, which Windows can recycle for the next LINE
			# session).
			AppModule._fileDialogWaitToken += 1
			self._inCallOcrRecognizer = None
			self._inCallOcrPixels = None
			self._inCallOcrImgInfo = None
		except Exception:
			log.debugWarning("LINE AppModule terminate cleanup failed", exc_info=True)
		super().terminate()

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""Name the few buttons LINE still leaves unlabelled.

		LINE 26.6.0.4004+ names its list items, edit fields (including the
		login, search and message boxes) and nearly every button itself, so
		only nameless buttons get an overlay.
		"""
		if _suppressAddon:
			return
		if not isinstance(obj, UIA):
			return

		try:
			role = obj.role
		except Exception:
			log.debugWarning(
				"Error getting role in chooseNVDAObjectOverlayClasses",
				exc_info=True,
			)
			return

		if role == controlTypes.Role.BUTTON:
			try:
				btnName = obj.name
				if not btnName or not btnName.strip():
					clsList.insert(0, LineToolbarButton)
			except Exception:
				clsList.insert(0, LineToolbarButton)

	@script(
		# Translators: Description of a debug script that logs UIA tree info
		description=_("Debug: log UIA tree info for the focused element"),
		gesture="kb:NVDA+shift+k",
		category=SCRIPT_CATEGORY,
	)
	def script_debugUIATree(self, gesture):
		"""Debug helper: probes focused element properties + display model.

		Uses GetFocusedElement() (safe, same as navigation).
		Also tries NVDAHelper display model to read screen text.
		"""
		info = []

		try:
			handler = UIAHandler.handler
			if not handler:
				ui.message("No UIA handler")
				return

			rawEl = handler.clientObject.GetFocusedElement()
			if not rawEl:
				ui.message("No focused element")
				return

			# Basic Current* properties
			try:
				info.append(f"Name: {rawEl.CurrentName!r}")
			except Exception:
				info.append("Name: <error>")
			try:
				info.append(f"ControlType: {rawEl.CurrentControlType}")
			except Exception:
				info.append("ControlType: <error>")
			try:
				info.append(f"ClassName: {rawEl.CurrentClassName!r}")
			except Exception:
				pass
			try:
				info.append(f"AutomationId: {rawEl.CurrentAutomationId!r}")
			except Exception:
				pass
			try:
				rid = rawEl.GetRuntimeId()
				info.append(f"RuntimeId: {rid}")
			except Exception:
				pass

			# Bounding rectangle
			try:
				rect = rawEl.CurrentBoundingRectangle
				info.append(
					f"BoundingRect: left={rect.left}, top={rect.top}, right={rect.right}, bottom={rect.bottom}",
				)
			except Exception as e:
				info.append(f"BoundingRect: <error: {e}>")

			# Extract text via our safe helper (same one navigation uses)
			try:
				texts = _extractTextFromUIAElement(rawEl)
				info.append(f"ExtractedTexts: {texts}")
			except Exception as e:
				info.append(f"ExtractedTexts: <error: {e}>")

			# Try NVDAHelper display model text extraction
			info.append("--- Display Model ---")
			try:
				import ctypes

				rect = rawEl.CurrentBoundingRectangle
				# Get a window handle for the display model
				windowHandle = None
				try:
					windowHandle = rawEl.CurrentNativeWindowHandle
				except Exception:
					pass
				if not windowHandle:
					try:
						windowHandle = ctypes.windll.user32.GetForegroundWindow()
					except Exception:
						pass
				if windowHandle:
					try:
						import displayModel

						class _MinimalObj:
							def __init__(self, hwnd, location, appMod):
								self.windowHandle = hwnd
								self.location = location
								self.appModule = appMod
								self.windowClassName = "Qt663QWindowIcon"

						left = int(rect.left)
						top = int(rect.top)
						width = int(rect.right - rect.left)
						height = int(rect.bottom - rect.top)

						if width > 0 and height > 0:
							from locationHelper import RectLTWH

							location = RectLTWH(left, top, width, height)
							minObj = _MinimalObj(windowHandle, location, self)

							try:
								dmInfo = displayModel.DisplayModelTextInfo(minObj, textInfos.POSITION_ALL)
								dmText = dmInfo.text
								if dmText and dmText.strip():
									info.append(f"  DisplayModel text: {dmText.strip()!r}")
								else:
									info.append("  DisplayModel text: (empty)")
							except Exception as e:
								info.append(f"  DisplayModel error: {e}")
						else:
							info.append("  DisplayModel: invalid rect")
					except Exception as e:
						info.append(f"  DisplayModel import error: {e}")
				else:
					info.append("  No LINE window handle")
			except Exception as e:
				info.append(f"  Display model error: {e}")

			# Try OCR on the bounding rectangle (works for GPU-rendered content)
			info.append("--- OCR ---")
			ocrStarted = False
			try:
				rect = rawEl.CurrentBoundingRectangle
				left = int(rect.left)
				top = int(rect.top)
				width = int(rect.right - rect.left)
				height = int(rect.bottom - rect.top)

				if width > 0 and height > 0:
					import screenBitmap

					sb = screenBitmap.ScreenBitmap(width, height)
					pixels = sb.captureImage(left, top, width, height)
					info.append(f"  ScreenBitmap captured: {width}x{height}")

					try:
						from contentRecog import uwpOcr

						langs = uwpOcr.getLanguages()
						info.append(f"  OCR languages: {langs}")

						# Pick language: prefer Traditional Chinese
						ocrLang = None
						for candidate in ["zh-Hant-TW", "zh-TW", "zh-Hant"]:
							if candidate in langs:
								ocrLang = candidate
								break
						if not ocrLang:
							for lang in langs:
								if lang.startswith("zh"):
									ocrLang = lang
									break
						if not ocrLang and langs:
							ocrLang = langs[0]

						if ocrLang:
							recognizer = uwpOcr.UwpOcr(language=ocrLang)
							info.append(f"  OCR recognizer: {ocrLang}")

							resizeFactor = recognizer.getResizeFactor(width, height)

							class _ImgInfo:
								def __init__(self, w, h, factor, sLeft, sTop):
									self.recogWidth = w * factor
									self.recogHeight = h * factor
									self.resizeFactor = factor
									self._screenLeft = sLeft
									self._screenTop = sTop

								def convertXToScreen(self, x):
									return self._screenLeft + int(x / self.resizeFactor)

								def convertYToScreen(self, y):
									return self._screenTop + int(y / self.resizeFactor)

								def convertWidthToScreen(self, width):
									return int(width / self.resizeFactor)

								def convertHeightToScreen(self, height):
									return int(height / self.resizeFactor)

							imgInfo = _ImgInfo(width, height, resizeFactor, left, top)

							if resizeFactor > 1:
								sb2 = screenBitmap.ScreenBitmap(
									width * resizeFactor,
									height * resizeFactor,
								)
								ocrPixels = sb2.captureImage(
									left,
									top,
									width,
									height,
								)
							else:
								ocrPixels = pixels

							info.append("  OCR: started (async)...")
							ocrStarted = True

							# CRITICAL: Store recognizer, pixels, imgInfo on self
							# to prevent garbage collection while native OCR runs.
							# If these are collected, the native callback crashes NVDA.
							self._ocrRecognizer = recognizer
							self._ocrPixels = ocrPixels
							self._ocrImgInfo = imgInfo

							# Fully async OCR — callback fires on background thread
							appModRef = self  # prevent 'self' confusion in closure

							def _onOcrResult(result):
								"""Handle OCR result on background thread, dispatch to main."""
								import wx

								def _handleOnMain():
									try:
										if isinstance(result, Exception):
											ocrMsg = f"OCR error: {result}"
										else:
											# LinesWordsResult has .text with the full recognized string
											ocrText = getattr(result, "text", "") or ""
											ocrText = _removeCJKSpaces(ocrText.strip())
											if ocrText:
												ocrMsg = f"OCR: {ocrText}"
											else:
												ocrMsg = "OCR: (no text found)"

										log.info(f"LINE Debug OCR result: {ocrMsg}")
										ui.message(ocrMsg)
									except Exception as e:
										log.warning(f"OCR result handler error: {e}", exc_info=True)
									finally:
										# Clean up references now that OCR is done
										appModRef._ocrRecognizer = None
										appModRef._ocrPixels = None
										appModRef._ocrImgInfo = None

								wx.CallAfter(_handleOnMain)

							try:
								recognizer.recognize(ocrPixels, imgInfo, _onOcrResult)
							except Exception as e:
								info.append(f"  OCR recognize error: {e}")
								ocrStarted = False
								self._ocrRecognizer = None
								self._ocrPixels = None
								self._ocrImgInfo = None
						else:
							info.append("  OCR: no language available")
					except Exception as e:
						info.append(f"  OCR setup error: {e}")
				else:
					info.append("  OCR: invalid rect")
			except Exception as e:
				info.append(f"  OCR error: {e}")

		except Exception as e:
			info.append(f"Error: {e}")

		debug_output = "\n".join(info)
		log.info(f"LINE Debug (v28):\n{debug_output}")
		if api.copyToClip(debug_output):
			suffix = " (OCR pending...)" if ocrStarted else ""
			ui.message(f"Copied.{suffix} {debug_output}")

	def _collectAllElements(self, rootElement, handler):
		"""Collect all UIA elements from the tree using multiple strategies.

		LINE's Qt6 UIA implementation often doesn't respond to FindAll
		with specific conditions. This method tries several approaches.
		"""
		allElements = []

		# Strategy 1: FindAll with TrueCondition (finds everything)
		try:
			trueCondition = handler.clientObject.CreateTrueCondition()
			elements = rootElement.FindAll(
				UIAHandler.TreeScope_Descendants,
				trueCondition,
			)
			if elements and elements.Length > 0:
				log.info(f"LINE: FindAll(TrueCondition) found {elements.Length} elements")
				for i in range(elements.Length):
					try:
						allElements.append(elements.GetElement(i))
					except Exception:
						pass
				return allElements
		except Exception as e:
			log.debug(f"LINE: FindAll(TrueCondition) failed: {e}")

		# Strategy 2: Use RawViewWalker to traverse the tree
		try:
			walker = handler.clientObject.RawViewWalker
			if walker:
				self._walkTree(walker, rootElement, allElements, maxDepth=10)
				log.info(f"LINE: RawViewWalker found {len(allElements)} elements")
		except Exception as e:
			log.debug(f"LINE: RawViewWalker failed: {e}")

		return allElements

	def _walkTree(self, walker, parent, result, maxDepth=10, currentDepth=0):
		"""Recursively walk the UIA tree using a TreeWalker."""
		if currentDepth >= maxDepth or len(result) > 500:
			return
		try:
			child = walker.GetFirstChildElement(parent)
			while child:
				result.append(child)
				self._walkTree(walker, child, result, maxDepth, currentDepth + 1)
				try:
					child = walker.GetNextSiblingElement(child)
				except Exception:
					break
		except Exception:
			pass

	def _findButtonByKeywords(self, elements, includeKeywords, excludeKeywords=None):
		"""Search a list of UIA elements for an element matching keywords.

		LINE Qt6 does not use standard Button ControlType, so we search
		ALL elements regardless of their type.

		Returns the matching element or None.
		"""
		if excludeKeywords is None:
			excludeKeywords = []

		# First pass: log all elements with non-empty names for diagnostics
		for el in elements:
			try:
				ctType = 0
				autoId = ""
				name = ""
				try:
					ctType = el.CurrentControlType
				except Exception:
					pass
				try:
					autoId = el.CurrentAutomationId or ""
				except Exception:
					pass
				try:
					name = el.CurrentName or ""
				except Exception:
					pass
				if name or autoId:
					log.debug(
						f"LINE elem: ct={ctType}, autoId={autoId!r}, name={name!r}",
					)
			except Exception:
				pass

		# Second pass: search for matching keywords
		for el in elements:
			try:
				# Get properties
				autoId = ""
				try:
					autoId = el.CurrentAutomationId or ""
				except Exception:
					pass

				name = ""
				try:
					name = el.CurrentName or ""
				except Exception:
					pass

				helpText = ""
				try:
					helpText = str(el.GetCurrentPropertyValue(30048) or "")
				except Exception:
					pass

				className = ""
				try:
					className = el.CurrentClassName or ""
				except Exception:
					pass

				combined = f"{autoId} {name} {helpText} {className}".lower()

				# Skip if any exclude keyword matches
				excluded = False
				for exkw in excludeKeywords:
					if exkw.lower() in combined:
						excluded = True
						break
				if excluded:
					continue

				# Check include keywords
				for keyword in includeKeywords:
					if keyword.lower() in combined:
						log.info(
							f"LINE: found matching element: "
							f"ctType={ctType}, autoId={autoId!r}, "
							f"name={name!r}, help={helpText!r}, class={className!r}",
						)
						return el
			except Exception:
				continue

		return None

	def _invokeElement(self, element, actionName, announce=True):
		"""Invoke a UIA element using InvokePattern or mouse click fallback.

		Safe to call from worker threads (announcements are queued).
		"""

		# Try InvokePattern
		try:
			if _tryInvokeUIAElement(element):
				if announce:
					_queueUiMessage(actionName)
				return True
		except Exception as e:
			log.debug(f"LINE: InvokePattern failed: {e}")

		# Fallback: click the button center
		try:
			rect = element.CurrentBoundingRectangle
			cx = int((rect.left + rect.right) / 2)
			cy = int((rect.top + rect.bottom) / 2)

			if cx > 0 and cy > 0:
				ctypes.windll.user32.SetCursorPos(cx, cy)
				ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
				ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
				if announce:
					_queueUiMessage(actionName)
				return True
		except Exception as e:
			log.debug(f"LINE: click fallback failed: {e}")

		return False

	def _clickAtPosition(self, x, y, hwnd=None):
		"""Perform a mouse click at the given screen coordinates.

		Uses SetForegroundWindow to ensure LINE receives the click,
		then SetCursorPos + mouse_event for the physical click.
		If hwnd is provided, that window is brought to foreground first.
		"""
		import ctypes

		if hwnd is None:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
		if hwnd:
			ctypes.windll.user32.SetForegroundWindow(hwnd)

		_sendMouseClick(int(x), int(y))

	def _getHeaderIconPosition(self):
		"""Get the screen position of the phone icon in the LINE chat header.

		LINE's Qt6 UI does not expose header toolbar buttons via UIA.
		We use the window client geometry to calculate where the icons are.
		All pixel offsets are scaled by window DPI so positions adapt to
		different display scaling settings.

		The chat header has icons from right to left:
		  Index 0: More options (⋮ three dots menu)
		  Index 1: Notes/Keep (📝)
		  Index 2: Phone/Voice call (📞)
		  Index 3: Search (🔍)

		Returns:
			(phoneX, phoneY, winRight) tuple, or None if window not found.
		"""
		import ctypes

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			log.debug("LINE: no foreground window for header click")
			return None

		iconInfo = _getChatHeaderIconPointInfo(hwnd, iconIndex=2)
		if not iconInfo:
			return None
		iconPoint, contentRect = iconInfo

		return (iconPoint[0], iconPoint[1], contentRect[2])

	def _makeCallByType(self, callType):
		"""Click phone icon, wait for popup menu, then click voice or video.

		Full flow (3 steps):
		  1. Click phone icon → popup menu appears (wait 500ms)
		  2. Click voice/video menu item
		  3. OCR the confirmation dialog, announce it, auto-click "開始"

		From the screenshot, clicking the phone icon shows a popup menu:
		  - 語音通話 (voice call): 1st item
		  - 視訊通話 (video call): 2nd item

		The confirmation dialog is centered on the window with:
		  - Text: "確定要與XXX進行語音通話？"
		  - "開始" button (green, left) and "取消" button (gray, right)

		Note: This function only works from the friends tab. Call initiation
		from the chat tab is not currently supported.
		"""
		import ctypes
		import ctypes.wintypes

		pos = self._getHeaderIconPosition()
		if not pos:
			return False

		phoneX, phoneY, winRight = pos

		# Get full window rect for dialog position calculation.
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		winRect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(winRect))
		winLeft = winRect.left
		winTop = winRect.top
		winW = winRect.right - winRect.left
		winH = winRect.bottom - winRect.top

		scale = _getDpiScale()

		log.info(
			f"LINE: clicking phone icon at ({phoneX}, {phoneY})",
		)

		appModRef = self

		# Step 1: Click the phone icon
		self._clickAtPosition(phoneX, phoneY, hwnd)

		# Step 2: After delay, click the voice/video menu item
		def _clickMenuItem():
			menuX = phoneX
			if callType == "voice":
				menuY = phoneY + int(23 * scale)
			else:
				menuY = phoneY + int(70 * scale)

			log.info(
				f"LINE: clicking menu item '{callType}' at ({menuX}, {menuY})",
			)
			appModRef._clickAtPosition(menuX, menuY, hwnd)

			# Step 3: After delay, handle confirmation
			# Video calls show a camera preview (longer load time)
			# Voice calls show a simple confirmation dialog
			if callType == "video":
				core.callLater(1500, _clickVideoStart)
			else:
				core.callLater(800, _handleConfirmDialog)

		core.callLater(500, _clickMenuItem)

		def _clickVideoStart():
			"""Click the green phone button on video call camera preview.

			The video call shows a camera preview screen (相機畫面預覽)
			with a green phone icon at the bottom center of the window.
			"""
			try:
				import ctypes
				import ctypes.wintypes

				# Re-read the current foreground window rect,
				# as the camera preview may be a different window.
				curHwnd = ctypes.windll.user32.GetForegroundWindow()
				curRect = ctypes.wintypes.RECT()
				ctypes.windll.user32.GetWindowRect(
					curHwnd,
					ctypes.byref(curRect),
				)
				curLeft = curRect.left
				curTop = curRect.top
				curW = curRect.right - curRect.left
				_curH = curRect.bottom - curRect.top

				vScale = _getDpiScale()

				# The green phone button is at horizontal center,
				# near the bottom of the camera preview window.
				# From the screenshot: roughly center-x, ~50px from bottom.
				btnX = curLeft + curW // 2
				btnY = curRect.bottom - int(45 * vScale)

				log.info(
					f"LINE: clicking video call start button at "
					f"({btnX}, {btnY}), window=({curLeft},{curTop},"
					f"{curRect.right},{curRect.bottom})",
				)

				ui.message(_("視訊通話確認"))
				appModRef._clickAtPosition(btnX, btnY, curHwnd)

				ui.message(_("已開始視訊通話"))
			except Exception as e:
				log.warning(
					f"LINE: click video start failed: {e}",
					exc_info=True,
				)
				ui.message(_("無法點擊開始按鈕"))

		def _handleConfirmDialog():
			"""OCR the confirmation dialog, announce it, and auto-click 開始.

			This is used for voice calls only. Voice calls show a
			confirmation dialog centered on the window. Group calls have
			a taller dialog (with member avatars) than personal calls.
			"""
			try:
				cScale = _getDpiScale()
				dialogW = int(320 * cScale)
				dialogH = int(200 * cScale)
				winCenterX = winLeft + winW // 2
				winCenterY = winTop + winH // 2
				dialogLeft = winCenterX - dialogW // 2
				dialogTop = winCenterY - dialogH // 2

				log.info(
					f"LINE: OCR confirmation dialog area: ({dialogLeft},{dialogTop}) {dialogW}x{dialogH}",
				)

				import screenBitmap

				sb = screenBitmap.ScreenBitmap(dialogW, dialogH)
				_pixels = sb.captureImage(
					dialogLeft,
					dialogTop,
					dialogW,
					dialogH,
				)

				from contentRecog import uwpOcr

				langs = uwpOcr.getLanguages()
				ocrLang = None
				for candidate in ["zh-Hant-TW", "zh-TW", "zh-Hant"]:
					if candidate in langs:
						ocrLang = candidate
						break
				if not ocrLang:
					for lang in langs:
						if lang.startswith("zh"):
							ocrLang = lang
							break
				if not ocrLang and langs:
					ocrLang = langs[0]

				if not ocrLang:
					log.warning("LINE: no OCR language for dialog")
					_clickStart()
					return

				recognizer = uwpOcr.UwpOcr(language=ocrLang)
				resizeFactor = recognizer.getResizeFactor(
					dialogW,
					dialogH,
				)
				if resizeFactor < 2:
					resizeFactor = 2

				sb2 = screenBitmap.ScreenBitmap(
					dialogW * resizeFactor,
					dialogH * resizeFactor,
				)
				ocrPixels = sb2.captureImage(
					dialogLeft,
					dialogTop,
					dialogW,
					dialogH,
				)

				appModRef._callOcrRecognizer = recognizer
				appModRef._callOcrPixels = ocrPixels

				class _ImgInfo:
					def __init__(self, w, h, factor, sLeft, sTop):
						self.recogWidth = w * factor
						self.recogHeight = h * factor
						self.resizeFactor = factor
						self._screenLeft = sLeft
						self._screenTop = sTop

					def convertXToScreen(self, x):
						return self._screenLeft + int(
							x / self.resizeFactor,
						)

					def convertYToScreen(self, y):
						return self._screenTop + int(
							y / self.resizeFactor,
						)

					def convertWidthToScreen(self, width):
						return int(width / self.resizeFactor)

					def convertHeightToScreen(self, height):
						return int(height / self.resizeFactor)

				imgInfo = _ImgInfo(
					dialogW,
					dialogH,
					resizeFactor,
					dialogLeft,
					dialogTop,
				)
				appModRef._callOcrImgInfo = imgInfo

				def _onOcrResult(result):
					import wx

					def _handleOnMain():
						try:
							ocrText = ""
							ocrActionTargets = {}
							if not isinstance(result, Exception):
								ocrText = (
									getattr(
										result,
										"text",
										"",
									)
									or ""
								)
								ocrText = _removeCJKSpaces(
									ocrText.strip(),
								)
								ocrActionTargets = _extractVoiceCallConfirmationActionClickPoints(
									_extractOcrLines(result),
									(
										dialogLeft,
										dialogTop,
										dialogLeft + dialogW,
										dialogTop + dialogH,
									),
								)

							log.info(
								f"LINE: confirmation dialog OCR: {ocrText!r}",
							)

							callState = _getVoiceCallConfirmationState(ocrText)
							if "join" in ocrActionTargets:
								callState["action"] = "join"
								callState["isGroup"] = True
							elif "start" in ocrActionTargets and callState["action"] != "join":
								callState["action"] = "start"
							actionTarget = _selectVoiceCallConfirmationActionTarget(
								ocrActionTargets,
								callState["action"],
							)
							if ocrText:
								ui.message(ocrText)
							else:
								ui.message(_("語音通話確認"))

							core.callLater(
								300,
								_clickStart,
								callState["isGroup"],
								callState["action"],
								actionTarget.get("clickPoint"),
							)
						except Exception as e:
							log.warning(
								f"LINE: dialog OCR handler error: {e}",
								exc_info=True,
							)
							_clickStart()
						finally:
							appModRef._callOcrRecognizer = None
							appModRef._callOcrPixels = None
							appModRef._callOcrImgInfo = None

					wx.CallAfter(_handleOnMain)

				recognizer.recognize(ocrPixels, imgInfo, _onOcrResult)

			except Exception as e:
				log.warning(
					f"LINE: dialog handling error: {e}",
					exc_info=True,
				)
				_clickStart()

		def _clickStart(isGroup=False, action="start", clickPoint=None):
			"""Click the primary action on the voice call confirmation dialog.

			Group call dialogs are taller (member avatars), so the button
			is further below the window centre than for personal calls.
			"""
			try:
				sScale = _getDpiScale()
				winCenterX = winLeft + winW // 2
				winCenterY = winTop + winH // 2
				action = "join" if action == "join" else "start"
				isGroupLayout = bool(isGroup or action == "join")
				if clickPoint and len(clickPoint) >= 2:
					startBtnX = int(clickPoint[0])
					startBtnY = int(clickPoint[1])
				else:
					startBtnX = winCenterX - int(43 * sScale)
					if isGroupLayout:
						startBtnY = winCenterY + int(50 * sScale)
					else:
						startBtnY = winCenterY + int(17 * sScale)

				log.info(
					f"LINE: clicking voice call {action} at "
					f"({startBtnX}, {startBtnY}) group={isGroupLayout} "
					f"ocrPoint={bool(clickPoint)}",
				)
				appModRef._clickAtPosition(startBtnX, startBtnY, hwnd)

				if action == "join":
					ui.message(_("已加入群組語音通話"))
				elif isGroupLayout:
					ui.message(_("已開始群組語音通話"))
				else:
					ui.message(_("已開始語音通話"))
			except Exception as e:
				log.warning(
					f"LINE: click voice call primary action failed: {e}",
					exc_info=True,
				)
				ui.message(_("無法點擊開始按鈕"))

		return True

	# ── Incoming call handling ──────────────────────────────────────────

	def _findIncomingCallWindow(self):
		"""Find LINE's incoming call window by enumerating all top-level windows.

		LINE incoming calls may appear in:
		- The same process as the main LINE window
		- A separate child process (e.g. LineCall)

		We search ALL visible windows for call-related keywords, then
		verify ownership via executable name.

		Returns the HWND of the call window, or None.
		"""
		import ctypes
		import ctypes.wintypes
		import os

		lineProcessId = self.processID
		callHwnd = None

		# Keywords to match in window titles (case-insensitive)
		_CALL_KEYWORDS = [
			"來電",
			"通話",
			"linecall",
			"call",
			"ringing",
			"着信",
			"สาย",
		]

		# Executable names that belong to LINE
		_LINE_EXES = {"line.exe", "line_app.exe", "linecall.exe", "linelauncher.exe"}

		WNDENUMPROC = ctypes.WINFUNCTYPE(
			ctypes.wintypes.BOOL,
			ctypes.wintypes.HWND,
			ctypes.wintypes.LPARAM,
		)

		def _getExeName(pid):
			"""Get the executable name for a process ID."""
			try:
				PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
				hProc = ctypes.windll.kernel32.OpenProcess(
					PROCESS_QUERY_LIMITED_INFORMATION,
					False,
					pid,
				)
				if not hProc:
					return ""
				try:
					buf = ctypes.create_unicode_buffer(260)
					size = ctypes.wintypes.DWORD(260)
					ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(
						hProc,
						0,
						buf,
						ctypes.byref(size),
					)
					if ok:
						return os.path.basename(buf.value).lower()
					return ""
				finally:
					ctypes.windll.kernel32.CloseHandle(hProc)
			except Exception:
				return ""

		# Determine main window HWND to skip
		mainHwnd = None
		try:
			mainHwnd = self.windowHandle
		except Exception:
			pass

		# ── Pass 1: search ALL visible windows by title ──────────────
		allWindows = []

		def _enumAll(hwnd, lParam):
			if not ctypes.windll.user32.IsWindowVisible(hwnd):
				return True
			buf = ctypes.create_unicode_buffer(512)
			ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
			title = buf.value or ""
			pid = ctypes.wintypes.DWORD()
			ctypes.windll.user32.GetWindowThreadProcessId(
				hwnd,
				ctypes.byref(pid),
			)
			allWindows.append((hwnd, title, pid.value))
			return True

		ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enumAll), 0)

		log.debug(
			f"LINE: _findIncomingCallWindow scanning {len(allWindows)} "
			f"visible windows, mainHwnd={mainHwnd}, linePID={lineProcessId}",
		)

		for hwnd, title, pid in allWindows:
			if hwnd == mainHwnd:
				continue
			titleLower = title.lower()
			for kw in _CALL_KEYWORDS:
				if kw.lower() in titleLower:
					# Verify this window belongs to LINE
					if pid == lineProcessId:
						log.info(
							f"LINE: found call window (same process) hwnd={hwnd}, title={title!r}, pid={pid}",
						)
						callHwnd = hwnd
						break
					# Check if it's a LINE child process
					exeName = _getExeName(pid)
					if exeName in _LINE_EXES:
						log.info(
							f"LINE: found call window (child process) "
							f"hwnd={hwnd}, title={title!r}, pid={pid}, "
							f"exe={exeName}",
						)
						callHwnd = hwnd
						break
					else:
						log.debug(
							f"LINE: title matched but exe mismatch: "
							f"hwnd={hwnd}, title={title!r}, exe={exeName}",
						)
			if callHwnd:
				break

		# ── Pass 2: OCR fallback on non-main LINE windows ───────────
		if not callHwnd:
			fgHwnd = ctypes.windll.user32.GetForegroundWindow()
			skipHwnds = set()
			if mainHwnd:
				skipHwnds.add(mainHwnd)
			if fgHwnd:
				fgPid = ctypes.wintypes.DWORD()
				ctypes.windll.user32.GetWindowThreadProcessId(
					fgHwnd,
					ctypes.byref(fgPid),
				)
				if fgPid.value == lineProcessId:
					skipHwnds.add(fgHwnd)

			candidateHwnds = []
			for hwnd, title, pid in allWindows:
				if hwnd in skipHwnds:
					continue
				# Check both same-process and child-process windows
				isLine = pid == lineProcessId
				if not isLine:
					exeName = _getExeName(pid)
					isLine = exeName in _LINE_EXES
				if not isLine:
					continue
				rect = ctypes.wintypes.RECT()
				ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
				w = rect.right - rect.left
				h = rect.bottom - rect.top
				if w > 50 and h > 30:
					candidateHwnds.append((hwnd, rect))

			log.debug(
				f"LINE: OCR fallback has {len(candidateHwnds)} candidates",
			)

			for hwnd, rect in candidateHwnds:
				try:
					ocrText = self._ocrWindowArea(
						hwnd,
						sync=True,
						timeout=2.0,
					)
					if ocrText:
						ocrLower = ocrText.lower()
						checkRegion = ocrLower[:150] if len(ocrLower) > 200 else ocrLower
						for kw in _CALL_KEYWORDS:
							if kw.lower() in checkRegion:
								log.info(
									f"LINE: found call window via OCR hwnd={hwnd}, text={ocrText!r}",
								)
								callHwnd = hwnd
								break
					if callHwnd:
						break
				except Exception as e:
					log.debug(
						f"LINE: OCR probe on hwnd={hwnd} failed: {e}",
					)

		if not callHwnd:
			log.debug("LINE: no incoming call window found")

		return callHwnd

	def _ocrWindowAreaResult(self, hwnd, region=None, sync=False, timeout=3.0):
		"""OCR a window (or part of it) and return the raw OCR result object.

		Args:
			hwnd: Window handle to capture.
			region: Optional (left, top, width, height) tuple in screen
				coordinates.  If None, uses the full window rect.
			sync: If True, block until OCR completes (up to timeout).
			timeout: Max seconds to wait when sync=True.

		Returns:
			The OCR result object, or None on failure.
		"""
		import ctypes
		import ctypes.wintypes
		import threading

		if not region:
			rect = ctypes.wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
			left = rect.left
			top = rect.top
			width = rect.right - rect.left
			height = rect.bottom - rect.top
		else:
			left, top, width, height = region

		if width <= 0 or height <= 0:
			return None

		try:
			import screenBitmap
			from contentRecog import uwpOcr

			sb = screenBitmap.ScreenBitmap(width, height)
			pixels = sb.captureImage(left, top, width, height)

			langs = uwpOcr.getLanguages()
			ocrLang = pickOcrLanguage(langs)
			if not ocrLang:
				log.warning("LINE: no OCR language available")
				return None

			recognizer = uwpOcr.UwpOcr(language=ocrLang)
			resizeFactor = recognizer.getResizeFactor(width, height)
			if resizeFactor < 2:
				resizeFactor = 2

			if resizeFactor > 1:
				sb2 = screenBitmap.ScreenBitmap(
					width * resizeFactor,
					height * resizeFactor,
				)
				ocrPixels = sb2.captureImage(left, top, width, height)
			else:
				ocrPixels = pixels

			class _ImgInfo:
				def __init__(self, w, h, factor, sLeft, sTop):
					self.recogWidth = w * factor
					self.recogHeight = h * factor
					self.resizeFactor = factor
					self._screenLeft = sLeft
					self._screenTop = sTop

				def convertXToScreen(self, x):
					return self._screenLeft + int(x / self.resizeFactor)

				def convertYToScreen(self, y):
					return self._screenTop + int(y / self.resizeFactor)

				def convertWidthToScreen(self, w):
					return int(w / self.resizeFactor)

				def convertHeightToScreen(self, h):
					return int(h / self.resizeFactor)

			imgInfo = _ImgInfo(width, height, resizeFactor, left, top)

			if sync:
				resultHolder = [None]
				event = threading.Event()

				# Keep references alive
				self._inCallOcrRecognizer = recognizer
				self._inCallOcrPixels = ocrPixels
				self._inCallOcrImgInfo = imgInfo

				def _onResult(result):
					resultHolder[0] = result
					event.set()

				recognizer.recognize(ocrPixels, imgInfo, _onResult)
				event.wait(timeout=timeout)

				self._inCallOcrRecognizer = None
				self._inCallOcrPixels = None
				self._inCallOcrImgInfo = None

				result = resultHolder[0]
				if result is None or isinstance(result, Exception):
					return None
				return result
			else:
				# Async — not used for incoming call detection
				return None
		except Exception as e:
			log.debug(f"LINE: _ocrWindowAreaResult failed: {e}", exc_info=True)
			return None

	def _ocrWindowArea(self, hwnd, region=None, sync=False, timeout=3.0):
		"""OCR a window (or part of it) and return the recognized text."""
		result = self._ocrWindowAreaResult(
			hwnd,
			region=region,
			sync=sync,
			timeout=timeout,
		)
		if result is None or isinstance(result, Exception):
			return ""
		text = getattr(result, "text", "") or ""
		return _removeCJKSpaces(text.strip())

	def _getCallButtonElements(self, hwnd):
		"""Collect UIA elements from the call window and log their properties.

		Returns (allElements, handler, rootEl) tuple, or ([], None, None).
		"""
		import ctypes
		import ctypes.wintypes

		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winW = rect.right - rect.left
		winH = rect.bottom - rect.top

		log.info(
			f"LINE: call window rect=({rect.left},{rect.top},{rect.right},{rect.bottom}), size={winW}x{winH}",
		)

		try:
			handler = UIAHandler.handler
			if not handler or not handler.clientObject:
				return ([], None, None)

			rootEl = None
			try:
				rootEl = handler.clientObject.ElementFromHandle(hwnd)
			except Exception:
				pass
			if not rootEl:
				return ([], None, None)

			allElements = self._collectAllElements(rootEl, handler)

			# Log ALL elements with detailed info for debugging
			for i, el in enumerate(allElements):
				try:
					ct = 0
					name = ""
					autoId = ""
					elRectStr = "?"
					try:
						ct = el.CurrentControlType
					except Exception:
						pass
					try:
						name = el.CurrentName or ""
					except Exception:
						pass
					try:
						autoId = el.CurrentAutomationId or ""
					except Exception:
						pass
					try:
						elRect = el.CurrentBoundingRectangle
						elRectStr = f"({elRect.left},{elRect.top},{elRect.right},{elRect.bottom})"
					except Exception:
						pass
					# Check InvokePattern support
					hasInvoke = False
					try:
						pat = el.GetCurrentPattern(10000)
						hasInvoke = pat is not None
					except Exception:
						pass
					log.info(
						f"LINE call elem[{i}]: ct={ct}, "
						f"name={name!r}, autoId={autoId!r}, "
						f"rect={elRectStr}, invoke={hasInvoke}",
					)
				except Exception:
					log.debug(f"LINE call elem[{i}]: error reading")

			return (allElements, handler, rootEl)
		except Exception as e:
			log.debug(f"LINE: call element collection failed: {e}")
			return ([], None, None)

	def _findCallButtonByRect(self, hwnd, allElements, side="right"):
		"""Find a button-like element by its position in the call window.

		LINE's call window has button-like elements with no names.
		We identify them by bounding rectangle position:
		  - 'right' side = answer button (green)
		  - 'left' side = decline button (red)

		IMPORTANT: Only considers elements whose center is INSIDE the
		window rect.  LINE's Qt6 window exposes border/frame elements
		that are OUTSIDE the window bounds and must be filtered out.

		Returns (element, centerX, centerY) or None.
		"""
		import ctypes
		import ctypes.wintypes

		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winW = rect.right - rect.left
		winH = rect.bottom - rect.top
		winCenterX = rect.left + winW // 2

		# Collect elements with valid bounding rects INSIDE the window
		candidates = []
		outsideCount = 0
		for el in allElements:
			try:
				elRect = el.CurrentBoundingRectangle
				elW = elRect.right - elRect.left
				elH = elRect.bottom - elRect.top
				elCX = (elRect.left + elRect.right) // 2
				elCY = (elRect.top + elRect.bottom) // 2

				# CRITICAL: element center must be INSIDE the window
				if elCX < rect.left or elCX > rect.right or elCY < rect.top or elCY > rect.bottom:
					outsideCount += 1
					continue

				# Filter: must be visible, reasonably sized (like a button)
				if elW < 10 or elH < 10:
					continue
				# Skip elements that span the full window width
				if elW > winW * 0.8:
					continue
				# Skip root / container elements
				if elW > winW * 0.6 and elH > winH * 0.6:
					continue

				candidates.append((el, elRect, elCX, elCY, elW, elH))
			except Exception:
				continue

		if outsideCount:
			log.info(
				f"LINE: filtered out {outsideCount} border elements outside window rect",
			)

		if not candidates:
			log.info("LINE: no button candidates INSIDE window")
			return None

		log.info(
			f"LINE: {len(candidates)} button candidates found, looking for '{side}' button",
		)

		if side == "right":
			rightCandidates = [c for c in candidates if c[2] > winCenterX]
			if rightCandidates:
				rightCandidates.sort(key=lambda c: c[2], reverse=True)
				best = rightCandidates[0]
				log.info(
					f"LINE: selected right button at ({best[2]},{best[3]}), size={best[4]}x{best[5]}",
				)
				return (best[0], best[2], best[3])
		else:
			leftCandidates = [c for c in candidates if c[2] < winCenterX]
			if leftCandidates:
				leftCandidates.sort(key=lambda c: c[2])
				best = leftCandidates[0]
				log.info(
					f"LINE: selected left button at ({best[2]},{best[3]}), size={best[4]}x{best[5]}",
				)
				return (best[0], best[2], best[3])

		# Fallback: any candidate sorted by position
		if candidates:
			if side == "right":
				candidates.sort(key=lambda c: c[2], reverse=True)
			else:
				candidates.sort(key=lambda c: c[2])
			best = candidates[0]
			log.info(
				f"LINE: fallback button at ({best[2]},{best[3]}), size={best[4]}x{best[5]}",
			)
			return (best[0], best[2], best[3])

		return None

	def _ocrFindButtonKeyword(self, hwnd, keywords):
		"""Use OCR to check if any keyword appears in the call window.

		Returns (matched: bool, ocrText: str) tuple.
		  matched = True if any keyword is found in the OCR text.
		  ocrText = the raw OCR text (for further analysis by caller).
		Note: NVDA's uwpOcr result only provides flat text,
		not per-word positions, so we just confirm presence.
		"""
		try:
			ocrText = self._ocrWindowArea(hwnd, sync=True, timeout=3.0)
			if not ocrText:
				log.info("LINE: OCR returned no text for call window")
				return (False, "")

			ocrTextLower = ocrText.lower()
			log.info(
				f"LINE: OCR call window text: '{ocrText}'",
			)
			for kw in keywords:
				if kw.lower() in ocrTextLower:
					log.info(f"LINE: OCR found keyword '{kw}'")
					return (True, ocrText)

			log.info("LINE: OCR no keyword match")
			return (False, ocrText)
		except Exception as e:
			log.debug(
				f"LINE: _ocrFindButtonKeyword failed: {e}",
				exc_info=True,
			)
			return (False, "")

	_VIDEO_KEYWORDS = ["視訊", "video", "ビデオ", "วิดีโอ"]

	def _isVideoCallWindow(self, hwnd, ocrText=None):
		"""Check if the call window is a video call (vs voice call).

		First checks window title (fast). If title is generic (e.g.
		"LineCall"), falls back to checking the provided ocrText,
		or performs OCR if no ocrText is given.
		"""
		import ctypes

		buf = ctypes.create_unicode_buffer(512)
		ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
		title = (buf.value or "").lower()
		isVideo = any(kw in title for kw in self._VIDEO_KEYWORDS)
		if isVideo:
			log.debug(f"LINE: _isVideoCallWindow title={title!r} → True")
			return True

		# Title is generic (e.g. "LineCall") — check OCR text
		if ocrText is None:
			# Perform a quick OCR to check
			try:
				ocrText = (
					self._ocrWindowArea(
						hwnd,
						sync=True,
						timeout=2.0,
					)
					or ""
				)
			except Exception:
				ocrText = ""

		if ocrText:
			ocrLower = ocrText.lower()
			isVideo = any(kw in ocrLower for kw in self._VIDEO_KEYWORDS)
			log.debug(
				f"LINE: _isVideoCallWindow title={title!r}, ocrCheck → {isVideo}",
			)
			return isVideo

		log.debug(f"LINE: _isVideoCallWindow title={title!r} → False")
		return False

	def _answerIncomingCall(self, hwnd):
		"""Answer an incoming call by clicking the answer (green) button.

		Multi-strategy approach:
		  1. Bring the call window to the foreground
		  2. Try UIA keyword search for answer button
		  3. Try UIA bounding-rect analysis (buttons inside window)
		  4. OCR: find "接聽" text and click its position
		  5. Fallback: click at proportional position inside window
		"""
		import ctypes
		import ctypes.wintypes

		# Step 0: Bring call window to foreground
		try:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
			_waitForForegroundWindow(hwnd)
		except Exception:
			pass

		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winW = rect.right - rect.left
		winH = rect.bottom - rect.top

		isVideoCall = self._isVideoCallWindow(hwnd)
		allElements, handler, rootEl = self._getCallButtonElements(hwnd)

		# Strategy 1: UIA keyword search
		if allElements:
			answerEl = self._findButtonByKeywords(
				allElements,
				["接聽", "accept", "answer", "応答", "รับสาย", "接受"],
				excludeKeywords=["拒絕", "decline", "reject"],
			)
			if answerEl:
				if self._invokeElement(answerEl, "已接聽"):
					return True

		# Strategy 2: UIA bounding-rect analysis (inside window only)
		# Voice call popup: answer (green) on the LEFT
		# Video call popup: answer (green camera) on the RIGHT
		if allElements:
			btnSide = "right" if isVideoCall else "left"
			result = self._findCallButtonByRect(
				hwnd,
				allElements,
				side=btnSide,
			)
			if result:
				el, cx, cy = result
				invoked = False
				try:
					if _tryInvokeUIAElement(el):
						invoked = True
						log.info("LINE: answered via InvokePattern")
				except Exception as e:
					log.debug(f"LINE: InvokePattern failed: {e}")

				if not invoked:
					log.info(
						f"LINE: clicking answer button at ({cx}, {cy})",
					)
					self._clickAtPosition(cx, cy)

				_queueUiMessage(_("已接聽"))
				return True

		# Strategy 3: OCR confirms call window, then click at position
		log.info("LINE: trying OCR to confirm call window")
		ocrConfirmed, ocrText = self._ocrFindButtonKeyword(
			hwnd,
			[
				"接聽",
				"accept",
				"answer",
				"応答",
				"รับสาย",
				"拒絕",
				"decline",
				"reject",
				"來電",
			],
		)

		# Re-check video call status using OCR text
		# (window title may be generic like "LineCall")
		if not isVideoCall and ocrText:
			isVideoCall = self._isVideoCallWindow(hwnd, ocrText=ocrText)
			if isVideoCall:
				log.info(
					"LINE: OCR confirms this is a VIDEO call, adjusting click position",
				)

		# Strategy 4: Click at position inside the window
		# Screenshot layout: [avatar][caller text][red reject ~80%][green answer ~92%]
		# Answer (green phone) is the RIGHTMOST button
		if isVideoCall:
			# Video call: answer (green camera) at top-right
			answerX = rect.left + int(winW * 0.92)
			answerY = rect.top + int(winH * 0.08)
		elif winH > 200:
			# Full call window — answer button
			answerX = rect.left + int(winW * 0.65)
			answerY = rect.top + int(winH * 0.75)
		else:
			# Small notification popup (e.g. 456x99)
			# Answer button is at far right edge
			answerX = rect.left + int(winW * 0.92)
			answerY = rect.top + int(winH * 0.35)

		log.info(
			f"LINE: clicking answer (fallback) at "
			f"({answerX}, {answerY}), isVideo={isVideoCall}, "
			f"winRect=({rect.left},"
			f"{rect.top},{rect.right},{rect.bottom})",
		)
		self._clickAtPosition(answerX, answerY)
		_queueUiMessage(_("已接聽"))
		return True

	def _rejectIncomingCall(self, hwnd):
		"""Reject an incoming call by clicking the decline (red) button.

		Multi-strategy approach (mirrors _answerIncomingCall):
		  1. Bring the call window to the foreground
		  2. Try UIA keyword search for decline button
		  3. Try UIA bounding-rect analysis (buttons inside window)
		  4. OCR: find "拒絕" text and click its position
		  5. Fallback: click at proportional position inside window
		"""
		import ctypes
		import ctypes.wintypes

		# Step 0: Bring call window to foreground
		try:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
			_waitForForegroundWindow(hwnd)
		except Exception:
			pass

		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winW = rect.right - rect.left
		winH = rect.bottom - rect.top

		isVideoCall = self._isVideoCallWindow(hwnd)
		allElements, handler, rootEl = self._getCallButtonElements(hwnd)

		# Strategy 1: UIA keyword search
		if allElements:
			rejectEl = self._findButtonByKeywords(
				allElements,
				["拒絕", "decline", "reject", "拒否", "ปฏิเสธ"],
				excludeKeywords=["接聽", "accept", "answer"],
			)
			if rejectEl:
				if self._invokeElement(rejectEl, "已拒絕"):
					return True

		# Strategy 2: UIA bounding-rect analysis (inside window only)
		# Voice call popup: reject (red) on the RIGHT
		# Video call popup: reject (red) on the LEFT of the answer button
		if allElements:
			btnSide = "left" if isVideoCall else "right"
			result = self._findCallButtonByRect(
				hwnd,
				allElements,
				side=btnSide,
			)
			if result:
				el, cx, cy = result
				invoked = False
				try:
					if _tryInvokeUIAElement(el):
						invoked = True
						log.info("LINE: rejected via InvokePattern")
				except Exception as e:
					log.debug(f"LINE: InvokePattern failed: {e}")

				if not invoked:
					log.info(
						f"LINE: clicking reject button at ({cx}, {cy})",
					)
					self._clickAtPosition(cx, cy)

				_queueUiMessage(_("已拒絕"))
				return True

		# Strategy 3: OCR confirms call window, then click at position
		log.info("LINE: trying OCR to confirm call window")
		ocrConfirmed, ocrText = self._ocrFindButtonKeyword(
			hwnd,
			[
				"拒絕",
				"decline",
				"reject",
				"拒否",
				"ปฏิเสธ",
				"接聽",
				"accept",
				"answer",
				"來電",
			],
		)

		# Re-check video call status using OCR text
		if not isVideoCall and ocrText:
			isVideoCall = self._isVideoCallWindow(hwnd, ocrText=ocrText)
			if isVideoCall:
				log.info(
					"LINE: OCR confirms this is a VIDEO call, adjusting click position",
				)

		# Strategy 4: Click at position inside the window
		# Screenshot layout: [avatar][caller text][red reject ~80%][green answer ~92%]
		# Reject (red phone) is second from right
		if isVideoCall:
			# Video call: decline (red) at top-right area, left of answer
			rejectX = rect.left + int(winW * 0.80)
			rejectY = rect.top + int(winH * 0.08)
		elif winH > 200:
			# Full call window — decline button
			rejectX = rect.left + int(winW * 0.35)
			rejectY = rect.top + int(winH * 0.75)
		else:
			# Small notification popup
			# Reject button is second from right
			rejectX = rect.left + int(winW * 0.80)
			rejectY = rect.top + int(winH * 0.35)

		log.info(
			f"LINE: clicking reject (fallback) at "
			f"({rejectX}, {rejectY}), isVideo={isVideoCall}, "
			f"winRect=({rect.left},"
			f"{rect.top},{rect.right},{rect.bottom})",
		)
		self._clickAtPosition(rejectX, rejectY)
		_queueUiMessage(_("已拒絕"))
		return True

	def _getCallerInfo(self, hwnd):
		"""OCR the call window to extract and announce the caller's name."""
		import ctypes
		import ctypes.wintypes

		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))

		ocrText = self._ocrWindowArea(hwnd, sync=True, timeout=3.0)
		if ocrText:
			# Clean up: remove "來電" and other system labels to extract
			# just the caller name
			callerName = ocrText
			for removeKw in [
				"來電",
				"着信",
				"ringing",
				"incoming call",
				"สายเรียกเข้า",
			]:
				callerName = callerName.replace(removeKw, "")
			callerName = callerName.strip()
			if callerName:
				_queueUiMessage(_("來電：{callerName}").format(callerName=callerName))
			else:
				_queueUiMessage(_("來電（OCR: {ocrText}）").format(ocrText=ocrText))
			log.info(f"LINE: caller info OCR: {ocrText!r}")
		else:
			_queueUiMessage(_("無法辨識來電者"))
			log.info("LINE: caller info OCR returned empty")

	# ── Incoming call scripts ──────────────────────────────────────────
	# Note: gesture bindings are registered in the GlobalPlugin
	# (lineDesktopHelper.py) so they work even when LINE isn't focused.

	# Serializes the call-window tasks: each one does window enumeration and
	# screen OCR that takes seconds, so they run on a worker thread and must
	# not overlap.
	_callTaskInProgress = False
	# Serializes the in-call mic/camera status checks, which move the mouse to
	# reveal the control bar; overlapping checks would fight over the cursor.
	_callStatusCheckInProgress = False

	def _runCallTask(self, taskName, task, errorMessage):
		"""Run a blocking call-window task off the main thread.

		The task may only announce via _queueUiMessage (it runs on a worker
		thread where ui.message is unsafe).
		"""
		if AppModule._callTaskInProgress:
			log.debug(f"LINE: {taskName} skipped; another call task is running")
			return
		AppModule._callTaskInProgress = True

		def _work():
			try:
				task()
			except Exception as e:
				log.warning(f"LINE {taskName} error: {e}", exc_info=True)
				_queueUiMessage(errorMessage.format(error=e))

		def _done(result):
			AppModule._callTaskInProgress = False

		_runOnWorkerThread(_work, onDone=_done, taskName=taskName)

	def script_answerCall(self, gesture):
		"""Answer an incoming LINE call."""

		def _task():
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._answerIncomingCall(hwnd)
			else:
				_queueUiMessage(_("未偵測到來電"))

		self._runCallTask("answerCall", _task, _("接聽功能錯誤: {error}"))

	def script_rejectCall(self, gesture):
		"""Reject an incoming LINE call."""

		def _task():
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._rejectIncomingCall(hwnd)
			else:
				_queueUiMessage(_("未偵測到來電"))

		self._runCallTask("rejectCall", _task, _("拒絕功能錯誤: {error}"))

	def script_checkCaller(self, gesture):
		"""Announce who is calling."""

		def _task():
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._getCallerInfo(hwnd)
			else:
				_queueUiMessage(_("未偵測到來電"))

		self._runCallTask("checkCaller", _task, _("來電查看功能錯誤: {error}"))

	def script_focusCallWindow(self, gesture):
		"""Find the LineCall window, bring it to foreground, and OCR its content."""
		import ctypes

		def _task():
			hwnd = self._findIncomingCallWindow()
			if not hwnd:
				_queueUiMessage(_("未偵測到通話視窗"))
				return

			try:
				ctypes.windll.user32.SetForegroundWindow(hwnd)
			except Exception:
				pass

			# Wait for the window to actually reach the foreground, then OCR
			# it. Blocking is fine here: this runs on a worker thread.
			_waitForForegroundWindow(hwnd)
			try:
				ocrText = self._ocrWindowArea(hwnd, sync=True, timeout=3.0)
				if ocrText:
					_queueUiMessage(ocrText)
					log.info(f"LINE: call window OCR: {ocrText!r}")
				else:
					_queueUiMessage(_("通話視窗（無法辨識內容）"))
			except Exception as e:
				log.warning(f"LINE: call window OCR error: {e}", exc_info=True)
				_queueUiMessage(_("通話視窗"))

		self._runCallTask("focusCallWindow", _task, _("通話視窗功能錯誤: {error}"))

	# ── Outgoing call scripts ──────────────────────────────────────────

	@script(
		# Translators: Description of a script to make a voice call
		description=_("撥打語音通話"),
		gesture="kb:NVDA+windows+c",
		category=SCRIPT_CATEGORY,
	)
	def script_makeCall(self, gesture):
		"""Click the phone icon, then auto-select voice call from the popup menu."""
		try:
			if not self._makeCallByType("voice"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeCall error: {e}", exc_info=True)
			ui.message(_("通話功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to make a video call
		description=_("撥打視訊通話"),
		gesture="kb:NVDA+windows+v",
		category=SCRIPT_CATEGORY,
	)
	def script_makeVideoCall(self, gesture):
		"""Click the phone icon, then auto-select video call from the popup menu."""
		try:
			if not self._makeCallByType("video"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeVideoCall error: {e}", exc_info=True)
			ui.message(_("視訊通話功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to click the more options button
		description=_("LINE: 點擊更多選項按鈕"),
		gesture="kb:NVDA+windows+o",
		category=SCRIPT_CATEGORY,
	)
	def script_clickMoreOptions(self, gesture):
		"""Open the chat's "More options" menu; LINE's own menu takes the focus."""

		def _afterOpen(result):
			if not result or not result[0]:
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))

		_runOnWorkerThread(self._openChatMoreOptionsMenu, onDone=_afterOpen, taskName="openMoreOptions")

	def _openChatMoreOptionsMenu(self):
		"""Worker thread: click the chat header's "More options" button.

		Clicks the button LINE names (see _lineLabels._MORE_OPTIONS_BUTTON) when
		UIA finds it, otherwise the header icon at its usual position. Returns
		``(clicked, runtimeId)``, where ``runtimeId`` is the element that had the
		focus before the click, so a caller can wait for the menu to take it.
		Nothing is clicked unless a LINE window is in the foreground.
		"""
		handler = UIAHandler.handler
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not handler or not hwnd or not _isLineForeground(self.processID):
			return False, None
		try:
			previousRuntimeId = _getElementRuntimeId(handler.clientObject.GetFocusedElement())
		except Exception:
			previousRuntimeId = None
		try:
			button = _findNamedButton(handler, hwnd, _lineLabels._MORE_OPTIONS_BUTTON)
		except Exception:
			log.debug("LINE: searching for the More options button failed", exc_info=True)
			button = None
		if button and _clickElementCenter(button):
			return True, previousRuntimeId
		clickPoint = _getChatHeaderIconPoint(hwnd, iconIndex=0)
		if not clickPoint:
			return False, previousRuntimeId
		self._clickAtPosition(clickPoint[0], clickPoint[1], hwnd)
		return True, previousRuntimeId

	# ── Read chat room name ────────────────────────────────────────────

	def _readChatRoomName(self):
		"""Read the current chat room name by OCR of the chat header area."""
		import ctypes
		import ctypes.wintypes

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			ui.message(_("找不到 LINE 視窗"))
			return

		scale = _getDpiScale(hwnd)

		# Get window rect
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winLeft = rect.left
		winTop = rect.top
		winRight = rect.right
		_winWidth = winRight - winLeft

		# Try DWM extended frame bounds for accuracy
		dwmRect = ctypes.wintypes.RECT()
		try:
			DWMWA_EXTENDED_FRAME_BOUNDS = 9
			hr = ctypes.windll.dwmapi.DwmGetWindowAttribute(
				hwnd,
				DWMWA_EXTENDED_FRAME_BOUNDS,
				ctypes.byref(dwmRect),
				ctypes.sizeof(dwmRect),
			)
			if hr == 0:
				if dwmRect.right != rect.right or dwmRect.top != rect.top:
					winLeft = dwmRect.left
					winTop = dwmRect.top
					winRight = dwmRect.right
					_winWidth = winRight - winLeft
		except Exception:
			pass

		# The chat room name is in the header area.
		# Sidebar width is roughly 280px at 96 DPI.
		# Header name starts after the avatar/back button (~80px from sidebar edge)
		# and ends before the toolbar icons (~120px from right edge).
		# The name text is vertically in the range ~30px to ~60px from the top.
		sidebarWidth = int(280 * scale)
		nameLeft = winLeft + sidebarWidth + int(10 * scale)
		nameTop = winTop + int(30 * scale)
		nameRight = winRight - int(120 * scale)
		nameBottom = winTop + int(65 * scale)

		nameW = nameRight - nameLeft
		nameH = nameBottom - nameTop

		if nameW <= 0 or nameH <= 0:
			ui.message(_("無法取得聊天室標題區域"))
			return

		log.info(
			f"LINE: OCR chat room name area: ({nameLeft},{nameTop}) {nameW}x{nameH}, scale={scale:.2f}",
		)

		def _ocrName():
			# The header OCR blocks up to 3 s; run it on a worker thread.
			ocrText = self._ocrWindowArea(
				hwnd,
				region=(nameLeft, nameTop, nameW, nameH),
				sync=True,
				timeout=3.0,
			)
			ocrText = _removeCJKSpaces(ocrText.strip()) if ocrText else ""
			if ocrText:
				_queueUiMessage(ocrText)
			else:
				_queueUiMessage(_("無法讀取聊天室名稱"))

		_runOnWorkerThread(_ocrName, taskName="readChatRoomName")

	@script(
		# Translators: Description of a script to read the current chat room name
		description=_("讀出目前聊天室名稱"),
		gesture="kb:NVDA+windows+t",
		category=SCRIPT_CATEGORY,
	)
	def script_readChatRoomName(self, gesture):
		"""Read the current chat room name.

		Always uses OCR on the header area for reliable reading.
		"""
		try:
			self._readChatRoomName()
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to report LINE version and window type
		description=_("報告 LINE 版本與視窗類型"),
		gesture="kb:NVDA+shift+v",
		category=SCRIPT_CATEGORY,
		speakOnDemand=True,
	)
	def script_reportLineInfo(self, gesture):
		"""Report LINE version, language, and current window type."""
		parts = []
		ver = self._lineVersion
		if ver:
			# Translators: Reported LINE version
			parts.append(_("LINE 版本: {version}").format(version=ver))
		else:
			parts.append(_("LINE 版本: 未知"))
		lang = self._lineLanguage
		if lang:
			parts.append(_("語言: {language}").format(language=lang))
		winType = _classifyLineWindow()
		typeNames = {
			"allinone": _("主視窗（含側邊欄）"),
			"chat": _("獨立聊天視窗"),
			"unknown": _("未知視窗"),
		}
		# Translators: Reported LINE window type
		parts.append(
			_("視窗類型: {type}").format(
				type=typeNames.get(winType, winType),
			),
		)
		qtA11y = _isQtAccessibleSet()
		parts.append(
			_("Qt 無障礙: 已啟用") if qtA11y else _("Qt 無障礙: 未啟用"),
		)
		msg = ", ".join(parts)
		ui.message(msg)
		log.info(f"LINE info: {msg}")

	def script_toggleQtAccessible(self, gesture):
		"""Toggle QT_ACCESSIBILITY=1 user environment variable."""
		currentlySet = _isQtAccessibleSet()
		if currentlySet:
			if _setQtAccessible(False):
				ui.message(_("已移除 QT_ACCESSIBILITY 環境變數，重啟 LINE 後生效"))
			else:
				ui.message(_("移除 QT_ACCESSIBILITY 環境變數失敗"))
		else:
			if _setQtAccessible(True):
				ui.message(_("已設定 QT_ACCESSIBILITY=1，重啟 LINE 後生效"))
			else:
				ui.message(_("設定 QT_ACCESSIBILITY 環境變數失敗"))

	def _lineFileDialogExists(self):
		"""True if a #32770 dialog owned by LINE's process is currently open.

		We enumerate all #32770 windows and check if any belong to LINE's
		process. Using FindWindowW("#32770", None) is wrong because it finds
		ANY #32770 window in the system (e.g. battery warning dialogs).
		"""
		import ctypes
		import ctypes.wintypes

		lineProcessId = self.processID
		found = False

		WNDENUMPROC = ctypes.WINFUNCTYPE(
			ctypes.wintypes.BOOL,
			ctypes.wintypes.HWND,
			ctypes.wintypes.LPARAM,
		)

		def _enumCallback(hwnd, lParam):
			nonlocal found
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(
				hwnd,
				buf,
				256,
			)
			if buf.value == "#32770":
				pid = ctypes.wintypes.DWORD()
				ctypes.windll.user32.GetWindowThreadProcessId(
					hwnd,
					ctypes.byref(pid),
				)
				if pid.value == lineProcessId:
					found = True
					return False  # stop enumeration
			return True  # continue enumeration

		ctypes.windll.user32.EnumWindows(
			WNDENUMPROC(_enumCallback),
			0,
		)
		return found

	# Increments on every file-dialog watch so a stale watcher (e.g. from a
	# previous LINE session) cannot clear a newer session's suppression.
	_fileDialogWaitToken = 0

	def _suppressAddonForFileDialog(self, reason):
		"""Pause addon behavior until LINE's file dialog closes.

		The dialog watch runs on a worker thread (the EnumWindows scan every
		500 ms used to run indefinitely on the main thread); the resume flips
		_suppressAddon back on the main thread via _runOnWorkerThread's onDone.
		"""
		global _suppressAddon
		_suppressAddon = True
		log.info(f"LINE: {reason}, addon suppressed, waiting for file dialog...")
		AppModule._fileDialogWaitToken += 1
		token = AppModule._fileDialogWaitToken

		def _waitForDialogGone():
			# A single stateful poll: "done" only once the dialog has been
			# observed open and then closed. A prior two-phase version (wait
			# to appear, then wait to be gone) would give up on a slow-to-
			# appear dialog and immediately treat "never seen" as "gone",
			# clearing _suppressAddon while the real dialog was still about
			# to open.
			seenDialog = False

			def _dialogSeenThenClosed():
				nonlocal seenDialog
				if token != AppModule._fileDialogWaitToken:
					# Superseded by a newer watch, or LINE exited (terminate
					# bumps the token). Stop now instead of polling a dead
					# process — with a recycled PID — for up to the full
					# give-up window, which would keep this thread and the
					# old AppModule alive across the next LINE session.
					return True
				if self._lineFileDialogExists():
					seenDialog = True
					return False
				return seenDialog

			try:
				from utils.blockUntilConditionMet import blockUntilConditionMet

				# Self-terminates when the dialog closes, when the token is
				# invalidated, or after giving up if no dialog ever appeared.
				blockUntilConditionMet(
					getValue=_dialogSeenThenClosed,
					giveUpAfterSeconds=1800.0,
					intervalBetweenSeconds=0.5,
				)
			except Exception as e:
				log.warning(f"LINE: file dialog poll error: {e}")

		def _resume(result):
			global _suppressAddon
			if token != AppModule._fileDialogWaitToken:
				# A newer suppress request owns the flag now.
				return
			_suppressAddon = False
			log.info("LINE: file dialog closed, addon resumed")

		_runOnWorkerThread(_waitForDialogGone, onDone=_resume, taskName="fileDialogWatch")

	def script_openMessageReader(self, gesture):
		"""Open message reader: choose "Save chat" in LINE's More options menu, parse, and display."""
		if getattr(self, "_messageReaderPending", False):
			ui.message(_("訊息閱讀器正在執行中，請稍候"))
			return
		ui.message(_("正在開啟訊息閱讀器…"))
		self._messageReaderPending = True

		def _openAndChooseSaveChat():
			opened, previousRuntimeId = self._openChatMoreOptionsMenu()
			if not opened:
				return None
			return _chooseNativeMenuItem(
				UIAHandler.handler,
				self.processID,
				previousRuntimeId,
				_lineLabels._MORE_OPTIONS_LABELS,
				"儲存聊天",
			)

		def _afterChoice(result):
			if result and result[0] == _MENU_ITEM_CLICKED:
				speech.cancelSpeech()
				core.callLater(800, self._messageReaderHandleSaveDialog)
				return
			self._messageReaderPending = False
			if result is None:
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
				return
			if result[0] == _MENU_ABANDONED:
				log.debug("LINE: message reader stopped; LINE lost the foreground")
				return
			if result[0] == _MENU_ITEM_NOT_FOUND and _isLineForeground(self.processID):
				from keyboardHandler import KeyboardInputGesture

				KeyboardInputGesture.fromName("escape").send()
			ui.message(_("找不到儲存聊天選項"))

		_runOnWorkerThread(_openAndChooseSaveChat, onDone=_afterChoice, taskName="messageReaderSaveChat")

	def _messageReaderHandleSaveDialog(self, retriesLeft=10):
		"""Find the Save As dialog, set filename to temp folder, and save."""
		import ctypes
		import ctypes.wintypes as wintypes
		import tempfile

		lineProcessId = self.processID
		WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
		dialogHwnd = None

		def _enumCallback(hwnd, lParam):
			nonlocal dialogHwnd
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
			if buf.value == "#32770":
				pid = wintypes.DWORD()
				ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
				if pid.value == lineProcessId:
					if ctypes.windll.user32.IsWindowVisible(hwnd):
						dialogHwnd = hwnd
						return False
			return True

		ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enumCallback), 0)

		if not dialogHwnd:
			if retriesLeft > 0:
				core.callLater(300, lambda: self._messageReaderHandleSaveDialog(retriesLeft - 1))
			else:
				self._messageReaderPending = False
				ui.message(_("未偵測到儲存對話框"))
			return

		# Find the filename edit control in the Save dialog
		# The file dialog has a ComboBoxEx32 > ComboBox > Edit hierarchy
		editHwnd = self._findSaveDialogEdit(dialogHwnd)
		if not editHwnd:
			self._messageReaderPending = False
			ui.message(_("無法操作儲存對話框"))
			return

		# Build temp file path (use system temp dir to avoid locking the addon folder).
		savePath = os.path.join(tempfile.gettempdir(), "lineDesktop_chat_export.txt")
		self._messageReaderSavePath = savePath

		# Pre-delete existing file to suppress the overwrite confirmation dialog
		try:
			if os.path.isfile(savePath):
				os.remove(savePath)
		except Exception as e:
			log.warning(f"LINE: could not pre-delete chat export: {e}")

		# Set the filename
		WM_SETTEXT = 0x000C
		pathBuffer = ctypes.create_unicode_buffer(savePath)
		ctypes.windll.user32.SendMessageW(
			editHwnd,
			WM_SETTEXT,
			0,
			pathBuffer,
		)

		# Press Enter to save (send BN_CLICKED to Save button, or press Enter)
		core.callLater(200, lambda: self._messageReaderPressSave(dialogHwnd))

	def _findSaveDialogEdit(self, dialogHwnd):
		"""Find the filename Edit control inside a standard Windows Save dialog.

		Uses EnumChildWindows to recursively search all descendants, looking for
		an Edit inside a ComboBoxEx32 > ComboBox chain (the filename field).
		Falls back to the first visible+enabled Edit control found.
		"""
		import ctypes
		import ctypes.wintypes as wintypes

		WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
		foundEdit = [None]

		# Primary: find Edit inside any ComboBoxEx32 > ComboBox (filename field)
		def _searchComboBoxEx(hwnd, lParam):
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
			if buf.value == "ComboBoxEx32":
				combo = ctypes.windll.user32.FindWindowExW(hwnd, None, "ComboBox", None)
				if combo:
					edit = ctypes.windll.user32.FindWindowExW(combo, None, "Edit", None)
					if edit and ctypes.windll.user32.IsWindowVisible(edit):
						foundEdit[0] = edit
						return False  # stop enumeration
			return True

		ctypes.windll.user32.EnumChildWindows(
			dialogHwnd,
			WNDENUMPROC(_searchComboBoxEx),
			0,
		)
		if foundEdit[0]:
			log.debug(f"LINE: found save dialog Edit in ComboBoxEx32: {foundEdit[0]}")
			return foundEdit[0]

		# Fallback: first visible+enabled Edit control in the dialog
		def _searchFirstEdit(hwnd, lParam):
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
			if buf.value == "Edit":
				if ctypes.windll.user32.IsWindowVisible(hwnd) and ctypes.windll.user32.IsWindowEnabled(hwnd):
					foundEdit[0] = hwnd
					return False
			return True

		ctypes.windll.user32.EnumChildWindows(
			dialogHwnd,
			WNDENUMPROC(_searchFirstEdit),
			0,
		)
		if foundEdit[0]:
			log.debug(f"LINE: found save dialog Edit (fallback): {foundEdit[0]}")
		else:
			log.debug("LINE: could not find Edit in save dialog")
		return foundEdit[0]

	def _messageReaderPressSave(self, dialogHwnd):
		"""Press Enter in the Save dialog to trigger save, then wait for file."""
		import ctypes
		import ctypes.wintypes as wintypes

		# Find the Save/存檔 button and click it
		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
		saveBtn = [None]

		def _findSaveBtn(hwnd, lParam):
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
			if buf.value == "Button":
				titleBuf = ctypes.create_unicode_buffer(64)
				ctypes.windll.user32.GetWindowTextW(hwnd, titleBuf, 64)
				title = titleBuf.value
				if any(k in title for k in ("存", "Save", "儲存")):
					if ctypes.windll.user32.IsWindowVisible(hwnd):
						saveBtn[0] = hwnd
						return False
			return True

		ctypes.windll.user32.EnumChildWindows(dialogHwnd, WNDENUMPROC(_findSaveBtn), 0)

		BM_CLICK = 0x00F5
		if saveBtn[0]:
			log.debug(f"LINE: clicking Save button hwnd={saveBtn[0]}")
			ctypes.windll.user32.SendMessageW(saveBtn[0], BM_CLICK, 0, 0)
		else:
			# Fallback: send Enter to the dialog
			log.debug("LINE: Save button not found, sending Enter to dialog")
			WM_KEYDOWN = 0x0100
			WM_KEYUP = 0x0101
			VK_RETURN = 0x0D
			ctypes.windll.user32.SendMessageW(dialogHwnd, WM_KEYDOWN, VK_RETURN, 0)
			ctypes.windll.user32.SendMessageW(dialogHwnd, WM_KEYUP, VK_RETURN, 0)

		# Handle possible "overwrite?" confirmation dialog
		core.callLater(500, self._messageReaderHandleOverwrite)

	def _messageReaderHandleOverwrite(self, retriesLeft=5):
		"""Handle the overwrite confirmation dialog if it appears, then read the file."""
		import ctypes
		import ctypes.wintypes as wintypes

		lineProcessId = self.processID
		WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
		dialogHwnd = None

		def _enumCallback(hwnd, lParam):
			nonlocal dialogHwnd
			buf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
			if buf.value == "#32770":
				pid = wintypes.DWORD()
				ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
				if pid.value == lineProcessId:
					if ctypes.windll.user32.IsWindowVisible(hwnd):
						dialogHwnd = hwnd
						return False
			return True

		ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enumCallback), 0)

		if dialogHwnd:
			# There's still a dialog — might be overwrite confirmation
			# Try clicking Yes / pressing Enter
			WM_KEYDOWN = 0x0100
			WM_KEYUP = 0x0101
			VK_RETURN = 0x0D
			ctypes.windll.user32.SendMessageW(dialogHwnd, WM_KEYDOWN, VK_RETURN, 0)
			ctypes.windll.user32.SendMessageW(dialogHwnd, WM_KEYUP, VK_RETURN, 0)
			# Check again after a delay
			if retriesLeft > 0:
				core.callLater(500, lambda: self._messageReaderHandleOverwrite(retriesLeft - 1))
			return

		# No dialog — file should be saved, read it
		core.callLater(300, self._messageReaderOpenFile)

	def _messageReaderOpenFile(self):
		"""Read the saved chat file and open the reader."""
		savePath = getattr(self, "_messageReaderSavePath", None)
		if not savePath or not os.path.isfile(savePath):
			ui.message(_("找不到儲存的聊天紀錄檔案"))
			self._messageReaderPending = False
			return

		try:
			from ._chatParser import parseChatFile

			messages = parseChatFile(savePath)
			if not messages:
				ui.message(_("聊天紀錄中沒有訊息"))
				self._messageReaderPending = False
				return

			from ._messageReader import openMessageReader

			log.info(
				f"LINE: message reader parsed {len(messages)} messages from {savePath}",
			)
			openMessageReader(messages, cleanupPath=savePath)
			self._messageReaderPending = False
		except Exception as e:
			log.warning(f"LINE: message reader parse error: {e}", exc_info=True)
			ui.message(_("訊息閱讀器開啟錯誤"))
			self._messageReaderPending = False

	@script(
		# Translators: Description of the script bound to control+O in LINE
		description=_("開啟檔案選擇對話方塊，並暫停附加元件輔助直到對話方塊關閉"),
		category=SCRIPT_CATEGORY,
	)
	def script_openFileDialog(self, gesture):
		"""Pass Ctrl+O to LINE, suppress addon while file dialog is open."""
		self._suppressAddonForFileDialog("Ctrl+O pressed")
		gesture.send()

	@script(
		# Translators: Description of the script bound to enter in LINE's message input
		description=_("傳送訊息，成功送出時播放提示音"),
		category=SCRIPT_CATEGORY,
	)
	def script_sendMessageAndPlaySound(self, gesture):
		"""Pass Enter to LINE. If a message was sent, play a sound.

		Uses a delayed outcome check: after sending the Enter key,
		waits briefly then checks if the input field is now empty
		(message was sent) or still has text (IME confirmation).
		This works with both IMM32 and TSF input methods.
		"""
		if _suppressAddon:
			gesture.send()
			return
		# Check if the currently focused UIA element is the message input
		isMessageInput = False
		hadTextBeforeEnter = False
		try:
			handler = UIAHandler.handler
			if handler:
				rawElement = handler.clientObject.GetFocusedElement()
				if rawElement:
					ct = rawElement.CurrentControlType
					if ct == 50004:  # Edit control
						if _isMessageInputElement(rawElement, handler):
							isMessageInput = True
							try:
								preVal = rawElement.GetCurrentPropertyValue(30045)  # ValueValue
								log.debug(f"LINE pre-Enter value: {preVal!r}")
								if preVal and isinstance(preVal, str) and preVal.strip():
									hadTextBeforeEnter = True
							except Exception:
								log.debugWarning("Failed to read pre-Enter value", exc_info=True)
		except Exception:
			log.debugWarning("Error detecting message input for send sound", exc_info=True)
		# Pass Enter key through to LINE
		gesture.send()
		# After Enter, schedule a delayed check only when there was text before Enter.
		# This prevents empty input Enter from incorrectly playing the send sound.
		if isMessageInput and hadTextBeforeEnter:

			def _checkFieldAndPlaySound():
				try:
					handler = UIAHandler.handler
					if not handler:
						return
					el = handler.clientObject.GetFocusedElement()
					if not el:
						return
					ct = el.CurrentControlType
					fieldEmpty = True
					if ct == 50004:  # Still on an edit control
						try:
							val = el.GetCurrentPropertyValue(30045)  # ValueValue
							log.debug(f"LINE post-Enter value: {val!r}")
							if val and isinstance(val, str) and val.strip():
								fieldEmpty = False
						except Exception:
							pass
					# Play sound only if the field is now empty (message was sent)
					if fieldEmpty:
						log.debug("LINE: field empty after Enter, playing send sound")
						if os.path.isfile(_SEND_SOUND_PATH):
							try:
								nvwave.playWaveFile(_SEND_SOUND_PATH, asynchronous=True)
							except Exception:
								log.debugWarning("Failed to play send sound", exc_info=True)
					else:
						log.debug("LINE: field still has text after Enter, skipping sound (IME confirm)")
				except Exception:
					log.debugWarning("Error in delayed send sound check", exc_info=True)

			# Wait 300ms for LINE to process the Enter key
			core.callLater(300, _checkFieldAndPlaySound)
		elif isMessageInput:
			log.debug("LINE: Enter on empty message input, skipping send sound")

	# ── Message actions via LINE's own context menu ──────────────────────

	def _nativeMessageMenuAction(self, labelKey, afterCallback=None):
		"""Choose one item from LINE's own context menu for the focused message.

		Sends Shift+F10, which opens LINE's menu of named buttons for the
		focused message, waits for that menu to take the focus, and clicks the
		item that ``labelKey`` names in _lineLabels._MESSAGE_MENU_LABELS.
		"""
		from keyboardHandler import KeyboardInputGesture

		handler = UIAHandler.handler
		try:
			message = handler.clientObject.GetFocusedElement() if handler else None
		except Exception:
			message = None
		if not message:
			ui.message(_("找不到目前的訊息"))
			return
		messageRuntimeId = _getElementRuntimeId(message)

		def _waitUntilReady():
			if not _waitForModifierRelease():
				# Still held after the wait: Shift+F10 would reach LINE combined
				# with them (e.g. Windows+Shift+F10), so do nothing instead.
				return False
			try:
				return _getElementRuntimeId(handler.clientObject.GetFocusedElement()) == messageRuntimeId
			except Exception:
				return False

		def _openMenu(ready):
			if not ready or not _isLineForeground(self.processID):
				log.debug(
					f"LINE: not opening the menu for {labelKey}; keys still held or focus left the message",
				)
				return
			KeyboardInputGesture.fromName("shift+f10").send()
			_runOnWorkerThread(
				lambda: _chooseNativeMenuItem(
					handler,
					self.processID,
					messageRuntimeId,
					_lineLabels._MESSAGE_MENU_LABELS,
					labelKey,
				),
				onDone=_afterChoice,
				taskName="messageMenuChoice",
			)

		def _afterChoice(result):
			status, itemName = result or (_MENU_NOT_OPENED, "")
			if status == _MENU_ITEM_CLICKED:
				# Drop the menu's own focus announcements; what LINE focuses
				# next (the input after Reply, its dialog after Recall) follows.
				speech.cancelSpeech()
				ui.message(itemName)
				if afterCallback:
					core.callLater(500, afterCallback)
				return
			if status == _MENU_ABANDONED:
				log.debug(f"LINE: menu action {labelKey} stopped; LINE lost the foreground")
				return
			if status == _MENU_ITEM_NOT_FOUND and _isLineForeground(self.processID):
				KeyboardInputGesture.fromName("escape").send()
			ui.message(_("找不到「{actionName}」選項").format(actionName=labelKey))

		_runOnWorkerThread(_waitUntilReady, onDone=_openMenu, taskName="messageMenuWait")

	@script(
		gesture="kb:NVDA+windows+r",
		description=_("Reply to the current message"),
		category=SCRIPT_CATEGORY,
	)
	def script_replyMessage(self, gesture):
		"""Reply to the current message via LINE's context menu."""
		self._nativeMessageMenuAction("回覆")

	@script(
		gesture="kb:control+c",
		description=_("Copy the current message"),
		category=SCRIPT_CATEGORY,
	)
	def script_copyMessage(self, gesture):
		"""Copy the current message via LINE's context menu.

		Edit fields and buttons (chat list items, toolbar buttons) get a normal
		Control+C, as does everything while a file dialog has the add-on paused.
		"""
		if _suppressAddon:
			gesture.send()
			return
		try:
			rawEl = UIAHandler.handler.clientObject.GetFocusedElement()
			if rawEl and rawEl.CurrentControlType in (
				_UIA_EDIT_CONTROL_TYPE_ID,
				_UIA_BUTTON_CONTROL_TYPE_ID,
			):
				gesture.send()
				return
		except Exception:
			pass
		self._nativeMessageMenuAction("複製")

	@script(
		gesture="kb:NVDA+windows+k",
		description=_("Save the current message as a file"),
		category=SCRIPT_CATEGORY,
	)
	def script_saveAsMessage(self, gesture):
		"""Save the current message (file, image, voice) via LINE's context menu."""
		if _suppressAddon:
			return

		def _afterSaveAs():
			self._suppressAddonForFileDialog("Save As selected")

		self._nativeMessageMenuAction("另存新檔", afterCallback=_afterSaveAs)

	@script(
		gesture="kb:NVDA+windows+i",
		# Translators: Input help message for the image-description command.
		description=_("使用設定的 AI 服務描述目前的圖片訊息"),
		category=SCRIPT_CATEGORY,
	)
	def script_describeImage(self, gesture):
		"""Send the focused message bubble to Google AI and speak the result.

		Captures the UIA bounding rect of the currently focused message,
		encodes it as PNG, and asks the configured Gemma vision model to
		describe the image.  The HTTP request runs on a worker thread so
		NVDA's main loop stays responsive; the reply is announced via
		``wx.CallAfter(ui.message, ...)`` to stay on the UI thread.
		"""
		if _suppressAddon:
			return

		# Deferred like every dialog-opening script in NVDA's own
		# globalCommands.py: this script runs synchronously on the keyboard
		# hook's call stack, so a modal consent dialog must not be shown
		# directly from here (showing one crashed on the first NVDA+Windows+I
		# use, before consent had ever been granted).
		import wx

		wx.CallAfter(self._describeImageAfterConsent)

	def _describeImageAfterConsent(self):
		if not _ensureImageDescriptionConsent():
			# Translators: Announced when the user declines the AI upload consent dialog
			ui.message(_("已取消，未傳送圖片"))
			return

		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if not rawEl:
				ui.message(_("找不到目前的訊息"))
				return
			rect = rawEl.CurrentBoundingRectangle
			left = int(rect.left)
			top = int(rect.top)
			width = int(rect.right - rect.left)
			height = int(rect.bottom - rect.top)
		except Exception as e:
			log.debug(
				f"LINE: script_describeImage getElement failed: {e}",
				exc_info=True,
			)
			ui.message(_("找不到目前的訊息"))
			return

		if width <= 0 or height <= 0:
			ui.message(_("找不到目前的訊息"))
			return

		pngBytes = _captureRegionAsPng(left, top, width, height)
		if not pngBytes:
			ui.message(_("擷取圖片失敗"))
			return

		ui.message(_("描述圖片中，請稍候"))

		import threading

		def _worker():
			import wx

			prompt = _getRuntimeImagePrompt()
			initialContents = _buildInitialImageContents(pngBytes, prompt)
			description, errMsg = _callImageDescriptionApi(initialContents)
			if description:
				from ._imageDescriptionDialog import openImageDescriptionDialog

				openImageDescriptionDialog(
					_callImageDescriptionApi,
					initialContents,
					prompt,
					description,
				)
			else:
				wx.CallAfter(ui.message, errMsg or _("圖片描述失敗"))

		threading.Thread(target=_worker, daemon=True).start()

	def _isVoiceDurationLine(self, text):
		"""Return True when OCR text looks like a voice duration label."""
		if not text:
			return False
		normalized = str(text).strip().replace("：", ":")
		normalized = re.sub(r"\s+", "", normalized)
		return bool(re.fullmatch(r"\d{1,2}:\d{2}", normalized))

	def _looksLikeVoiceMessageOcr(self, text):
		"""Heuristic for LINE voice messages based on OCR text."""
		if not text:
			return False

		normalized = _removeCJKSpaces(text.strip())
		normalizedLower = normalized.lower()
		lines = [line.strip(" \t,|") for line in normalized.split("\n") if line and line.strip(" \t,|")]
		if not lines:
			return False

		hasDurationLine = any(self._isVoiceDurationLine(line) for line in lines[:4])
		hasActionHint = any(
			keyword in normalized
			for keyword in (
				"另存新檔",
				"分享",
				"Keep",
				"儲存",
			)
		)
		hasFileHint = any(keyword in normalized for keyword in ("下載期限",)) or any(
			unit in normalizedLower for unit in ("kb", "mb", "gb")
		)

		return hasDurationLine and hasActionHint and not hasFileHint

	def _playVoiceMessageViaOcr(self, rawEl, hwnd):
		"""Use message OCR + screenshot-derived ratios to click Play."""
		try:
			rect = rawEl.CurrentBoundingRectangle
			left = int(rect.left)
			top = int(rect.top)
			right = int(rect.right)
			bottom = int(rect.bottom)
			width = right - left
			height = bottom - top
			if width <= 0 or height <= 0:
				return False

			ocrText = self._ocrWindowArea(
				hwnd,
				region=(left, top, width, height),
				sync=True,
				timeout=3.0,
			)
			ocrText = _removeCJKSpaces(ocrText.strip()) if ocrText else ""
			log.info(f"LINE: play voice message OCR: {ocrText!r}")
			if not self._looksLikeVoiceMessageOcr(ocrText):
				return False

			lines = [line.strip(" \t,|") for line in ocrText.split("\n") if line and line.strip(" \t,|")]
			durationIdx = -1
			for idx, line in enumerate(lines[:4]):
				if self._isVoiceDurationLine(line):
					durationIdx = idx
					break

			# Screenshot-based layout:
			# the play icon sits in the upper-middle portion of the message row,
			# roughly one quarter in from the side where the voice bubble lives.
			candidates = []
			if durationIdx > 0:
				candidates.append((0.26, 0.42, "ocr-left"))
			else:
				candidates.append((0.26, 0.42, "ocr-left-ambiguous"))
				candidates.append((0.74, 0.42, "ocr-right-ambiguous"))

			for xRatio, yRatio, label in candidates:
				clickX = left + int(width * xRatio)
				clickY = top + int(height * yRatio)
				clickX = max(left + 8, min(clickX, right - 8))
				clickY = max(top + 8, min(clickY, bottom - 8))
				log.info(
					f"LINE: play voice message clicking {label} at ({clickX}, {clickY})",
				)
				self._clickAtPosition(clickX, clickY, hwnd)
				if len(candidates) > 1:
					time.sleep(0.12)
			return True
		except Exception as e:
			log.debug(
				f"LINE: play voice message OCR fallback failed: {e}",
				exc_info=True,
			)
			return False

	@script(
		gesture="kb:NVDA+windows+p",
		description=_("Play the current voice message"),
		category=SCRIPT_CATEGORY,
	)
	def script_playVoiceMessage(self, gesture):
		"""Play the current voice message by clicking the Play button."""
		if _suppressAddon:
			return
		try:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
			handler = UIAHandler.handler
			if handler:
				rawEl = handler.clientObject.GetFocusedElement()
				if rawEl:
					walker = handler.clientObject.RawViewWalker
					target = None
					queue = [(rawEl, 0)]
					visited = 0
					while queue and visited < 80:
						el, depth = queue.pop(0)
						visited += 1
						try:
							name = el.CurrentName or ""
						except Exception:
							name = ""
						if name and ("播放" in name or "Play" in name):
							target = el
							break
						if depth < 4:
							try:
								child = walker.GetFirstChildElement(el)
								while child:
									queue.append((child, depth + 1))
									child = walker.GetNextSiblingElement(child)
							except Exception:
								pass
					if target:
						try:
							rect = target.CurrentBoundingRectangle
							cx = int((rect.left + rect.right) / 2)
							cy = int((rect.top + rect.bottom) / 2)
							if cx > 0 and cy > 0:
								self._clickAtPosition(cx, cy, hwnd)
								ui.message(_("播放"))
								return
						except Exception:
							pass

					# The OCR fallback blocks (sync OCR up to 3 s); run it on a
					# worker thread so NVDA stays responsive.
					def _ocrFallback():
						if self._playVoiceMessageViaOcr(rawEl, hwnd):
							_queueUiMessage(_("播放"))
						else:
							_queueUiMessage(_("找不到語音訊息的播放按鈕"))

					_runOnWorkerThread(_ocrFallback, taskName="playVoiceMessageOcr")
					return
		except Exception as e:
			log.debug(
				f"LINE: play voice message UIA search failed: {e}",
				exc_info=True,
			)
		ui.message(_("找不到語音訊息的播放按鈕"))

	@script(
		gesture="kb:NVDA+windows+delete",
		description=_("Recall (unsend) the current message"),
		category=SCRIPT_CATEGORY,
	)
	def script_recallMessage(self, gesture):
		"""Recall (unsend) the current message via LINE's context menu.

		LINE then shows its own confirmation dialog, which takes the focus and
		is fully readable, so the add-on leaves it to NVDA.
		"""
		self._nativeMessageMenuAction("收回")

	@script(
		# Translators: Description of the script bound to control+shift+A during LINE calls
		description=_("切換麥克風並讀出開關狀態"),
		category=SCRIPT_CATEGORY,
	)
	def script_toggleMicAndAnnounce(self, gesture):
		"""Pass Ctrl+Shift+A to LINE, then OCR the call window to announce mic status."""
		# Always pass the gesture through first
		gesture.send()

		import ctypes
		import ctypes.wintypes
		import os
		import threading

		# LINE executable names
		_LINE_EXES = {"line.exe", "line_app.exe", "linecall.exe", "linelauncher.exe"}

		def _isLineProcess(pid):
			"""Check if a PID belongs to a LINE process."""
			if pid == self.processID:
				return True
			try:
				PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
				hProc = ctypes.windll.kernel32.OpenProcess(
					PROCESS_QUERY_LIMITED_INFORMATION,
					False,
					pid,
				)
				if hProc:
					try:
						buf = ctypes.create_unicode_buffer(260)
						size = ctypes.wintypes.DWORD(260)
						ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(
							hProc,
							0,
							buf,
							ctypes.byref(size),
						)
						if ok:
							exeName = os.path.basename(buf.value).lower()
							return exeName in _LINE_EXES
					finally:
						ctypes.windll.kernel32.CloseHandle(hProc)
			except Exception:
				pass
			return False

		# ── Find the call window ──────────────────────────────────
		# Strategy: check foreground window FIRST, because during an
		# active video call the call window IS the foreground, and
		# _findIncomingCallWindow() may skip it or pick a wrong overlay.
		hwnd = None
		mainHwnd = None
		try:
			mainHwnd = self.windowHandle
		except Exception:
			pass

		fgHwnd = ctypes.windll.user32.GetForegroundWindow()
		if fgHwnd and fgHwnd != mainHwnd:
			fgPid = ctypes.wintypes.DWORD()
			ctypes.windll.user32.GetWindowThreadProcessId(
				fgHwnd,
				ctypes.byref(fgPid),
			)
			if _isLineProcess(fgPid.value):
				hwnd = fgHwnd
				log.info(
					f"LINE: mic status using foreground window as call window: hwnd={fgHwnd}",
				)

		# Fallback to _findIncomingCallWindow
		if not hwnd:
			hwnd = self._findIncomingCallWindow()

		if not hwnd:
			# Not in a call, just let the keystroke go through silently
			return

		appModRef = self

		def _checkMicStatus():
			"""After a short delay, OCR the call window to detect mic status.

			For video calls, LINE auto-hides the control bar. We move the
			mouse into the window to trigger it to appear, then OCR.
			"""
			import time

			try:
				rect = ctypes.wintypes.RECT()
				ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
				winW = rect.right - rect.left
				winH = rect.bottom - rect.top

				if winW <= 0 or winH <= 0:
					return

				# Save original mouse position
				origPt = ctypes.wintypes.POINT()
				ctypes.windll.user32.GetCursorPos(ctypes.byref(origPt))

				# Move mouse to bottom-center of the call window.
				# This triggers LINE's auto-hiding control bar in video calls.
				# For voice calls, controls are always visible so this is harmless.
				centerX = (rect.left + rect.right) // 2
				bottomY = rect.top + int(winH * 0.80)
				ctypes.windll.user32.SetCursorPos(centerX, bottomY)

				# Wait for controls to appear + mic toggle animation
				time.sleep(0.8)

				# Re-read window rect (may have changed)
				ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
				winW = rect.right - rect.left
				winH = rect.bottom - rect.top

				# Attempt 1: OCR bottom 30% (standard for voice calls)
				bottomRegion = (
					int(rect.left),
					int(rect.top + int(winH * 0.70)),
					winW,
					int(winH * 0.30),
				)
				ocrText = appModRef._ocrWindowArea(
					hwnd,
					region=bottomRegion,
					sync=True,
					timeout=3.0,
				)
				ocrText = ocrText.strip() if ocrText else ""

				if ocrText:
					log.info(f"LINE: mic status OCR (bottom 30%): {ocrText!r}")
					if _announceMicFromOcr(ocrText):
						ctypes.windll.user32.SetCursorPos(origPt.x, origPt.y)
						return

				# Attempt 2: OCR entire window (video call controls
				# may appear in the center or as an overlay)
				ocrText = appModRef._ocrWindowArea(
					hwnd,
					sync=True,
					timeout=3.0,
				)
				ocrText = ocrText.strip() if ocrText else ""

				# Restore mouse position
				ctypes.windll.user32.SetCursorPos(origPt.x, origPt.y)

				if ocrText:
					log.info(f"LINE: mic status OCR (full window): {ocrText!r}")
					_announceMicFromOcr(ocrText)
				else:
					log.debug("LINE: mic status OCR returned empty")

			except Exception as e:
				log.debug(f"LINE: mic status check failed: {e}", exc_info=True)

		def _announceMicFromOcr(ocrText):
			"""Detect mic on/off from OCR text and announce. Returns True if detected."""
			import wx

			# The control-bar tooltip shows the toggle action, so a "mute"
			# label means the mic is currently ON, and vice-versa.
			if _lineLabels.containsAny(ocrText, _lineLabels._MIC_CURRENTLY_ON):
				wx.CallAfter(ui.message, _("麥克風已開啟"))
				return True
			elif _lineLabels.containsAny(ocrText, _lineLabels._MIC_CURRENTLY_OFF):
				wx.CallAfter(ui.message, _("麥克風已關閉"))
				return True
			else:
				log.debug(f"LINE: mic status not detected in OCR: {ocrText!r}")
				return False

		if AppModule._callStatusCheckInProgress:
			# The gesture already toggled the mic; skip a second overlapping
			# OCR/announce pass rather than fighting over the cursor.
			return
		AppModule._callStatusCheckInProgress = True

		def _runAndClear():
			try:
				_checkMicStatus()
			finally:
				AppModule._callStatusCheckInProgress = False

		t = threading.Thread(target=_runAndClear, daemon=True)
		t.start()

	@script(
		# Translators: Description of the script bound to control+shift+V during LINE calls
		description=_("切換鏡頭並讀出開關狀態"),
		category=SCRIPT_CATEGORY,
	)
	def script_toggleCameraAndAnnounce(self, gesture):
		"""Pass Ctrl+Shift+V to LINE, then OCR the call window to announce camera status."""
		# Always pass the gesture through first
		gesture.send()

		import ctypes
		import ctypes.wintypes
		import os
		import threading

		# LINE executable names
		_LINE_EXES = {"line.exe", "line_app.exe", "linecall.exe", "linelauncher.exe"}

		def _isLineProcess(pid):
			"""Check if a PID belongs to a LINE process."""
			if pid == self.processID:
				return True
			try:
				PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
				hProc = ctypes.windll.kernel32.OpenProcess(
					PROCESS_QUERY_LIMITED_INFORMATION,
					False,
					pid,
				)
				if hProc:
					try:
						buf = ctypes.create_unicode_buffer(260)
						size = ctypes.wintypes.DWORD(260)
						ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(
							hProc,
							0,
							buf,
							ctypes.byref(size),
						)
						if ok:
							exeName = os.path.basename(buf.value).lower()
							return exeName in _LINE_EXES
					finally:
						ctypes.windll.kernel32.CloseHandle(hProc)
			except Exception:
				pass
			return False

		# ── Find the call window ──────────────────────────────────
		hwnd = None
		mainHwnd = None
		try:
			mainHwnd = self.windowHandle
		except Exception:
			pass

		fgHwnd = ctypes.windll.user32.GetForegroundWindow()
		if fgHwnd and fgHwnd != mainHwnd:
			fgPid = ctypes.wintypes.DWORD()
			ctypes.windll.user32.GetWindowThreadProcessId(
				fgHwnd,
				ctypes.byref(fgPid),
			)
			if _isLineProcess(fgPid.value):
				hwnd = fgHwnd
				log.info(
					f"LINE: camera status using foreground window as call window: hwnd={fgHwnd}",
				)

		# Fallback to _findIncomingCallWindow
		if not hwnd:
			hwnd = self._findIncomingCallWindow()

		if not hwnd:
			# Not in a call, just let the keystroke go through silently
			return

		appModRef = self

		def _checkCameraStatus():
			"""After a short delay, OCR the call window to detect camera status.

			For video calls, LINE auto-hides the control bar. We move the
			mouse into the window to trigger it to appear, then OCR.
			"""
			import time

			try:
				rect = ctypes.wintypes.RECT()
				ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
				winW = rect.right - rect.left
				winH = rect.bottom - rect.top

				if winW <= 0 or winH <= 0:
					return

				# Save original mouse position
				origPt = ctypes.wintypes.POINT()
				ctypes.windll.user32.GetCursorPos(ctypes.byref(origPt))

				# Move mouse to bottom-center of the call window.
				# This triggers LINE's auto-hiding control bar in video calls.
				centerX = (rect.left + rect.right) // 2
				bottomY = rect.top + int(winH * 0.80)
				ctypes.windll.user32.SetCursorPos(centerX, bottomY)

				# Wait for controls to appear + camera toggle animation
				time.sleep(0.8)

				# Re-read window rect (may have changed)
				ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
				winW = rect.right - rect.left
				winH = rect.bottom - rect.top

				# Attempt 1: OCR bottom 30% (standard for voice calls)
				bottomRegion = (
					int(rect.left),
					int(rect.top + int(winH * 0.70)),
					winW,
					int(winH * 0.30),
				)
				ocrText = appModRef._ocrWindowArea(
					hwnd,
					region=bottomRegion,
					sync=True,
					timeout=3.0,
				)
				ocrText = ocrText.strip() if ocrText else ""

				if ocrText:
					log.info(f"LINE: camera status OCR (bottom 30%): {ocrText!r}")
					if _announceCameraFromOcr(ocrText):
						ctypes.windll.user32.SetCursorPos(origPt.x, origPt.y)
						return

				# Attempt 2: OCR entire window (video call controls
				# may appear in the center or as an overlay)
				ocrText = appModRef._ocrWindowArea(
					hwnd,
					sync=True,
					timeout=3.0,
				)
				ocrText = ocrText.strip() if ocrText else ""

				# Restore mouse position
				ctypes.windll.user32.SetCursorPos(origPt.x, origPt.y)

				if ocrText:
					log.info(f"LINE: camera status OCR (full window): {ocrText!r}")
					_announceCameraFromOcr(ocrText)
				else:
					log.debug("LINE: camera status OCR returned empty")

			except Exception as e:
				log.debug(f"LINE: camera status check failed: {e}", exc_info=True)

		def _announceCameraFromOcr(ocrText):
			"""Detect camera on/off from OCR text and announce. Returns True if detected."""
			import wx

			# "關鏡頭"/"關相機"/"關閉相機" means tooltip says "turn off" → camera is ON
			if any(
				kw in ocrText
				for kw in [
					"關鏡頭",
					"關相機",
					"關閉相機",
					"Turn off camera",
					"turn off camera",
					"カメラをオフ",
					"ปิดกล้อง",
				]
			):
				wx.CallAfter(ui.message, _("相機已開啟"))
				return True
			# "開鏡頭"/"開相機"/"開啟相機" means tooltip says "turn on" → camera is OFF
			elif any(
				kw in ocrText
				for kw in [
					"開鏡頭",
					"開相機",
					"開啟相機",
					"Turn on camera",
					"turn on camera",
					"カメラをオン",
					"เปิดกล้อง",
				]
			):
				wx.CallAfter(ui.message, _("相機已關閉"))
				return True
			else:
				log.debug(f"LINE: camera status not detected in OCR: {ocrText!r}")
				return False

		if AppModule._callStatusCheckInProgress:
			# The gesture already toggled the camera; skip a second overlapping
			# OCR/announce pass rather than fighting over the cursor.
			return
		AppModule._callStatusCheckInProgress = True

		def _runAndClear():
			try:
				_checkCameraStatus()
			finally:
				AppModule._callStatusCheckInProgress = False

		t = threading.Thread(target=_runAndClear, daemon=True)
		t.start()

	# ── Navigate to chat room tabs ─────────────────────────────────────

	def _navigateToChatTab(self, tabName):
		"""Navigate to a specific chat room tab by clicking on it.

		Args:
			tabName: The name of the tab to navigate to
			        (全部, 好友, 群組, 社群, 官方帳號)

		Returns:
			True if successful, False otherwise

		The tab bar in LINE Desktop (Qt6) is at the top of the chat list
		panel.  When a tab is already selected, LINE renders it with a
		highlight / underline style that OCR often cannot read.  Therefore
		we do NOT gate on OCR — we always click at the known position.
		"""
		import ctypes
		import ctypes.wintypes
		import api

		# Known tab positions in client coordinates at 96 DPI.
		# These are measured from the top-left of the window's
		# client area.  The sidebar (icons column) is to the left,
		# X values account for the sidebar width.
		# Y=35 is the vertical centre of the tab bar text.
		_TAB_POSITIONS = {
			"全部": (100, 35),
			"好友": (140, 35),
			"群組": (180, 35),
			"社群": (225, 35),
			"官方帳號": (285, 35),
		}

		if tabName not in _TAB_POSITIONS:
			log.warning(f"LINE: Unknown tab name: {tabName}")
			return False

		try:
			# ── Find the main LINE window ──
			obj = api.getFocusObject()
			if obj and obj.appModule and obj.appModule.appName.lower() in ("line", "line_app", "linecall"):
				hwnd = obj.windowHandle
				# Walk up to the top-level window
				while hwnd:
					parent = ctypes.windll.user32.GetParent(hwnd)
					if not parent:
						break
					hwnd = parent
			else:
				# Fallback: find by window class
				hwnd = ctypes.windll.user32.FindWindowW(
					"Qt663QWindowIcon",
					None,
				)
				if not hwnd:
					hwnd = ctypes.windll.user32.FindWindowW(
						"Qt66QWindowIcon",
						None,
					)
				if not hwnd:
					hwnd = ctypes.windll.user32.FindWindowW(
						"Qt65QWindowIcon",
						None,
					)
				if not hwnd:
					hwnd = ctypes.windll.user32.FindWindowW(
						"Qt5QWindowIcon",
						None,
					)

			if not hwnd:
				log.warning(
					"LINE: Cannot find main window for tab navigation",
				)
				return False

			# ── Calculate click position ──
			baseX, baseY = _TAB_POSITIONS[tabName]

			# Apply DPI scaling
			dpiScale = _getDpiScale(hwnd)
			clientX = int(baseX * dpiScale)
			clientY = int(baseY * dpiScale)

			# Convert client coordinates to screen coordinates
			point = ctypes.wintypes.POINT(clientX, clientY)
			ctypes.windll.user32.ClientToScreen(
				hwnd,
				ctypes.byref(point),
			)
			screenX = point.x
			screenY = point.y

			# ── Click the tab ──
			self._clickAtPosition(screenX, screenY, hwnd)
			log.info(
				f"LINE: Clicked tab '{tabName}' at screen position ({screenX}, {screenY})",
			)

			return True

		except Exception as e:
			log.warning(
				f"LINE: Error navigating to tab {tabName}: {e}",
				exc_info=True,
			)
			return False

	__gestures = {
		"kb:control+o": "openFileDialog",
		"kb:enter": "sendMessageAndPlaySound",
		"kb:control+shift+a": "toggleMicAndAnnounce",
		"kb:control+shift+v": "toggleCameraAndAnnounce",
	}
