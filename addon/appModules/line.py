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
import braille
import core
from logHandler import log
from NVDAObjects.UIA import UIA
import UIAHandler
import ctypes
import ctypes.wintypes
import comtypes
import re
import time
import configparser
import winreg
from ._virtualWindow import VirtualWindow
import addonHandler

addonHandler.initTranslation()


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


def _isChatWindowMode(hwnd=None):
	"""Return True if the current LINE window is a standalone ChatWindow.

	In ChatWindow mode, there is no sidebar — all list items are message items.
	"""
	return _classifyLineWindow(hwnd) == "chat"


# Sound file to play after a message is successfully sent
_SEND_SOUND_PATH = os.path.join(
	os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
	"sounds",
	"sent.wav",
)

# Sound file to play when copy falls back to OCR (result may not be 100% accurate)
_OCR_SOUND_PATH = os.path.join(
	os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
	"sounds",
	"ocr.wav",
)


def _isImeComposing():
	"""Check if an IME composition is currently in progress.

	Uses Windows IMM32 API to detect active composition string.
	Returns True if the user is in the middle of IME character selection.
	"""
	try:
		user32 = ctypes.windll.user32
		imm32 = ctypes.windll.imm32
		hwnd = user32.GetForegroundWindow()
		himc = imm32.ImmGetContext(hwnd)
		if not himc:
			return False
		# GCS_COMPSTR = 0x0008 - composition string
		comp_size = imm32.ImmGetCompositionStringW(himc, 0x0008, None, 0)
		imm32.ImmReleaseContext(hwnd, himc)
		return comp_size > 0
	except Exception:
		return False


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


_CALL_OCR_LOG_NOISE_MARKERS = (
	"INFO -",
	"DEBUG -",
	"WARNING -",
	"ERROR -",
	"Traceback",
	'File "',
	"ConfigManager",
	"Loading config",
	"Config loaded",
	"Confiq loaded",
	"AppData\\Roaming",
	"\\Users\\",
	".pyc",
	"__main__",
	"addonHandler.",
	"external:",
)

_CALL_CHAT_CLOCK_RE = re.compile(
	r"^(?:(?:[上下][午牛年干])|午|am|pm)[0-9IiLlOo]{1,2}:\d{2}$",
	re.IGNORECASE,
)
_CALL_DURATION_RE = re.compile(r"^\d{1,2}:\d{2}(?::\d{2})?$")
_CALL_DURATION_NOISE_LINES = {
	"全選",
	"已",
	"已讀",
	"已謴",
}


def _looksLikeOcrLogNoise(text):
	"""Return True for OCR snippets that are clearly logs, not LINE call records."""
	if not text:
		return False

	normalized = _removeCJKSpaces(str(text).strip())
	if not normalized:
		return False

	lower = normalized.lower()
	if any(marker.lower() in lower for marker in _CALL_OCR_LOG_NOISE_MARKERS):
		return True
	if re.search(r"[A-Za-z]:\\", normalized):
		return True
	return bool(re.search(r"^(?:INFO|DEBUG|WARNING|ERROR)\s*[-:]", normalized, re.MULTILINE))


def _normalizeCallOcrLine(rawLine):
	line = _removeCJKSpaces(str(rawLine).strip())
	line = re.sub(
		r"(?<=\d)\s*[:：•\.。．·･]+\s*(?=\d)",
		":",
		line,
	)
	line = re.sub(r"\s+", "", line)
	return re.sub(r":{2,}", ":", line)


def _isChatClockTimeLine(line):
	return bool(_CALL_CHAT_CLOCK_RE.fullmatch(line or ""))


def _isCallDurationFallbackNoiseLine(line):
	if not line:
		return True
	if line in _CALL_DURATION_NOISE_LINES:
		return True
	if line.startswith("已") and len(line) <= 3:
		return True
	return bool(re.fullmatch(r"[\d:：•\.。．·･\W_]+", line))


def _extractCallDuration(text):
	"""Extract a normalized call duration from OCR text, ignoring clock timestamps."""
	if not text:
		return None

	normalized = _removeCJKSpaces(str(text).strip())
	if _looksLikeOcrLogNoise(normalized):
		return None
	lines = []
	for rawLine in normalized.splitlines():
		line = _normalizeCallOcrLine(rawLine)
		if not line:
			continue
		lines.append(line)

	match = None
	for line in lines:
		if _isChatClockTimeLine(line):
			continue
		if _CALL_DURATION_RE.fullmatch(line):
			match = _CALL_DURATION_RE.fullmatch(line)
			break
	else:
		remainingLines = [line for line in lines if not _isChatClockTimeLine(line)]
		if any(not _isCallDurationFallbackNoiseLine(line) for line in remainingLines):
			return None
		collapsed = "".join(remainingLines)
		match = re.search(r"(?<!\d)(\d{1,2}(?::\d{2}){1,2})(?!\d)", collapsed)
	if not match:
		return None

	parts = match.group(0).split(":")
	if len(parts) == 2:
		return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
	if len(parts) == 3:
		return f"{int(parts[0]):02d}:{int(parts[1]):02d}:{int(parts[2]):02d}"
	return None


def _getCallAnnouncementFromOcr(text):
	"""Return the spoken announcement for a call record OCR snippet."""
	if not text:
		return None

	normalized = _removeCJKSpaces(str(text).strip())
	if _looksLikeOcrLogNoise(normalized):
		return None
	if re.search(r"取消(?:的)?通話", normalized):
		return "取消的通話"
	if "取消" in normalized and re.search(
		r"(?:(?:[上下]午)|午|am|pm)?\s*\d{1,2}\s*:\s*\d{2}",
		normalized,
		re.IGNORECASE,
	):
		return "取消的通話"
	if "無應答" in normalized:
		return "無應答"
	if "未接來電" in normalized:
		return "未接來電"

	duration = _extractCallDuration(normalized)
	if duration:
		return f"通話時間：{duration}"
	return None


_IMAGE_ATTACHMENT_MENU_KEYWORDS = (
	"轉為文字",
	"掃描行動條碼",
	"新增至相簿",
	"設為聊天室背景",
)

_STICKER_MESSAGE_MENU_KEYWORDS = (
	"貼圖小舖",
	"貼圖小鋪",
)


def _looksLikeImageAttachmentMenu(text):
	"""Return True when menu text matches LINE image/photo attachment actions."""
	if not text:
		return False

	normalized = _removeCJKSpaces(str(text).strip())
	if "另存新檔" not in normalized:
		return False

	return any(keyword in normalized for keyword in _IMAGE_ATTACHMENT_MENU_KEYWORDS)


def _looksLikeStickerMessageMenu(text):
	"""Return True when menu text matches LINE sticker-message actions."""
	if not text:
		return False

	normalized = _removeCJKSpaces(str(text).strip()).replace(" ", "")
	return any(keyword in normalized for keyword in _STICKER_MESSAGE_MENU_KEYWORDS)


def _extractDownloadDeadlineAnnouncement(text):
	"""Extract a spoken download-deadline hint from OCR text."""
	if not text:
		return None

	lines = [
		_removeCJKSpaces(str(rawLine).strip())
		for rawLine in str(text).splitlines()
		if rawLine and str(rawLine).strip()
	]
	actionKeywords = ("儲存", "另存新檔", "分享", "Keep")
	for index, line in enumerate(lines):
		if "下載期限" not in line:
			continue
		suffix = line.split("下載期限", 1)[1].lstrip("：: ").strip(" |")
		if suffix:
			return f"下載期限：{suffix}"
		if index + 1 < len(lines):
			nextLine = lines[index + 1].lstrip("：: ").strip(" |")
			if nextLine and not any(keyword in nextLine for keyword in actionKeywords):
				return f"下載期限：{nextLine}"
		return "下載期限未明"
	return None


def _normalizeRecallDialogLine(text):
	"""Normalize a recall-dialog OCR/UIA line for action matching."""
	return _removeCJKSpaces((text or "").strip()).replace(" ", "").lower()


def _matchRecallDialogActionLabel(text):
	"""Map a short recall-dialog action label to its canonical action name."""
	normalized = _normalizeRecallDialogLine(text)
	if not normalized:
		return None

	if normalized in ("無痕收回", "无痕收回"):
		return "無痕收回"
	if normalized.startswith("無痕收回premium") or normalized.startswith("无痕收回premium"):
		return "無痕收回"
	if normalized == "收回":
		return "收回"
	if normalized in {"取消", "關閉", "关闭"}:
		return "取消"
	return None


def _extractRecallDialogActionLabels(text):
	"""Extract actionable recall-dialog buttons from OCR text in reading order."""
	labels = []
	for rawLine in str(text or "").splitlines():
		label = _matchRecallDialogActionLabel(rawLine)
		if label and (not labels or labels[-1] != label):
			labels.append(label)
	return labels


def _isModernRecallDialogText(text, actionLabels=()):
	"""Return True when OCR hints this is the newer recall-dialog layout."""
	actionSet = set(actionLabels or ())
	if "無痕收回" in actionSet:
		return True

	normalized = _removeCJKSpaces(str(text or "").strip()).replace(" ", "").lower()
	if (
		"收回" in actionSet
		and "取消" in actionSet
		and any(keyword in normalized for keyword in ("關閉", "关闭"))
		and "收回已讀訊息時" in normalized
		and "收到通知" in normalized
		and "有可能無法" in normalized
	):
		return True

	return any(
		keyword in normalized
		for keyword in (
			"無痕收回",
			"无痕收回",
			"premium",
			"未讀訊息",
			"未读讯息",
			"任何提醒",
		)
	)


def _isCompactModernRecallDialog(actionLabels=(), isModernDialog=False):
	"""Return True for the newer two-button recall dialog without Premium recall."""
	if not isModernDialog:
		return False
	actionSet = set(actionLabels or ())
	return "收回" in actionSet and "取消" in actionSet and "無痕收回" not in actionSet


def _getRecallConfirmationPrompt(availableActions, isModernDialog=False):
	"""Return the spoken prompt for the current recall confirmation dialog."""
	actionSet = set(availableActions or ())
	if "無痕收回" in actionSet:
		return _("確認要收回嗎？按 Y 收回，按 N 取消，按 P 無痕收回，需要 Premium")
	return _("確認要收回嗎？按 Y 收回，按 N 取消")


def _normalizePhotoTextConsentDialogLine(text):
	"""Normalize a photo-to-text consent OCR/UIA line for action matching."""
	normalized = _removeCJKSpaces((text or "").strip()).replace(" ", "").lower()
	return normalized.strip(".,:;!?|/\\()[]{}<>。！？、．：；")


def _matchPhotoTextConsentActionLabel(text):
	"""Map a short photo-to-text consent action label to its canonical action name."""
	normalized = _normalizePhotoTextConsentDialogLine(text)
	if not normalized:
		return None

	if normalized == "同意":
		return "同意"
	if normalized == "不同意":
		return "不同意"
	return None


def _extractPhotoTextConsentActionLabels(text):
	"""Extract actionable photo-to-text consent buttons from OCR text."""
	labels = []
	for rawLine in str(text or "").splitlines():
		label = _matchPhotoTextConsentActionLabel(rawLine)
		if label and (not labels or labels[-1] != label):
			labels.append(label)
	return labels


def _isPhotoTextConsentDialogText(text, actionLabels=()):
	"""Return True when OCR/UIA text matches LINE's first-run photo consent dialog."""
	normalized = _removeCJKSpaces(str(text or "").strip()).replace(" ", "").lower()
	actionSet = set(actionLabels or ())
	hasButtons = {"同意", "不同意"}.issubset(actionSet)
	if not hasButtons:
		return False

	hasTitle = any(keyword in normalized for keyword in ("同意提供照片", "提供照片"))
	hasUploadNotice = any(keyword in normalized for keyword in ("照片", "相片")) and any(
		keyword in normalized for keyword in ("伺服器", "服务器", "上傳", "上传", "進行處理", "进行处理")
	)
	hasServiceNotice = any(
		keyword in normalized for keyword in ("服務規定", "服务规定", "開始使用", "开始使用")
	)
	return sum(bool(flag) for flag in (hasTitle, hasUploadNotice, hasServiceNotice)) >= 2


def _getPhotoTextConsentPrompt():
	"""Return the spoken prompt for LINE's first-run photo consent dialog."""
	return _("轉為文字會將照片上傳到 LINE 伺服器處理。按 A 同意，按 D 不同意")


def _extractPhotoTextConsentActionClickPoints(ocrLines, dialogRect):
	"""Map photo-to-text consent actions to OCR-derived click points."""
	if not ocrLines or not dialogRect:
		return {}

	dialogLeft, dialogTop, dialogRight, dialogBottom = dialogRect
	dialogWidth = dialogRight - dialogLeft
	dialogHeight = dialogBottom - dialogTop
	if dialogWidth <= 0 or dialogHeight <= 0:
		return {}

	buttonBandTop = dialogTop + (dialogHeight * 0.55)
	preferredCenterX = {
		"同意": dialogLeft + (dialogWidth * 0.38),
		"不同意": dialogLeft + (dialogWidth * 0.62),
	}
	matched = {}

	for index, line in enumerate(ocrLines):
		label = _matchPhotoTextConsentActionLabel(line.get("text", ""))
		rect = line.get("rect")
		if not label or not rect or not _rectsIntersect(rect, dialogRect):
			continue

		rectLeft, rectTop, rectRight, rectBottom = rect
		if rectRight <= rectLeft or rectBottom <= rectTop:
			continue

		centerX = (rectLeft + rectRight) / 2.0
		centerY = (rectTop + rectBottom) / 2.0
		score = (
			1 if centerY >= buttonBandTop else 0,
			centerY,
			-abs(centerX - preferredCenterX[label]),
			(rectRight - rectLeft) * (rectBottom - rectTop),
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


def _getPhotoTextConsentDialogFallbackClickPoint(actionName, dialogRect):
	"""Return a best-effort click point for the photo consent dialog buttons."""
	if not dialogRect:
		return None

	ratios = {
		"同意": (0.38, 0.79),
		"不同意": (0.62, 0.79),
	}
	if actionName not in ratios:
		return None

	left, top, right, bottom = dialogRect
	width = right - left
	height = bottom - top
	if width <= 0 or height <= 0:
		return None

	xRatio, yRatio = ratios[actionName]
	return (
		int(left + width * xRatio),
		int(top + height * yRatio),
	)


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


def _extractRecallDialogActionClickPoints(ocrLines, dialogRect):
	"""Map recall-dialog action labels to OCR-derived click points."""
	if not ocrLines or not dialogRect:
		return {}

	dialogLeft, _dialogTop, dialogRight, _dialogBottom = dialogRect
	dialogCenterX = (dialogLeft + dialogRight) / 2.0
	matched = {}

	for index, line in enumerate(ocrLines):
		label = _matchRecallDialogActionLabel(line.get("text", ""))
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


def _rectIntersectionArea(rectA, rectB):
	"""Return the intersection area of two rectangles."""
	if not rectA or not rectB:
		return 0
	left = max(rectA[0], rectB[0])
	top = max(rectA[1], rectB[1])
	right = min(rectA[2], rectB[2])
	bottom = min(rectA[3], rectB[3])
	if right <= left or bottom <= top:
		return 0
	return (right - left) * (bottom - top)


def _rectIoU(rectA, rectB):
	"""Return the overlap ratio of two rectangles."""
	intersection = _rectIntersectionArea(rectA, rectB)
	if intersection <= 0:
		return 0.0
	areaA = max((rectA[2] - rectA[0]) * (rectA[3] - rectA[1]), 1)
	areaB = max((rectB[2] - rectB[0]) * (rectB[3] - rectB[1]), 1)
	union = areaA + areaB - intersection
	return float(intersection) / float(union or 1)


def _inferRecallDialogTargetsByGeometry(candidates, dialogRect, actionLabels, isModernDialog=False):
	"""Infer unlabeled recall-dialog button targets from geometry."""
	if not dialogRect or not candidates:
		return {}

	actionSet = set(actionLabels or ())
	isCompactModernDialog = _isCompactModernRecallDialog(
		actionLabels,
		isModernDialog=isModernDialog,
	)
	left, top, right, bottom = dialogRect
	dialogWidth = right - left
	dialogHeight = bottom - top
	if dialogWidth <= 0 or dialogHeight <= 0:
		return {}

	dialogCenterX = left + (dialogWidth / 2.0)
	minWidth = max(120, int(dialogWidth * 0.26))
	minHeight = max(24, int(dialogHeight * 0.05))
	maxHeight = max(52, int(dialogHeight * 0.20))
	maxCenterOffset = max(90, int(dialogWidth * 0.22))
	minCenterY = top + (dialogHeight * 0.36)
	maxCenterY = top + (dialogHeight * 0.84)
	expectedCenterY = top + (
		dialogHeight * (0.64 if isCompactModernDialog else 0.56 if isModernDialog else 0.60)
	)

	filtered = []
	for candidate in candidates:
		rect = candidate.get("rect")
		if not rect:
			continue
		width = rect[2] - rect[0]
		height = rect[3] - rect[1]
		if width < minWidth or height < minHeight or height > maxHeight:
			continue

		centerX = (rect[0] + rect[2]) / 2.0
		centerY = (rect[1] + rect[3]) / 2.0
		if abs(centerX - dialogCenterX) > maxCenterOffset:
			continue
		if centerY < minCenterY or centerY > maxCenterY:
			continue

		score = (
			1 if candidate.get("hasInvoke") else 0,
			1 if candidate.get("controlType") == 50000 else 0,
			width,
			-abs(centerX - dialogCenterX),
			-abs(centerY - expectedCenterY),
		)
		filtered.append(
			{
				**candidate,
				"centerX": centerX,
				"centerY": centerY,
				"score": score,
			},
		)

	if not filtered:
		return {}

	filtered.sort(key=lambda item: item["score"], reverse=True)
	deduped = []
	for candidate in filtered:
		isDuplicate = False
		for kept in deduped:
			if _rectIoU(candidate["rect"], kept["rect"]) >= 0.55 or (
				abs(candidate["centerX"] - kept["centerX"]) <= 18
				and abs(candidate["centerY"] - kept["centerY"]) <= 18
			):
				isDuplicate = True
				break
		if not isDuplicate:
			deduped.append(candidate)

	if not deduped:
		return {}

	deduped.sort(
		key=lambda item: (
			item["centerY"],
			-(item["rect"][2] - item["rect"][0]),
		),
	)

	inferred = {}
	if isModernDialog:
		if isCompactModernDialog:
			if "收回" in actionSet and len(deduped) >= 1:
				inferred["收回"] = max(
					deduped,
					key=lambda item: (
						1 if item.get("hasInvoke") else 0,
						1 if item.get("controlType") == 50000 else 0,
						(item["rect"][2] - item["rect"][0]) * (item["rect"][3] - item["rect"][1]),
						-abs(item["centerY"] - expectedCenterY),
						-abs(item["centerX"] - dialogCenterX),
					),
				)
		elif "無痕收回" in actionSet and len(deduped) >= 1:
			inferred["無痕收回"] = deduped[0]
		if "收回" in actionSet:
			if "無痕收回" in actionSet and len(deduped) >= 2:
				inferred["收回"] = deduped[1]
			elif len(deduped) >= 1 and not isCompactModernDialog:
				inferred["收回"] = deduped[0]
	else:
		if "收回" in actionSet and len(deduped) >= 1:
			inferred["收回"] = max(
				deduped,
				key=lambda item: (
					(item["rect"][2] - item["rect"][0]) * (item["rect"][3] - item["rect"][1]),
					-abs(item["centerY"] - expectedCenterY),
				),
			)

	return inferred


def _getRecallDialogFallbackClickPoint(
	actionName,
	dialogRect,
	isModernDialog=False,
	availableActions=(),
):
	"""Return a best-effort click point for the modern recall dialog buttons."""
	if not dialogRect:
		return None

	if _isCompactModernRecallDialog(availableActions, isModernDialog=isModernDialog):
		ratios = {
			"收回": 0.64,
		}
	elif isModernDialog:
		ratios = {
			"無痕收回": 0.49,
			"收回": 0.59,
		}
	else:
		ratios = {
			"收回": 0.58,
		}
	if actionName not in ratios:
		return None

	left, top, right, bottom = dialogRect
	width = right - left
	height = bottom - top
	if width <= 0 or height <= 0:
		return None

	return (
		int(left + width / 2),
		int(top + height * ratios[actionName]),
	)


def _collectPopupMenuRowRects(
	popupHwnd,
	popupRect: tuple[int, int, int, int],
	maxDepth: int = 5,
) -> list[tuple[int, int, int, int]]:
	"""Collect clickable menu-row rectangles from a popup window via UIA."""
	left, top, right, bottom = popupRect
	popupWidth = max(1, right - left)
	popupHeight = max(1, bottom - top)

	try:
		handler = UIAHandler.handler
		client = getattr(handler, "clientObject", None)
		if not client:
			return []
		rootElement = client.ElementFromHandle(popupHwnd)
		if not rootElement:
			return []
		walker = client.RawViewWalker
	except Exception as e:
		log.debug(f"LINE: failed to initialize popup row rect collector: {e}")
		return []

	rowRects: list[tuple[int, int, int, int]] = []
	seen = set()

	def _normalizeRect(rect):
		try:
			rowLeft = max(left, int(rect.left))
			rowTop = max(top, int(rect.top))
			rowRight = min(right, int(rect.right))
			rowBottom = min(bottom, int(rect.bottom))
		except Exception:
			return None
		if rowRight <= rowLeft or rowBottom <= rowTop:
			return None
		return (rowLeft, rowTop, rowRight, rowBottom)

	def _visit(parent, depth=0):
		try:
			child = walker.GetFirstChildElement(parent)
		except Exception:
			return

		idx = 0
		while child and idx < 40:
			try:
				rect = _normalizeRect(child.CurrentBoundingRectangle)
				if rect:
					rowLeft, rowTop, rowRight, rowBottom = rect
					rowWidth = rowRight - rowLeft
					rowHeight = rowBottom - rowTop
					isRowLike = 24 <= rowHeight <= 90 and rowWidth >= int(popupWidth * 0.55)
					isContainerLike = depth < maxDepth and (
						rowHeight > 90
						or (rowWidth >= int(popupWidth * 0.75) and rowHeight >= int(popupHeight * 0.20))
					)
					if isRowLike:
						if rect not in seen:
							seen.add(rect)
							rowRects.append(rect)
					elif isContainerLike:
						_visit(child, depth + 1)
			except Exception:
				pass

			try:
				child = walker.GetNextSiblingElement(child)
			except Exception:
				break
			idx += 1

	_visit(rootElement)
	rowRects.sort(key=lambda rect: (((rect[1] + rect[3]) / 2), rect[0]))
	log.debug(f"LINE: collected {len(rowRects)} popup row rects: {rowRects}")
	return rowRects


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


def _buildMessageBubbleClickPositions(
	rect,
	winTop,
	winBottom,
	includeVerticalOffsets=False,
):
	"""Build conservative right-click probes for a focused message row."""
	try:
		elLeft, elTop, elRight, elBottom = [int(value) for value in rect]
	except Exception:
		return []

	elWidth = elRight - elLeft
	if elWidth <= 0:
		return []

	cx = int((elLeft + elRight) / 2)
	cy = int((elTop + elBottom) / 2)
	clampedCenter = max(winTop + 10, min(cy, winBottom - 10))
	clampedTop = max(elTop + 2, winTop + 10)
	clampedBottom = min(elBottom - 2, winBottom - 10)
	winHeight = max(int(winBottom - winTop), 1)
	isLowerBubble = clampedCenter >= (winTop + int(winHeight * 0.45))

	# Keep the historic six probes first so currently working messages
	# behave the same, then add narrower left-edge fallbacks for short
	# received bubbles whose clickable padding sits further left. Keep
	# the synthetic "center" probe last so copy-read can stop after the
	# dedicated edge/padding probes are exhausted.
	clickPositions = [
		(elLeft + elWidth // 6, clampedCenter, "1/6-left"),
		(elLeft + 5 * elWidth // 6, clampedCenter, "5/6-right"),
		(elLeft + elWidth // 4, clampedCenter, "1/4-left"),
		(elLeft + 3 * elWidth // 4, clampedCenter, "3/4-right"),
		(elLeft + 9 * elWidth // 10, clampedCenter, "9/10-right"),
		(elLeft + 7 * elWidth // 8, clampedCenter, "7/8-right"),
		(elLeft + elWidth // 10, clampedCenter, "1/10-left"),
		(elLeft + elWidth // 8, clampedCenter, "1/8-left"),
	]

	if includeVerticalOffsets:
		if isLowerBubble:
			clickPositions.extend(
				[
					(elLeft + 5 * elWidth // 6, clampedTop, "5/6-top"),
					(elLeft + 5 * elWidth // 6, clampedBottom, "5/6-bottom"),
				],
			)
		clickPositions.extend(
			[
				(elLeft + elWidth // 10, clampedTop, "1/10-top"),
				(elLeft + elWidth // 8, clampedTop, "1/8-top"),
				(elLeft + elWidth // 6, clampedTop, "1/6-top"),
				(elLeft + elWidth // 10, clampedBottom, "1/10-bottom"),
				(elLeft + elWidth // 8, clampedBottom, "1/8-bottom"),
			],
		)
		if not isLowerBubble:
			clickPositions.append(
				(elLeft + 5 * elWidth // 6, clampedBottom, "5/6-bottom"),
			)

	clickPositions.append((cx, clampedCenter, "center"))
	return clickPositions


def _mergeClickPositions(primaryPositions, fallbackPositions):
	"""Prepend unique primary probes before fallback probes."""
	merged = []
	seenPoints = set()
	for positions in (primaryPositions or (), fallbackPositions or ()):
		for position in positions:
			try:
				x, y, label = position
				key = (int(x), int(y))
			except Exception:
				continue
			if key in seenPoints:
				continue
			seenPoints.add(key)
			merged.append((int(x), int(y), label))
	return merged


def _normalizeMessageBubbleOcrLine(text):
	"""Normalize OCR line text for message-bubble metadata detection."""
	normalized = _removeCJKSpaces((text or "").strip())
	normalized = re.sub(
		r"(?<=\d)\s*[:：•\.。．·･℃]+\s*(?=\d)",
		":",
		normalized,
	)
	normalized = re.sub(r"\s+", "", normalized)
	return normalized


def _isMessageBubbleMetadataOcrLine(text):
	"""Return True when an OCR line looks like read/time metadata, not content."""
	normalized = _normalizeMessageBubbleOcrLine(text)
	if not normalized:
		return True

	stripped = normalized.strip(" ,，、.。:：;；!！?？'\"`()[]{}<>-~～")
	if not stripped:
		return True
	if stripped in {"已讀", "未讀"}:
		return True

	timePattern = r"(?:(?:[上下]午)|午|am|pm)?\d{1,2}:\d{1,2}(?::\d{1,2})?"
	if re.fullmatch(timePattern, stripped, re.IGNORECASE):
		return True
	if re.fullmatch(rf"(?:已讀|未讀){timePattern}", stripped, re.IGNORECASE):
		return True
	return False


_LINE_DATE_SEPARATOR_RE = re.compile(
	r"^(?:"
	r"(?:今天|昨天)"
	r"|"
	r"(?:(?:19|20)\d{2}(?:[./-]|年))?"
	r"\d{1,2}(?:[./-]|月)"
	r"\d{1,2}(?:日)?"
	r"(?:(?:\([一二三四五六日天]\))|(?:星期[一二三四五六日天])|(?:週[一二三四五六日天]))?"
	r")$",
)


def _normalizeLineDateSeparatorOcrText(text):
	"""Normalize OCR text for LINE's centered date separator rows."""
	normalized = _removeCJKSpaces((text or "").strip())
	if not normalized:
		return ""
	normalized = normalized.replace("（", "(").replace("）", ")").replace("／", "/").replace("．", ".")
	normalized = re.sub(r"\s+", "", normalized)
	# OCR can misread the standalone date chip "昨天" as visually similar text.
	if normalized == "阼天":
		return "昨天"
	return normalized


def _looksLikeLineDateSeparatorText(text):
	"""Return True when OCR text matches LINE's standalone date separator."""
	normalized = _normalizeLineDateSeparatorOcrText(text)
	if not normalized:
		return False
	return bool(_LINE_DATE_SEPARATOR_RE.fullmatch(normalized))


def _getSpokenLineDateSeparatorText(text):
	"""Return the user-facing text for a recognized LINE date separator chip."""
	rawText = (text or "").strip()
	if not rawText:
		return ""

	normalized = _normalizeLineDateSeparatorOcrText(rawText)
	if normalized in {"今天", "昨天"}:
		return normalized
	return rawText


def _isCenteredLineDateSeparatorOcr(text, ocrLines, rect):
	"""Return True only for compact centered date chips, not normal bubbles."""
	if not _looksLikeLineDateSeparatorText(text):
		return False
	if not ocrLines or not rect:
		return False
	try:
		elLeft, elTop, elRight, elBottom = [int(value) for value in rect]
	except Exception:
		return False
	if elRight <= elLeft or elBottom <= elTop:
		return False

	contentRects = []
	normalizedLines = []
	for line in ocrLines:
		if not isinstance(line, dict):
			continue
		lineText = _normalizeLineDateSeparatorOcrText(line.get("text", ""))
		if not lineText:
			continue
		lineRect = line.get("rect")
		if not lineRect:
			return False
		try:
			lineLeft, lineTop, lineRight, lineBottom = [int(value) for value in lineRect]
		except Exception:
			return False
		lineLeft = max(elLeft, lineLeft)
		lineTop = max(elTop, lineTop)
		lineRight = min(elRight, lineRight)
		lineBottom = min(elBottom, lineBottom)
		if lineRight <= lineLeft or lineBottom <= lineTop:
			return False
		normalizedLines.append(lineText)
		contentRects.append((lineLeft, lineTop, lineRight, lineBottom))

	if not contentRects:
		return False
	if "".join(normalizedLines) != _normalizeLineDateSeparatorOcrText(text):
		return False

	contentLeft = min(lineRect[0] for lineRect in contentRects)
	contentTop = min(lineRect[1] for lineRect in contentRects)
	contentRight = max(lineRect[2] for lineRect in contentRects)
	contentBottom = max(lineRect[3] for lineRect in contentRects)
	rowWidth = max(1, elRight - elLeft)
	rowHeight = max(1, elBottom - elTop)
	contentWidth = max(1, contentRight - contentLeft)
	contentHeight = max(1, contentBottom - contentTop)
	rowCenterX = (elLeft + elRight) / 2.0
	contentCenterX = (contentLeft + contentRight) / 2.0
	centerTolerance = max(24, int(rowWidth * 0.12))
	if abs(contentCenterX - rowCenterX) > centerTolerance:
		return False
	if contentWidth > int(rowWidth * 0.45):
		return False
	if contentHeight > max(40, int(rowHeight * 0.75)):
		return False
	return True


def _buildMessageBubbleOcrClickPositions(ocrLines, rect, winTop, winBottom):
	"""Build OCR-derived probes near the bubble padding instead of the text body."""
	if not ocrLines or not rect:
		return []
	try:
		elLeft, elTop, elRight, elBottom = [int(value) for value in rect]
	except Exception:
		return []
	if elRight <= elLeft or elBottom <= elTop:
		return []

	contentRects = []
	for line in ocrLines:
		if not isinstance(line, dict):
			continue
		if _isMessageBubbleMetadataOcrLine(line.get("text", "")):
			continue
		lineRect = line.get("rect")
		if not lineRect:
			continue
		try:
			lineLeft, lineTop, lineRight, lineBottom = [int(value) for value in lineRect]
		except Exception:
			continue
		lineLeft = max(elLeft, lineLeft)
		lineTop = max(elTop, lineTop)
		lineRight = min(elRight, lineRight)
		lineBottom = min(elBottom, lineBottom)
		if lineRight <= lineLeft or lineBottom <= lineTop:
			continue
		contentRects.append((lineLeft, lineTop, lineRight, lineBottom))

	if not contentRects:
		return []

	contentLeft = min(rect[0] for rect in contentRects)
	contentTop = min(rect[1] for rect in contentRects)
	contentRight = max(rect[2] for rect in contentRects)
	contentBottom = max(rect[3] for rect in contentRects)
	leftGap = max(0, contentLeft - elLeft)
	rightGap = max(0, elRight - contentRight)
	if max(leftGap, rightGap) < 8:
		return []

	contentWidth = max(1, contentRight - contentLeft)
	contentHeight = max(1, contentBottom - contentTop)
	rowWidth = max(1, elRight - elLeft)
	paddingX = max(12, min(28, int(contentWidth * 0.40)))
	edgeInset = max(8, min(18, int(rowWidth * 0.04)))
	preferRight = leftGap > rightGap

	if preferRight:
		anchorX = min(elRight - edgeInset, contentRight + paddingX)
		if anchorX <= contentRight:
			anchorX = min(elRight - edgeInset, contentRight + 8)
		sideLabel = "ocr-right"
	else:
		anchorX = max(elLeft + edgeInset, contentLeft - paddingX)
		if anchorX >= contentLeft:
			anchorX = max(elLeft + edgeInset, contentLeft - 8)
		sideLabel = "ocr-left"

	def _clampY(value):
		return max(winTop + 10, min(int(value), winBottom - 10))

	yCandidates = [
		(_clampY((contentTop + contentBottom) / 2), "center"),
	]
	if contentHeight >= 18:
		yCandidates.extend(
			[
				(_clampY(contentTop + (contentHeight * 0.35)), "upper"),
				(_clampY(contentTop + (contentHeight * 0.65)), "lower"),
			],
		)

	positions = []
	seenY = set()
	for yValue, suffix in yCandidates:
		if yValue in seenY:
			continue
		seenY.add(yValue)
		positions.append((int(anchorX), int(yValue), f"{sideLabel}-{suffix}"))
	return positions


def _hasExhaustedMessageBubbleFallbackProbes(posIdx, clickPositions):
	"""Whether the non-center bubble probes have already been tried."""
	try:
		if posIdx < 0 or not clickPositions:
			return False
		lastRelevantIdx = len(clickPositions) - 1
		if clickPositions[lastRelevantIdx][2] == "center":
			lastRelevantIdx -= 1
		return lastRelevantIdx >= 0 and posIdx >= lastRelevantIdx
	except Exception:
		return False


def _getElementRuntimeId(element):
	if element is None:
		return None
	try:
		return _normalizeRuntimeId(element.GetRuntimeId())
	except Exception:
		return None


def _getFocusedElementRuntimeId():
	try:
		handler = UIAHandler.handler
		client = getattr(handler, "clientObject", None)
		if not client:
			return None
		return _getElementRuntimeId(client.GetFocusedElement())
	except Exception:
		return None


def _messageProbePointHitsTargetElement(handler, targetElement, x, y, targetRuntimeId):
	"""Return False only when UIA can prove a probe point targets an edit field.

	LINE/Qt often reports bubble hit-test elements outside the focused
	ListItem's UIA subtree.  Treat those as usable unless the point clearly
	lands in an Edit control such as the message input or chat search field.
	"""
	if targetRuntimeId is None:
		targetRuntimeId = _getElementRuntimeId(targetElement)
	try:
		client = getattr(handler, "clientObject", None)
		if client is None or not hasattr(client, "ElementFromPoint"):
			return True
		point = ctypes.wintypes.POINT(int(x), int(y))
		hitElement = client.ElementFromPoint(point)
	except Exception:
		log.debug("LINE: copy-read probe hit-test failed", exc_info=True)
		return True
	if hitElement is None:
		return True

	try:
		walker = client.RawViewWalker
	except Exception:
		walker = None

	current = hitElement
	for _depth in range(12):
		if targetRuntimeId is not None and _getElementRuntimeId(current) == targetRuntimeId:
			return True
		try:
			if current.CurrentControlType == 50004:  # Edit
				return False
		except Exception:
			pass
		if walker is None:
			break
		try:
			current = walker.GetParentElement(current)
		except Exception:
			break
		if current is None:
			break
	return True


def _restoreFocusToElement(element, expectedRuntimeId=None):
	"""Best-effort focus restore after an unsafe copy-read probe."""
	try:
		if expectedRuntimeId is not None and _getElementRuntimeId(element) != expectedRuntimeId:
			return
		element.SetFocus()
	except Exception:
		log.debug("LINE: copy-read focus restore failed", exc_info=True)


# Global variable to track the last focused object
# This is needed because api.getFocusObject() sometimes returns the main Window
# even when we handled a gainFocus event for a ListItem.
lastFocusedObject = None

# Track the last UIA element we announced, to avoid re-announcing the same thing
_lastAnnouncedUIAElement = None
_lastAnnouncedUIAName = None
_lastOCRElement = None
# Track the raw UIA focused element (e.g. the edit field) separately from
# the announced element (which could be a list item found via selection detection).
# This allows us to detect stuck-focus even after announcing a list item.
_lastRawFocusedElement = None

# Tracks the newest async request for opening the message context menu so
# stale callLater callbacks do not speak after the user has moved on.
_messageContextMenuRequestId = 0
# Tracks the newest async request for copy-first message reading.
_copyReadRequestId = 0
# Owns clipboard restore for the active copy-read request.
_copyReadClipboardOwnerId = 0
# Debounces delayed focus queries after navigation gestures.
_focusQueryRequestId = 0

# Chat list navigation state.
# When True, up/down arrows will be handled as chat list navigation
# even if UIA focus has moved away (e.g. to message input).
_chatListMode = False
# Cached reference to the search field UIA element, used to find
# the chat list even when focus is elsewhere.
_chatListSearchField = None
# Cached chat room name, set when navigating the chat list.
# Used by NVDA+Windows+T to instantly read the name without OCR.
_currentChatRoomName = None

# Flag to suppress addon while a file dialog is open
_suppressAddon = False

# Image description backend providers exposed in the settings panel.
_IMAGE_DESCRIPTION_PROVIDER_GOOGLE = "google"
_IMAGE_DESCRIPTION_PROVIDER_OLLAMA = "ollama"
_IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS = (
	_IMAGE_DESCRIPTION_PROVIDER_GOOGLE,
	_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
)
_IMAGE_DESCRIPTION_DEFAULT_PROVIDER = _IMAGE_DESCRIPTION_PROVIDER_GOOGLE
_IMAGE_DESCRIPTION_PROVIDER_LABELS = {
	_IMAGE_DESCRIPTION_PROVIDER_GOOGLE: "Google AI",
	_IMAGE_DESCRIPTION_PROVIDER_OLLAMA: "Ollama Cloud",
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
_IMAGE_DESCRIPTION_USER_KEY_FILENAME = "line_desktop_image_api_key.dat"
_IMAGE_DESCRIPTION_USER_OLLAMA_KEY_FILENAME = "line_desktop_ollama_api_key.dat"
_IMAGE_DESCRIPTION_USER_MODEL_FILENAME = "line_desktop_image_model.txt"
_IMAGE_DESCRIPTION_USER_OLLAMA_MODEL_FILENAME = "line_desktop_ollama_model.txt"
_IMAGE_DESCRIPTION_USER_PROMPT_FILENAME = "line_desktop_image_prompt.txt"
# Default Google model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_DEFAULT_MODEL = "gemini-3.1-flash-lite-preview"
# Default Ollama Cloud model used when the user has not picked one in the settings panel.
_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL = "gemini-3-flash-preview:cloud"
# Models exposed in the settings panel dropdown. The order here is the order
# the user sees. Strings are the actual model IDs sent to the API endpoint.
# VERIFY all IDs against: GET https://generativelanguage.googleapis.com/v1beta/models
# Notable uncertainties: "gemini-3-flash-preview" (vs "gemini-3.1-…" naming style),
# "gemini-flash-latest"/"gemini-flash-lite-latest" (versionless aliases),
# "gemini-2.5-flash-lite" (may need a "-preview" suffix),
# "gemma-4-26b-a4b-it"/"gemma-4-31b-it" (availability in Generative Language API),
# "gemma-3-27b-it" (Gemma 3 27B — confirm exact API model ID).
_IMAGE_DESCRIPTION_AVAILABLE_MODELS = (
	"gemini-3-flash-preview",
	"gemini-3.1-flash-lite-preview",
	"gemini-flash-latest",
	"gemini-flash-lite-latest",
	"gemini-2.5-flash",
	"gemini-2.5-flash-lite",
	"gemma-4-26b-a4b-it",
	"gemma-4-31b-it",
	"gemma-3-27b-it",
)
# Vision-capable models commonly available on Ollama Cloud. Users can pick
# whichever they have access to; unavailable IDs surface as an HTTP error.
_IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS = (
	"gemini-3-flash-preview:cloud",
	"qwen3.5:397b-cloud",
	"gemma4:31b-cloud",
)
_IMAGE_DESCRIPTION_ENDPOINT = (
	"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
)
_IMAGE_DESCRIPTION_OLLAMA_ENDPOINT = "https://ollama.com/api/chat"
_IMAGE_DESCRIPTION_DEFAULT_PROMPT = "請用繁體中文簡要描述這張圖片的內容。"
# Maximum length accepted from user-supplied prompts, to avoid pathological
# payloads being sent to the API.
_IMAGE_DESCRIPTION_PROMPT_MAX_LEN = 2000

_IMAGE_API_KEY_SALT_LEN = 16
_IMAGE_API_KEY_MAC_LEN = 16
_IMAGE_API_KEY_PBKDF2_ITERS = 100000

# Sentinel indicating the effective API key has not been computed yet.
_NOT_COMPUTED = object()
# Populated once at AppModule.__init__ via _initEffectiveImageApiKey().
# After that, _getEffectiveImageApiKey() is a pure memory read (no PBKDF2).
_cachedEffectiveImageApiKey = _NOT_COMPUTED
_cachedEffectiveOllamaApiKey = _NOT_COMPUTED
_cachedEffectiveImageProvider = _NOT_COMPUTED
_cachedEffectiveOllamaModel = _NOT_COMPUTED


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


_NOTES_WINDOW_KEYWORDS = ("記事本", "note", "keep", "ノート", "บันทึก", "노트")
_NOTES_OCR_KEYWORDS = (
	"記事本",
	"相簿",
	"已儲存",
	"note",
	"album",
	"saved",
	"keep",
	"ノート",
	"アルバム",
	"保存済み",
	"บันทึก",
	"노트",
)
_NOTES_OCR_CACHE_TTL = 3.0
_notesWindowDetectionCache = {
	"key": None,
	"expiresAt": 0.0,
	"isNotesWindow": False,
}

# Recursion guard to prevent infinite _get_name → _getDeepText → _get_name loops
# Uses UIA runtime IDs (stable across Python wrapper recreation) instead of id(self)
_nameRecursionGuard = set()

# Thread-level recursion depth counter as ultimate safety net
_nameRecursionDepth = 0
_MAX_NAME_RECURSION_DEPTH = 5


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
				{"text": prompt if prompt is not None else _getEffectiveImagePrompt()},
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
	if provider == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
		return _callOllamaImageDescriptionApi(
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


def _describeImageBytes(pngBytes, timeout=None):
	"""Send PNG image bytes to the active image-description provider and return (description, error_msg).

	Returns (text, None) on success or (None, error_msg) on failure.
	"""
	if not pngBytes:
		return None, _("無法取得圖片資料")
	contents = _buildInitialImageContents(pngBytes)
	return _callImageDescriptionApi(contents, timeout=timeout)


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


# Physical-pixel cap for the right edge of LINE's left sidebar (icons column +
# chat list panel). Empirically, LINE renders the sidebar at a fixed physical
# pixel width regardless of DPI scaling (a Qt quirk on Windows): observed at
# both 96 DPI and 144 DPI maximized the sidebar still ends near X=547. We cap
# conservatively at 540 (observed edge ≈ 547, 7 px safety margin) so we do NOT
# multiply by DPI scale — that double-counts and lets the cap drift back into
# the message area when the window is maximized at high DPI.
_LINE_SIDEBAR_RIGHT_BASE = 540


def _getSidebarRightBoundary(wndRect, _hwnd=None, percentage=0.45):
	"""Return the screen-space X separating LINE's sidebar from the message area.

	LINE keeps the sidebar at a fixed physical-pixel width regardless of window
	size or DPI, so a pure ``wndWidth * percentage`` threshold lands deep in the
	message area when the window is maximized. We cap the percentage at the
	empirical sidebar edge so the same code path works for standard windows,
	maximized windows, and high-DPI configurations.

	``_hwnd`` is reserved for future per-monitor DPI queries and is not used yet.
	"""
	try:
		wndLeft = int(wndRect.left)
		wndWidth = int(wndRect.right - wndRect.left)
	except Exception:
		return None
	if wndWidth <= 0:
		return None
	return wndLeft + min(int(wndWidth * percentage), _LINE_SIDEBAR_RIGHT_BASE)


def _scheduleQueryAndSpeakUIAFocus(delay=100):
	"""Schedule a focus query, dropping stale callbacks when navigation repeats quickly."""
	global _focusQueryRequestId
	_invalidateActiveCopyRead()
	_invalidateActiveMessageContextMenu()
	_focusQueryRequestId += 1
	requestId = _focusQueryRequestId

	def _run():
		if requestId != _focusQueryRequestId:
			return
		_queryAndSpeakUIAFocus()

	core.callLater(delay, _run)


def _invalidateActiveCopyRead():
	"""Expire any in-flight copy-read chain before new focus work begins."""
	global _copyReadRequestId
	_copyReadRequestId += 1


def _invalidateActiveMessageContextMenu():
	"""Expire any in-flight message context-menu chain before new focus work begins."""
	global _messageContextMenuRequestId
	_messageContextMenuRequestId += 1


def _getForegroundWindowInfo():
	"""Return foreground hwnd, lowercased title, and screen rect."""
	try:
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			return None, "", None
		buf = ctypes.create_unicode_buffer(512)
		ctypes.windll.user32.GetWindowTextW(hwnd, buf, 512)
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		return (
			hwnd,
			(buf.value or "").lower(),
			(
				int(rect.left),
				int(rect.top),
				int(rect.right),
				int(rect.bottom),
			),
		)
	except Exception:
		return None, "", None


def _getWindowProcessId(hwnd):
	"""Return the owning process ID for a window handle."""
	try:
		if not hwnd:
			return 0
		pid = ctypes.wintypes.DWORD()
		ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
		return int(pid.value)
	except Exception:
		return 0


def _shouldDismissCopyReadMenu(targetHwnd):
	"""Return True only while the active popup still belongs to LINE."""
	try:
		currentHwnd = ctypes.windll.user32.GetForegroundWindow()
	except Exception:
		return False
	if not currentHwnd or not targetHwnd:
		return False
	if currentHwnd == targetHwnd:
		return True
	targetPid = _getWindowProcessId(targetHwnd)
	if not targetPid:
		return False
	return _getWindowProcessId(currentHwnd) == targetPid


def _rectsIntersect(rectA, rectB):
	if not rectA or not rectB:
		return False
	return not (rectA[2] <= rectB[0] or rectA[0] >= rectB[2] or rectA[3] <= rectB[1] or rectA[1] >= rectB[3])


def _isRectVisibleInForegroundWindow(left, top, right, bottom):
	"""Return True when a screen rect overlaps the current foreground window."""
	_fgHwnd, _fgTitle, fgRect = _getForegroundWindowInfo()
	if not fgRect:
		return True
	return _rectsIntersect((left, top, right, bottom), fgRect)


def _isElementVisibleInForegroundWindow(element):
	"""Return False when UIA gives a virtualized/off-window element rect."""
	try:
		rect = element.CurrentBoundingRectangle
		left = int(rect.left)
		top = int(rect.top)
		right = int(rect.right)
		bottom = int(rect.bottom)
	except Exception:
		return True
	if right <= left or bottom <= top:
		return False
	return _isRectVisibleInForegroundWindow(left, top, right, bottom)


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


def _matchMessageContextMenuLabel(text):
	"""Map OCR text to a known LINE message context-menu label."""
	try:
		from ._virtualWindows import messageContextMenu as messageContextMenuModule

		return messageContextMenuModule._matchMenuLabel(text)
	except Exception:
		return None


def _extractMatchedMessageContextMenuLabels(ocrText):
	"""Return OCR lines plus any labels that look like real message menu items."""
	popupLines = [_removeCJKSpaces(line.strip()) for line in (ocrText or "").split("\n") if line.strip()]
	lineMatches = []
	matchedLabels = []
	for line in popupLines:
		label = _matchMessageContextMenuLabel(line)
		lineMatches.append((line, label))
		if label:
			matchedLabels.append(label)
	return popupLines, lineMatches, matchedLabels


def _resolvePopupMenuLabelClickPoint(
	targetLabel,
	ocrLines,
	popupRect,
	popupRowRects=None,
	buildMenuElements=None,
):
	"""Resolve a popup-menu label to a stable click point via OCR + row rects."""
	if not targetLabel or not ocrLines or not popupRect:
		return None

	if buildMenuElements is None:
		try:
			from ._virtualWindows.messageContextMenu import _buildMenuElements as buildMenuElements
		except Exception as e:
			log.debug(
				f"LINE: failed to load popup OCR menu builder: {e}",
				exc_info=True,
			)
			return None

	try:
		elements = buildMenuElements(
			ocrLines,
			popupRect,
			rowRects=popupRowRects,
		)
	except Exception as e:
		log.debug(
			f"LINE: popup OCR menu builder failed for {targetLabel!r}: {e}",
			exc_info=True,
		)
		return None

	for elementIndex, element in enumerate(elements):
		if element.get("name") != targetLabel:
			continue
		clickPoint = element.get("clickPoint")
		if not clickPoint or len(clickPoint) != 2:
			continue
		try:
			clickX = int(clickPoint[0])
			clickY = int(clickPoint[1])
		except Exception:
			continue
		return {
			"clickPoint": (clickX, clickY),
			"index": elementIndex,
			"count": len(elements),
		}
	return None


def _sendGestureWithAddonSuppressed(gestureName):
	"""Send a key directly to LINE without routing it back through addon scripts."""
	from keyboardHandler import KeyboardInputGesture

	global _suppressAddon
	previousSuppressState = _suppressAddon
	_suppressAddon = True
	try:
		KeyboardInputGesture.fromName(gestureName).send()
	finally:
		_suppressAddon = previousSuppressState


def _popupLooksLikeMessageContextMenu(appMod, hwnd, popupRect):
	"""Return True only when popup OCR matches known LINE message-menu labels."""
	left, top, right, bottom = popupRect
	popupW = right - left
	popupH = bottom - top
	if popupW <= 0 or popupH <= 0:
		return False
	try:
		ocrText = appMod._ocrWindowArea(
			hwnd,
			region=(left, top, popupW, popupH),
			sync=True,
			timeout=2.0,
		)
		popupLines, _popupLineMatches, matchedLabels = _extractMatchedMessageContextMenuLabels(ocrText)
		if matchedLabels:
			log.info(
				f"LINE: message context menu confirmed via popup OCR: "
				f"lines={popupLines}, matched={matchedLabels}",
			)
			return True
		log.debug(
			f"LINE: popup OCR did not resemble a message context menu: {popupLines}",
		)
	except Exception as e:
		log.debug(
			f"LINE: popup OCR validation failed: {e}",
			exc_info=True,
		)
	return False


def _isNotesWindowContext(element, walker, allowOcr=True):
	"""Detect whether the foreground LINE window is currently showing a notes panel."""
	global _notesWindowDetectionCache
	hwnd, windowTitle, windowRect = _getForegroundWindowInfo()
	isNotesWindow = any(kw in windowTitle for kw in _NOTES_WINDOW_KEYWORDS)

	if not isNotesWindow:
		try:
			ancestor = element
			for _depth in range(5):
				ancestor = walker.GetParentElement(ancestor)
				if not ancestor:
					break
				try:
					ancestorName = ancestor.CurrentName
					if ancestorName and isinstance(ancestorName, str):
						ancestorNameLower = ancestorName.strip().lower()
						if any(kw in ancestorNameLower for kw in _NOTES_WINDOW_KEYWORDS):
							isNotesWindow = True
							log.info(
								f"LINE: Detected notes window via UIA ancestor name: {ancestorName!r}",
							)
							break
				except Exception:
					continue
		except Exception:
			log.debug("LINE: UIA ancestor notes detection failed", exc_info=True)

	cacheKey = None
	if hwnd and windowRect:
		cacheKey = (int(hwnd), windowTitle, windowRect)
		cache = _notesWindowDetectionCache
		if not isNotesWindow and cache["key"] == cacheKey and cache["expiresAt"] > time.monotonic():
			return cache["isNotesWindow"], windowTitle

	if isNotesWindow or not allowOcr or not cacheKey or not windowRect:
		return isNotesWindow, windowTitle

	left, top, right, bottom = windowRect
	winWidth = right - left
	winHeight = bottom - top
	try:
		if winWidth > 0 and winHeight > 0:
			log.debug(
				f"LINE: Attempting OCR notes detection, window size: {winWidth}x{winHeight}",
			)
			ocrWidth = winWidth
			ocrHeight = int(winHeight * 0.20)
			ocrLeft = left
			ocrTop = top
			log.debug(
				f"LINE: OCR region: left={ocrLeft}, top={ocrTop}, width={ocrWidth}, height={ocrHeight}",
			)

			import screenBitmap
			from contentRecog import uwpOcr
			import threading

			langs = uwpOcr.getLanguages()
			if langs:
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
				if not ocrLang:
					ocrLang = langs[0]
				log.debug(f"LINE: Using OCR language: {ocrLang}")

				sb = screenBitmap.ScreenBitmap(ocrWidth, ocrHeight)
				pixels = sb.captureImage(ocrLeft, ocrTop, ocrWidth, ocrHeight)
				recognizer = uwpOcr.UwpOcr(language=ocrLang)
				resultHolder = [None]
				event = threading.Event()

				class _ImgInfo:
					def __init__(self, w, h):
						self.recogWidth = w
						self.recogHeight = h
						self.resizeFactor = 1

					def convertXToScreen(self, x):
						return ocrLeft + x

					def convertYToScreen(self, y):
						return ocrTop + y

					def convertWidthToScreen(self, w):
						return w

					def convertHeightToScreen(self, h):
						return h

				def _onOcr(result):
					resultHolder[0] = result
					event.set()

				recognizer.recognize(pixels, _ImgInfo(ocrWidth, ocrHeight), _onOcr)
				event.wait(timeout=2.0)
				result = resultHolder[0]
				if result and not isinstance(result, Exception):
					ocrText = getattr(result, "text", "") or ""
					ocrText = _removeCJKSpaces(ocrText.strip())
					log.debug(f"LINE: OCR result text: {ocrText!r}")
					isNotesWindow = any(kw in ocrText.lower() for kw in _NOTES_OCR_KEYWORDS)
					if isNotesWindow:
						log.info(f"LINE: Detected notes window via OCR: {ocrText!r}")
					else:
						log.debug("LINE: No notes keywords found in OCR text")
				else:
					log.debug(f"LINE: OCR returned no result or error: {result}")
			else:
				log.debug("LINE: No OCR languages available")
	except Exception:
		log.debug("LINE: notes window OCR detection failed", exc_info=True)
	finally:
		_notesWindowDetectionCache = {
			"key": cacheKey,
			"expiresAt": time.monotonic() + _NOTES_OCR_CACHE_TTL,
			"isNotesWindow": isNotesWindow,
		}

	return isNotesWindow, windowTitle


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


def _getObjectNameDirect(obj):
	"""Get an object's name WITHOUT triggering _get_name overrides.

	This prevents the infinite recursion where _get_name → _getDeepText
	→ child.name → child._get_name → _getDeepText (depth resets to 0).

	We access the raw UIA element's CurrentName directly when possible.
	"""
	# Prefer raw UIA name (bypasses Python _get_name completely)
	if hasattr(obj, "UIAElement") and obj.UIAElement is not None:
		try:
			name = obj.UIAElement.CurrentName
			if name and name.strip():
				return name.strip()
		except Exception:
			pass
	# Fallback: try the base NVDAObject.name (skip overlay _get_name)
	try:
		name = UIA.name.fget(obj) if isinstance(obj, UIA) else obj.name
		if name and name.strip():
			return name.strip()
	except Exception:
		pass
	return ""


def _getDeepText(obj, maxDepth=3, _depth=0):
	"""Recursively collect non-empty text from an object and its children.

	Falls back to _getTextViaUIAFindAll if childCount is 0.
	Uses _getObjectNameDirect to avoid triggering _get_name overrides
	which would cause infinite recursion.
	"""
	if _depth > maxDepth or obj is None:
		return []
	texts = []
	# Get this object's name directly (bypassing _get_name overrides)
	name = _getObjectNameDirect(obj)
	if name:
		texts.append(name)
	# Try children via NVDA's normal enumeration
	childCount = 0
	try:
		childCount = obj.childCount
	except Exception:
		pass
	if childCount > 0:
		# If we already found text at this level and it's not a container,
		# don't recurse deeper to avoid duplication
		if texts and obj.role not in (
			controlTypes.Role.LIST,
			controlTypes.Role.LISTITEM,
			controlTypes.Role.GROUPING,
			controlTypes.Role.SECTION,
			controlTypes.Role.TREEVIEWITEM,
			controlTypes.Role.PANE,
			controlTypes.Role.WINDOW,
		):
			return texts

		# If it's a generic container with children, recurse
		try:
			children = obj.children
			for child in children:
				childTexts = _getDeepText(child, maxDepth, _depth + 1)
				texts.extend(childTexts)
		except Exception:
			pass
	else:
		# No children exposed to NVDA? Try UIA FindAll
		uiaTexts = _getTextViaUIAFindAll(obj)
		if uiaTexts:
			texts.extend(uiaTexts)

	# Deduplicate while preserving order
	seen = set()
	unique_texts = []
	for t in texts:
		if t not in seen:
			unique_texts.append(t)
			seen.add(t)
	return unique_texts


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


def _ocrReadElementText(rawElement, appModuleRef=None, preferCallAnnouncement=False):
	"""Perform OCR on a raw UIA element's bounding rect and speak the result.

	This is used as a fallback when all UIA text extraction strategies
	return empty. LINE's Qt6 renders text via GPU, so OCR is the only
	way to read it.

	The OCR is asynchronous — result is spoken via wx.CallAfter on main thread.
	"""
	# Skip OCR if addon is suppressed (e.g. file dialog is open)
	if _suppressAddon:
		log.debug("LINE OCR: suppressed (addon paused)")
		return
	try:
		rect = rawElement.CurrentBoundingRectangle
		left = int(rect.left)
		top = int(rect.top)
		width = int(rect.right - rect.left)
		height = int(rect.bottom - rect.top)
		right = int(rect.right)
		bottom = int(rect.bottom)

		if width <= 0 or height <= 0:
			return
		if not _isRectVisibleInForegroundWindow(left, top, right, bottom):
			log.debug(
				f"LINE OCR skipped for off-window element at ({left},{top}) {width}x{height}",
			)
			return

		import screenBitmap

		sb = screenBitmap.ScreenBitmap(width, height)
		pixels = sb.captureImage(left, top, width, height)

		from contentRecog import uwpOcr

		langs = uwpOcr.getLanguages()
		if not langs:
			return

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
		if not ocrLang:
			ocrLang = langs[0]

		recognizer = uwpOcr.UwpOcr(language=ocrLang)
		resizeFactor = recognizer.getResizeFactor(width, height)
		# 確保至少 2x 放大，提高小字辨識精度
		if resizeFactor < 2:
			resizeFactor = 2

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

		# Store references to prevent garbage collection during async OCR
		_ocrReadElementText._recognizer = recognizer
		_ocrReadElementText._pixels = ocrPixels
		_ocrReadElementText._imgInfo = imgInfo

		def _onOcrResult(result):
			"""Handle OCR result on background thread, dispatch to main."""
			import wx

			def _handleOnMain():
				try:
					if isinstance(result, Exception):
						log.debug(f"LINE OCR error: {result}")
						return
					# LinesWordsResult has .text with the full recognized string
					ocrText = getattr(result, "text", "") or ""
					ocrText = _removeCJKSpaces(ocrText.strip())
					if ocrText:
						announcement = None
						if preferCallAnnouncement:
							announcement = _getCallAnnouncementFromOcr(ocrText)
						log.info(f"LINE OCR nav result: {ocrText!r}")
						if not preferCallAnnouncement:
							_storeChatNameFromText(ocrText)
						speech.cancelSpeech()
						ui.message(announcement or ocrText)
					else:
						log.debug("LINE OCR: no text found in element")
				except Exception as e:
					log.debug(f"LINE OCR result handler error: {e}")
				finally:
					_ocrReadElementText._recognizer = None
					_ocrReadElementText._pixels = None
					_ocrReadElementText._imgInfo = None

			wx.CallAfter(_handleOnMain)

		try:
			recognizer.recognize(ocrPixels, imgInfo, _onOcrResult)
			log.debug(f"LINE OCR started for element at ({left},{top}) {width}x{height}")
		except Exception as e:
			log.debug(f"LINE OCR recognize error: {e}")
			_ocrReadElementText._recognizer = None
			_ocrReadElementText._pixels = None
			_ocrReadElementText._imgInfo = None
	except Exception:
		log.debug("_ocrReadElementText failed", exc_info=True)


def _findSelectedItemInList(handler, focusedElement):
	"""Walk up from focusedElement to find a parent List, then find the selected item.

	LINE's Qt6 keeps UIA focus on the edit field even when arrows move
	selection in a list. We walk up to find the List, then use
	SelectionItem property or walk children to find the selected ListItem.
	Returns the selected item's UIA element, or None.
	"""
	try:
		walker = handler.clientObject.RawViewWalker
		parent = walker.GetParentElement(focusedElement)
		depth = 0
		while parent and depth < 10:
			try:
				ct = parent.CurrentControlType
				if ct == 50008:  # UIA List
					# Found a List - walk children to find selected item
					# Use LegacyIAccessibleState (propId=30094 is description,
					# 30100 is LegacyIAccessibleState)
					try:
						condition = handler.clientObject.CreatePropertyCondition(
							30003,
							50007,  # ControlType == ListItem
						)
						items = parent.FindAll(UIAHandler.TreeScope_Children, condition)
						if items:
							for i in range(items.Length):
								item = items.GetElement(i)
								try:
									# LegacyIAccessibleState: 0x2 = SELECTED
									state = item.GetCurrentPropertyValue(30100)
									if isinstance(state, int) and (state & 0x2):
										log.info(f"LINE: found selected list item via state={state}")
										return item
								except Exception:
									pass
					except Exception:
						pass

					# Also try SelectionItemPattern.IsSelected (propId=30079)
					try:
						condition = handler.clientObject.CreatePropertyCondition(
							30003,
							50007,  # ControlType == ListItem
						)
						items = parent.FindAll(UIAHandler.TreeScope_Children, condition)
						if items:
							for i in range(items.Length):
								item = items.GetElement(i)
								try:
									isSelected = item.GetCurrentPropertyValue(30079)
									if isSelected:
										log.info(
											"LINE: found selected list item via SelectionItemPattern.IsSelected",
										)
										return item
								except Exception:
									pass
					except Exception:
						pass

					# Fallback: try HasKeyboardFocus on each ListItem
					try:
						condition = handler.clientObject.CreatePropertyCondition(
							30003,
							50007,  # ControlType == ListItem
						)
						items = parent.FindAll(UIAHandler.TreeScope_Children, condition)
						if items:
							for i in range(items.Length):
								item = items.GetElement(i)
								try:
									if item.CurrentHasKeyboardFocus:
										log.info("LINE: found selected list item via HasKeyboardFocus")
										return item
								except Exception:
									pass
					except Exception:
						pass

					# Fallback: check LegacyIAccessibleState for FOCUSED (0x4)
					try:
						condition = handler.clientObject.CreatePropertyCondition(
							30003,
							50007,  # ControlType == ListItem
						)
						items = parent.FindAll(UIAHandler.TreeScope_Children, condition)
						if items:
							for i in range(items.Length):
								item = items.GetElement(i)
								try:
									state = item.GetCurrentPropertyValue(30100)
									# 0x4 = STATE_SYSTEM_FOCUSED
									if isinstance(state, int) and (state & 0x4):
										log.info(f"LINE: found selected list item via FOCUSED state={state}")
										return item
								except Exception:
									pass
					except Exception:
						pass

					# Store the parent list for OCR fallback
					_findSelectedItemInList._lastListElement = parent
					break
			except Exception:
				pass
			try:
				parent = walker.GetParentElement(parent)
			except Exception:
				break
			depth += 1
	except Exception:
		pass
	return None


# Track the index of the last navigated chat list item.
# This is used when UIA cannot detect which item is selected — we
# fall back to positional tracking.
_chatListCurrentIndex = -1


def _findListElement(handler, startElement):
	"""Walk up from startElement to find a parent List element.

	Returns (listElement, walker) or (None, None).
	"""
	try:
		walker = handler.clientObject.RawViewWalker
		parent = walker.GetParentElement(startElement)
		depth = 0
		while parent and depth < 10:
			try:
				ct = parent.CurrentControlType
				if ct == 50008:  # UIA List
					return parent, walker
			except Exception:
				pass
			try:
				parent = walker.GetParentElement(parent)
			except Exception:
				break
			depth += 1
	except Exception:
		pass
	return None, None


def _getListItems(handler, listElement):
	"""Get all ListItem children of a List element.

	Returns a UIA element array, or None.
	"""
	try:
		condition = handler.clientObject.CreatePropertyCondition(
			30003,
			50007,  # ControlType == ListItem
		)
		items = listElement.FindAll(UIAHandler.TreeScope_Children, condition)
		return items
	except Exception:
		return None


def _findCurrentItemIndex(items):
	"""Find the currently selected/focused item index in a list.

	Tries multiple UIA strategies. Returns the index (0-based), or -1.
	"""
	if not items:
		return -1
	count = items.Length
	for i in range(count):
		try:
			item = items.GetElement(i)
			# Check SELECTED state (0x2)
			state = item.GetCurrentPropertyValue(30100)
			if isinstance(state, int) and (state & 0x2):
				return i
		except Exception:
			pass
	for i in range(count):
		try:
			item = items.GetElement(i)
			if item.GetCurrentPropertyValue(30079):  # SelectionItemPattern.IsSelected
				return i
		except Exception:
			pass
	for i in range(count):
		try:
			item = items.GetElement(i)
			if item.CurrentHasKeyboardFocus:
				return i
		except Exception:
			pass
	for i in range(count):
		try:
			item = items.GetElement(i)
			state = item.GetCurrentPropertyValue(30100)
			if isinstance(state, int) and (state & 0x4):  # FOCUSED
				return i
		except Exception:
			pass
	return -1


def _clickElement(element):
	"""Click the center of a UIA element's bounding rectangle."""
	try:
		rect = element.CurrentBoundingRectangle
		cx = int((rect.left + rect.right) / 2)
		cy = int((rect.top + rect.bottom) / 2)
		if cx <= 0 or cy <= 0:
			return False
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if hwnd:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
		ctypes.windll.user32.SetCursorPos(cx, cy)
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
		return True
	except Exception:
		log.debug("_clickElement failed", exc_info=True)
		return False


def _announceElement(element):
	"""Extract text from a UIA element and speak it.

	Also attempts to extract the chat room name and store it.
	Returns the announcement text (or None if OCR fallback used).
	"""
	global _currentChatRoomName
	textParts = _extractTextFromUIAElement(element)
	if textParts:
		announcement = " ".join(textParts)
		speech.cancelSpeech()
		ui.message(announcement)
		# Try to extract chat name from the text
		_storeChatNameFromText(announcement)
		return announcement
	else:
		# Try OCR as fallback
		speech.cancelSpeech()
		_ocrAndStoreChatName(element)
		return None


def _storeChatNameFromText(text):
	"""Extract and store the chat room name from OCR/UIA text.

	The first line of the text typically contains the chat name,
	possibly followed by unread count like '( 5 )'.
	"""
	global _currentChatRoomName
	if not text:
		return
	import re

	# Take the first line
	lines = text.strip().split("\n")
	if lines:
		firstLine = lines[0].strip()
		# Remove trailing unread count like ( 123 )
		firstLine = re.sub(r"\s*\(\s*\d+\s*\)\s*$", "", firstLine)
		# Remove leading time patterns like '上午 11:08' or '下午 3:52'
		firstLine = re.sub(r"^[上下]午\s*\d+\s*[:：]\s*\d+\s*", "", firstLine)
		firstLine = firstLine.strip()
		if firstLine:
			previousRoom = _currentChatRoomName
			_currentChatRoomName = firstLine
			log.info(f"LINE: stored chat room name: {_currentChatRoomName}")
			if previousRoom != firstLine:
				try:
					from . import _chatCache

					_chatCache.onChatRoomChanged(firstLine)
				except Exception:
					log.debug("LINE: chat cache room-change hook failed", exc_info=True)


def _ocrAndStoreChatName(element):
	"""OCR an element and store the chat room name from the result."""
	global _currentChatRoomName
	try:
		rect = element.CurrentBoundingRectangle
		left = int(rect.left)
		top = int(rect.top)
		width = int(rect.right - rect.left)
		height = int(rect.bottom - rect.top)
		if width <= 0 or height <= 0 or left < -width or top < -height:
			ui.message(_("List item"))
			return

		log.debug(f"LINE OCR+name started for element at ({left},{top}) {width}x{height}")

		# Use the same OCR mechanism as _ocrReadElementText
		try:
			import screenBitmap
			from contentRecog import uwpOcr

			langs = uwpOcr.getLanguages()
			if not langs:
				ui.message(_("List item"))
				return

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
			if not ocrLang:
				ocrLang = langs[0]

			recognizer = uwpOcr.UwpOcr(language=ocrLang)
			resizeFactor = recognizer.getResizeFactor(width, height)
			# Use higher resize factor for small elements to improve OCR accuracy
			minFactor = 2
			if width < 100 or height < 100:
				minFactor = max(3, int(200 / max(min(width, height), 1)))
			if resizeFactor < minFactor:
				resizeFactor = minFactor

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

			if resizeFactor > 1:
				sb = screenBitmap.ScreenBitmap(
					width * resizeFactor,
					height * resizeFactor,
				)
				ocrPixels = sb.captureImage(left, top, width, height)
			else:
				sb = screenBitmap.ScreenBitmap(width, height)
				ocrPixels = sb.captureImage(left, top, width, height)

			# Store references to prevent garbage collection during async OCR
			_ocrAndStoreChatName._recognizer = recognizer
			_ocrAndStoreChatName._pixels = ocrPixels
			_ocrAndStoreChatName._imgInfo = imgInfo

			def _onOcrResult(result):
				import wx

				def _handleOnMain():
					global _currentChatRoomName
					try:
						if isinstance(result, Exception):
							log.debug(f"LINE OCR+name error: {result}")
							ui.message(_("List item"))
							return
						ocrText = getattr(result, "text", "") or ""
						ocrText = _removeCJKSpaces(ocrText.strip())
						if ocrText:
							log.info(f"LINE OCR+name result: {ocrText!r}")
							speech.cancelSpeech()
							ui.message(ocrText)
							_storeChatNameFromText(ocrText)
						else:
							log.debug("LINE OCR+name: no text found")
							ui.message(_("List item"))
					except Exception as e:
						log.debug(f"LINE OCR+name result handler error: {e}")
					finally:
						_ocrAndStoreChatName._recognizer = None
						_ocrAndStoreChatName._pixels = None
						_ocrAndStoreChatName._imgInfo = None

				wx.CallAfter(_handleOnMain)

			recognizer.recognize(ocrPixels, imgInfo, _onOcrResult)
		except Exception:
			log.debug("OCR fallback failed", exc_info=True)
			ui.message(_("List item"))
	except Exception:
		log.debug("_ocrAndStoreChatName failed", exc_info=True)
		ui.message(_("List item"))


def _isInChatListContext(handler):
	"""Check if the current UIA focus is near a chat list.

	Returns (True, listElement, items, currentIndex) if we're in a
	chat list context, or (False, None, None, -1) otherwise.
	Also caches the search field element for later use.
	"""
	global _chatListCurrentIndex, _chatListSearchField
	try:
		rawElement = handler.clientObject.GetFocusedElement()
		if rawElement is None:
			return False, None, None, -1

		# Check if focus is directly on a list item
		try:
			ct = rawElement.CurrentControlType
		except Exception:
			ct = 0

		if ct == 50007:  # ListItem - focus is on a list item already
			listEl, walker = _findListElement(handler, rawElement)
			if listEl:
				# In standalone ChatWindow mode there is no sidebar —
				# all lists are message lists, not chat lists.
				if _isChatWindowMode():
					return False, None, None, -1
				# Check if this list is in the left sidebar (chat list area)
				# vs the right side (message list area)
				try:
					hwnd = ctypes.windll.user32.GetForegroundWindow()
					if hwnd:
						wndRect = ctypes.wintypes.RECT()
						ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(wndRect))
						sidebarRight = _getSidebarRightBoundary(wndRect, hwnd)
						if sidebarRight is not None:
							listRect = listEl.CurrentBoundingRectangle
							if listRect.left >= sidebarRight:
								# List is on the right side — message list, not chat list
								return False, None, None, -1
				except Exception:
					log.debug("Sidebar position check failed", exc_info=True)
				items = _getListItems(handler, listEl)
				if items and items.Length > 0:
					idx = _findCurrentItemIndex(items)
					if idx < 0:
						idx = _chatListCurrentIndex
					return True, listEl, items, idx

		# Check if focus is on the search edit field (near a list)
		if ct == 50004:  # Edit control
			label = _detectEditFieldLabel(rawElement, handler, allowNotesOcr=False)
			# If focus is on message input, we're NOT in chat list context
			# Translators: Label for the LINE message input field
			if label == _("Message input"):
				return False, None, None, -1
			# Translators: Label for the search chat rooms field in LINE
			if label == _("Search chat rooms"):
				# Cache the search field for later
				_chatListSearchField = rawElement
				listEl, walker = _findListElement(handler, rawElement)
				if listEl:
					items = _getListItems(handler, listEl)
					if items and items.Length > 0:
						idx = _findCurrentItemIndex(items)
						if idx < 0:
							idx = _chatListCurrentIndex
						return True, listEl, items, idx
	except Exception:
		log.debug("_isInChatListContext failed", exc_info=True)
	return False, None, None, -1


def _findChatListFromCache(handler):
	"""Find the chat list using the cached search field element.

	Used when _chatListMode is True but focus has moved elsewhere.
	Returns (listElement, items) or (None, None).
	"""
	global _chatListSearchField
	if _chatListSearchField is None:
		return None, None
	try:
		listEl, walker = _findListElement(handler, _chatListSearchField)
		if listEl:
			items = _getListItems(handler, listEl)
			if items and items.Length > 0:
				return listEl, items
	except Exception:
		log.debug("_findChatListFromCache failed", exc_info=True)
		_chatListSearchField = None
	return None, None


def _findChatListFromWindow(handler):
	"""Find the chat list by walking the UIA tree from the LINE window root.

	Fallback when _findChatListFromCache fails (stale COM ref).
	Finds List elements and picks the one in the left sidebar area.
	In standalone ChatWindow mode there is no sidebar, so returns (None, None).
	Returns (listElement, items) or (None, None).
	"""
	global _chatListSearchField
	try:
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			return None, None

		# Standalone ChatWindow has no sidebar — no chat list to find
		if _isChatWindowMode(hwnd):
			return None, None

		# Get window rect to identify sidebar area
		wndRect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(wndRect))
		sidebarRight = _getSidebarRightBoundary(wndRect, hwnd)
		if sidebarRight is None:
			return None, None

		# Get root UIA element for the window
		rootEl = handler.clientObject.ElementFromHandle(hwnd)
		if not rootEl:
			return None, None

		# Find all List elements
		listCondition = handler.clientObject.CreatePropertyCondition(
			30003,
			50008,  # ControlType == List
		)
		lists = rootEl.FindAll(UIAHandler.TreeScope_Descendants, listCondition)
		if not lists:
			return None, None

		# Pick the list in the sidebar (left side of window)
		for i in range(lists.Length):
			try:
				listEl = lists.GetElement(i)
				rect = listEl.CurrentBoundingRectangle
				# List must be in the left sidebar area
				if rect.left < sidebarRight and rect.right <= sidebarRight + 50:
					items = _getListItems(handler, listEl)
					if items and items.Length > 0:
						log.info(f"LINE: found chat list from window root, {items.Length} items")
						# Try to cache the search field for next time
						_tryCacheSearchField(handler, listEl)
						return listEl, items
			except Exception:
				continue
	except Exception:
		log.debug("_findChatListFromWindow failed", exc_info=True)
	return None, None


def _tryCacheSearchField(handler, listElement):
	"""Try to find and cache the search field near a list element."""
	global _chatListSearchField
	try:
		walker = handler.clientObject.RawViewWalker
		parent = walker.GetParentElement(listElement)
		if not parent:
			return
		# Look for an Edit control sibling
		editCondition = handler.clientObject.CreatePropertyCondition(
			30003,
			50004,  # ControlType == Edit
		)
		edits = parent.FindAll(UIAHandler.TreeScope_Children, editCondition)
		if edits and edits.Length > 0:
			_chatListSearchField = edits.GetElement(0)
			log.debug("LINE: cached search field from list parent")
	except Exception:
		pass


def _detectEditFieldLabel(element, handler, allowNotesOcr=True):
	"""Detect the type of a raw UIA edit element and return an appropriate label."""
	try:
		walker = handler.clientObject.RawViewWalker
		parentEl = walker.GetParentElement(element)
		if not parentEl:
			return ""

		editCondition = handler.clientObject.CreatePropertyCondition(
			30003,
			50004,
		)
		editElements = parentEl.FindAll(
			UIAHandler.TreeScope_Children,
			editCondition,
		)
		editCount = editElements.Length if editElements else 0

		if editCount >= 2:
			try:
				myRect = element.CurrentBoundingRectangle
				myTop = myRect.top
			except Exception:
				return _("Login field")

			editTops = []
			for i in range(editCount):
				try:
					siblingRect = editElements.GetElement(i).CurrentBoundingRectangle
					editTops.append(siblingRect.top)
				except Exception:
					continue

			if len(editTops) >= 2:
				if myTop <= min(editTops):
					return _("Email")
				return _("Password")
			return _("Login field")

		placeholder = _getEditPlaceholder(element)
		searchKeywords = ("搜尋", "search", "検索", "ค้นหา", "찾기")
		messageKeywords = ("輸入訊息", "message", "メッセージ", "ข้อความ", "입력")
		hasSearchHint = bool(
			placeholder and any(kw in placeholder for kw in searchKeywords),
		)
		hasMessageHint = bool(
			placeholder and any(kw in placeholder for kw in messageKeywords),
		)

		windowTitle = ""
		isNotesWindow = False
		positionSuggestsSearch = False
		positionSuggestsMessage = False
		try:
			myRect = element.CurrentBoundingRectangle
			myLeft = myRect.left
			myTop = myRect.top
			myBottom = myRect.bottom

			hwnd = ctypes.windll.user32.GetForegroundWindow()
			wndRect = ctypes.wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(wndRect))
			wndWidth = wndRect.right - wndRect.left
			wndHeight = wndRect.bottom - wndRect.top

			if wndWidth > 0 and wndHeight > 0:
				relativeX = (myLeft - wndRect.left) / wndWidth
				relativeY = (myTop - wndRect.top) / wndHeight
				relativeBottom = (myBottom - wndRect.top) / wndHeight
				positionSuggestsSearch = relativeX < 0.5 and relativeY < 0.3
				positionSuggestsMessage = relativeX >= 0.5 or relativeBottom > 0.7
		except Exception:
			log.debug("_detectEditFieldLabel position detection failed", exc_info=True)

		if hasMessageHint or (positionSuggestsMessage and not hasSearchHint):
			return _("Message input")

		needsSearchLabel = hasSearchHint or positionSuggestsSearch
		if needsSearchLabel:
			hasQueryLikeText = bool(
				placeholder and not hasSearchHint and not hasMessageHint,
			)
			needsNotesOcr = allowNotesOcr and not hasQueryLikeText
			isNotesWindow, windowTitle = _isNotesWindowContext(
				element,
				walker,
				allowOcr=needsNotesOcr,
			)
			log.debug(
				f"LINE edit field classification: placeholder={placeholder!r}, "
				f"windowTitle={windowTitle!r}, isNotesWindow={isNotesWindow}, "
				f"searchHint={hasSearchHint}, messageHint={hasMessageHint}, "
				f"queryLikeText={hasQueryLikeText}, "
				f"notesOcrEnabled={needsNotesOcr}, "
				f"positionSuggestsSearch={positionSuggestsSearch}, "
				f"positionSuggestsMessage={positionSuggestsMessage}",
			)
			if isNotesWindow:
				return _("Search notes content")
			return _("Search chat rooms")

		if positionSuggestsMessage:
			return _("Message input")

		return ""
	except Exception:
		log.debug("_detectEditFieldLabel failed", exc_info=True)
		return ""


def _queryAndSpeakUIAFocus():
	"""Query UIA for the currently focused element and speak it.

	Called after passing a navigation gesture through to LINE,
	because LINE's Qt6 does NOT fire UIA focus change events.
	We poll the UIA focused element directly and extract text
	using only safe, read-only COM property access.

	NOTE: We do NOT create NVDA UIA objects or call
	NormalizeElementBuildCache — those cause cross-process COM
	calls that crash LINE's Qt6 process.
	"""
	if _suppressAddon:
		return
	global _lastAnnouncedUIAElement, _lastAnnouncedUIAName, _lastRawFocusedElement
	try:
		handler = UIAHandler.handler
		if handler is None:
			return
		rawElement = handler.clientObject.GetFocusedElement()
		if rawElement is None:
			return

		# Build a unique identifier to avoid re-announcing
		try:
			runtimeId = rawElement.GetRuntimeId()
			elementId = str(runtimeId) if runtimeId else None
		except Exception:
			elementId = None

		targetElement = rawElement

		# Detect if UIA focus is stuck on the same element (e.g. edit field)
		# by comparing with _lastRawFocusedElement (not _lastAnnouncedUIAElement,
		# which may have been updated to a selected list item's ID).
		rawFocusStuck = elementId and elementId == _lastRawFocusedElement
		_lastRawFocusedElement = elementId

		if rawFocusStuck:
			# UIA focus hasn't moved - the edit field still has focus.
			# Try to find the selected item in a nearby list.
			selectedItem = _findSelectedItemInList(handler, rawElement)
			if selectedItem:
				targetElement = selectedItem
				try:
					runtimeId = selectedItem.GetRuntimeId()
					elementId = str(runtimeId) if runtimeId else None
				except Exception:
					elementId = None
				# Check if this is the same item we already announced
				if elementId and elementId == _lastAnnouncedUIAElement:
					return
			else:
				# Could not find selected item via UIA properties.
				# Try OCR on the list area as a fallback.
				listEl = getattr(_findSelectedItemInList, "_lastListElement", None)
				if listEl:
					log.info("LINE: selected item not found via UIA, trying OCR on list")
					_findSelectedItemInList._lastListElement = None
					_lastAnnouncedUIAElement = None
					_ocrReadElementText(listEl)
					return
				return
		else:
			# UIA focus has moved to a new element.
			# Check if this is the same item we already announced
			if elementId and elementId == _lastAnnouncedUIAElement:
				return

		_lastAnnouncedUIAElement = elementId

		# Get control type for role name
		try:
			ct = targetElement.CurrentControlType
		except Exception:
			ct = 0

		# For ListItem elements, check if it's in the message area
		# (right side of window) vs chat list sidebar (left side).
		# Only use copy-first for message list items.
		if ct == 50007:  # ListItem
			isMessageItem = False
			# In standalone ChatWindow mode, all list items are message items
			if _isChatWindowMode():
				isMessageItem = True
			else:
				try:
					elRect = targetElement.CurrentBoundingRectangle
					elLeft = int(elRect.left)
					# Get the LINE window rect
					lineHwnd = ctypes.windll.user32.GetForegroundWindow()
					import ctypes.wintypes as _wt

					wr = _wt.RECT()
					ctypes.windll.user32.GetWindowRect(
						lineHwnd,
						ctypes.byref(wr),
					)
					# Message list items sit to the right of LINE's sidebar
					# (icons + chat list). The boundary uses a fixed pixel cap
					# so the test still works when the window is maximized.
					# >= is intentional: an element whose left edge lands exactly
					# on the boundary belongs to the message area, not the sidebar.
					boundary = _getSidebarRightBoundary(wr, lineHwnd, percentage=0.35)
					if boundary is not None and elLeft >= boundary:
						isMessageItem = True
				except Exception:
					pass
			if isMessageItem:
				if not _isElementVisibleInForegroundWindow(targetElement):
					log.debug(
						f"LINE UIA focus: skipping off-window message ListItem runtimeId={elementId}",
					)
					return
				log.info(
					f"LINE UIA focus: ct={ct} (message ListItem), "
					f"runtimeId={elementId}, using copy-first read",
				)
				_copyAndReadMessage(targetElement)
				return
			# Not a message item — fall through to standard handling

		# For non-ListItem elements, use standard UIA text extraction
		textParts = _extractTextFromUIAElement(targetElement)

		# For edit fields, try to detect field labels (login / search / message input)
		if ct == 50004:  # Edit control
			label = _detectEditFieldLabel(targetElement, handler)
			if label:
				# Filter out generic app name from text parts
				textParts = [t for t in textParts if t.strip() not in ("LINE", "line")]
				textParts.insert(0, label)

		controlTypeNames = {
			50000: "按鈕",
			50004: "編輯",
			50005: "超連結",
			50007: "清單項目",
			50008: "清單",
			50011: "項目",
			50016: "索引標籤項目",
			50018: "文字",
			50025: "群組",
			50033: "窗格",
		}
		roleName = controlTypeNames.get(ct, "")

		log.info(
			f"LINE UIA focus: ct={ct}, texts={textParts}, runtimeId={elementId}",
		)

		if textParts:
			announcement = " ".join(textParts)
			if roleName:
				announcement = f"{announcement} {roleName}"
			_lastAnnouncedUIAName = announcement
			speech.cancelSpeech()
			ui.message(announcement)
		elif roleName:
			_lastAnnouncedUIAName = roleName
			speech.cancelSpeech()
			ui.message(roleName)
			# For non-message ListItems (chat sidebar), fall back to OCR
			# when UIA text is empty
			if ct == 50007:  # ListItem
				global _lastOCRElement
				try:
					rid = str(targetElement.GetRuntimeId())
				except Exception:
					rid = None
				if not rid or rid != _lastOCRElement:
					_lastOCRElement = rid
					_ocrReadElementText(targetElement)
		else:
			return

	except Exception:
		log.debugWarning("_queryAndSpeakUIAFocus failed", exc_info=True)


def _copyAndReadMessage(targetElement):
	"""Read a message by right-clicking → Copy → reading clipboard content.

	This is the preferred method for reading messages because it returns
	the exact text that LINE stores, unlike OCR which may be inaccurate.

	Falls back to OCR (with ocr.wav sound alert) if copy fails.
	Restores the original clipboard content after reading.
	"""
	global _copyReadRequestId, _copyReadClipboardOwnerId
	try:
		rect = targetElement.CurrentBoundingRectangle
		cx = int((rect.left + rect.right) / 2)
		cy = int((rect.top + rect.bottom) / 2)
		if cx <= 0 or cy <= 0:
			log.debug("LINE: _copyAndReadMessage: invalid element position")
			_ocrReadMessageFallback(targetElement)
			return
	except Exception as e:
		log.debug(f"LINE: _copyAndReadMessage: failed to get rect: {e}")
		_ocrReadMessageFallback(targetElement)
		return

	hwnd = ctypes.windll.user32.GetForegroundWindow()
	_copyReadRequestId += 1
	requestId = _copyReadRequestId
	targetRuntimeId = _getElementRuntimeId(targetElement)

	# Save current clipboard content
	origClip = None
	try:
		origClip = api.getClipData()
	except Exception:
		pass

	# Clear clipboard so we can detect if copy succeeded
	try:
		_copyReadClipboardOwnerId = requestId
		api.copyToClip("")
	except Exception:
		pass

	# Get LINE window rect to clamp click positions
	try:
		import ctypes.wintypes as wintypes

		winRect = wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(
			hwnd,
			ctypes.byref(winRect),
		)
		winTop = int(winRect.top)
		winBottom = int(winRect.bottom)
	except Exception:
		winTop = 0
		winBottom = 9999

	elLeft = int(rect.left)
	elRight = int(rect.right)
	elTop = int(rect.top)
	elBottom = int(rect.bottom)
	baseClickPositions = _buildMessageBubbleClickPositions(
		(elLeft, elTop, elRight, elBottom),
		winTop,
		winBottom,
		includeVerticalOffsets=True,
	)
	clickState = {"positions": baseClickPositions}
	selectAllCount = [0]  # mutable counter for ≤2-item menus
	keyboardFallbackTried = [False]
	messageOcrCache = {
		"done": False,
		"text": "",
		"lines": [],
		"isDateSeparator": False,
		"callAnnouncement": None,
	}

	def _isCurrentRequest():
		if requestId != _copyReadRequestId:
			return False
		try:
			if ctypes.windll.user32.GetForegroundWindow() != hwnd:
				return False
		except Exception:
			pass
		if targetRuntimeId is None:
			return True
		currentRuntimeId = _getFocusedElementRuntimeId()
		return currentRuntimeId is None or currentRuntimeId == targetRuntimeId

	def _abortIfStale(stage, restoreClipboard=False, dismissMenu=False):
		if _isCurrentRequest():
			return False
		sameRequest = requestId == _copyReadRequestId
		shouldRestoreFocus = sameRequest and _shouldDismissCopyReadMenu(hwnd)
		log.debug(
			f"LINE: abandoning stale copy-read during {stage}; "
			f"requestId={requestId}, currentRequestId={_copyReadRequestId}, "
			f"targetRuntimeId={targetRuntimeId}, "
			f"currentRuntimeId={_getFocusedElementRuntimeId()}",
		)
		if dismissMenu:
			_dismissMenu()
		if restoreClipboard:
			_restoreClipboard(origClip)
		if shouldRestoreFocus:
			_restoreFocusToElement(targetElement, targetRuntimeId)
		return True

	def _getPositionLabel(posIdx):
		try:
			if posIdx < 0:
				return "keyboard-fallback"
			return clickState["positions"][posIdx][2]
		except Exception:
			return "keyboard-fallback"

	def _getMessageOcrText():
		"""OCR the message bubble to detect file/voice message hints."""
		if _abortIfStale("message OCR", restoreClipboard=True):
			return ""
		if messageOcrCache["done"]:
			return messageOcrCache["text"]
		messageOcrCache["done"] = True
		try:
			msgRect = targetElement.CurrentBoundingRectangle
			msgW = int(msgRect.right - msgRect.left)
			msgH = int(msgRect.bottom - msgRect.top)
			if msgW <= 0 or msgH <= 0:
				messageOcrCache["text"] = ""
				return ""
			import screenBitmap
			from contentRecog import uwpOcr
			import threading

			mLeft = int(msgRect.left)
			mTop = int(msgRect.top)
			sb = screenBitmap.ScreenBitmap(msgW, msgH)
			pixels = sb.captureImage(mLeft, mTop, msgW, msgH)

			langs = uwpOcr.getLanguages()
			ocrLang = None
			for cand in ["zh-Hant-TW", "zh-TW", "zh-Hant"]:
				if cand in langs:
					ocrLang = cand
					break
			if not ocrLang:
				for lang in langs:
					if lang.startswith("zh"):
						ocrLang = lang
						break
			if not ocrLang and langs:
				ocrLang = langs[0]
			if not ocrLang:
				messageOcrCache["text"] = ""
				return ""

			recognizer = uwpOcr.UwpOcr(language=ocrLang)
			resizeFactor = recognizer.getResizeFactor(msgW, msgH)
			if resizeFactor <= 0:
				resizeFactor = 1
			if resizeFactor < 2:
				resizeFactor = 2

			if resizeFactor > 1:
				sb2 = screenBitmap.ScreenBitmap(
					msgW * resizeFactor,
					msgH * resizeFactor,
				)
				ocrPixels = sb2.captureImage(mLeft, mTop, msgW, msgH)
			else:
				ocrPixels = pixels

			class _ImgInfo:
				def __init__(self, w, h, factor, sL, sT):
					self.recogWidth = w * factor
					self.recogHeight = h * factor
					self.resizeFactor = factor
					self._screenLeft = sL
					self._screenTop = sT

				def convertXToScreen(self, x):
					return self._screenLeft + int(x / self.resizeFactor)

				def convertYToScreen(self, y):
					return self._screenTop + int(y / self.resizeFactor)

				def convertWidthToScreen(self, w):
					return int(w / self.resizeFactor)

				def convertHeightToScreen(self, h):
					return int(h / self.resizeFactor)

			imgInfo = _ImgInfo(msgW, msgH, resizeFactor, mLeft, mTop)

			resultHolder = [None]
			event = threading.Event()

			def _onOcr(result):
				resultHolder[0] = result
				event.set()

			recognizer.recognize(ocrPixels, imgInfo, _onOcr)
			event.wait(timeout=2.5)

			msgText = ""
			ocrLines = []
			result = resultHolder[0]
			if result and not isinstance(result, Exception):
				msgText = getattr(result, "text", "") or ""
				msgText = _removeCJKSpaces(msgText.strip())
				ocrLines = _extractOcrLines(result)

			messageOcrCache["text"] = msgText
			messageOcrCache["lines"] = ocrLines
			messageOcrCache["isDateSeparator"] = _isCenteredLineDateSeparatorOcr(
				msgText,
				ocrLines,
				(elLeft, elTop, elRight, elBottom),
			)
			messageOcrCache["callAnnouncement"] = _getCallAnnouncementFromOcr(
				msgText,
			)
			if messageOcrCache["isDateSeparator"]:
				spokenDateText = _getSpokenLineDateSeparatorText(msgText)
				log.info(
					f"LINE: copy-read detected centered date separator via OCR: "
					f"{msgText!r} -> {spokenDateText!r}",
				)
				return spokenDateText
			ocrClickPositions = _buildMessageBubbleOcrClickPositions(
				ocrLines,
				(elLeft, elTop, elRight, elBottom),
				winTop,
				winBottom,
			)
			if ocrClickPositions:
				log.debug(
					f"LINE: copy-read OCR-derived bubble probes: "
					f"{[label for _x, _y, label in ocrClickPositions]}",
				)
				clickState["positions"] = _mergeClickPositions(
					ocrClickPositions,
					baseClickPositions,
				)
			elif ocrLines:
				log.debug(
					f"LINE: copy-read OCR lines yielded no bubble probes: "
					f"{[(line.get('text', ''), bool(line.get('rect'))) for line in ocrLines]}",
				)
			log.debug(f"LINE: copy-read message OCR: {msgText!r}")
			return msgText
		except Exception as e:
			log.debug(
				f"LINE: copy-read message OCR failed: {e}",
				exc_info=True,
			)
			messageOcrCache["text"] = ""
			return ""

	def _startMouseCopyProbes(reason, dismissMenu=False):
		if _abortIfStale(
			f"mouseCopyProbeFallback:{reason}",
			restoreClipboard=True,
			dismissMenu=dismissMenu,
		):
			return
		if dismissMenu:
			_dismissMenu()
		log.info(
			f"LINE: copy-read keyboard menu failed ({reason}); trying mouse probes",
		)
		core.callLater(300, lambda: _attemptCopyAtOffset(0))

	def _tryKeyboardCopyFallback():
		if keyboardFallbackTried[0]:
			log.info("LINE: copy-read keyboard fallback exhausted, falling back to OCR")
			_restoreClipboard(origClip)
			_ocrReadMessageFallback(targetElement)
			return
		if _abortIfStale(
			"keyboardCopyFallback",
			restoreClipboard=True,
			dismissMenu=True,
		):
			return
		keyboardFallbackTried[0] = True
		log.info(
			"LINE: copy-read trying keyboard context menu before mouse probes",
		)
		try:
			_sendGestureWithAddonSuppressed("applications")
		except Exception as e:
			log.debug(
				f"LINE: copy-read keyboard fallback applications send failed: {e}",
				exc_info=True,
			)
			_startMouseCopyProbes("applications send failed")
			return

		core.callLater(
			500,
			lambda: _findCopyMenuItem(-1, retriesLeft=4),
		)

	def _attemptCopyAtOffset(posIdx=0):
		"""Right-click at clickPositions[posIdx] and try to copy."""
		if _abortIfStale("attemptCopyAtOffset", restoreClipboard=True):
			return
		positions = clickState["positions"]
		if posIdx >= len(positions):
			# All mouse positions exhausted; keyboard menu can still target the row.
			_tryKeyboardCopyFallback()
			return

		clickX, clickY, posLabel = positions[posIdx]
		log.info(
			f"LINE: copy-read right-clicking at ({clickX}, {clickY}) [{posLabel}]",
		)
		if _abortIfStale("right click", restoreClipboard=True):
			return

		if not _messageProbePointHitsTargetElement(
			UIAHandler.handler,
			targetElement,
			clickX,
			clickY,
			targetRuntimeId,
		):
			log.info(
				f"LINE: copy-read skipping probe over edit field at ({clickX}, {clickY}) [{posLabel}]",
			)
			core.callLater(50, lambda: _attemptCopyAtOffset(posIdx + 1))
			return

		# Perform right-click
		if hwnd:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
		ctypes.windll.user32.SetCursorPos(int(clickX), int(clickY))
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0008, 0, 0, 0, 0)  # RIGHTDOWN
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0010, 0, 0, 0, 0)  # RIGHTUP

		# Wait for context menu, then find Copy
		core.callLater(300, lambda: _findCopyMenuItem(posIdx, retriesLeft=4))

	def _findCopyMenuItem(posIdx, retriesLeft=4):
		"""Find the context menu and click Copy."""
		if _abortIfStale("findCopyMenuItem", restoreClipboard=True, dismissMenu=True):
			return
		try:
			uiaHandler = UIAHandler.handler
			if not uiaHandler:
				_restoreClipboard(origClip)
				_ocrReadMessageFallback(targetElement)
				return

			import ctypes.wintypes as wintypes

			# Find popup window
			pid = wintypes.DWORD()
			tid = ctypes.windll.user32.GetWindowThreadProcessId(
				hwnd,
				ctypes.byref(pid),
			)
			popupCandidates = []

			WNDENUMPROC = ctypes.WINFUNCTYPE(
				ctypes.c_bool,
				wintypes.HWND,
				wintypes.LPARAM,
			)

			def _enumCb(enumHwnd, lParam):
				if enumHwnd != hwnd and ctypes.windll.user32.IsWindowVisible(enumHwnd):
					wRect = wintypes.RECT()
					ctypes.windll.user32.GetWindowRect(
						enumHwnd,
						ctypes.byref(wRect),
					)
					w = wRect.right - wRect.left
					h = wRect.bottom - wRect.top
					if w >= 50 and h >= 30:
						popupCandidates.append(enumHwnd)
				return True

			ctypes.windll.user32.EnumThreadWindows(
				tid,
				WNDENUMPROC(_enumCb),
				0,
			)

			popupHwnd = None
			if popupCandidates:
				popupHwnd = popupCandidates[0]
			else:
				if 0 <= posIdx < len(clickState["positions"]):
					clickX, clickY, _positionLabel = clickState["positions"][posIdx]
					for dy in [0, -40, -80, 40, 80]:
						pt = wintypes.POINT(clickX, clickY + dy)
						candHwnd = ctypes.windll.user32.WindowFromPoint(pt)
						if candHwnd and candHwnd != hwnd:
							popupHwnd = candHwnd
							break

			if not popupHwnd:
				if retriesLeft > 0:
					core.callLater(
						200,
						lambda: _findCopyMenuItem(posIdx, retriesLeft - 1),
					)
					return
				# No popup found, try next position
				if posIdx < 0:
					_startMouseCopyProbes("keyboard popup not found")
					return
				core.callLater(300, lambda: _attemptCopyAtOffset(posIdx + 1))
				return

			element = uiaHandler.clientObject.ElementFromHandle(popupHwnd)
			if not element:
				if retriesLeft > 0:
					core.callLater(
						200,
						lambda: _findCopyMenuItem(posIdx, retriesLeft - 1),
					)
					return
				if posIdx < 0:
					_startMouseCopyProbes("keyboard popup element unavailable", dismissMenu=True)
					return
				core.callLater(300, lambda: _attemptCopyAtOffset(posIdx + 1))
				return

			# Validate popup is a real context menu
			try:
				pCt = element.CurrentControlType
				eRect = element.CurrentBoundingRectangle
				eW = int(eRect.right - eRect.left)
				eH = int(eRect.bottom - eRect.top)
				if pCt == 50033 or eW < 50 or eH < 30:
					if posIdx < 0:
						_startMouseCopyProbes("keyboard popup was not a context menu", dismissMenu=True)
						return
					core.callLater(300, lambda: _attemptCopyAtOffset(posIdx + 1))
					return
			except Exception:
				pass

			walker = uiaHandler.clientObject.RawViewWalker

			# Collect menu items
			menuItems = []

			def _collectItems(parent, depth=0):
				child = walker.GetFirstChildElement(parent)
				idx = 0
				while child and idx < 30:
					try:
						cRect = child.CurrentBoundingRectangle
						cH = int(cRect.bottom - cRect.top)
						cW = int(cRect.right - cRect.left)
						if cW > 0 and cH > 0:
							if 20 <= cH <= 80 and cW >= cH * 2:
								# Get text from children
								itemText = ""
								tc = walker.GetFirstChildElement(child)
								ci = 0
								while tc and ci < 10:
									try:
										n = tc.CurrentName
										if n and n.strip():
											itemText = n.strip()
											break
									except Exception:
										pass
									try:
										tc = walker.GetNextSiblingElement(tc)
									except Exception:
										break
									ci += 1
								menuItems.append((child, itemText))
							elif cH > 80 and depth < 5:
								_collectItems(child, depth + 1)
							elif cH >= 20:
								menuItems.append((child, ""))
					except Exception:
						pass
					try:
						child = walker.GetNextSiblingElement(child)
					except Exception:
						break
					idx += 1

			_collectItems(element)
			log.debug(
				f"LINE: copy-read found {len(menuItems)} menu items: {[t for _, t in menuItems]}",
			)

			if not menuItems:
				if retriesLeft > 0:
					core.callLater(
						200,
						lambda: _findCopyMenuItem(posIdx, retriesLeft - 1),
					)
					return
				if posIdx < 0:
					_startMouseCopyProbes("keyboard popup had no menu items", dismissMenu=True)
					return

			# Find "複製" item by UIA text
			copyItem = None
			copyClickPoint = None
			for item, text in menuItems:
				if text and "複製" in text:
					copyItem = item
					log.info(f"LINE: copy-read matched '複製' by UIA: {text!r}")
					break

			# Strategy 2: OCR the popup when UIA text is all empty
			popupOcrText = ""
			popupOcrMatchedLabels = []
			if not copyItem and menuItems and len(menuItems) >= 3:
				log.debug("LINE: copy-read no UIA text, trying popup OCR")
				try:
					popupRect = element.CurrentBoundingRectangle
					popupW = int(popupRect.right - popupRect.left)
					popupH = int(popupRect.bottom - popupRect.top)
					if popupW > 0 and popupH > 0:
						popupRectTuple = (
							int(popupRect.left),
							int(popupRect.top),
							int(popupRect.right),
							int(popupRect.bottom),
						)
						popupRowRects = _collectPopupMenuRowRects(
							popupHwnd,
							popupRectTuple,
						)
						import screenBitmap
						from contentRecog import uwpOcr
						import threading

						pLeft = int(popupRect.left)
						pTop = int(popupRect.top)
						sb = screenBitmap.ScreenBitmap(popupW, popupH)
						pixels = sb.captureImage(pLeft, pTop, popupW, popupH)

						langs = uwpOcr.getLanguages()
						ocrLang = None
						for cand in ["zh-Hant-TW", "zh-TW", "zh-Hant"]:
							if cand in langs:
								ocrLang = cand
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
							resizeFactor = recognizer.getResizeFactor(popupW, popupH)

							if resizeFactor > 1:
								sb2 = screenBitmap.ScreenBitmap(
									popupW * resizeFactor,
									popupH * resizeFactor,
								)
								ocrPixels = sb2.captureImage(
									pLeft,
									pTop,
									popupW,
									popupH,
								)
							else:
								ocrPixels = pixels

							class _ImgInfo:
								def __init__(self, w, h, factor, sL, sT):
									self.recogWidth = w * factor
									self.recogHeight = h * factor
									self.resizeFactor = factor
									self._screenLeft = sL
									self._screenTop = sT

								def convertXToScreen(self, x):
									return self._screenLeft + int(x / self.resizeFactor)

								def convertYToScreen(self, y):
									return self._screenTop + int(y / self.resizeFactor)

								def convertWidthToScreen(self, w):
									return int(w / self.resizeFactor)

								def convertHeightToScreen(self, h):
									return int(h / self.resizeFactor)

							imgInfo = _ImgInfo(popupW, popupH, resizeFactor, pLeft, pTop)

							resultHolder = [None]
							event = threading.Event()

							def _onOcr(result):
								resultHolder[0] = result
								event.set()

							recognizer.recognize(ocrPixels, imgInfo, _onOcr)
							event.wait(timeout=3.0)

							ocrText = ""
							popupOcrLines = []
							result = resultHolder[0]
							if result and not isinstance(result, Exception):
								ocrText = getattr(result, "text", "") or ""
								ocrText = _removeCJKSpaces(ocrText.strip())
								popupOcrLines = _extractOcrLines(result)

							popupOcrText = ocrText
							popupLines, popupLineMatches, popupOcrMatchedLabels = (
								_extractMatchedMessageContextMenuLabels(ocrText)
							)
							log.debug(f"LINE: copy-read popup OCR: {ocrText!r}")
							copyTarget = _resolvePopupMenuLabelClickPoint(
								"複製",
								popupOcrLines,
								popupRectTuple,
								popupRowRects=popupRowRects,
							)
							if copyTarget:
								copyClickPoint = copyTarget["clickPoint"]
								log.info(
									f"LINE: copy-read matched '複製' via popup OCR rows, "
									f"item {copyTarget['index'] + 1}/{copyTarget['count']}",
								)
							elif "複製" in popupOcrMatchedLabels:
								targetMatchIdx = popupOcrMatchedLabels.index("複製")
								nItems = len(menuItems)
								itemIdx = min(targetMatchIdx, nItems - 1)
								copyItem = menuItems[itemIdx][0]
								log.info(
									f"LINE: copy-read matched '複製' via popup OCR order, "
									f"matched item {itemIdx + 1}/{nItems}",
								)
							elif ocrText:
								log.debug(
									f"LINE: copy-read popup OCR did not resemble a message "
									f"context menu: {popupLines}",
								)
				except Exception as e:
					log.debug(
						f"LINE: copy-read popup OCR failed: {e}",
						exc_info=True,
					)

			# Detect file / voice messages when Copy is unavailable
			if not copyItem and not copyClickPoint and menuItems and len(menuItems) >= 3:
				popupOcrLooksLikeMenu = bool(popupOcrMatchedLabels)
				menuTextBlob = "\n".join(text for _item, text in menuItems if text)
				if popupOcrLooksLikeMenu and popupOcrText:
					menuTextBlob = "\n".join(part for part in (menuTextBlob, popupOcrText) if part)

				def _menuHasText(keyword):
					for _item, text in menuItems:
						if text and keyword in text:
							return True
					if popupOcrLooksLikeMenu and popupOcrText and keyword in popupOcrText:
						return True
					return False

				if _looksLikeStickerMessageMenu(menuTextBlob):
					log.info("LINE: copy-read detected sticker message")
					ui.message(_("貼圖，請按 NVDA+Windows+I 描述"))
					_dismissMenu()
					_restoreClipboard(origClip)
					return

				menuHasSaveAs = _menuHasText("另存新檔")
				menuHasSave = _menuHasText("儲存")
				if menuHasSaveAs and menuHasSave:
					if _looksLikeImageAttachmentMenu(menuTextBlob):
						log.info("LINE: copy-read detected image message")
						ui.message(_("圖片，請按 NVDA+Windows+K 下載。"))
						_dismissMenu()
						_restoreClipboard(origClip)
						return
					msgOcrText = _getMessageOcrText()
					if msgOcrText:
						msgHasSaveAs = "另存新檔" in msgOcrText
						msgHasSave = "儲存" in msgOcrText
						if msgHasSaveAs and msgHasSave:
							downloadDeadline = _extractDownloadDeadlineAnnouncement(
								msgOcrText,
							)
							if downloadDeadline:
								log.info("LINE: copy-read detected file message")
								ui.message(
									_("檔案，{deadline}。請按 NVDA+Windows+K 下載。").format(
										deadline=downloadDeadline,
									),
								)
								_dismissMenu()
								_restoreClipboard(origClip)
								return
							voiceDuration = _extractCallDuration(msgOcrText)
							if voiceDuration:
								log.info("LINE: copy-read detected voice message")
								ui.message(
									"語音訊息，請按 NVDA+Windows+P 播放，按 NVDA+Windows+K 下載。",
								)
								_dismissMenu()
								_restoreClipboard(origClip)
								return
							log.info("LINE: copy-read detected image message from OCR")
							ui.message(_("圖片，請按 NVDA+Windows+K 下載。"))
							_dismissMenu()
							_restoreClipboard(origClip)
							return

			# If still no match, check if we got the wrong menu
			if not copyItem and not copyClickPoint:
				if posIdx < 0:
					_startMouseCopyProbes("keyboard menu lacked Copy", dismissMenu=True)
					return
				isSelectAll = len(menuItems) <= 2
				if isSelectAll and messageOcrCache["callAnnouncement"]:
					log.info(
						f"LINE: copy-read detected call record early: "
						f"{messageOcrCache['callAnnouncement']!r} "
						f"after small menu at "
						f"[{_getPositionLabel(posIdx)}]",
					)
					_dismissMenu()
					_restoreClipboard(origClip)
					speech.cancelSpeech()
					ui.message(messageOcrCache["callAnnouncement"])
					return
				if isSelectAll:
					selectAllCount[0] += 1
				if (
					isSelectAll
					and selectAllCount[0] >= 6
					and _hasExhaustedMessageBubbleFallbackProbes(
						posIdx,
						clickState["positions"],
					)
				):
					# After the extra-left probes have also failed,
					# treat repeated ≤2-item menus as a likely call
					# record or otherwise non-copyable message.
					_dismissMenu()
					msgOcrText = _getMessageOcrText()
					callAnnouncement = messageOcrCache["callAnnouncement"] or _getCallAnnouncementFromOcr(
						msgOcrText,
					)
					messageOcrCache["callAnnouncement"] = callAnnouncement
					if callAnnouncement:
						log.info(
							f"LINE: copy-read detected call "
							f"record: {callAnnouncement!r} "
							f"from OCR: {msgOcrText!r}",
						)
						_restoreClipboard(origClip)
						ui.message(callAnnouncement)
						return
					log.info(
						f"LINE: copy-read wrong menu "
						f"(≤2 items) seen "
						f"{selectAllCount[0]} times, last "
						f"at [{_getPositionLabel(posIdx)}]"
						f", skipping to OCR",
					)
					core.callLater(
						300,
						lambda: _attemptCopyAtOffset(len(clickState["positions"])),
					)
					return
				if len(menuItems) >= 3:
					# Got a correct context menu (≥3 items)
					# but 複製 isn't in it.
					# Check if this is a self-sent message:
					# right-clicking on the bubble edge (not
					# the text) gives a menu with 收回 but
					# no 複製. Try other positions first.
					ocrHasRecall = popupOcrLooksLikeMenu and popupOcrText and "收回" in popupOcrText
					if ocrHasRecall:
						log.info(
							f"LINE: copy-read self-sent "
							f"message edge hit "
							f"({len(menuItems)} items, "
							f"has 收回 but no 複製) at "
							f"[{_getPositionLabel(posIdx)}]"
							f", trying next position",
						)
						_dismissMenu()
						core.callLater(
							300,
							lambda: _attemptCopyAtOffset(
								posIdx + 1,
							),
						)
						return
					# No 收回 either — this message type
					# doesn't support copy (sticker/image).
					# Bail immediately to content OCR.
					log.info(
						f"LINE: copy-read correct menu "
						f"({len(menuItems)} items) but no "
						f"'複製' at [{_getPositionLabel(posIdx)}]"
						f", skipping to OCR",
					)
					_dismissMenu()
					core.callLater(
						300,
						lambda: _attemptCopyAtOffset(len(clickState["positions"])),
					)
					return
				# Wrong menu or OCR couldn't find 複製, try next position
				log.info(
					f"LINE: copy-read '複製' not found at [{_getPositionLabel(posIdx)}]",
				)
				_dismissMenu()
				core.callLater(300, lambda: _attemptCopyAtOffset(posIdx + 1))
				return

			# Click the Copy item
			if _abortIfStale("clickCopyItem", restoreClipboard=True, dismissMenu=True):
				return
			if copyClickPoint:
				itemCx, itemCy = copyClickPoint
			else:
				iRect = copyItem.CurrentBoundingRectangle
				itemCx = int((iRect.left + iRect.right) / 2)
				itemCy = int((iRect.top + iRect.bottom) / 2)
			log.info(f"LINE: copy-read clicking '複製' at ({itemCx}, {itemCy})")
			ctypes.windll.user32.SetCursorPos(itemCx, itemCy)
			time.sleep(0.05)
			ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
			time.sleep(0.05)
			ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP

			# Wait for clipboard to update, then read
			core.callLater(200, _readClipboardAndSpeak)

		except Exception as e:
			log.debug(f"LINE: copy-read menu detection failed: {e}", exc_info=True)
			if posIdx < 0:
				_startMouseCopyProbes("keyboard menu detection failed", dismissMenu=True)
				return
			_dismissMenu()
			_restoreClipboard(origClip)
			_ocrReadMessageFallback(targetElement)

	def _readClipboardAndSpeak():
		"""Read clipboard content and speak it."""
		if _abortIfStale("readClipboard", restoreClipboard=True):
			return
		try:
			clipText = api.getClipData()
		except Exception:
			clipText = ""

		if clipText and clipText.strip():
			log.info(f"LINE: copy-read success: {clipText!r}")
			speech.cancelSpeech()
			ui.message(clipText.strip())
			_restoreClipboard(origClip)
		else:
			log.info("LINE: copy-read clipboard empty, falling back to OCR")
			_restoreClipboard(origClip)
			_ocrReadMessageFallback(targetElement)

	def _dismissMenu():
		"""Send Escape only while LINE still owns the active popup/menu."""
		if not _shouldDismissCopyReadMenu(hwnd):
			return
		try:
			from keyboardHandler import KeyboardInputGesture

			KeyboardInputGesture.fromName("escape").send()
		except Exception:
			pass

	def _restoreClipboard(original):
		"""Restore the original clipboard content."""
		global _copyReadClipboardOwnerId
		try:
			if _copyReadClipboardOwnerId != requestId:
				return
			if original is not None:
				api.copyToClip(original)
		except Exception:
			pass
		finally:
			if _copyReadClipboardOwnerId == requestId:
				_copyReadClipboardOwnerId = 0

	# Start the copy-read process
	initialOcrText = _getMessageOcrText()
	if messageOcrCache["isDateSeparator"] and initialOcrText:
		_restoreClipboard(origClip)
		speech.cancelSpeech()
		ui.message(initialOcrText.strip())
		return
	# Background chat cache: if NVDA+Windows+U cached this room's text,
	# look the focused bubble up by OCR snippet and read it directly,
	# skipping the right-click copy-read flow entirely.
	if initialOcrText:
		try:
			from . import _chatCache

			cacheActive = _chatCache.isActive()
			log.debug(
				f"LINE: chat cache check: active={cacheActive}"
				f" room={_chatCache.getChatRoomName()!r}"
				f" entries={_chatCache.getMessageCount()}"
				f" ocr={initialOcrText!r}",
			)
			if cacheActive:
				cachedText, _cacheIdx = _chatCache.lookupMessage(initialOcrText)
				if cachedText:
					log.info(
						f"LINE: copy-read served from chat cache: {cachedText!r}",
					)
					_restoreClipboard(origClip)
					speech.cancelSpeech()
					ui.message(cachedText)
					return
				else:
					log.debug(
						f"LINE: chat cache lookup returned no match for OCR {initialOcrText!r}",
					)
		except Exception:
			log.debug("LINE: chat cache lookup failed", exc_info=True)
	_tryKeyboardCopyFallback()


def _ocrReadMessageFallback(targetElement):
	"""Fall back to OCR for reading a message, with ocr.wav sound alert.

	Plays ocr.wav to indicate the result may not be 100% accurate,
	then reads the element text via OCR.
	"""
	try:
		nvwave.playWaveFile(_OCR_SOUND_PATH, asynchronous=True)
	except Exception:
		log.debugWarning("Failed to play OCR sound", exc_info=True)
	_ocrReadElementText(targetElement, preferCallAnnouncement=True)


class LineChatListItem(UIA):
	"""Overlay class for chat/contact list items in the sidebar.

	Qt6 list items typically have empty name AND childCount=0.
	We use UIA FindAll and display model as fallbacks.
	"""

	def _get_name(self):
		# Prevent infinite recursion using global depth counter
		global _nameRecursionDepth
		if _nameRecursionDepth >= _MAX_NAME_RECURSION_DEPTH:
			return super().name or ""
		# Also guard using UIA runtime ID (stable across Python wrapper recreation)
		guardKey = None
		try:
			if hasattr(self, "UIAElement") and self.UIAElement:
				rid = self.UIAElement.GetRuntimeId()
				guardKey = ("LineChatListItem", str(rid))
		except Exception:
			guardKey = ("LineChatListItem", id(self))
		if guardKey and guardKey in _nameRecursionGuard:
			return super().name or ""
		if guardKey:
			_nameRecursionGuard.add(guardKey)
		_nameRecursionDepth += 1
		try:
			# First try the native name
			name = super().name
			if name and name.strip():
				return name
			# Try deep text (includes UIA FindAll fallback)
			# _getDeepText now uses _getObjectNameDirect to avoid re-entering _get_name
			texts = _getDeepText(self, maxDepth=4)
			if texts:
				return " - ".join(texts)
			# Last resort: read from display
			displayText = _getTextFromDisplay(self)
			if displayText:
				return displayText
			return ""
		finally:
			_nameRecursionDepth -= 1
			if guardKey:
				_nameRecursionGuard.discard(guardKey)

	def event_gainFocus(self):
		super().event_gainFocus()


class LineChatMessage(UIA):
	"""Overlay class for individual chat messages."""

	def _get_name(self):
		global _nameRecursionDepth
		if _nameRecursionDepth >= _MAX_NAME_RECURSION_DEPTH:
			return super().name or ""
		guardKey = None
		try:
			if hasattr(self, "UIAElement") and self.UIAElement:
				rid = self.UIAElement.GetRuntimeId()
				guardKey = ("LineChatMessage", str(rid))
		except Exception:
			guardKey = ("LineChatMessage", id(self))
		if guardKey and guardKey in _nameRecursionGuard:
			return super().name or ""
		if guardKey:
			_nameRecursionGuard.add(guardKey)
		_nameRecursionDepth += 1
		try:
			name = super().name
			if name and name.strip():
				return name
			texts = _getDeepText(self, maxDepth=3)
			if texts:
				return ": ".join(texts)
			displayText = _getTextFromDisplay(self)
			if displayText:
				return displayText
			return ""
		finally:
			_nameRecursionDepth -= 1
			if guardKey:
				_nameRecursionGuard.discard(guardKey)

	def _get_description(self):
		desc = super().description
		return desc or ""


class LineMessageInput(UIA):
	"""Overlay class for the message input/composition area."""

	def _get_name(self):
		try:
			name = super().name
		except Exception:
			log.debugWarning(
				"Error in LineMessageInput._get_name",
				exc_info=True,
			)
			name = ""
		if not name or not name.strip():
			# Translators: Label for the LINE message input field
			return _("Message input")
		return name


class LineSearchField(UIA):
	"""Overlay class for the search/filter field in the sidebar."""

	def _get_name(self):
		try:
			name = super().name
		except Exception:
			log.debugWarning(
				"Error in LineSearchField._get_name",
				exc_info=True,
			)
			name = ""
		if not name or not name.strip() or name.strip().lower() in ("line",):
			# Translators: Label for the search chat rooms field in LINE
			return _("Search chat rooms")
		return name


class LineLoginEditField(UIA):
	"""Overlay class for login window edit fields (email / password).

	The LINE login window has two unlabelled edit fields.
	We identify them via display-model / OCR placeholder text,
	or fall back to vertical position (upper = email, lower = password).
	"""

	def _get_name(self):
		# If UIA already provides a meaningful name (not just the app name), use it.
		try:
			name = super().name
			if name and name.strip():
				# Filter out generic names that are just the app/window name
				if name.strip() not in ("LINE", "line"):
					return name
		except Exception:
			pass

		# Strategy 1: read display / placeholder text to identify field type
		label = self._detectFieldLabel()
		if label:
			return label

		# Strategy 2: use vertical position among sibling edit fields
		label = self._detectByPosition()
		if label:
			return label

		# Translators: Generic label for a login field when position cannot be determined
		return _("Login field")

	def _detectFieldLabel(self):
		"""Read placeholder / visible text via display model to guess field type."""
		try:
			text = _getTextFromDisplay(self)
			if text:
				t = text.lower()
				if any(
					kw in t
					for kw in (
						"email",
						"mail",
						"電子郵件",
						"メール",
						"อีเมล",
					)
				):
					# Translators: Label for the email input field on the LINE login screen
					return _("Email")
				if any(
					kw in t
					for kw in (
						"password",
						"密碼",
						"パスワード",
						"รหัสผ่าน",
					)
				):
					# Translators: Label for the password input field on the LINE login screen
					return _("Password")
		except Exception:
			log.debugWarning("LineLoginEditField._detectFieldLabel error", exc_info=True)
		return ""

	def _detectByPosition(self):
		"""Use vertical position among siblings to guess email vs password."""
		try:
			if not hasattr(self, "UIAElement") or self.UIAElement is None:
				return ""
			rect = self.UIAElement.CurrentBoundingRectangle
			myTop = rect.top

			parent = self.parent
			if not parent:
				return ""

			# Gather top-coordinates of sibling edit fields
			editTops = []
			try:
				for child in parent.children:
					try:
						if child.role in (
							controlTypes.Role.EDITABLETEXT,
							controlTypes.Role.DOCUMENT,
						):
							if hasattr(child, "UIAElement") and child.UIAElement:
								cr = child.UIAElement.CurrentBoundingRectangle
								editTops.append(cr.top)
					except Exception:
						continue
			except Exception:
				pass

			if len(editTops) >= 2:
				if myTop <= min(editTops):
					# Translators: Label for the email input field on the LINE login screen
					return _("Email")
				else:
					# Translators: Label for the password input field on the LINE login screen
					return _("Password")
		except Exception:
			log.debugWarning("LineLoginEditField._detectByPosition error", exc_info=True)
		return ""


class LineContactItem(UIA):
	"""Overlay class for contact list items."""

	def _get_name(self):
		global _nameRecursionDepth
		if _nameRecursionDepth >= _MAX_NAME_RECURSION_DEPTH:
			return super().name or ""
		guardKey = None
		try:
			if hasattr(self, "UIAElement") and self.UIAElement:
				rid = self.UIAElement.GetRuntimeId()
				guardKey = ("LineContactItem", str(rid))
		except Exception:
			guardKey = ("LineContactItem", id(self))
		if guardKey and guardKey in _nameRecursionGuard:
			return super().name or ""
		if guardKey:
			_nameRecursionGuard.add(guardKey)
		_nameRecursionDepth += 1
		try:
			name = super().name
			if name and name.strip():
				return name
			texts = _getDeepText(self, maxDepth=3)
			if texts:
				return " - ".join(texts)
			displayText = _getTextFromDisplay(self)
			if displayText:
				return displayText
			return ""
		finally:
			_nameRecursionDepth -= 1
			if guardKey:
				_nameRecursionGuard.discard(guardKey)


class LineGenericList(UIA):
	"""Overlay class for list containers in LINE."""

	def _get_positionInfo(self):
		try:
			info = super().positionInfo
		except Exception:
			log.debugWarning(
				"Error in LineGenericList._get_positionInfo",
				exc_info=True,
			)
			info = {}
		return info


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
		VirtualWindow.initialize()
		# Read and cache LINE installation info
		self._lineVersion = _readLineVersion()
		self._lineLanguage = _readLineLanguage()
		self._qtAccessibleSet = _isQtAccessibleSet()
		# Decrypt the image API keys once at startup so PBKDF2 never runs
		# during an actual image-description request.
		_initEffectiveImageApiKey()
		_initEffectiveOllamaApiKey()
		log.info(
			f"LINE AppModule loaded for process: {self.processID}, "
			f"exe: {self.appName}, "
			f"lineVersion: {self._lineVersion}, "
			f"lineLanguage: {self._lineLanguage}, "
			f"qtAccessible: {self._qtAccessibleSet}",
		)

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""Apply custom overlay classes based on role and hierarchy."""
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

		try:
			automationId = obj.UIAAutomationId or ""
		except Exception:
			automationId = ""

		try:
			className = obj.UIAClassName or ""
		except Exception:
			className = ""

		# Qt6 specific class name patterns
		isQt = "qt" in className.lower() or "Qt" in className

		# --- Chat/contact list items ---
		if role == controlTypes.Role.LISTITEM:
			parent = obj.parent
			if parent and parent.role == controlTypes.Role.LIST:
				# Check if this looks like a chat list or contact list
				if any(
					keyword in automationId.lower()
					for keyword in (
						"chat",
						"room",
						"talk",
						"conversation",
						"friend",
						"contact",
						"message",
						"buddy",
					)
				):
					clsList.insert(0, LineChatListItem)
				elif any(
					keyword in (parent.name or "").lower()
					for keyword in (
						"chat",
						"聊天",
						"トーク",
						"好友",
						"友だち",
						"friend",
						"contact",
						"message",
						"訊息",
						"メッセージ",
					)
				):
					clsList.insert(0, LineChatListItem)
				else:
					# Default: treat any list item as potentially a chat item
					clsList.insert(0, LineChatListItem)
			log.debug(
				f"LINE listitem: name={obj.name!r}, automationId={automationId!r}, children={obj.childCount}",
			)

		# --- Editable text fields ---
		elif role in (
			controlTypes.Role.EDITABLETEXT,
			controlTypes.Role.DOCUMENT,
		):
			# Check if this is a login-window edit field first
			if self._isLoginEditField(obj):
				clsList.insert(0, LineLoginEditField)
			elif self._isSearchField(obj):
				clsList.insert(0, LineSearchField)
			elif automationId and any(
				kw in automationId.lower()
				for kw in (
					"input",
					"compose",
					"message",
					"send",
					"chat",
					"editor",
					"textbox",
					"edit",
				)
			):
				clsList.insert(0, LineMessageInput)
			# Qt6 text edits
			elif isQt and "edit" in className.lower():
				clsList.insert(0, LineMessageInput)
			elif controlTypes.State.FOCUSABLE in obj.states:
				clsList.insert(0, LineMessageInput)

		# --- Individual messages in chat view ---
		elif role in (
			controlTypes.Role.GROUPING,
			controlTypes.Role.SECTION,
			controlTypes.Role.PARAGRAPH,
			controlTypes.Role.STATICTEXT,
		):
			if automationId and any(
				kw in automationId.lower()
				for kw in (
					"message",
					"bubble",
					"chat_content",
					"msg",
				)
			):
				clsList.insert(0, LineChatMessage)

		# --- List containers ---
		elif role == controlTypes.Role.LIST:
			clsList.insert(0, LineGenericList)

		# --- Toolbar/sidebar buttons without labels ---
		elif role == controlTypes.Role.BUTTON:
			try:
				btnName = obj.name
				if not btnName or not btnName.strip():
					clsList.insert(0, LineToolbarButton)
			except Exception:
				clsList.insert(0, LineToolbarButton)

	def _isLoginEditField(self, obj):
		"""Detect if an edit field belongs to the login window.

		The login window typically has 2 sibling edit fields (email + password),
		while the chat window has only 1 (message input).

		Uses raw UIA FindAll instead of parent.children to avoid triggering
		chooseNVDAObjectOverlayClasses recursively (which causes RecursionError).
		"""
		try:
			if not hasattr(obj, "UIAElement") or obj.UIAElement is None:
				return False
			parent = obj.parent
			if not parent or not hasattr(parent, "UIAElement") or parent.UIAElement is None:
				return False
			# Use raw UIA to count edit-type children without creating NVDA objects
			handler = UIAHandler.handler
			if handler is None:
				return False
			# UIA ControlType for Edit = 50004
			editCondition = handler.clientObject.CreatePropertyCondition(
				30003,  # UIA_ControlTypePropertyId
				50004,  # UIA_EditControlTypeId
			)
			editElements = parent.UIAElement.FindAll(
				UIAHandler.TreeScope_Children,
				editCondition,
			)
			editCount = editElements.Length if editElements else 0
			return editCount >= 2
		except Exception:
			return False

	def _isSearchField(self, obj):
		"""Detect if an edit field is the search/filter field in the sidebar.

		Checks UIA placeholder text for search keywords and falls back
		to position-based heuristic (left sidebar, near the top).
		"""
		try:
			if not hasattr(obj, "UIAElement") or obj.UIAElement is None:
				return False
			el = obj.UIAElement

			# Check UIA Name for search keywords
			SEARCH_KEYWORDS = ("搜尋", "search", "検索", "ค้นหา", "찾기")
			try:
				name = el.CurrentName
				if name and name.strip():
					nameLower = name.strip().lower()
					if any(kw in nameLower for kw in SEARCH_KEYWORDS):
						return True
					# If name is meaningful but not search-related, it's not a search field
					if nameLower not in ("line",):
						return False
			except Exception:
				pass

			# Check UIA ValueValue (30045) for placeholder
			try:
				val = el.GetCurrentPropertyValue(30045)
				if val and isinstance(val, str) and val.strip():
					valLower = val.strip().lower()
					if any(kw in valLower for kw in SEARCH_KEYWORDS):
						return True
			except Exception:
				pass

			# Position heuristic: search field is in the left sidebar and near the top
			try:
				rect = el.CurrentBoundingRectangle
				hwnd = ctypes.windll.user32.GetForegroundWindow()
				if hwnd:
					wndRect = ctypes.wintypes.RECT()
					ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(wndRect))
					wndWidth = wndRect.right - wndRect.left
					wndHeight = wndRect.bottom - wndRect.top
					if wndWidth > 0 and wndHeight > 0:
						relX = (rect.left - wndRect.left) / wndWidth
						relY = (rect.top - wndRect.top) / wndHeight
						# Search field: left sidebar (< 50% width) and near top (< 30% height)
						if relX < 0.5 and relY < 0.3:
							return True
			except Exception:
				pass
		except Exception:
			pass
		return False

	def event_NVDAObject_init(self, obj):
		"""Log object initialization at debug level."""
		# Intentionally minimal — accessing obj.name here triggers cross-process
		# COM calls for EVERY object, which can crash LINE's Qt6 UIA provider.
		pass

	def event_gainFocus(self, obj, nextHandler):
		"""Handle focus changes with enhanced text extraction."""
		global lastFocusedObject

		try:
			# Update lastFocusedObject only if it's a specific element, not the generic window
			# This prevents the "focus bounce" (ListItem -> Window) from hiding the ListItem
			role = obj.role
			if role not in (
				controlTypes.Role.WINDOW,
				controlTypes.Role.APPLICATION,
				controlTypes.Role.PANE,
			):
				lastFocusedObject = obj
		except Exception:
			pass

		VirtualWindow.onFocusChanged(obj)
		nextHandler()

	def event_UIA_elementSelected(self, obj, nextHandler):
		"""Handle UIA element selection events.

		Qt6 apps sometimes fire elementSelected instead of focus for list items.
		"""
		if _suppressAddon:
			nextHandler()
			return
		try:
			log.info(
				f"LINE UIA_elementSelected: role={obj.role}, name={obj.name!r}, class={obj.windowClassName}",
			)
		except Exception:
			pass
		# If we get a selection event for a list item, treat it as focus
		if obj.role == controlTypes.Role.LISTITEM:
			try:
				obj.setFocus()
				api.setFocusObject(obj)
				api.setNavigatorObject(obj)
				speech.cancelSpeech()
				speech.speakObject(obj, reason=controlTypes.OutputReason.FOCUS)
				braille.handler.handleGainFocus(obj)
			except Exception:
				log.debugWarning("Error handling elementSelected", exc_info=True)
		nextHandler()

	def event_UIA_notification(self, obj, nextHandler, **kwargs):
		"""Handle UIA notification events."""
		if _suppressAddon:
			nextHandler()
			return
		try:
			log.info(
				f"LINE UIA_notification: role={obj.role}, name={obj.name!r}, kwargs={kwargs}",
			)
		except Exception:
			pass
		nextHandler()

	def event_stateChange(self, obj, nextHandler):
		"""Track state changes for potentially focusable elements."""
		if _suppressAddon:
			nextHandler()
			return
		try:
			if isinstance(obj, UIA) and obj.role == controlTypes.Role.LISTITEM:
				if controlTypes.State.SELECTED in obj.states:
					log.info(
						f"LINE stateChange SELECTED: role={obj.role}, "
						f"name={obj.name!r}, class={obj.windowClassName}",
					)
		except Exception:
			pass
		nextHandler()

	def event_nameChange(self, obj, nextHandler):
		"""Track name changes which may indicate content update."""
		if _suppressAddon:
			nextHandler()
			return
		try:
			log.debug(
				f"LINE nameChange: role={obj.role}, name={obj.name!r}, class={obj.windowClassName}",
			)
		except Exception:
			pass
		nextHandler()

	@script(
		# Translators: Description of a debug script that logs UIA tree info
		description=_("Debug: log UIA tree info for the focused element"),
		gesture="kb:NVDA+shift+k",
		category="LINE Desktop",
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
		"""Invoke a UIA element using InvokePattern or mouse click fallback."""

		# Try InvokePattern
		try:
			if _tryInvokeUIAElement(element):
				if announce:
					ui.message(actionName)
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
				import time

				time.sleep(0.05)
				ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
				if announce:
					ui.message(actionName)
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
		import time

		if hwnd is None:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
		if hwnd:
			ctypes.windll.user32.SetForegroundWindow(hwnd)

		ctypes.windll.user32.SetCursorPos(int(x), int(y))
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP

	def _getHeaderIconPosition(self):
		"""Get the screen position of the phone icon in the LINE chat header.

		LINE's Qt6 UI does not expose header toolbar buttons via UIA.
		We use the window geometry to calculate where the icons are.
		All pixel offsets are scaled by system DPI so positions adapt to
		different display scaling settings (100%–300%).

		The chat header has icons from right to left:
		  Index 0: More options (⋮ three dots menu)
		  Index 1: Notes/Keep (📝)
		  Index 2: Phone/Voice call (📞)
		  Index 3: Search (🔍)

		Returns:
			(phoneX, phoneY, winRight) tuple, or None if window not found.
		"""
		import ctypes
		import ctypes.wintypes

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			log.debug("LINE: no foreground window for header click")
			return None

		scale = _getDpiScale(hwnd)

		# Get window rect — try DwmGetWindowAttribute first for
		# accurate extended frame bounds (avoids DPI virtualization).
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winLeft = rect.left
		winTop = rect.top
		winRight = rect.right
		winWidth = winRight - winLeft

		# Also try DWM extended frame bounds
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
				log.info(
					f"LINE: DWM frame bounds=({dwmRect.left},{dwmRect.top},{dwmRect.right},{dwmRect.bottom})",
				)
				# Use DWM bounds if different (more accurate)
				if dwmRect.right != rect.right or dwmRect.top != rect.top:
					log.info("LINE: using DWM bounds instead of GetWindowRect")
					winLeft = dwmRect.left
					winTop = dwmRect.top
					winRight = dwmRect.right
					winWidth = winRight - winLeft
		except Exception:
			pass

		log.info(
			f"LINE: window rect=({winLeft},{winTop},{winRight},{rect.bottom}), "
			f"width={winWidth}, dpiScale={scale:.2f}",
		)

		# Reference values at 96 DPI (100% scaling), scaled by DPI factor.
		# Icon spacing ~27px at 96 DPI, first icon offset ~15px from right edge.
		iconY = winTop + int(55 * scale)
		iconSpacing = int(27 * scale)
		firstIconOffset = int(15 * scale)
		iconX = winRight - firstIconOffset - (2 * iconSpacing)

		log.info(
			f"LINE: header icon pos: iconX={iconX}, iconY={iconY}, "
			f"spacing={iconSpacing}, offset={firstIconOffset}",
		)

		# Verify position is within window bounds
		if iconX < winLeft or iconX > winRight:
			log.warning(f"LINE: icon position {iconX} outside window bounds")
			return None

		return (iconX, iconY, winRight)

	def _clickMoreOptionsButton(self):
		"""Click the more options (⋮) button in the chat header.

		The more options button is the rightmost icon in the header (index 0).
		Returns True if successful, False if window not found.
		"""
		import ctypes
		import ctypes.wintypes

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			log.debug("LINE: no foreground window for more options click")
			return False

		scale = _getDpiScale(hwnd)

		# Get window rect
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winLeft = rect.left
		winTop = rect.top
		winRight = rect.right

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
			if hr == 0 and (dwmRect.right != rect.right or dwmRect.top != rect.top):
				winLeft = dwmRect.left
				winTop = dwmRect.top
				winRight = dwmRect.right
		except Exception:
			pass

		# Calculate more options button position (rightmost icon, index 0)
		iconY = winTop + int(55 * scale)
		firstIconOffset = int(15 * scale)
		moreOptionsX = winRight - firstIconOffset

		log.info(
			f"LINE: clicking more options button at ({moreOptionsX}, {iconY}), dpiScale={scale:.2f}",
		)

		# Verify position is within window bounds
		if moreOptionsX < winLeft or moreOptionsX > winRight:
			log.warning(f"LINE: more options position {moreOptionsX} outside window bounds")
			return False

		# Click the button
		self._clickAtPosition(moreOptionsX, iconY, hwnd)

		return True

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

							log.info(
								f"LINE: confirmation dialog OCR: {ocrText!r}",
							)

							isGroup = "群組" in ocrText
							if ocrText:
								ui.message(ocrText)
							else:
								ui.message(_("語音通話確認"))

							core.callLater(
								300,
								_clickStart,
								isGroup,
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

		def _clickStart(isGroup=False):
			"""Click the 開始 (Start) button on the voice call confirmation dialog.

			Group call dialogs are taller (member avatars), so the button
			is further below the window centre than for personal calls.
			"""
			try:
				sScale = _getDpiScale()
				winCenterX = winLeft + winW // 2
				winCenterY = winTop + winH // 2
				startBtnX = winCenterX - int(43 * sScale)
				if isGroup:
					startBtnY = winCenterY + int(50 * sScale)
				else:
					startBtnY = winCenterY + int(17 * sScale)

				log.info(
					f"LINE: clicking 開始 at ({startBtnX}, {startBtnY}) group={isGroup}",
				)
				appModRef._clickAtPosition(startBtnX, startBtnY, hwnd)

				if isGroup:
					ui.message(_("已開始群組語音通話"))
				else:
					ui.message(_("已開始語音通話"))
			except Exception as e:
				log.warning(
					f"LINE: click 開始 failed: {e}",
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
		import time

		# Step 0: Bring call window to foreground
		try:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
			time.sleep(0.3)
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

				ui.message(_("已接聽"))
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
		ui.message(_("已接聽"))
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
		import time

		# Step 0: Bring call window to foreground
		try:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
			time.sleep(0.3)
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

				ui.message(_("已拒絕"))
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
		ui.message(_("已拒絕"))
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
				ui.message(_("來電：{callerName}").format(callerName=callerName))
			else:
				ui.message(_("來電（OCR: {ocrText}）").format(ocrText=ocrText))
			log.info(f"LINE: caller info OCR: {ocrText!r}")
		else:
			ui.message(_("無法辨識來電者"))
			log.info("LINE: caller info OCR returned empty")

	# ── Incoming call scripts ──────────────────────────────────────────
	# Note: gesture bindings are registered in the GlobalPlugin
	# (lineDesktopHelper.py) so they work even when LINE isn't focused.

	def script_answerCall(self, gesture):
		"""Answer an incoming LINE call."""
		try:
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._answerIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(_("接聽功能錯誤: {error}").format(error=e))

	def script_rejectCall(self, gesture):
		"""Reject an incoming LINE call."""
		try:
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._rejectIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(_("拒絕功能錯誤: {error}").format(error=e))

	def script_checkCaller(self, gesture):
		"""Announce who is calling."""
		try:
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._getCallerInfo(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(_("來電查看功能錯誤: {error}").format(error=e))

	def script_focusCallWindow(self, gesture):
		"""Find the LineCall window, bring it to foreground, and OCR its content."""
		import ctypes
		import ctypes.wintypes

		hwnd = self._findIncomingCallWindow()
		if not hwnd:
			ui.message(_("未偵測到通話視窗"))
			return

		# Bring the call window to the foreground
		try:
			ctypes.windll.user32.SetForegroundWindow(hwnd)
		except Exception:
			pass

		# Give the window time to come to foreground, then OCR it
		def _announceCallWindow():
			try:
				ocrText = self._ocrWindowArea(hwnd, sync=True, timeout=3.0)
				if ocrText:
					speech.cancelSpeech()
					ui.message(ocrText)
					log.info(f"LINE: call window OCR: {ocrText!r}")
				else:
					ui.message(_("通話視窗（無法辨識內容）"))
			except Exception as e:
				log.warning(f"LINE: call window OCR error: {e}", exc_info=True)
				ui.message(_("通話視窗"))

		core.callLater(300, _announceCallWindow)

	# ── Outgoing call scripts ──────────────────────────────────────────

	@script(
		# Translators: Description of a script to make a voice call
		description=_("撥打語音通話"),
		gesture="kb:NVDA+windows+c",
		category="LINE Desktop",
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
		category="LINE Desktop",
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
		category="LINE Desktop",
	)
	def script_clickMoreOptions(self, gesture):
		"""Click the more options (⋮) button in the chat header."""
		try:
			if not self._clickMoreOptionsButton():
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
				return
			import core

			core.callLater(500, self._activateMoreOptionsMenu)
		except Exception as e:
			log.warning(f"LINE clickMoreOptions error: {e}", exc_info=True)
			ui.message(_("更多選項功能錯誤: {error}").format(error=e))

	def _activateMoreOptionsMenu(self, retriesLeft=3):
		"""Find the more options popup and activate the virtual window for browsing."""
		import ctypes
		import ctypes.wintypes as wintypes

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			return

		pid = wintypes.DWORD()
		tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
		threadWindows = []
		seenWindows = set()

		def _captureWindowInfo(targetHwnd):
			if (
				not targetHwnd
				or targetHwnd in seenWindows
				or not ctypes.windll.user32.IsWindow(targetHwnd)
				or not ctypes.windll.user32.IsWindowVisible(targetHwnd)
			):
				return

			wRect = wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(targetHwnd, ctypes.byref(wRect))
			width = wRect.right - wRect.left
			height = wRect.bottom - wRect.top
			if width < 50 or height < 100:
				return

			classBuf = ctypes.create_unicode_buffer(256)
			ctypes.windll.user32.GetClassNameW(targetHwnd, classBuf, 256)
			seenWindows.add(targetHwnd)
			threadWindows.append(
				{
					"hwnd": targetHwnd,
					"left": wRect.left,
					"top": wRect.top,
					"right": wRect.right,
					"bottom": wRect.bottom,
					"width": width,
					"height": height,
					"area": width * height,
					"className": classBuf.value,
				},
			)

		def _enumCb(enumHwnd, lParam):
			_captureWindowInfo(enumHwnd)
			return True

		ctypes.windll.user32.EnumThreadWindows(tid, WNDENUMPROC(_enumCb), 0)
		_captureWindowInfo(hwnd)

		if not threadWindows:
			if retriesLeft > 0:
				import core

				core.callLater(300, lambda: self._activateMoreOptionsMenu(retriesLeft - 1))
			return

		mainWindow = max(threadWindows, key=lambda item: item["area"])
		scale = _getDpiScale(mainWindow["hwnd"])
		anchorX = mainWindow["right"] - int(15 * scale)
		anchorY = mainWindow["top"] + int(55 * scale)
		mainArea = max(mainWindow["area"], 1)

		def _distanceToRect(pointX, pointY, rectInfo):
			dx = 0
			if pointX < rectInfo["left"]:
				dx = rectInfo["left"] - pointX
			elif pointX > rectInfo["right"]:
				dx = pointX - rectInfo["right"]
			dy = 0
			if pointY < rectInfo["top"]:
				dy = rectInfo["top"] - pointY
			elif pointY > rectInfo["bottom"]:
				dy = pointY - rectInfo["bottom"]
			return dx + dy

		candidates = []
		for info in threadWindows:
			if info["width"] >= int(mainWindow["width"] * 0.85) and info["height"] >= int(
				mainWindow["height"] * 0.85,
			):
				log.debug(
					f"LINE: skipping near-main window hwnd={info['hwnd']} "
					f"rect=({info['left']}, {info['top']}, {info['right']}, {info['bottom']}) "
					f"class={info['className']!r}",
				)
				continue

			distance = _distanceToRect(anchorX, anchorY, info)
			score = distance * 4
			score += abs(info["right"] - mainWindow["right"])
			score += abs(info["top"] - anchorY)
			score += int(info["area"] / mainArea * 500)
			if info["left"] < mainWindow["left"] + int(mainWindow["width"] * 0.5):
				score += 250
			if info["top"] > mainWindow["top"] + int(220 * scale):
				score += 200
			containsAnchor = (
				info["left"] <= anchorX <= info["right"]
				and info["top"] - int(24 * scale) <= anchorY <= info["bottom"]
			)
			if containsAnchor:
				score -= 150

			info["geometryScore"] = score
			info["containsAnchor"] = containsAnchor
			candidates.append(info)
			log.debug(
				f"LINE: more options candidate hwnd={info['hwnd']} "
				f"class={info['className']!r} "
				f"rect=({info['left']}, {info['top']}, {info['right']}, {info['bottom']}) "
				f"score={score} containsAnchor={containsAnchor}",
			)

		if not candidates:
			if retriesLeft > 0:
				import core

				core.callLater(300, lambda: self._activateMoreOptionsMenu(retriesLeft - 1))
			return

		candidates.sort(key=lambda item: item["geometryScore"])

		from ._virtualWindows.chatMoreOptions import ChatMoreOptions, _matchMenuLabel

		bestCandidate = None
		bestMatchCount = -1
		topCandidates = candidates[: min(3, len(candidates))]
		for info in topCandidates:
			ocrText = self._ocrWindowArea(
				info["hwnd"],
				region=(
					info["left"],
					info["top"],
					info["width"],
					info["height"],
				),
				sync=True,
				timeout=1.5,
			)
			labels = []
			for line in ocrText.splitlines():
				label = _matchMenuLabel(line)
				if label and label not in labels:
					labels.append(label)
			info["ocrMenuLabels"] = labels
			log.debug(
				f"LINE: more options candidate hwnd={info['hwnd']} OCR labels={labels}",
			)
			if len(labels) > bestMatchCount:
				bestCandidate = info
				bestMatchCount = len(labels)
			elif (
				len(labels) == bestMatchCount
				and bestCandidate
				and info["geometryScore"] < bestCandidate["geometryScore"]
			):
				bestCandidate = info

		if bestCandidate is None:
			bestCandidate = candidates[0]

		if bestMatchCount <= 0 and retriesLeft > 0:
			import core

			log.debug("LINE: more options popup not verified by OCR yet, retrying")
			core.callLater(300, lambda: self._activateMoreOptionsMenu(retriesLeft - 1))
			return

		left = bestCandidate["left"]
		top = bestCandidate["top"]
		right = bestCandidate["right"]
		bottom = bestCandidate["bottom"]
		popupRect = (left, top, right, bottom)
		popupRowRects = _collectPopupMenuRowRects(
			bestCandidate["hwnd"],
			popupRect,
		)
		log.info(f"LINE: more options popup found at ({left}, {top}, {right}, {bottom})")
		VirtualWindow.currentWindow = ChatMoreOptions(
			popupRect,
			rowRects=popupRowRects,
			onAction=self._handleChatMoreOptionsAction,
		)

	# ── Message context menu (right-click / Shift+F10) ────────────────

	def _activateMessageContextMenu(
		self,
		retriesLeft=3,
		onAction=None,
		onFailure=None,
		shouldAbort=None,
	):
		"""Find a keyboard-opened message context menu and activate the virtual window."""
		import ctypes
		import ctypes.wintypes as wintypes

		if callable(shouldAbort) and shouldAbort():
			return

		hwnd = ctypes.windll.user32.GetForegroundWindow()
		if not hwnd:
			if callable(onFailure):
				onFailure()
			return

		pid = wintypes.DWORD()
		tid = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
		WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
		popupCandidates = []
		mainRect = wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(mainRect))
		mainArea = max(
			(mainRect.right - mainRect.left) * (mainRect.bottom - mainRect.top),
			1,
		)

		def _enumCb(enumHwnd, lParam):
			if enumHwnd != hwnd and ctypes.windll.user32.IsWindowVisible(enumHwnd):
				wRect = wintypes.RECT()
				ctypes.windll.user32.GetWindowRect(enumHwnd, ctypes.byref(wRect))
				w = wRect.right - wRect.left
				h = wRect.bottom - wRect.top
				area = w * h
				if w >= 50 and h >= 30 and area < mainArea * 0.5:
					popupCandidates.append(
						{
							"hwnd": enumHwnd,
							"left": wRect.left,
							"top": wRect.top,
							"right": wRect.right,
							"bottom": wRect.bottom,
							"width": w,
							"height": h,
						},
					)
			return True

		ctypes.windll.user32.EnumThreadWindows(tid, WNDENUMPROC(_enumCb), 0)

		if not popupCandidates:
			if retriesLeft > 0:
				import core

				core.callLater(
					300,
					lambda: self._activateMessageContextMenu(
						retriesLeft - 1,
						onAction=onAction,
						onFailure=onFailure,
						shouldAbort=shouldAbort,
					),
				)
			elif callable(onFailure):
				onFailure()
			return

		if callable(shouldAbort) and shouldAbort():
			return

		best = max(popupCandidates, key=lambda c: c["height"])
		popupRect = (best["left"], best["top"], best["right"], best["bottom"])
		popupRowRects = _collectPopupMenuRowRects(best["hwnd"], popupRect)
		if not _popupLooksLikeMessageContextMenu(self, hwnd, popupRect):
			log.debug(
				f"LINE: keyboard-opened popup did not resemble a message context menu "
				f"at {popupRect}; dismissing",
			)
			try:
				_sendGestureWithAddonSuppressed("escape")
			except Exception:
				pass
			if callable(onFailure):
				onFailure()
			return

		log.info(
			f"LINE: message context menu popup found at "
			f"({best['left']}, {best['top']}, {best['right']}, {best['bottom']})",
		)

		from ._virtualWindows.messageContextMenu import MessageContextMenu

		VirtualWindow.currentWindow = MessageContextMenu(
			popupRect,
			rowRects=popupRowRects,
			onAction=onAction,
		)
		return

	# ── Read chat room name ────────────────────────────────────────────

	def _readChatRoomName(self):
		"""Read the current chat room name.

		Priority order:
		  1. Cache file name [LINE]姓名.txt (after NVDA+Windows+U) — extract the
		     name segment between the [LINE] prefix and the .txt suffix.
		  2. Header-area OCR — fallback when no cache exists.
		"""
		import ctypes
		import ctypes.wintypes

		# 1. Background cache file name (NVDA+Windows+U was used for this room)
		try:
			from . import _chatCache

			if _chatCache.isActive():
				tempPath = _chatCache.getTempPath()
				if tempPath:
					base = os.path.basename(tempPath)
					stem, ext = os.path.splitext(base)
					if ext.lower() == ".txt" and stem.startswith("[LINE]"):
						nameFromFile = stem[len("[LINE]") :].strip()
						if nameFromFile:
							log.info(f"LINE: using chat name from cache file {nameFromFile!r}")
							ui.message(nameFromFile)
							return
		except Exception:
			log.debug("LINE: chat cache file-name lookup failed", exc_info=True)

		# 2. Header-area OCR (fallback)
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

		try:
			ocrText = self._ocrWindowArea(
				hwnd,
				region=(nameLeft, nameTop, nameW, nameH),
				sync=True,
				timeout=3.0,
			)
			ocrText = _removeCJKSpaces(ocrText.strip()) if ocrText else ""
			if ocrText:
				ui.message(ocrText)
			else:
				ui.message(_("無法讀取聊天室名稱"))
		except Exception as e:
			log.warning(f"LINE: readChatRoomName OCR error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to read the current chat room name
		description=_("讀出目前聊天室名稱"),
		gesture="kb:NVDA+windows+t",
		category="LINE Desktop",
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
		category="LINE Desktop",
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

	def _pollFileDialog(self):
		"""Poll to detect when the file dialog closes, then resume addon.

		We enumerate all #32770 windows and check if any belong to LINE's
		process. Using FindWindowW("#32770", None) is wrong because it finds
		ANY #32770 window in the system (e.g. battery warning dialogs).
		"""
		global _suppressAddon
		import ctypes
		import ctypes.wintypes

		lineProcessId = self.processID

		try:
			foundOurDialog = False

			# Callback for EnumWindows
			WNDENUMPROC = ctypes.WINFUNCTYPE(
				ctypes.wintypes.BOOL,
				ctypes.wintypes.HWND,
				ctypes.wintypes.LPARAM,
			)

			def _enumCallback(hwnd, lParam):
				nonlocal foundOurDialog
				# Get the class name of this window
				buf = ctypes.create_unicode_buffer(256)
				ctypes.windll.user32.GetClassNameW(
					hwnd,
					buf,
					256,
				)
				if buf.value == "#32770":
					# Check if this dialog belongs to LINE's process
					pid = ctypes.wintypes.DWORD()
					ctypes.windll.user32.GetWindowThreadProcessId(
						hwnd,
						ctypes.byref(pid),
					)
					if pid.value == lineProcessId:
						foundOurDialog = True
						return False  # stop enumeration
				return True  # continue enumeration

			ctypes.windll.user32.EnumWindows(
				WNDENUMPROC(_enumCallback),
				0,
			)

			if foundOurDialog:
				log.debug("LINE: file dialog still open, polling...")
				core.callLater(500, self._pollFileDialog)
			else:
				_suppressAddon = False
				log.info("LINE: file dialog closed, addon resumed")
		except Exception as e:
			log.warning(f"LINE: file dialog poll error: {e}")
			_suppressAddon = False

	def _suppressAddonForFileDialog(self, reason):
		"""Pause addon behavior until LINE's file dialog closes."""
		global _suppressAddon
		_suppressAddon = True
		log.info(f"LINE: {reason}, addon suppressed, waiting for file dialog...")
		core.callLater(1000, self._pollFileDialog)

	def _handleChatMoreOptionsAction(self, actionName):
		"""Handle post-click actions from the chat more-options virtual window."""
		if actionName == "儲存聊天":
			if getattr(self, "_messageReaderPending", False):
				core.callLater(800, self._messageReaderHandleSaveDialog)
			else:
				self._suppressAddonForFileDialog("Save chat selected")

	def script_openMessageReader(self, gesture):
		"""Open message reader: click more options, auto-click save chat, parse, and display."""
		if getattr(self, "_messageReaderPending", False):
			ui.message(_("訊息閱讀器正在執行中，請稍候"))
			return
		ui.message(_("正在開啟訊息閱讀器…"))
		try:
			if not self._clickMoreOptionsButton():
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
				return
			self._messageReaderPending = True
			self._messageReaderBackgroundCache = False
			core.callLater(500, self._activateMoreOptionsMenu)
			# Poll until virtual window is ready, then auto-click 儲存聊天
			core.callLater(1500, self._messageReaderAutoClickSaveChat)
		except Exception as e:
			log.warning(f"LINE openMessageReader error: {e}", exc_info=True)
			self._messageReaderPending = False
			ui.message(_("訊息閱讀器功能錯誤"))

	def script_cacheChatToBackground(self, gesture):
		"""Save chat to a temp file and cache it for background message lookup.

		Behaves like the message reader save flow, but skips the reader dialog
		so focus stays in the chat room. Subsequent message-list navigation
		looks up the focused bubble in the cached text via OCR matching
		instead of running the copy-read flow.
		"""
		if getattr(self, "_messageReaderPending", False):
			ui.message(_("訊息閱讀器正在執行中，請稍候"))
			return
		ui.message(_("正在儲存聊天到背景…"))
		try:
			if not self._clickMoreOptionsButton():
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
				return
			self._messageReaderPending = True
			self._messageReaderBackgroundCache = True
			core.callLater(500, self._activateMoreOptionsMenu)
			core.callLater(1500, self._messageReaderAutoClickSaveChat)
		except Exception as e:
			log.warning(f"LINE cacheChatToBackground error: {e}", exc_info=True)
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False
			ui.message(_("背景快取功能錯誤"))

	def _messageReaderAutoClickSaveChat(self, retriesLeft=15):
		"""Poll the ChatMoreOptions virtual window until 儲存聊天 is found, then click it."""
		if not getattr(self, "_messageReaderPending", False):
			return

		from ._virtualWindows.chatMoreOptions import ChatMoreOptions

		window = VirtualWindow.currentWindow
		if not isinstance(window, ChatMoreOptions) or not window.elements:
			if retriesLeft > 0:
				core.callLater(300, lambda: self._messageReaderAutoClickSaveChat(retriesLeft - 1))
			else:
				self._messageReaderPending = False
				self._messageReaderBackgroundCache = False
				ui.message(_("找不到儲存聊天選項"))
			return

		for i, elem in enumerate(window.elements):
			if elem.get("name") == "儲存聊天":
				window.pos = i
				window.click()
				log.info("LINE: message reader auto-clicked 儲存聊天")
				return

		if retriesLeft > 0:
			core.callLater(300, lambda: self._messageReaderAutoClickSaveChat(retriesLeft - 1))
		else:
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False
			ui.message(_("找不到儲存聊天選項"))

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
				self._messageReaderBackgroundCache = False
				ui.message(_("未偵測到儲存對話框"))
			return

		# Find the filename edit control in the Save dialog
		# The file dialog has a ComboBoxEx32 > ComboBox > Edit hierarchy
		editHwnd = self._findSaveDialogEdit(dialogHwnd)
		if not editHwnd:
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False
			ui.message(_("無法操作儲存對話框"))
			return

		# Read LINE's pre-filled filename — it's [LINE]{chatRoomName}.txt and
		# carries the authoritative room name straight from LINE (much cleaner
		# than OCR-derived names that include unread badges like "( ) 9").
		WM_GETTEXT = 0x000D
		WM_GETTEXTLENGTH = 0x000E
		dialogStem = ""
		try:
			textLen = ctypes.windll.user32.SendMessageW(editHwnd, WM_GETTEXTLENGTH, 0, 0)
			if textLen > 0:
				readBuf = ctypes.create_unicode_buffer(textLen + 1)
				ctypes.windll.user32.SendMessageW(editHwnd, WM_GETTEXT, textLen + 1, readBuf)
				suggested = readBuf.value or ""
				stem, ext = os.path.splitext(os.path.basename(suggested))
				if ext.lower() == ".txt" and stem.startswith("[LINE]"):
					dialogStem = stem[len("[LINE]") :].strip()
				log.info(
					f"LINE: save dialog suggested filename {suggested!r}, extracted stem {dialogStem!r}",
				)
		except Exception as e:
			log.debug(f"LINE: could not read save dialog edit text: {e}", exc_info=True)

		# Adopt the dialog's chat name so the cache and any later NVDA+Windows+T
		# read use the LINE-authoritative name, not a stale list-nav OCR value.
		if dialogStem:
			global _currentChatRoomName
			if _currentChatRoomName != dialogStem:
				log.info(
					f"LINE: updating chat room name from save dialog: {_currentChatRoomName!r} -> {dialogStem!r}",
				)
				_currentChatRoomName = dialogStem

		# Build temp file path (use system temp dir to avoid locking the addon folder).
		# Use a separate file when caching to background so an open reader window
		# (cleanupPath bound) does not race with the cache file.
		if getattr(self, "_messageReaderBackgroundCache", False):
			# Name cache files [LINE]姓名.txt so users can identify which room
			# each cache belongs to. Prefer the dialog stem (LINE's own name)
			# over _currentChatRoomName which may be empty or list-nav noise.
			candidate = dialogStem or (_currentChatRoomName or "").strip()
			safeName = re.sub(r'[\\/:*?"<>|]', "_", candidate) if candidate else ""
			safeName = safeName.rstrip(" .")
			if safeName:
				fileName = f"[LINE]{safeName}.txt"
			else:
				fileName = "lineDesktop_chat_cache.txt"
		else:
			fileName = "lineDesktop_chat_export.txt"
		savePath = os.path.join(tempfile.gettempdir(), fileName)
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
		"""Read the saved chat file, then either open the reader or cache silently."""
		savePath = getattr(self, "_messageReaderSavePath", None)
		backgroundCache = getattr(self, "_messageReaderBackgroundCache", False)
		if not savePath or not os.path.isfile(savePath):
			ui.message(_("找不到儲存的聊天紀錄檔案"))
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False
			return

		try:
			from ._chatParser import parseChatFile

			messages = parseChatFile(savePath)
			if not messages:
				ui.message(_("聊天紀錄中沒有訊息"))
				self._messageReaderPending = False
				self._messageReaderBackgroundCache = False
				try:
					if backgroundCache and os.path.isfile(savePath):
						os.remove(savePath)
				except Exception:
					pass
				return

			if backgroundCache:
				from . import _chatCache

				_chatCache.setCache(messages, savePath, _currentChatRoomName)
				log.info(
					f"LINE: cached {len(messages)} messages from {savePath}"
					f" for room {_currentChatRoomName!r}",
				)
				ui.message(
					_("聊天已快取到背景，瀏覽訊息時將直接讀取（離開聊天室後清除）"),
				)
			else:
				from ._messageReader import openMessageReader

				log.info(
					f"LINE: message reader parsed {len(messages)} messages from {savePath}",
				)
				openMessageReader(messages, cleanupPath=savePath)
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False
		except Exception as e:
			log.warning(f"LINE: message reader parse error: {e}", exc_info=True)
			ui.message(
				_("背景快取開啟錯誤") if backgroundCache else _("訊息閱讀器開啟錯誤"),
			)
			self._messageReaderPending = False
			self._messageReaderBackgroundCache = False

	def script_openFileDialog(self, gesture):
		"""Pass Ctrl+O to LINE, suppress addon while file dialog is open."""
		self._suppressAddonForFileDialog("Ctrl+O pressed")
		gesture.send()

	def script_navigateAndTrack(self, gesture):
		"""Pass navigation key to LINE, then poll UIA focused element.

		LINE's Qt6 framework does not fire UIA focus change events when
		navigating with Tab/arrows. This script sends the key through,
		waits briefly for LINE to process it, then queries the UIA
		focused element directly and announces it.
		"""
		if _suppressAddon:
			gesture.send()
			return
		global _lastOCRElement, _chatListMode
		# Exiting chat list mode on Tab/Shift+Tab navigation
		_chatListMode = False
		_lastOCRElement = None
		gesture.send()
		_scheduleQueryAndSpeakUIAFocus(100)

	def script_chatListArrow(self, gesture):
		"""Navigate the chat list with up/down arrows.

		When in chat list context (search field or chat list item),
		passes the arrow key to LINE, then sends Tab twice to reach
		the chatroom name and reads it aloud.

		When NOT in chat list context (e.g. message list), passes
		the arrow key through normally with standard announcement.
		"""
		if _suppressAddon:
			gesture.send()
			return

		global _lastOCRElement
		_lastOCRElement = None

		# Check if we're in the chat list context
		try:
			handler = UIAHandler.handler
			if handler:
				inList, listEl, items, currentIdx = _isInChatListContext(handler)
				if inList:
					# In chat list — arrow + Tab x2 to read chatroom name
					gesture.send()

					def _sendFirstTab():
						"""Send first Tab key, then schedule second Tab."""
						try:
							from keyboardHandler import KeyboardInputGesture

							tabGesture = KeyboardInputGesture.fromName("tab")
							tabGesture.send()
						except Exception:
							log.debugWarning("First Tab simulation failed", exc_info=True)
						core.callLater(50, _sendSecondTabAndRead)

					def _sendSecondTabAndRead():
						"""Send second Tab key, then read the focused element."""
						try:
							from keyboardHandler import KeyboardInputGesture

							tabGesture = KeyboardInputGesture.fromName("tab")
							tabGesture.send()
						except Exception:
							log.debugWarning("Second Tab simulation failed", exc_info=True)
						# Read the focused element after Tab x2
						_scheduleQueryAndSpeakUIAFocus(100)

					core.callLater(200, _sendFirstTab)
					return
		except Exception:
			log.debugWarning("Chat list context check failed", exc_info=True)

		# Not in chat list — pass through normally (message list, etc.)
		gesture.send()
		_scheduleQueryAndSpeakUIAFocus(100)

	# Map Control+number keys to LINE tab names
	_TAB_NAMES = {
		"1": _("Friends"),
		"2": _("Chats"),
		"3": _("Add friends"),
	}

	def script_switchTabAndAnnounce(self, gesture):
		"""Pass Control+1/2/3 to LINE and announce the tab name.

		LINE's built-in shortcuts: Control+1=Friends, Control+2=Chats,
		Control+3=Add friends. This script passes the key through and
		announces the tab being switched to.
		"""
		if _suppressAddon:
			gesture.send()
			return
		global _currentChatRoomName, _chatListMode, _chatListSearchField
		# Clear chat state when switching tabs
		_currentChatRoomName = None
		_chatListMode = False
		_chatListSearchField = None
		try:
			from . import _chatCache

			_chatCache.clearCache()
		except Exception:
			log.debug("LINE: chat cache clear-on-tab-switch failed", exc_info=True)
		gesture.send()
		# Extract the number key from the gesture identifier
		# gesture.mainKeyName gives us the key name like "1", "2", "3"
		keyName = gesture.mainKeyName
		tabName = self._TAB_NAMES.get(keyName, "")
		if tabName:
			ui.message(tabName)

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
						label = _detectEditFieldLabel(rawElement, handler, allowNotesOcr=False)
						log.debug(f"LINE Enter key: edit field label={label!r}")
						if label == _("Message input"):
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

	# ── Right-click context menu actions ─────────────────────────────────

	def _rightClickAtPosition(self, x, y, hwnd=None):
		"""Perform a right-click at the given screen coordinates."""
		import ctypes
		import time

		if hwnd is None:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
		if hwnd:
			ctypes.windll.user32.SetForegroundWindow(hwnd)

		ctypes.windll.user32.SetCursorPos(int(x), int(y))
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0008, 0, 0, 0, 0)  # RIGHTDOWN
		time.sleep(0.05)
		ctypes.windll.user32.mouse_event(0x0010, 0, 0, 0, 0)  # RIGHTUP

	def _getMessageCenter(self):
		"""Get the center coordinates of the currently focused message.

		Returns (cx, cy, hwnd) or None.
		"""
		import ctypes

		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if not rawEl:
				return None

			rect = rawEl.CurrentBoundingRectangle
			cx = int((rect.left + rect.right) / 2)
			cy = int((rect.top + rect.bottom) / 2)

			if cx <= 0 or cy <= 0:
				return None

			hwnd = ctypes.windll.user32.GetForegroundWindow()
			return (cx, cy, hwnd)
		except Exception as e:
			log.debug(f"LINE: _getMessageCenter failed: {e}")
			return None

	# Reference offsets from right-click point to each context menu item
	# center, measured at 100% DPI (1920x1080, 17-inch screen).
	# Message at (697, 255); menu items at:
	#   複製 (715, 288) → offset (+18, +33)
	#   分享 (714, 323) → offset (+17, +68)
	#   收回 (720, 350) → offset (+23, +95)
	#   回覆 (647, 359) → offset (-50, +104)
	# At runtime these are scaled by _getDpiScale().
	_MENU_OFFSETS = {
		"複製": (18, 33),
		"分享": (17, 68),
		"收回": (23, 95),
		"回覆": (-50, 104),
	}

	def _contextMenuAction(self, itemIndex, actionName, afterCallback=None):
		"""Right-click current message and select an item from the context menu.

		Menu items top to bottom: 0=回覆, 1=複製, 2=分享, 3=收回
		Uses UIA to find the popup menu and click directly on the target item.
		Falls back to offset-based clicking if UIA detection fails.

		If the right-click hits the text content area (producing a text
		selection menu like '全選'), dismisses and retries at a different
		Y offset on the message bubble.
		"""
		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if not rawEl:
				ui.message(_("找不到目前的訊息"))
				return

			rect = rawEl.CurrentBoundingRectangle
			cx = int((rect.left + rect.right) / 2)
			cy = int((rect.top + rect.bottom) / 2)
			elTop = int(rect.top)
			elBottom = int(rect.bottom)

			if cx <= 0 or cy <= 0:
				ui.message(_("找不到目前的訊息"))
				return

			hwnd = ctypes.windll.user32.GetForegroundWindow()
			targetRuntimeId = _getElementRuntimeId(rawEl)

			# Get LINE window rect to clamp click positions
			import ctypes.wintypes as wintypes

			winRect = wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(
				hwnd,
				ctypes.byref(winRect),
			)
			winTop = int(winRect.top)
			winBottom = int(winRect.bottom)
		except Exception as e:
			log.debug(f"LINE: _contextMenuAction getElement failed: {e}")
			ui.message(_("找不到目前的訊息"))
			return

		elLeft = int(rect.left)
		elRight = int(rect.right)

		# Wait for modifier keys (Ctrl, Shift, Alt) to be physically
		# released before proceeding.  The user likely just pressed
		# Ctrl+C, and if we send any keystrokes while Ctrl is still
		# held, NVDA's keyboard hook will combine them (e.g.
		# Escape → Ctrl+Escape → Windows Start Menu).
		VK_CONTROL = 0x11
		VK_SHIFT = 0x10
		VK_MENU = 0x12  # Alt
		log.debug("LINE: waiting for modifier keys to be released")
		GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState
		for _wait in range(40):  # up to ~2 seconds
			held = any(GetAsyncKeyState(vk) & 0x8000 for vk in (VK_CONTROL, VK_SHIFT, VK_MENU))
			if not held:
				break
			time.sleep(0.05)
		log.debug("LINE: modifiers released, proceeding with right-click")

		ocrLines = []
		try:
			ocrResult = self._ocrWindowAreaResult(
				hwnd,
				region=(
					elLeft,
					elTop,
					max(elRight - elLeft, 0),
					max(elBottom - elTop, 0),
				),
				sync=True,
				timeout=2.0,
			)
			if ocrResult is not None and not isinstance(ocrResult, Exception):
				ocrLines = _extractOcrLines(ocrResult)
		except Exception as e:
			log.debug(
				f"LINE: failed to derive OCR bubble probes for {actionName}: {e}",
				exc_info=True,
			)
		ocrClickPositions = _buildMessageBubbleOcrClickPositions(
			ocrLines,
			(elLeft, elTop, elRight, elBottom),
			winTop,
			winBottom,
		)
		if ocrClickPositions:
			log.debug(
				f"LINE: {actionName} OCR-derived bubble probes: "
				f"{[label for _x, _y, label in ocrClickPositions]}",
			)
		elif ocrLines:
			log.debug(
				f"LINE: {actionName} OCR lines yielded no bubble probes: "
				f"{[(line.get('text', ''), bool(line.get('rect'))) for line in ocrLines]}",
			)
		clickPositions = _mergeClickPositions(
			ocrClickPositions,
			_buildMessageBubbleClickPositions(
				(elLeft, elTop, elRight, elBottom),
				winTop,
				winBottom,
				includeVerticalOffsets=True,
			),
		)

		appModRef = self

		def _attemptAtOffset(posIdx=0):
			"""Right-click at clickPositions[posIdx] and find the menu item."""
			if posIdx >= len(clickPositions):
				# All click positions exhausted.
				# For copy action, fall back to OCR + direct clipboard.
				if actionName == "複製":
					log.info(
						"LINE: all click positions failed, falling back to OCR copy",
					)
					try:
						rect = rawEl.CurrentBoundingRectangle
						elW = int(rect.right - rect.left)
						elH = int(rect.bottom - rect.top)
						if elW > 0 and elH > 0:
							ocrText = appModRef._ocrWindowArea(
								hwnd,
								region=(
									int(rect.left),
									int(rect.top),
									elW,
									elH,
								),
								sync=True,
								timeout=3.0,
							)
							ocrText = (
								_removeCJKSpaces(
									ocrText.strip(),
								)
								if ocrText
								else ""
							)
							if ocrText:
								# Remove timestamp lines
								lines = ocrText.split("\n")
								content = []
								for ln in lines:
									stripped = ln.strip()
									# Skip timestamp lines
									if re.match(
										r"^[上下午]*\s*\d{1,2}\s*:\s*\d{2}",
										stripped,
									):
										continue
									if stripped:
										content.append(stripped)
								if content:
									copyText = "\n".join(content)
									import api

									api.copyToClip(copyText)
									log.info(
										f"LINE: OCR fallback copied: {copyText!r}",
									)
									ui.message(_("複製"))
									try:
										nvwave.playWaveFile(_OCR_SOUND_PATH, asynchronous=True)
									except Exception:
										log.debugWarning("Failed to play OCR sound", exc_info=True)
									if afterCallback:
										core.callLater(
											500,
											afterCallback,
										)
									return
					except Exception as e:
						log.debug(
							f"LINE: OCR fallback failed: {e}",
							exc_info=True,
						)
				ui.message(_("找不到「{actionName}」選項").format(actionName=actionName))
				return

			clickX, clickY, posLabel = clickPositions[posIdx]
			log.info(
				f"LINE: right-clicking message at ({clickX}, {clickY}) [{posLabel}] for {actionName}",
			)
			if not _messageProbePointHitsTargetElement(
				UIAHandler.handler,
				rawEl,
				clickX,
				clickY,
				targetRuntimeId,
			):
				log.info(
					f"LINE: skipping {actionName} probe over edit field at ({clickX}, {clickY}) [{posLabel}]",
				)
				core.callLater(50, lambda: _attemptAtOffset(posIdx + 1))
				return
			appModRef._rightClickAtPosition(
				clickX,
				clickY,
				hwnd,
			)

			def _findAndClickMenuItem(retriesLeft=4):
				"""Find the context menu popup and click the target item.

				If the popup or menu items aren't found yet, retries up to
				retriesLeft times with 200ms between attempts.
				If the popup is found but has the wrong menu (target action
				not present), dismisses it and tries the next click position.
				"""
				try:
					uiaHandler = UIAHandler.handler
					if not uiaHandler:
						log.debug("LINE: no UIA handler")
						return

					import ctypes.wintypes as wintypes

					# Find popup window via EnumThreadWindows
					pid = wintypes.DWORD()
					tid = ctypes.windll.user32.GetWindowThreadProcessId(
						hwnd,
						ctypes.byref(pid),
					)
					popupCandidates = []

					WNDENUMPROC = ctypes.WINFUNCTYPE(
						ctypes.c_bool,
						wintypes.HWND,
						wintypes.LPARAM,
					)

					def _enumCallback(enumHwnd, lParam):
						if enumHwnd != hwnd and ctypes.windll.user32.IsWindowVisible(enumHwnd):
							# Filter by window size to skip tooltips
							wRect = wintypes.RECT()
							ctypes.windll.user32.GetWindowRect(
								enumHwnd,
								ctypes.byref(wRect),
							)
							w = wRect.right - wRect.left
							h = wRect.bottom - wRect.top
							if w >= 50 and h >= 30:
								popupCandidates.append(enumHwnd)
							else:
								log.debug(
									f"LINE: skipping small window {enumHwnd}: {w}x{h}",
								)
						return True

					ctypes.windll.user32.EnumThreadWindows(
						tid,
						WNDENUMPROC(_enumCallback),
						0,
					)

					popupHwnd = None
					if popupCandidates:
						popupHwnd = popupCandidates[0]
						log.debug(
							f"LINE: found popup via EnumThreadWindows: "
							f"hwnd {popupHwnd} "
							f"(candidates: {len(popupCandidates)})",
						)
					else:
						for dy in [0, -40, -80, 40, 80]:
							pt = wintypes.POINT(clickX, clickY + dy)
							candidateHwnd = ctypes.windll.user32.WindowFromPoint(pt)
							if candidateHwnd and candidateHwnd != hwnd:
								popupHwnd = candidateHwnd
								log.debug(
									f"LINE: found popup via WindowFromPoint offset dy={dy}: hwnd {popupHwnd}",
								)
								break

					if not popupHwnd:
						if retriesLeft > 0:
							log.debug(
								f"LINE: no popup window found, retrying ({retriesLeft} left)",
							)
							core.callLater(
								200,
								lambda: _findAndClickMenuItem(
									retriesLeft - 1,
								),
							)
							return
						log.debug(
							"LINE: no popup window found after all retries, trying next position",
						)
						core.callLater(
							300,
							lambda: _attemptAtOffset(
								posIdx + 1,
							),
						)
						return

					log.debug(
						f"LINE: using popup hwnd {popupHwnd}, LINE main hwnd = {hwnd}",
					)

					element = uiaHandler.clientObject.ElementFromHandle(
						popupHwnd,
					)
					if not element:
						if retriesLeft > 0:
							log.debug(
								f"LINE: ElementFromHandle returned nothing, retrying ({retriesLeft} left)",
							)
							core.callLater(
								200,
								lambda: _findAndClickMenuItem(
									retriesLeft - 1,
								),
							)
							return
						log.debug(
							"LINE: ElementFromHandle returned nothing",
						)
						return

					# Validate popup is a real context menu,
					# not a tooltip (ct=50033) or other junk
					try:
						ct = element.CurrentControlType
						name = element.CurrentName or ""
						eRect = element.CurrentBoundingRectangle
						eW = int(eRect.right - eRect.left)
						eH = int(eRect.bottom - eRect.top)
						log.debug(
							f"LINE: popup element: ct={ct}, "
							f"name={name!r}, "
							f"rect=({eRect.left},{eRect.top})-"
							f"({eRect.right},{eRect.bottom}), "
							f"{eW}x{eH}",
						)
						# Reject tooltips and tiny popups
						if ct == 50033 or eW < 50 or eH < 30:
							log.debug(
								f"LINE: popup is not a context menu (ct={ct}, {eW}x{eH}), skipping",
							)
							# No real popup, go to next offset
							core.callLater(
								300,
								lambda: _attemptAtOffset(
									posIdx + 1,
								),
							)
							return
					except Exception:
						pass

					walker = uiaHandler.clientObject.RawViewWalker

					def _getMenuItemText(item):
						"""Extract text label from a menu item's children."""
						try:
							textChild = walker.GetFirstChildElement(item)
							childIdx = 0
							while textChild and childIdx < 10:
								try:
									n = textChild.CurrentName
									if n and n.strip():
										return n.strip()
								except Exception:
									pass
								try:
									gc = walker.GetFirstChildElement(
										textChild,
									)
									gcIdx = 0
									while gc and gcIdx < 5:
										try:
											gcN = gc.CurrentName
											if gcN and gcN.strip():
												return gcN.strip()
										except Exception:
											pass
										try:
											gc = walker.GetNextSiblingElement(
												gc,
											)
										except Exception:
											break
										gcIdx += 1
								except Exception:
									pass
								try:
									textChild = walker.GetNextSiblingElement(
										textChild,
									)
								except Exception:
									break
								childIdx += 1
						except Exception:
							pass
						return ""

					def _collectMenuItems(parent, depth=0, prefix=""):
						"""Walk UIA tree and collect menu item elements."""
						items = []
						child = walker.GetFirstChildElement(parent)
						idx = 0
						while child and idx < 30:
							try:
								ct = child.CurrentControlType
								childName = ""
								try:
									childName = child.CurrentName or ""
								except Exception:
									pass
								childRect = child.CurrentBoundingRectangle
								childH = int(
									childRect.bottom - childRect.top,
								)
								childW = int(
									childRect.right - childRect.left,
								)
								log.debug(
									f"LINE: {prefix}child[{idx}] "
									f"ct={ct}, name={childName!r}, "
									f"rect=({childRect.left},"
									f"{childRect.top})-"
									f"({childRect.right},"
									f"{childRect.bottom}), "
									f"{childW}x{childH}",
								)

								if childW <= 0 or childH <= 0:
									pass
								elif 20 <= childH <= 80 and childW >= childH * 2:
									itemText = _getMenuItemText(child)
									log.debug(
										f"LINE: {prefix}child[{idx}] menu item row, text={itemText!r}",
									)
									items.append((child, itemText))
								elif childH > 80 and depth < 5:
									log.debug(
										f"LINE: {prefix}child[{idx}] large container, recursing",
									)
									subItems = _collectMenuItems(
										child,
										depth + 1,
										prefix + "  ",
									)
									items.extend(subItems)
								elif childH >= 20:
									# Smaller items (e.g. 16px separators)
									# are still big enough to detect but
									# not clickable menu items.
									itemText = _getMenuItemText(child)
									items.append((child, itemText))
							except Exception:
								pass
							try:
								child = walker.GetNextSiblingElement(
									child,
								)
							except Exception:
								break
							idx += 1
						return items

					menuItems = _collectMenuItems(element)
					log.debug(
						f"LINE: found {len(menuItems)} menu items: {[t for _, t in menuItems]}",
					)

					if not menuItems:
						if retriesLeft > 0:
							log.debug(
								f"LINE: popup found but 0 menu items, retrying ({retriesLeft} left)",
							)
							core.callLater(
								200,
								lambda: _findAndClickMenuItem(
									retriesLeft - 1,
								),
							)
							return
						log.debug(
							"LINE: 0 menu items after all retries",
						)

					# Strategy 1: Match by UIA text label
					targetItem = None
					targetClickPoint = None
					popupOcrResult = ""
					for item, text in menuItems:
						if text and actionName in text:
							targetItem = item
							log.info(
								f"LINE: matched menu item '{actionName}' by UIA text: {text!r}",
							)
							break

					# Strategy 2: OCR the entire popup window
					# Run OCR BEFORE offset matching to verify the
					# menu actually contains the target action.
					if not targetItem and menuItems:
						log.debug(
							"LINE: no UIA text match, trying whole-popup OCR",
						)
						try:
							popupRect = element.CurrentBoundingRectangle
							popupW = int(popupRect.right - popupRect.left)
							popupH = int(popupRect.bottom - popupRect.top)
							if popupW > 0 and popupH > 0:
								popupRectTuple = (
									int(popupRect.left),
									int(popupRect.top),
									int(popupRect.right),
									int(popupRect.bottom),
								)
								popupRowRects = _collectPopupMenuRowRects(
									popupHwnd,
									popupRectTuple,
								)
								ocrResult = appModRef._ocrWindowAreaResult(
									popupHwnd,
									region=(
										int(popupRect.left),
										int(popupRect.top),
										popupW,
										popupH,
									),
									sync=True,
									timeout=3.0,
								)
								ocrText = ""
								popupOcrLines = []
								if ocrResult is not None and not isinstance(ocrResult, Exception):
									ocrText = getattr(ocrResult, "text", "") or ""
									ocrText = _removeCJKSpaces(ocrText.strip())
									popupOcrLines = _extractOcrLines(ocrResult)
								popupOcrResult = ocrText
								log.debug(
									f"LINE: popup OCR result: {ocrText!r}",
								)
								ocrTarget = _resolvePopupMenuLabelClickPoint(
									actionName,
									popupOcrLines,
									popupRectTuple,
									popupRowRects=popupRowRects,
								)
								if ocrTarget:
									targetClickPoint = ocrTarget["clickPoint"]
									log.info(
										f"LINE: matched '{actionName}' via popup OCR rows, "
										f"item {ocrTarget['index'] + 1}/{ocrTarget['count']}",
									)
						except Exception as e:
							log.debug(
								f"LINE: popup OCR failed: {e}",
							)

					# Strategy 1.5: Offset-based position matching
					# Use reference offsets from the right-click point
					# to estimate which menu item is the target,
					# verified against actual UIA bounding rects.
					# Only if OCR confirmed the action is in the menu.
					if (
						not targetItem
						and not targetClickPoint
						and menuItems
						and len(menuItems) >= 3
						and actionName in appModRef._MENU_OFFSETS
						and popupOcrResult
						and actionName in popupOcrResult
					):
						dx, dy = appModRef._MENU_OFFSETS[actionName]
						scale = _getDpiScale(hwnd)
						offsetX = clickX + int(dx * scale)
						offsetY = clickY + int(dy * scale)
						log.debug(
							f"LINE: trying offset-based match "
							f"for '{actionName}': click=("
							f"{clickX},{clickY}) + "
							f"offset({dx},{dy})*{scale:.2f} "
							f"= ({offsetX},{offsetY})",
						)
						for item, text in menuItems:
							try:
								iR = item.CurrentBoundingRectangle
								if iR.left <= offsetX <= iR.right and iR.top <= offsetY <= iR.bottom:
									targetItem = item
									log.info(
										f"LINE: matched "
										f"'{actionName}' by "
										f"offset at "
										f"({offsetX},"
										f"{offsetY}) "
										f"within rect "
										f"({iR.left},"
										f"{iR.top})-"
										f"({iR.right},"
										f"{iR.bottom})",
									)
									break
							except Exception:
								pass

					if targetItem or targetClickPoint:
						# Click the target item
						if targetClickPoint:
							itemCx, itemCy = targetClickPoint
						else:
							iRect = targetItem.CurrentBoundingRectangle
							itemCx = int(
								(iRect.left + iRect.right) / 2,
							)
							itemCy = int(
								(iRect.top + iRect.bottom) / 2,
							)
						log.info(
							f"LINE: clicking menu item '{actionName}' at ({itemCx}, {itemCy})",
						)
						ctypes.windll.user32.SetCursorPos(
							itemCx,
							itemCy,
						)
						time.sleep(0.05)
						ctypes.windll.user32.mouse_event(
							0x0002,
							0,
							0,
							0,
							0,
						)  # LEFTDOWN
						time.sleep(0.05)
						ctypes.windll.user32.mouse_event(
							0x0004,
							0,
							0,
							0,
							0,
						)  # LEFTUP
						log.info(
							f"LINE: context menu '{actionName}' selected",
						)
						ui.message(actionName)
						# Context menu auto-dismisses after
						# clicking a menu item. No Escape needed.
						if afterCallback:
							core.callLater(500, afterCallback)
					else:
						# Wrong menu (target item not found).
						# Dismiss the popup first.
						log.debug(
							f"LINE: wrong menu at [{posLabel}], dismissing",
						)
						from keyboardHandler import (
							KeyboardInputGesture,
						)

						KeyboardInputGesture.fromName(
							"escape",
						).send()
						# Got wrong menu (≤2 items = 全選 or
						# 背景設定). After the dedicated edge/padding
						# probes are exhausted, bail; otherwise try
						# the next position.
						isSelectAll = len(menuItems) <= 2
						if isSelectAll and _hasExhaustedMessageBubbleFallbackProbes(
							posIdx,
							clickPositions,
						):
							log.info(
								f"LINE: wrong menu (≤2) seen again at [{posLabel}], skipping to end",
							)
							core.callLater(
								300,
								lambda: _attemptAtOffset(
									len(clickPositions),
								),
							)
						else:
							log.info(
								f"LINE: '{actionName}' not found at [{posLabel}], trying next position",
							)
							core.callLater(
								300,
								lambda: _attemptAtOffset(
									posIdx + 1,
								),
							)

				except Exception as e:
					log.debug(
						f"LINE: context menu detection failed: {e}",
						exc_info=True,
					)
					try:
						from keyboardHandler import (
							KeyboardInputGesture,
						)

						KeyboardInputGesture.fromName(
							"escape",
						).send()
					except Exception:
						pass

			# Wait for context menu to appear, then find and click
			core.callLater(300, _findAndClickMenuItem)

		_attemptAtOffset(0)

	@script(
		gesture="kb:NVDA+windows+r",
		description=_("Reply to the current message"),
		category="LINE",
	)
	def script_replyMessage(self, gesture):
		"""Reply to the current message via right-click context menu."""
		self._contextMenuAction(0, "回覆")

	@script(
		gesture="kb:control+c",
		description=_("Copy the current message"),
		category="LINE",
	)
	def script_copyMessage(self, gesture):
		"""Copy the current message via right-click context menu.

		Only activates in message list context.
		In edit fields, passes Control+C through normally.

		NOTE: We must check the actual UIA focused element, not
		NVDA's focus object (api.getFocusObject()).  LINE's Qt6
		framework does not fire UIA focus change events during
		Tab/arrow navigation, so NVDA's internal focus object
		may be stale (e.g. still pointing to a search field
		even though the real focus moved to the message list).
		"""
		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if rawEl:
				ct = rawEl.CurrentControlType
				if ct == 50004:  # Edit control
					gesture.send()
					return
		except Exception:
			pass
		self._contextMenuAction(1, "複製")

	@script(
		gesture="kb:NVDA+windows+k",
		description=_("Save the current message as a file"),
		category="LINE",
	)
	def script_saveAsMessage(self, gesture):
		"""Save the current message via right-click context menu (Save As)."""
		if _suppressAddon:
			return

		def _afterSaveAs():
			self._suppressAddonForFileDialog("Save As selected")

		self._contextMenuAction(0, "另存新檔", afterCallback=_afterSaveAs)

	@script(
		gesture="kb:NVDA+windows+i",
		# Translators: Input help message for the image-description command.
		description=_("使用設定的 AI 服務描述目前的圖片訊息"),
		category="LINE",
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

			prompt = _getEffectiveImagePrompt()
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
		category="LINE",
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
					if self._playVoiceMessageViaOcr(rawEl, hwnd):
						ui.message(_("播放"))
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
		category="LINE",
	)
	def script_recallMessage(self, gesture):
		"""Recall (unsend) the current message via right-click context menu.

		After selecting recall, LINE shows a confirmation dialog.
		We speak a prompt and wait for user to press Y (recall),
		N (cancel), or P (stealth recall, requires Premium).
		"""
		self._contextMenuAction(
			3,
			"收回",
			afterCallback=self._watchForRecallConfirmationDialog,
		)

	def _getRecallConfirmationDialogRect(self, hwnd):
		"""Return the centered screen rect that contains the LINE recall dialog."""
		return self._getCenteredLineDialogRect(hwnd)

	def _getCenteredLineDialogRect(self, hwnd):
		"""Return a centered screen rect that covers LINE's modal dialogs."""
		try:
			winRect = ctypes.wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(winRect))
			winW = int(winRect.right - winRect.left)
			winH = int(winRect.bottom - winRect.top)
			if winW <= 0 or winH <= 0:
				return None

			scale = _getDpiScale(hwnd)
			dialogW = min(int(winW * 0.82), int(520 * scale))
			dialogH = min(int(winH * 0.74), int(430 * scale))
			dialogW = max(dialogW, int(min(winW, 320 * scale)))
			dialogH = max(dialogH, int(min(winH, 180 * scale)))
			dialogLeft = int(winRect.left + (winW - dialogW) / 2)
			dialogTop = int(winRect.top + (winH - dialogH) / 2)
			return (
				dialogLeft,
				dialogTop,
				dialogLeft + dialogW,
				dialogTop + dialogH,
			)
		except Exception as e:
			log.debug(f"LINE: failed to calculate centered dialog rect: {e}", exc_info=True)
			return None

	def _getPhotoTextConsentDialogRect(self, hwnd):
		"""Return the centered screen rect that contains the photo consent dialog."""
		return self._getCenteredLineDialogRect(hwnd)

	def _captureRecallConfirmationState(self):
		"""Capture OCR and UIA state for the currently visible recall dialog."""
		state = {
			"hwnd": None,
			"dialogRect": None,
			"ocrText": "",
			"actionLabels": [],
			"targets": {},
			"isModernDialog": False,
		}
		try:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
			if not hwnd:
				return state

			dialogRect = self._getRecallConfirmationDialogRect(hwnd)
			if not dialogRect:
				return state

			left, top, right, bottom = dialogRect
			ocrResult = self._ocrWindowAreaResult(
				hwnd,
				region=(left, top, right - left, bottom - top),
				sync=True,
				timeout=2.0,
			)
			ocrText = ""
			ocrLines = []
			if ocrResult is not None and not isinstance(ocrResult, Exception):
				ocrText = _removeCJKSpaces(
					(getattr(ocrResult, "text", "") or "").strip(),
				)
				ocrLines = _extractOcrLines(ocrResult)
			actionLabels = _extractRecallDialogActionLabels(ocrText)
			isModernDialog = _isModernRecallDialogText(ocrText, actionLabels)
			ocrActionTargets = _extractRecallDialogActionClickPoints(ocrLines, dialogRect)

			targetMatches = {}
			handler = UIAHandler.handler
			client = getattr(handler, "clientObject", None)
			if client:
				rootEl = client.ElementFromHandle(hwnd)
				if rootEl:
					allElements = self._collectAllElements(rootEl, handler)
					dialogCenterY = (top + bottom) / 2
					filteredElements = []
					geometryCandidates = []

					def _considerTarget(label, element, rectTuple):
						rectWidth = rectTuple[2] - rectTuple[0]
						rectHeight = rectTuple[3] - rectTuple[1]
						if rectWidth <= 0 or rectHeight <= 0:
							return

						hasInvoke = 0
						try:
							hasInvoke = 1 if element.GetCurrentPattern(10000) else 0
						except Exception:
							pass

						score = (
							hasInvoke,
							rectWidth * rectHeight,
							-abs(((rectTuple[1] + rectTuple[3]) / 2) - dialogCenterY),
						)
						current = targetMatches.get(label)
						if current is None or score > current["score"]:
							targetMatches[label] = {
								"score": score,
								"element": element,
								"rect": rectTuple,
							}

					for element in allElements:
						try:
							rect = element.CurrentBoundingRectangle
							rectTuple = (
								int(rect.left),
								int(rect.top),
								int(rect.right),
								int(rect.bottom),
							)
						except Exception:
							continue

						if (
							rectTuple[2] <= rectTuple[0]
							or rectTuple[3] <= rectTuple[1]
							or not _rectsIntersect(rectTuple, dialogRect)
						):
							continue

						filteredElements.append(element)
						controlType = 0
						try:
							controlType = element.CurrentControlType
						except Exception:
							pass
						hasInvoke = False
						try:
							pattern = element.GetCurrentPattern(10000)
							hasInvoke = bool(pattern)
						except Exception:
							pass
						geometryCandidates.append(
							{
								"element": element,
								"rect": rectTuple,
								"controlType": controlType,
								"hasInvoke": hasInvoke,
							},
						)
						candidateTexts = []
						try:
							name = element.CurrentName or ""
							if name:
								candidateTexts.append(name)
						except Exception:
							pass
						try:
							helpText = str(element.GetCurrentPropertyValue(30048) or "")
							if helpText:
								candidateTexts.append(helpText)
						except Exception:
							pass
						try:
							autoId = element.CurrentAutomationId or ""
							if autoId:
								candidateTexts.append(autoId)
						except Exception:
							pass

						for candidateText in candidateTexts:
							label = _matchRecallDialogActionLabel(candidateText)
							if label:
								_considerTarget(label, element, rectTuple)
								break

					for label, includeKeywords, excludeKeywords in (
						("無痕收回", ["無痕收回"], []),
						("收回", ["收回"], ["無痕"]),
						("取消", ["取消", "關閉", "关闭"], []),
					):
						if label in targetMatches or not filteredElements:
							continue
						element = self._findButtonByKeywords(
							filteredElements,
							includeKeywords,
							excludeKeywords,
						)
						if not element:
							continue
						try:
							rect = element.CurrentBoundingRectangle
							rectTuple = (
								int(rect.left),
								int(rect.top),
								int(rect.right),
								int(rect.bottom),
							)
						except Exception:
							continue
						if _rectsIntersect(rectTuple, dialogRect):
							_considerTarget(label, element, rectTuple)

					for label, candidate in _inferRecallDialogTargetsByGeometry(
						geometryCandidates,
						dialogRect,
						actionLabels,
						isModernDialog=isModernDialog,
					).items():
						if label in targetMatches:
							continue
						targetMatches[label] = {
							"score": candidate.get("score", ()),
							"element": candidate["element"],
							"rect": candidate["rect"],
						}

			stateTargets = {}
			for label, target in targetMatches.items():
				stateTargets[label] = {
					"element": target["element"],
					"rect": target["rect"],
					"clickPoint": None,
				}
			for label, ocrTarget in ocrActionTargets.items():
				targetInfo = stateTargets.setdefault(
					label,
					{
						"element": None,
						"rect": None,
						"clickPoint": None,
					},
				)
				if targetInfo.get("rect") is None and ocrTarget.get("rect") is not None:
					targetInfo["rect"] = ocrTarget["rect"]
				if ocrTarget.get("clickPoint") is not None:
					targetInfo["clickPoint"] = ocrTarget["clickPoint"]

			state.update(
				{
					"hwnd": hwnd,
					"dialogRect": dialogRect,
					"ocrText": ocrText,
					"actionLabels": actionLabels,
					"targets": stateTargets,
					"isModernDialog": isModernDialog,
				},
			)
			ocrTargetSummary = {
				label: target["clickPoint"] for label, target in sorted(ocrActionTargets.items())
			}
			log.debug(
				f"LINE: recall dialog state labels={actionLabels}, "
				f"targets={sorted(state['targets'])}, "
				f"ocrTargets={ocrTargetSummary}, "
				f"modern={isModernDialog}",
			)
		except Exception as e:
			log.debug(f"LINE: capture recall confirmation state failed: {e}", exc_info=True)

		return state

	def _refreshRecallConfirmationState(self):
		"""Merge fresh recall-dialog state into the pending confirmation session."""
		state = self._captureRecallConfirmationState()
		storedActions = set(getattr(self, "_recallDialogActions", ()) or ())
		storedTargets = dict(getattr(self, "_recallDialogTargets", {}) or {})
		storedRect = getattr(self, "_recallDialogRect", None)
		storedHwnd = getattr(self, "_recallDialogHwnd", None)
		storedModern = bool(getattr(self, "_recallDialogIsModern", False))

		mergedActions = storedActions | set(state["actionLabels"]) | set(state["targets"])
		if state["targets"]:
			for label, freshTarget in state["targets"].items():
				mergedTarget = dict(storedTargets.get(label, {}) or {})
				for key in ("element", "rect", "clickPoint"):
					value = freshTarget.get(key)
					if value is not None:
						mergedTarget[key] = value
				storedTargets[label] = mergedTarget

		self._recallDialogActions = mergedActions
		self._recallDialogTargets = storedTargets
		self._recallDialogRect = state["dialogRect"] or storedRect
		self._recallDialogHwnd = state["hwnd"] or storedHwnd
		self._recallDialogIsModern = storedModern or state["isModernDialog"]

		return {
			"hwnd": self._recallDialogHwnd,
			"dialogRect": self._recallDialogRect,
			"ocrText": state["ocrText"],
			"actionLabels": list(self._recallDialogActions),
			"targets": self._recallDialogTargets,
			"isModernDialog": self._recallDialogIsModern,
		}

	def _beginRecallConfirmation(self):
		"""Prompt for recall confirmation and bind Y/N/P shortcuts."""
		state = self._refreshRecallConfirmationState()
		prompt = _getRecallConfirmationPrompt(
			state["actionLabels"],
			isModernDialog=state["isModernDialog"],
		)
		if getattr(self, "_recallPending", False):
			ui.message(prompt)
			return

		token = getattr(self, "_recallConfirmationToken", 0) + 1
		self._recallConfirmationToken = token
		self._recallPending = True
		ui.message(prompt)
		self.bindGesture("kb:y", "confirmRecall")
		self.bindGesture("kb:n", "cancelRecall")
		self.bindGesture("kb:p", "stealthRecall")

		def _autoCancel():
			if (
				getattr(self, "_recallPending", False)
				and getattr(self, "_recallConfirmationToken", 0) == token
			):
				self._endRecallConfirmation("取消")

		core.callLater(10000, _autoCancel)

	def _isRecallConfirmationDialogVisible(self):
		"""Return True when the centered LINE recall confirmation dialog is visible."""
		try:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
			if not hwnd:
				return False

			dialogRect = self._getRecallConfirmationDialogRect(hwnd)
			if not dialogRect:
				return False

			dialogLeft, dialogTop, dialogRight, dialogBottom = dialogRect

			ocrText = self._ocrWindowArea(
				hwnd,
				region=(
					dialogLeft,
					dialogTop,
					dialogRight - dialogLeft,
					dialogBottom - dialogTop,
				),
				sync=True,
				timeout=2.0,
			)
			ocrText = _removeCJKSpaces((ocrText or "").strip())
			normalizedLines = [line.replace(" ", "") for line in ocrText.splitlines() if line.strip()]
			normalizedText = "".join(normalizedLines)
			actionLabels = _extractRecallDialogActionLabels(ocrText)
			actionSet = set(actionLabels)
			isModernDialog = _isModernRecallDialogText(ocrText, actionLabels)
			hasButtons = {"收回", "取消"}.issubset(actionSet) or (
				"收回" in normalizedText and any(label in actionSet for label in {"取消", "無痕收回"})
			)
			hasRecallBody = any(
				keyword in normalizedText
				for keyword in (
					"確定要收回訊息嗎",
					"收回訊息",
					"未讀訊息",
					"任何提醒",
					"可能無法",
					"聊天室",
					"對方",
					"line版本",
				)
			)
			looksLikeCompactDialog = len(normalizedLines) <= 4 and {"收回", "取消"}.issubset(actionSet)
			log.debug(
				f"LINE: recall confirmation OCR: text={ocrText!r}, "
				f"normalizedLines={normalizedLines}, actions={actionLabels}, "
				f"modern={isModernDialog}",
			)
			return hasButtons and (hasRecallBody or looksLikeCompactDialog or isModernDialog)
		except Exception as e:
			log.debug(f"LINE: recall confirmation detection failed: {e}", exc_info=True)
			return False

	def _watchForRecallConfirmationDialog(self, retriesLeft=6, delayMs=250):
		"""Poll briefly for the LINE recall confirmation dialog, then start Y/N/P mode."""
		watchId = getattr(self, "_recallDialogWatchId", 0) + 1
		self._recallDialogWatchId = watchId

		def _poll(remaining):
			if watchId != getattr(self, "_recallDialogWatchId", 0):
				return
			if getattr(self, "_recallPending", False):
				return
			try:
				foreground = api.getForegroundObject()
				if not foreground or foreground.appModule.appName != "line":
					return
			except Exception:
				return
			if self._isRecallConfirmationDialogVisible():
				self._beginRecallConfirmation()
				return
			if remaining > 0:
				core.callLater(delayMs, lambda: _poll(remaining - 1))

		core.callLater(delayMs, lambda: _poll(retriesLeft))

	def _handleMessageContextMenuAction(self, actionName):
		"""React to actions chosen from the message context virtual window."""
		if actionName == "收回":
			self._watchForRecallConfirmationDialog()
		elif actionName == "轉為文字":
			self._watchForPhotoTextConsentDialog()

	def _clearRecallConfirmationBindings(self):
		"""Clear the transient recall-confirmation bindings and cached dialog state."""
		self._recallPending = False
		self._recallActionInProgress = False
		self._recallDialogActions = set()
		self._recallDialogTargets = {}
		self._recallDialogRect = None
		self._recallDialogHwnd = None
		self._recallDialogIsModern = False

		# Remove dynamic confirmation gesture bindings.
		import inputCore

		for key in ("kb:y", "kb:n", "kb:p"):
			try:
				normalized = inputCore.normalizeGestureIdentifier(key)
				if normalized in self._gestureMap:
					del self._gestureMap[normalized]
			except Exception:
				pass

	def _performRecallConfirmationAction(self, actionName):
		"""Activate a specific action on the current recall confirmation dialog."""
		state = self._refreshRecallConfirmationState()
		targetInfo = state["targets"].get(actionName)
		isCompactModernDialog = _isCompactModernRecallDialog(
			state.get("actionLabels"),
			isModernDialog=state["isModernDialog"],
		)

		def _clickFallbackPoint():
			fallbackPoint = _getRecallDialogFallbackClickPoint(
				actionName,
				state["dialogRect"],
				isModernDialog=state["isModernDialog"],
				availableActions=state.get("actionLabels"),
			)
			if not fallbackPoint:
				return False
			self._clickAtPosition(
				fallbackPoint[0],
				fallbackPoint[1],
				hwnd=state["hwnd"],
			)
			log.info(
				f"LINE: fallback-clicking recall dialog action '{actionName}' "
				f"at ({fallbackPoint[0]}, {fallbackPoint[1]})",
			)
			return True

		if targetInfo:
			clickPoint = targetInfo.get("clickPoint")
			if clickPoint:
				log.info(
					f"LINE: OCR-clicking recall dialog action '{actionName}' "
					f"at ({clickPoint[0]}, {clickPoint[1]})",
				)
				self._clickAtPosition(
					clickPoint[0],
					clickPoint[1],
					hwnd=state["hwnd"],
				)
				return True
			if isCompactModernDialog and actionName == "收回" and _clickFallbackPoint():
				return True
			element = targetInfo.get("element")
			if element and self._invokeElement(element, actionName, announce=False):
				return True
			rect = targetInfo.get("rect")
			if rect:
				self._clickAtPosition(
					int((rect[0] + rect[2]) / 2),
					int((rect[1] + rect[3]) / 2),
					hwnd=state["hwnd"],
				)
				return True

		if _clickFallbackPoint():
			return True

		if actionName == "取消":
			from keyboardHandler import KeyboardInputGesture

			KeyboardInputGesture.fromName("escape").send()
			return True

		if actionName == "收回":
			ui.message(_("找不到收回按鈕"))
			return False

		if actionName == "無痕收回":
			ui.message(_("找不到無痕收回按鈕，需要 Premium"))
			return False

		return False

	def _scheduleRecallCompletionAnnouncement(self, actionName, token):
		"""Verify the dialog closed before announcing the recall result."""

		def _finish():
			if getattr(self, "_recallConfirmationToken", 0) != token:
				return
			stillVisible = self._isRecallConfirmationDialogVisible()
			self._clearRecallConfirmationBindings()
			if stillVisible:
				if actionName == "無痕收回":
					ui.message(_("無痕收回可能未成功"))
				elif actionName == "收回":
					ui.message(_("收回可能未成功"))
				else:
					ui.message(_("取消可能未成功"))
				return

			if actionName == "無痕收回":
				ui.message(_("已無痕收回"))
			elif actionName == "收回":
				ui.message(_("已收回"))
			else:
				ui.message(_("已取消"))

		core.callLater(650, _finish)

	def _endRecallConfirmation(self, actionName):
		"""End the recall confirmation by activating the requested dialog action."""
		if not getattr(self, "_recallPending", False) or getattr(self, "_recallActionInProgress", False):
			return
		self._recallActionInProgress = True
		if not self._performRecallConfirmationAction(actionName):
			self._recallActionInProgress = False
			return

		token = getattr(self, "_recallConfirmationToken", 0)
		self._scheduleRecallCompletionAnnouncement(actionName, token)

	def script_confirmRecall(self, gesture):
		"""User pressed Y to choose the standard recall action."""
		self._endRecallConfirmation("收回")

	def script_cancelRecall(self, gesture):
		"""User pressed N to cancel message recall."""
		self._endRecallConfirmation("取消")

	def script_stealthRecall(self, gesture):
		"""User pressed P to choose stealth recall (requires Premium)."""
		self._endRecallConfirmation("無痕收回")

	def _capturePhotoTextConsentState(self):
		"""Capture OCR and UIA state for the first-run photo-to-text consent dialog."""
		state = {
			"hwnd": None,
			"dialogRect": None,
			"ocrText": "",
			"actionLabels": [],
			"targets": {},
		}
		try:
			hwnd = ctypes.windll.user32.GetForegroundWindow()
			if not hwnd:
				return state

			dialogRect = self._getPhotoTextConsentDialogRect(hwnd)
			if not dialogRect:
				return state

			left, top, right, bottom = dialogRect
			ocrResult = self._ocrWindowAreaResult(
				hwnd,
				region=(left, top, right - left, bottom - top),
				sync=True,
				timeout=2.0,
			)
			ocrText = ""
			ocrLines = []
			if ocrResult is not None and not isinstance(ocrResult, Exception):
				ocrText = _removeCJKSpaces(
					(getattr(ocrResult, "text", "") or "").strip(),
				)
				ocrLines = _extractOcrLines(ocrResult)
			actionLabels = _extractPhotoTextConsentActionLabels(ocrText)
			ocrActionTargets = _extractPhotoTextConsentActionClickPoints(
				ocrLines,
				dialogRect,
			)

			targetMatches = {}
			handler = UIAHandler.handler
			client = getattr(handler, "clientObject", None)
			if client:
				rootEl = client.ElementFromHandle(hwnd)
				if rootEl:
					allElements = self._collectAllElements(rootEl, handler)
					dialogWidth = right - left
					dialogHeight = bottom - top
					expectedCenterY = top + (dialogHeight * 0.79)
					preferredCenterX = {
						"同意": left + (dialogWidth * 0.38),
						"不同意": left + (dialogWidth * 0.62),
					}
					filteredElements = []

					def _considerTarget(label, element, rectTuple):
						rectWidth = rectTuple[2] - rectTuple[0]
						rectHeight = rectTuple[3] - rectTuple[1]
						if rectWidth <= 0 or rectHeight <= 0:
							return

						hasInvoke = 0
						try:
							hasInvoke = 1 if element.GetCurrentPattern(10000) else 0
						except Exception:
							pass

						centerX = (rectTuple[0] + rectTuple[2]) / 2.0
						centerY = (rectTuple[1] + rectTuple[3]) / 2.0
						score = (
							hasInvoke,
							1 if centerY >= (top + (dialogHeight * 0.55)) else 0,
							-abs(centerX - preferredCenterX[label]),
							-abs(centerY - expectedCenterY),
							rectWidth * rectHeight,
						)
						current = targetMatches.get(label)
						if current is None or score > current["score"]:
							targetMatches[label] = {
								"score": score,
								"element": element,
								"rect": rectTuple,
							}

					for element in allElements:
						try:
							rect = element.CurrentBoundingRectangle
							rectTuple = (
								int(rect.left),
								int(rect.top),
								int(rect.right),
								int(rect.bottom),
							)
						except Exception:
							continue

						if (
							rectTuple[2] <= rectTuple[0]
							or rectTuple[3] <= rectTuple[1]
							or not _rectsIntersect(rectTuple, dialogRect)
						):
							continue

						filteredElements.append(element)
						candidateTexts = []
						try:
							name = element.CurrentName or ""
							if name:
								candidateTexts.append(name)
						except Exception:
							pass
						try:
							helpText = str(element.GetCurrentPropertyValue(30048) or "")
							if helpText:
								candidateTexts.append(helpText)
						except Exception:
							pass
						try:
							autoId = element.CurrentAutomationId or ""
							if autoId:
								candidateTexts.append(autoId)
						except Exception:
							pass

						for candidateText in candidateTexts:
							label = _matchPhotoTextConsentActionLabel(candidateText)
							if label:
								_considerTarget(label, element, rectTuple)
								break

					for label, includeKeywords, excludeKeywords in (
						("同意", ["同意"], ["提供照片", "服務規定"]),
						("不同意", ["不同意"], []),
					):
						if label in targetMatches or not filteredElements:
							continue
						element = self._findButtonByKeywords(
							filteredElements,
							includeKeywords,
							excludeKeywords,
						)
						if not element:
							continue
						try:
							rect = element.CurrentBoundingRectangle
							rectTuple = (
								int(rect.left),
								int(rect.top),
								int(rect.right),
								int(rect.bottom),
							)
						except Exception:
							continue
						if _rectsIntersect(rectTuple, dialogRect):
							_considerTarget(label, element, rectTuple)

			stateTargets = {}
			for label, target in targetMatches.items():
				stateTargets[label] = {
					"element": target["element"],
					"rect": target["rect"],
					"clickPoint": None,
				}
			for label, ocrTarget in ocrActionTargets.items():
				targetInfo = stateTargets.setdefault(
					label,
					{
						"element": None,
						"rect": None,
						"clickPoint": None,
					},
				)
				if targetInfo.get("rect") is None and ocrTarget.get("rect") is not None:
					targetInfo["rect"] = ocrTarget["rect"]
				if ocrTarget.get("clickPoint") is not None:
					targetInfo["clickPoint"] = ocrTarget["clickPoint"]

			state.update(
				{
					"hwnd": hwnd,
					"dialogRect": dialogRect,
					"ocrText": ocrText,
					"actionLabels": actionLabels,
					"targets": stateTargets,
				},
			)
			ocrTargetSummary = {
				label: target["clickPoint"] for label, target in sorted(ocrActionTargets.items())
			}
			log.debug(
				f"LINE: photo consent dialog state labels={actionLabels}, "
				f"targets={sorted(state['targets'])}, "
				f"ocrTargets={ocrTargetSummary}",
			)
		except Exception as e:
			log.debug(f"LINE: capture photo consent dialog state failed: {e}", exc_info=True)

		return state

	def _refreshPhotoTextConsentState(self):
		"""Merge fresh photo-consent dialog state into the pending session."""
		state = self._capturePhotoTextConsentState()
		storedActions = set(getattr(self, "_photoTextConsentDialogActions", ()) or ())
		storedTargets = dict(getattr(self, "_photoTextConsentDialogTargets", {}) or {})
		storedRect = getattr(self, "_photoTextConsentDialogRect", None)
		storedHwnd = getattr(self, "_photoTextConsentDialogHwnd", None)

		mergedActions = storedActions | set(state["actionLabels"]) | set(state["targets"])
		if state["targets"]:
			for label, freshTarget in state["targets"].items():
				mergedTarget = dict(storedTargets.get(label, {}) or {})
				for key in ("element", "rect", "clickPoint"):
					value = freshTarget.get(key)
					if value is not None:
						mergedTarget[key] = value
				storedTargets[label] = mergedTarget

		self._photoTextConsentDialogActions = mergedActions
		self._photoTextConsentDialogTargets = storedTargets
		self._photoTextConsentDialogRect = state["dialogRect"] or storedRect
		self._photoTextConsentDialogHwnd = state["hwnd"] or storedHwnd

		return {
			"hwnd": self._photoTextConsentDialogHwnd,
			"dialogRect": self._photoTextConsentDialogRect,
			"ocrText": state["ocrText"],
			"actionLabels": list(self._photoTextConsentDialogActions),
			"targets": self._photoTextConsentDialogTargets,
		}

	def _beginPhotoTextConsent(self):
		"""Prompt for photo upload consent and bind A/D shortcuts."""
		self._refreshPhotoTextConsentState()
		prompt = _getPhotoTextConsentPrompt()
		if getattr(self, "_photoTextConsentPending", False):
			ui.message(prompt)
			return

		token = getattr(self, "_photoTextConsentToken", 0) + 1
		self._photoTextConsentToken = token
		self._photoTextConsentPending = True
		ui.message(prompt)
		self.bindGesture("kb:a", "acceptPhotoTextConsent")
		self.bindGesture("kb:d", "declinePhotoTextConsent")

		def _autoDecline():
			if (
				getattr(self, "_photoTextConsentPending", False)
				and getattr(self, "_photoTextConsentToken", 0) == token
			):
				self._endPhotoTextConsent("不同意")

		core.callLater(10000, _autoDecline)

	def _isPhotoTextConsentDialogVisible(self):
		"""Return True when LINE's first-run photo consent dialog is visible."""
		try:
			state = self._capturePhotoTextConsentState()
			if not state["dialogRect"]:
				return False

			actionLabels = set(state["actionLabels"]) | set(state["targets"])
			isVisible = _isPhotoTextConsentDialogText(
				state["ocrText"],
				actionLabels,
			)
			log.debug(
				f"LINE: photo consent OCR: text={state['ocrText']!r}, "
				f"actions={sorted(actionLabels)}, visible={isVisible}",
			)
			return isVisible
		except Exception as e:
			log.debug(f"LINE: photo consent detection failed: {e}", exc_info=True)
			return False

	def _watchForPhotoTextConsentDialog(self, retriesLeft=8, delayMs=250):
		"""Poll briefly for LINE's first-run photo consent dialog, then start A/D mode."""
		watchId = getattr(self, "_photoTextConsentWatchId", 0) + 1
		self._photoTextConsentWatchId = watchId

		def _poll(remaining):
			if watchId != getattr(self, "_photoTextConsentWatchId", 0):
				return
			if getattr(self, "_photoTextConsentPending", False):
				return
			try:
				foreground = api.getForegroundObject()
				if not foreground or foreground.appModule.appName != "line":
					return
			except Exception:
				return
			if self._isPhotoTextConsentDialogVisible():
				self._beginPhotoTextConsent()
				return
			if remaining > 0:
				core.callLater(delayMs, lambda: _poll(remaining - 1))

		core.callLater(delayMs, lambda: _poll(retriesLeft))

	def _clearPhotoTextConsentBindings(self):
		"""Clear the transient photo-consent bindings and cached dialog state."""
		self._photoTextConsentPending = False
		self._photoTextConsentActionInProgress = False
		self._photoTextConsentDialogActions = set()
		self._photoTextConsentDialogTargets = {}
		self._photoTextConsentDialogRect = None
		self._photoTextConsentDialogHwnd = None

		import inputCore

		for key in ("kb:a", "kb:d"):
			try:
				normalized = inputCore.normalizeGestureIdentifier(key)
				if normalized in self._gestureMap:
					del self._gestureMap[normalized]
			except Exception:
				pass

	def _performPhotoTextConsentAction(self, actionName):
		"""Activate a specific action on the current photo consent dialog."""
		state = self._refreshPhotoTextConsentState()
		targetInfo = state["targets"].get(actionName)

		def _clickFallbackPoint():
			fallbackPoint = _getPhotoTextConsentDialogFallbackClickPoint(
				actionName,
				state["dialogRect"],
			)
			if not fallbackPoint:
				return False
			self._clickAtPosition(
				fallbackPoint[0],
				fallbackPoint[1],
				hwnd=state["hwnd"],
			)
			log.info(
				f"LINE: fallback-clicking photo consent action '{actionName}' "
				f"at ({fallbackPoint[0]}, {fallbackPoint[1]})",
			)
			return True

		if targetInfo:
			clickPoint = targetInfo.get("clickPoint")
			if clickPoint:
				log.info(
					f"LINE: OCR-clicking photo consent action '{actionName}' "
					f"at ({clickPoint[0]}, {clickPoint[1]})",
				)
				self._clickAtPosition(
					clickPoint[0],
					clickPoint[1],
					hwnd=state["hwnd"],
				)
				return True
			element = targetInfo.get("element")
			if element and self._invokeElement(element, actionName, announce=False):
				return True
			rect = targetInfo.get("rect")
			if rect:
				self._clickAtPosition(
					int((rect[0] + rect[2]) / 2),
					int((rect[1] + rect[3]) / 2),
					hwnd=state["hwnd"],
				)
				return True

		if _clickFallbackPoint():
			return True

		if actionName == "不同意":
			from keyboardHandler import KeyboardInputGesture

			KeyboardInputGesture.fromName("escape").send()
			return True

		if actionName == "同意":
			ui.message(_("找不到同意按鈕"))
			return False

		ui.message(_("找不到不同意按鈕"))
		return False

	def _schedulePhotoTextConsentCompletionAnnouncement(self, actionName, token):
		"""Verify the dialog closed before announcing the photo consent result."""

		def _finish():
			if getattr(self, "_photoTextConsentToken", 0) != token:
				return
			stillVisible = self._isPhotoTextConsentDialogVisible()
			self._clearPhotoTextConsentBindings()
			if stillVisible:
				if actionName == "同意":
					ui.message(_("同意提供照片可能未成功"))
				else:
					ui.message(_("不同意提供照片可能未成功"))
				return

			if actionName == "同意":
				ui.message(_("已同意提供照片"))
			else:
				ui.message(_("已不同意提供照片"))

		core.callLater(650, _finish)

	def _endPhotoTextConsent(self, actionName):
		"""End the photo consent flow by activating the requested dialog action."""
		if not getattr(self, "_photoTextConsentPending", False) or getattr(
			self,
			"_photoTextConsentActionInProgress",
			False,
		):
			return
		self._photoTextConsentActionInProgress = True
		if not self._performPhotoTextConsentAction(actionName):
			self._photoTextConsentActionInProgress = False
			return

		token = getattr(self, "_photoTextConsentToken", 0)
		self._schedulePhotoTextConsentCompletionAnnouncement(actionName, token)

	def script_acceptPhotoTextConsent(self, gesture):
		"""User pressed A to agree to LINE's first-run photo upload notice."""
		self._endPhotoTextConsent("同意")

	def script_declinePhotoTextConsent(self, gesture):
		"""User pressed D to decline LINE's first-run photo upload notice."""
		self._endPhotoTextConsent("不同意")

	@script(
		gesture="kb:applications",
		description=_("Open message context menu"),
		category="LINE",
	)
	def script_messageContextMenu(self, gesture):
		"""Right-click current message and open a virtual window for browsing the context menu."""
		global _messageContextMenuRequestId
		if _suppressAddon:
			gesture.send()
			return

		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if rawEl:
				ct = rawEl.CurrentControlType
				if ct == 50004:  # Edit control
					gesture.send()
					return
		except Exception:
			pass

		try:
			handler = UIAHandler.handler
			rawEl = handler.clientObject.GetFocusedElement()
			if not rawEl:
				ui.message(_("找不到目前的訊息"))
				return

			rect = rawEl.CurrentBoundingRectangle
			cx = int((rect.left + rect.right) / 2)
			cy = int((rect.top + rect.bottom) / 2)
			elLeft = int(rect.left)
			elRight = int(rect.right)
			elTop = int(rect.top)
			elBottom = int(rect.bottom)

			if cx <= 0 or cy <= 0:
				ui.message(_("找不到目前的訊息"))
				return

			hwnd = ctypes.windll.user32.GetForegroundWindow()

			import ctypes.wintypes as wintypes

			winRect = wintypes.RECT()
			ctypes.windll.user32.GetWindowRect(
				hwnd,
				ctypes.byref(winRect),
			)
			winTop = int(winRect.top)
			winBottom = int(winRect.bottom)
		except Exception as e:
			log.debug(f"LINE: messageContextMenu getElement failed: {e}")
			ui.message(_("找不到目前的訊息"))
			return

		_messageContextMenuRequestId += 1
		requestId = _messageContextMenuRequestId
		targetRuntimeId = _getElementRuntimeId(rawEl)

		def _isCurrentRequest():
			if requestId != _messageContextMenuRequestId:
				return False
			try:
				if ctypes.windll.user32.GetForegroundWindow() != hwnd:
					return False
			except Exception:
				pass
			if targetRuntimeId is None:
				return True
			currentRuntimeId = _getFocusedElementRuntimeId()
			return currentRuntimeId is None or currentRuntimeId == targetRuntimeId

		def _logAndAbortIfStale(stage):
			if _isCurrentRequest():
				return False
			if requestId == _messageContextMenuRequestId and _shouldDismissCopyReadMenu(hwnd):
				_restoreFocusToElement(rawEl, targetRuntimeId)
			log.debug(
				f"LINE: messageContextMenu abandoning stale request during "
				f"{stage}; requestId={requestId}, currentRequestId={_messageContextMenuRequestId}, "
				f"targetRuntimeId={targetRuntimeId}, currentRuntimeId={_getFocusedElementRuntimeId()}",
			)
			return True

		ocrLines = []
		try:
			ocrResult = self._ocrWindowAreaResult(
				hwnd,
				region=(
					elLeft,
					elTop,
					max(elRight - elLeft, 0),
					max(elBottom - elTop, 0),
				),
				sync=True,
				timeout=2.0,
			)
			if ocrResult is not None and not isinstance(ocrResult, Exception):
				ocrLines = _extractOcrLines(ocrResult)
		except Exception as e:
			log.debug(
				f"LINE: failed to derive OCR bubble probes for message context menu: {e}",
				exc_info=True,
			)
		ocrClickPositions = _buildMessageBubbleOcrClickPositions(
			ocrLines,
			(elLeft, elTop, elRight, elBottom),
			winTop,
			winBottom,
		)
		if ocrClickPositions:
			log.debug(
				"LINE: message context menu OCR-derived bubble probes: "
				f"{[label for _x, _y, label in ocrClickPositions]}",
			)
		elif ocrLines:
			log.debug(
				"LINE: message context menu OCR lines yielded no bubble probes: "
				f"{[(line.get('text', ''), bool(line.get('rect'))) for line in ocrLines]}",
			)
		clickPositions = _mergeClickPositions(
			ocrClickPositions,
			_buildMessageBubbleClickPositions(
				(elLeft, elTop, elRight, elBottom),
				winTop,
				winBottom,
				includeVerticalOffsets=True,
			),
		)

		VK_CONTROL = 0x11
		VK_SHIFT = 0x10
		VK_MENU = 0x12
		GetAsyncKeyState = ctypes.windll.user32.GetAsyncKeyState
		for _wait in range(40):
			held = any(GetAsyncKeyState(vk) & 0x8000 for vk in (VK_CONTROL, VK_SHIFT, VK_MENU))
			if not held:
				break
			time.sleep(0.05)

		appModRef = self
		keyboardFallbackTried = [False]

		def _startMouseContextMenuProbes(reason):
			if _logAndAbortIfStale(f"mouseContextMenuFallback:{reason}"):
				return
			log.info(
				f"LINE: keyboard message context menu failed ({reason}); trying mouse probes",
			)
			core.callLater(300, lambda: _attemptAtOffset(0))

		def _tryKeyboardFallback():
			if _logAndAbortIfStale("keyboardFallback"):
				return
			if keyboardFallbackTried[0]:
				ui.message(_("找不到訊息選單"))
				return
			keyboardFallbackTried[0] = True
			log.info(
				"LINE: trying keyboard message context menu before mouse probes",
			)
			try:
				_sendGestureWithAddonSuppressed("applications")
			except Exception as e:
				log.debug(
					f"LINE: keyboard fallback applications send failed: {e}",
					exc_info=True,
				)
				_startMouseContextMenuProbes("applications send failed")
				return

			def _activateKeyboardFallback():
				if _logAndAbortIfStale("keyboardFallbackActivate"):
					return
				self._activateMessageContextMenu(
					onAction=self._handleMessageContextMenuAction,
					onFailure=lambda: _startMouseContextMenuProbes(
						"keyboard popup not found",
					),
					shouldAbort=lambda: _logAndAbortIfStale(
						"keyboardFallbackPoll",
					),
				)

			core.callLater(500, _activateKeyboardFallback)

		def _attemptAtOffset(posIdx=0):
			if _logAndAbortIfStale(f"attemptAtOffset[{posIdx}]"):
				return
			if posIdx >= len(clickPositions):
				if keyboardFallbackTried[0]:
					ui.message(_("找不到訊息選單"))
				else:
					_tryKeyboardFallback()
				return

			clickX, clickY, posLabel = clickPositions[posIdx]
			log.info(
				f"LINE: right-clicking message at ({clickX}, {clickY}) [{posLabel}] for context menu",
			)
			if not _messageProbePointHitsTargetElement(
				UIAHandler.handler,
				rawEl,
				clickX,
				clickY,
				targetRuntimeId,
			):
				log.info(
					f"LINE: skipping context menu probe over edit field at ({clickX}, {clickY}) [{posLabel}]",
				)
				core.callLater(50, lambda: _attemptAtOffset(posIdx + 1))
				return
			appModRef._rightClickAtPosition(clickX, clickY, hwnd)

			def _findPopupAndActivate(retriesLeft=4):
				if _logAndAbortIfStale(
					f"findPopupAndActivate[{posIdx}] retriesLeft={retriesLeft}",
				):
					return
				try:
					import ctypes.wintypes as wintypes

					pid = wintypes.DWORD()
					tid = ctypes.windll.user32.GetWindowThreadProcessId(
						hwnd,
						ctypes.byref(pid),
					)
					popupCandidates = []

					WNDENUMPROC = ctypes.WINFUNCTYPE(
						ctypes.c_bool,
						wintypes.HWND,
						wintypes.LPARAM,
					)

					def _enumCallback(enumHwnd, lParam):
						if enumHwnd != hwnd and ctypes.windll.user32.IsWindowVisible(enumHwnd):
							wRect = wintypes.RECT()
							ctypes.windll.user32.GetWindowRect(
								enumHwnd,
								ctypes.byref(wRect),
							)
							w = wRect.right - wRect.left
							h = wRect.bottom - wRect.top
							if w >= 50 and h >= 30:
								popupCandidates.append(enumHwnd)
						return True

					ctypes.windll.user32.EnumThreadWindows(
						tid,
						WNDENUMPROC(_enumCallback),
						0,
					)

					popupHwnd = None
					if popupCandidates:
						popupHwnd = popupCandidates[0]
					else:
						for dy in [0, -40, -80, 40, 80]:
							pt = wintypes.POINT(clickX, clickY + dy)
							candidateHwnd = ctypes.windll.user32.WindowFromPoint(pt)
							if candidateHwnd and candidateHwnd != hwnd:
								popupHwnd = candidateHwnd
								break

					if not popupHwnd:
						if retriesLeft > 0:
							core.callLater(
								200,
								lambda: _findPopupAndActivate(retriesLeft - 1),
							)
							return
						core.callLater(
							300,
							lambda: _attemptAtOffset(posIdx + 1),
						)
						return

					uiaHandler = UIAHandler.handler
					element = uiaHandler.clientObject.ElementFromHandle(
						popupHwnd,
					)
					if not element:
						if retriesLeft > 0:
							core.callLater(
								200,
								lambda: _findPopupAndActivate(retriesLeft - 1),
							)
							return
						core.callLater(
							300,
							lambda: _attemptAtOffset(posIdx + 1),
						)
						return

					try:
						ct = element.CurrentControlType
						eRect = element.CurrentBoundingRectangle
						eW = int(eRect.right - eRect.left)
						eH = int(eRect.bottom - eRect.top)
						if ct == 50033 or eW < 50 or eH < 30:
							core.callLater(
								300,
								lambda: _attemptAtOffset(posIdx + 1),
							)
							return
					except Exception:
						pass

					# Check menu has enough items (≥3) to be a real context menu
					walker = uiaHandler.clientObject.RawViewWalker
					itemCount = 0
					child = walker.GetFirstChildElement(element)
					idx = 0
					while child and idx < 30:
						try:
							childRect = child.CurrentBoundingRectangle
							childH = int(childRect.bottom - childRect.top)
							childW = int(childRect.right - childRect.left)
							if 20 <= childH <= 80 and childW >= childH * 2:
								itemCount += 1
						except Exception:
							pass
						try:
							child = walker.GetNextSiblingElement(child)
						except Exception:
							break
						idx += 1

					if itemCount < 3:
						log.debug(
							f"LINE: context menu has only {itemCount} UIA items; "
							f"validating popup via OCR before dismissing",
						)

					eRect = element.CurrentBoundingRectangle
					popupRect = (
						int(eRect.left),
						int(eRect.top),
						int(eRect.right),
						int(eRect.bottom),
					)
					if itemCount < 3 and not _popupLooksLikeMessageContextMenu(
						appModRef,
						hwnd,
						popupRect,
					):
						# Wrong menu (e.g. 全選), dismiss and try next.
						log.debug(
							f"LINE: popup OCR did not confirm a message context menu "
							f"at {popupRect}; dismissing and trying next position",
						)
						_sendGestureWithAddonSuppressed("escape")
						core.callLater(
							300,
							lambda: _attemptAtOffset(posIdx + 1),
						)
						return

					popupRowRects = _collectPopupMenuRowRects(
						popupHwnd,
						popupRect,
					)
					log.info(
						f"LINE: message context menu popup found at {popupRect}",
					)

					from ._virtualWindows.messageContextMenu import MessageContextMenu

					if _logAndAbortIfStale(
						f"beforeVirtualWindow[{posIdx}]",
					):
						try:
							_sendGestureWithAddonSuppressed("escape")
						except Exception:
							pass
						return
					VirtualWindow.currentWindow = MessageContextMenu(
						popupRect,
						rowRects=popupRowRects,
						onAction=self._handleMessageContextMenuAction,
					)

				except Exception as e:
					log.debug(
						f"LINE: message context menu detection failed: {e}",
						exc_info=True,
					)
					try:
						_sendGestureWithAddonSuppressed("escape")
					except Exception:
						pass

			core.callLater(300, _findPopupAndActivate)

		_tryKeyboardFallback()

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

			if any(kw in ocrText for kw in ["關麥克風", "關閉麥克風", "Mute", "mute"]):
				wx.CallAfter(ui.message, "麥克風已開啟")
				return True
			elif any(kw in ocrText for kw in ["開麥克風", "開啟麥克風", "Unmute", "unmute"]):
				wx.CallAfter(ui.message, "麥克風已關閉")
				return True
			else:
				log.debug(f"LINE: mic status not detected in OCR: {ocrText!r}")
				return False

		t = threading.Thread(target=_checkMicStatus, daemon=True)
		t.start()

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
				kw in ocrText for kw in ["關鏡頭", "關相機", "關閉相機", "Turn off camera", "turn off camera"]
			):
				wx.CallAfter(ui.message, "相機已開啟")
				return True
			# "開鏡頭"/"開相機"/"開啟相機" means tooltip says "turn on" → camera is OFF
			elif any(
				kw in ocrText for kw in ["開鏡頭", "開相機", "開啟相機", "Turn on camera", "turn on camera"]
			):
				wx.CallAfter(ui.message, "相機已關閉")
				return True
			else:
				log.debug(f"LINE: camera status not detected in OCR: {ocrText!r}")
				return False

		t = threading.Thread(target=_checkCameraStatus, daemon=True)
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
		"kb:tab": "navigateAndTrack",
		"kb:shift+tab": "navigateAndTrack",
		"kb:upArrow": "chatListArrow",
		"kb:downArrow": "chatListArrow",
		"kb:leftArrow": "navigateAndTrack",
		"kb:rightArrow": "navigateAndTrack",
		"kb:control+o": "openFileDialog",
		"kb:control+1": "switchTabAndAnnounce",
		"kb:control+2": "switchTabAndAnnounce",
		"kb:control+3": "switchTabAndAnnounce",
		"kb:enter": "sendMessageAndPlaySound",
		"kb:control+shift+a": "toggleMicAndAnnounce",
		"kb:control+shift+v": "toggleCameraAndAnnounce",
		"kb:shift+f10": "messageContextMenu",
		"kb:applications": "messageContextMenu",
	}
