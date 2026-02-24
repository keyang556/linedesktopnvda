# LINE Desktop App Module for NVDA
# Provides accessibility enhancements for the LINE desktop application.
# LINE desktop uses Qt6 framework, which exposes UI via UIA on Windows.

import appModuleHandler
from scriptHandler import script
import controlTypes
import api
import ui
import textInfos
import speech
import braille
import core
from logHandler import log
from NVDAObjects.UIA import UIA
from NVDAObjects.IAccessible import IAccessible
from NVDAObjects import NVDAObject
import UIAHandler
import ctypes
import comtypes
import re

# Regex pattern to remove spurious spaces between CJK characters.
# Windows OCR inserts spaces between every CJK character.
_CJK_SPACE_RE = re.compile(
	r'(?<=[\u2E80-\u9FFF\uF900-\uFAFF\uFE30-\uFE4F])\ '
	r'(?=[\u2E80-\u9FFF\uF900-\uFAFF\uFE30-\uFE4F\u3000-\u303F\uFF00-\uFFEF])',
)

def _removeCJKSpaces(text):
	"""Remove spaces between CJK characters inserted by Windows OCR.

	'可 能 因 為' → '可能因為'
	Spaces between Latin characters are preserved.
	"""
	if not text:
		return text
	return _CJK_SPACE_RE.sub('', text)

# Global variable to track the last focused object
# This is needed because api.getFocusObject() sometimes returns the main Window
# even when we handled a gainFocus event for a ListItem.
lastFocusedObject = None

# Track the last UIA element we announced, to avoid re-announcing the same thing
_lastAnnouncedUIAElement = None
_lastAnnouncedUIAName = None

# Flag to suppress addon while a file dialog is open
_suppressAddon = False


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



def _getTextViaUIAFindAll(obj, maxElements=30):
	"""Use raw UIA FindAll to get text from descendants.

	Qt6 elements often report childCount=0 to NVDA but DO have
	UIA descendants accessible via FindAll. This method bypasses
	NVDA's child enumeration and queries UIA directly.
	"""
	texts = []
	if not hasattr(obj, 'UIAElement') or obj.UIAElement is None:
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
			condition
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
			if not hasattr(appMod, '_getBindingHandle'):
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


def _getDeepText(obj, maxDepth=3, _depth=0):
	"""Recursively collect non-empty text from an object and its children.

	Falls back to _getTextViaUIAFindAll if childCount is 0.
	"""
	if _depth > maxDepth or obj is None:
		return []
	texts = []
	# Get this object's name
	try:
		name = obj.name
		if name and name.strip():
			texts.append(name.strip())
	except Exception:
		pass
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
			controlTypes.Role.LIST, controlTypes.Role.LISTITEM,
			controlTypes.Role.GROUPING, controlTypes.Role.SECTION,
			controlTypes.Role.TREEVIEWITEM, controlTypes.Role.PANE,
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


def _ocrReadElementText(rawElement, appModuleRef=None):
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

		if width <= 0 or height <= 0:
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
				height * resizeFactor
			)
			ocrPixels = sb2.captureImage(
				left, top,
				width, height
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
					ocrText = getattr(result, 'text', '') or ''
					ocrText = _removeCJKSpaces(ocrText.strip())
					if ocrText:
						log.info(f"LINE OCR nav result: {ocrText!r}")
						speech.cancelSpeech()
						ui.message(ocrText)
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
							30003, 50007  # ControlType == ListItem
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
							30003, 50007  # ControlType == ListItem
						)
						items = parent.FindAll(UIAHandler.TreeScope_Children, condition)
						if items:
							for i in range(items.Length):
								item = items.GetElement(i)
								try:
									isSelected = item.GetCurrentPropertyValue(30079)
									if isSelected:
										log.info("LINE: found selected list item via SelectionItemPattern.IsSelected")
										return item
								except Exception:
									pass
					except Exception:
						pass
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
	global _lastAnnouncedUIAElement, _lastAnnouncedUIAName
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
		
		# If focus is stuck on the same element (e.g. edit field),
		# try to find the selected item in a nearby list
		if elementId and elementId == _lastAnnouncedUIAElement:
			selectedItem = _findSelectedItemInList(handler, rawElement)
			if selectedItem:
				targetElement = selectedItem
				try:
					runtimeId = selectedItem.GetRuntimeId()
					elementId = str(runtimeId) if runtimeId else None
				except Exception:
					elementId = None
				if elementId and elementId == _lastAnnouncedUIAElement:
					return
			else:
				return
		
		_lastAnnouncedUIAElement = elementId
		
		# Extract text using safe read-only COM properties only
		textParts = _extractTextFromUIAElement(targetElement)
		
		# Get control type for role name
		try:
			ct = targetElement.CurrentControlType
		except Exception:
			ct = 0
		
		controlTypeNames = {
			50000: "按鈕", 50004: "編輯", 50005: "超連結",
			50007: "清單項目", 50008: "清單", 50011: "項目",
			50016: "索引標籤項目", 50018: "文字", 50025: "群組",
			50033: "窗格",
		}
		roleName = controlTypeNames.get(ct, "")
		
		log.info(
			f"LINE UIA focus: ct={ct}, texts={textParts}, "
			f"runtimeId={elementId}"
		)
		
		if textParts:
			announcement = " ".join(textParts)
			if roleName:
				announcement = f"{announcement} {roleName}"
			_lastAnnouncedUIAName = announcement
			speech.cancelSpeech()
			ui.message(announcement)
		elif roleName:
			# UIA text is empty — announce role immediately, then try OCR
			_lastAnnouncedUIAName = roleName
			speech.cancelSpeech()
			ui.message(roleName)
			# Kick off async OCR to read the actual text content
			_ocrReadElementText(targetElement)
		else:
			return
		
	except Exception:
		log.debugWarning("_queryAndSpeakUIAFocus failed", exc_info=True)


class LineChatListItem(UIA):
	"""Overlay class for chat/contact list items in the sidebar.

	Qt6 list items typically have empty name AND childCount=0.
	We use UIA FindAll and display model as fallbacks.
	"""

	def _get_name(self):
		# First try the native name
		name = super().name
		if name and name.strip():
			return name
		# Try deep text (includes UIA FindAll fallback)
		texts = _getDeepText(self, maxDepth=4)
		if texts:
			return " - ".join(texts)
		# Last resort: read from display
		displayText = _getTextFromDisplay(self)
		if displayText:
			return displayText
		return ""

	def event_gainFocus(self):
		super().event_gainFocus()


class LineChatMessage(UIA):
	"""Overlay class for individual chat messages."""

	def _get_name(self):
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
				"Error in LineMessageInput._get_name", exc_info=True
			)
			name = ""
		if not name or not name.strip():
			# Translators: Label for the LINE message input field
			return "Message input"
		return name


class LineContactItem(UIA):
	"""Overlay class for contact list items."""

	def _get_name(self):
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


class LineGenericList(UIA):
	"""Overlay class for list containers in LINE."""

	def _get_positionInfo(self):
		try:
			info = super().positionInfo
		except Exception:
			log.debugWarning(
				"Error in LineGenericList._get_positionInfo", exc_info=True
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
		log.info(
			f"LINE AppModule loaded for process: {self.processID}, "
			f"exe: {self.appName}"
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
				if any(keyword in automationId.lower() for keyword in
					   ("chat", "room", "talk", "conversation", "friend",
						"contact", "message", "buddy")):
					clsList.insert(0, LineChatListItem)
				elif any(keyword in (parent.name or "").lower() for keyword in
						 ("chat", "聊天", "トーク", "好友", "友だち", "friend",
						  "contact", "message", "訊息", "メッセージ")):
					clsList.insert(0, LineChatListItem)
				else:
					# Default: treat any list item as potentially a chat item
					clsList.insert(0, LineChatListItem)
			log.debug(
				f"LINE listitem: name={obj.name!r}, "
				f"automationId={automationId!r}, children={obj.childCount}"
			)

		# --- Message input area ---
		elif role in (
			controlTypes.Role.EDITABLETEXT, controlTypes.Role.DOCUMENT
		):
			if automationId and any(
				kw in automationId.lower() for kw in (
					"input", "compose", "message", "send",
					"chat", "editor", "textbox", "edit",
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
			controlTypes.Role.GROUPING, controlTypes.Role.SECTION,
			controlTypes.Role.PARAGRAPH, controlTypes.Role.STATICTEXT,
		):
			if automationId and any(
				kw in automationId.lower() for kw in (
					"message", "bubble", "chat_content", "msg",
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
				f"LINE UIA_elementSelected: role={obj.role}, "
				f"name={obj.name!r}, class={obj.windowClassName}"
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
				f"LINE UIA_notification: role={obj.role}, "
				f"name={obj.name!r}, kwargs={kwargs}"
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
						f"name={obj.name!r}, class={obj.windowClassName}"
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
				f"LINE nameChange: role={obj.role}, "
				f"name={obj.name!r}, class={obj.windowClassName}"
			)
		except Exception:
			pass
		nextHandler()

	@script(
		description="Debug: log UIA tree info for the focused element",
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
				info.append(f"BoundingRect: left={rect.left}, top={rect.top}, right={rect.right}, bottom={rect.bottom}")
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
				import NVDAHelper
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
									height * resizeFactor
								)
								ocrPixels = sb2.captureImage(
									left, top,
									width, height
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
											ocrText = getattr(result, 'text', '') or ''
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
				UIAHandler.TreeScope_Descendants, trueCondition
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
						f"LINE elem: ct={ctType}, autoId={autoId!r}, name={name!r}"
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
							f"name={name!r}, help={helpText!r}, class={className!r}"
						)
						return el
			except Exception:
				continue
		
		return None
	
	def _invokeElement(self, element, actionName):
		"""Invoke a UIA element using InvokePattern or mouse click fallback."""
		import ctypes
		
		# Try InvokePattern
		try:
			invokePattern = element.GetCurrentPattern(10000)  # InvokePattern
			if invokePattern:
				invokePattern.QueryInterface(
					comtypes.gen.UIAutomationClient.IUIAutomationInvokePattern
				).Invoke()
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
				ui.message(actionName)
				return True
		except Exception as e:
			log.debug(f"LINE: click fallback failed: {e}")
		
		return False
	
	def _clickAtPosition(self, x, y):
		"""Perform a mouse click at the given screen coordinates."""
		import ctypes
		import time
		
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
		
		# Get complete window rect
		rect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
		winLeft = rect.left
		winTop = rect.top
		winRight = rect.right
		winWidth = winRight - winLeft
		
		log.info(
			f"LINE: window rect=({winLeft},{winTop},{winRight},{rect.bottom}), "
			f"width={winWidth}, dpiScale={scale:.2f}"
		)
		
		# Reference values at 96 DPI (100% scaling), scaled by DPI factor.
		# Back-calculated from working values at 150%: 83/1.5≈55, 40/1.5≈27, 23/1.5≈15
		iconY = winTop + int(55 * scale)
		iconSpacing = int(27 * scale)
		firstIconOffset = int(15 * scale)
		iconX = winRight - firstIconOffset - (2 * iconSpacing)
		
		log.info(
			f"LINE: header icon pos: iconX={iconX}, iconY={iconY}, "
			f"spacing={iconSpacing}, offset={firstIconOffset}"
		)
		
		# Verify position is within window bounds
		if iconX < winLeft or iconX > winRight:
			log.warning(f"LINE: icon position {iconX} outside window bounds")
			return None
		
		return (iconX, iconY, winRight)
	
	def _makeCallByType(self, callType):
		"""Click phone icon, wait for popup menu, then click voice or video.
		
		Full flow (3 steps):
		  1. Click phone icon → popup menu appears (wait 500ms)
		  2. Click voice/video menu item → confirmation dialog appears (wait 800ms)
		  3. OCR the confirmation dialog, announce it, auto-click "開始"
		
		From the screenshot, clicking the phone icon shows a popup menu:
		  - 語音通話 (voice call): 1st item
		  - 視訊通話 (video call): 2nd item
		
		The confirmation dialog is centered on the window with:
		  - Text: "確定要與XXX進行語音通話？"
		  - "開始" button (green, left) and "取消" button (gray, right)
		"""
		import ctypes
		import ctypes.wintypes
		
		pos = self._getHeaderIconPosition()
		if not pos:
			return False
		
		phoneX, phoneY, winRight = pos
		
		# Get full window rect for dialog position calculation
		hwnd = ctypes.windll.user32.GetForegroundWindow()
		winRect = ctypes.wintypes.RECT()
		ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(winRect))
		winLeft = winRect.left
		winTop = winRect.top
		winW = winRect.right - winRect.left
		winH = winRect.bottom - winRect.top
		
		log.info(f"LINE: clicking phone icon at ({phoneX}, {phoneY})")
		self._clickAtPosition(phoneX, phoneY)
		
		# Menu item positions — DPI-scaled from 96 DPI reference values.
		# Reference: menuX offset=40, voice menuY offset=23, video menuY offset=50
		scale = _getDpiScale()
		menuX = winRight - int(40 * scale)
		if callType == "voice":
			menuY = phoneY + int(23 * scale)
		else:
			menuY = phoneY + int(50 * scale)
		
		log.info(
			f"LINE: will click menu item '{callType}' at ({menuX}, {menuY})"
		)
		
		appModRef = self
		
		def _clickMenuItem():
			try:
				appModRef._clickAtPosition(menuX, menuY)
				log.info("LINE: menu item clicked, waiting for confirmation dialog...")
				# Step 3: Handle confirmation dialog after 800ms
				core.callLater(800, _handleConfirmDialog)
			except Exception as e:
				log.warning(f"LINE: menu click failed: {e}")
				ui.message("選單點擊失敗")
		
		def _handleConfirmDialog():
			"""OCR the confirmation dialog, announce it, and auto-click 開始."""
			try:
				# The confirmation dialog is centered on the LINE window
				# Dialog dimensions — DPI-scaled from 96 DPI reference.
				# Reference: 420/1.5≈280, 140/1.5≈93
				cScale = _getDpiScale()
				dialogW = int(280 * cScale)
				dialogH = int(93 * cScale)
				winCenterX = winLeft + winW // 2
				winCenterY = winTop + winH // 2
				dialogLeft = winCenterX - dialogW // 2
				dialogTop = winCenterY - dialogH // 2
				
				log.info(
					f"LINE: OCR confirmation dialog area: "
					f"({dialogLeft},{dialogTop}) {dialogW}x{dialogH}"
				)
				
				# OCR the dialog area
				import screenBitmap
				sb = screenBitmap.ScreenBitmap(dialogW, dialogH)
				pixels = sb.captureImage(
					dialogLeft, dialogTop, dialogW, dialogH
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
					log.warning("LINE: no OCR language for confirmation dialog")
					_clickStart()
					return
				
				recognizer = uwpOcr.UwpOcr(language=ocrLang)
				resizeFactor = recognizer.getResizeFactor(dialogW, dialogH)
				
				if resizeFactor > 1:
					sb2 = screenBitmap.ScreenBitmap(
						dialogW * resizeFactor,
						dialogH * resizeFactor
					)
					ocrPixels = sb2.captureImage(
						dialogLeft, dialogTop,
						dialogW, dialogH
					)
				else:
					ocrPixels = pixels
				
				# Keep references alive during async OCR
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
						return self._screenLeft + int(x / self.resizeFactor)
					def convertYToScreen(self, y):
						return self._screenTop + int(y / self.resizeFactor)
					def convertWidthToScreen(self, width):
						return int(width / self.resizeFactor)
					def convertHeightToScreen(self, height):
						return int(height / self.resizeFactor)
				
				imgInfo = _ImgInfo(
					dialogW, dialogH, resizeFactor, dialogLeft, dialogTop
				)
				appModRef._callOcrImgInfo = imgInfo
				
				def _onOcrResult(result):
					import wx
					def _handleOnMain():
						try:
							ocrText = ""
							if not isinstance(result, Exception):
								ocrText = getattr(result, 'text', '') or ''
								ocrText = _removeCJKSpaces(ocrText.strip())
							
							if ocrText:
								log.info(f"LINE: confirmation dialog OCR: {ocrText}")
								ui.message(ocrText)
							else:
								log.info("LINE: confirmation dialog OCR empty")
								if callType == "voice":
									ui.message("語音通話確認")
								else:
									ui.message("視訊通話確認")
							
							# Auto-click "開始" after a short delay for speech
							core.callLater(300, _clickStart)
						except Exception as e:
							log.warning(
								f"LINE: dialog OCR handler error: {e}",
								exc_info=True
							)
							_clickStart()
						finally:
							appModRef._callOcrRecognizer = None
							appModRef._callOcrPixels = None
							appModRef._callOcrImgInfo = None
					wx.CallAfter(_handleOnMain)
				
				recognizer.recognize(ocrPixels, imgInfo, _onOcrResult)
				
			except Exception as e:
				log.warning(f"LINE: dialog handling error: {e}", exc_info=True)
				# Even if OCR fails, try to click "開始"
				_clickStart()
		
		def _clickStart():
			"""Click the 開始 (Start) button on the confirmation dialog."""
			try:
				# "開始" button position — DPI-scaled from 96 DPI reference.
				# Reference: xOffset=65/1.5≈43, yOffset=25/1.5≈17
				sScale = _getDpiScale()
				winCenterX = winLeft + winW // 2
				winCenterY = winTop + winH // 2
				startBtnX = winCenterX - int(43 * sScale)
				startBtnY = winCenterY + int(17 * sScale)
				
				log.info(f"LINE: clicking 開始 at ({startBtnX}, {startBtnY})")
				appModRef._clickAtPosition(startBtnX, startBtnY)
				
				if callType == "voice":
					ui.message("已開始語音通話")
				else:
					ui.message("已開始視訊通話")
			except Exception as e:
				log.warning(f"LINE: click 開始 failed: {e}", exc_info=True)
				ui.message("無法點擊開始按鈕")
		
		# Step 1: Wait 500ms for popup menu, then click menu item
		core.callLater(500, _clickMenuItem)
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
			"來電", "通話", "linecall", "call", "ringing",
			"着信", "สาย",
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
					PROCESS_QUERY_LIMITED_INFORMATION, False, pid
				)
				if not hProc:
					return ""
				try:
					buf = ctypes.create_unicode_buffer(260)
					size = ctypes.wintypes.DWORD(260)
					ok = ctypes.windll.kernel32.QueryFullProcessImageNameW(
						hProc, 0, buf, ctypes.byref(size)
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
				hwnd, ctypes.byref(pid)
			)
			allWindows.append((hwnd, title, pid.value))
			return True
		
		ctypes.windll.user32.EnumWindows(WNDENUMPROC(_enumAll), 0)
		
		log.debug(
			f"LINE: _findIncomingCallWindow scanning {len(allWindows)} "
			f"visible windows, mainHwnd={mainHwnd}, linePID={lineProcessId}"
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
							f"LINE: found call window (same process) "
							f"hwnd={hwnd}, title={title!r}, pid={pid}"
						)
						callHwnd = hwnd
						break
					# Check if it's a LINE child process
					exeName = _getExeName(pid)
					if exeName in _LINE_EXES:
						log.info(
							f"LINE: found call window (child process) "
							f"hwnd={hwnd}, title={title!r}, pid={pid}, "
							f"exe={exeName}"
						)
						callHwnd = hwnd
						break
					else:
						log.debug(
							f"LINE: title matched but exe mismatch: "
							f"hwnd={hwnd}, title={title!r}, exe={exeName}"
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
					fgHwnd, ctypes.byref(fgPid)
				)
				if fgPid.value == lineProcessId:
					skipHwnds.add(fgHwnd)
			
			candidateHwnds = []
			for hwnd, title, pid in allWindows:
				if hwnd in skipHwnds:
					continue
				# Check both same-process and child-process windows
				isLine = (pid == lineProcessId)
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
				f"LINE: OCR fallback has {len(candidateHwnds)} candidates"
			)
			
			for hwnd, rect in candidateHwnds:
				try:
					ocrText = self._ocrWindowArea(
						hwnd, sync=True, timeout=2.0
					)
					if ocrText:
						ocrLower = ocrText.lower()
						checkRegion = (
							ocrLower[:150] if len(ocrLower) > 200
							else ocrLower
						)
						for kw in _CALL_KEYWORDS:
							if kw.lower() in checkRegion:
								log.info(
									f"LINE: found call window via OCR "
									f"hwnd={hwnd}, text={ocrText!r}"
								)
								callHwnd = hwnd
								break
					if callHwnd:
						break
				except Exception as e:
					log.debug(
						f"LINE: OCR probe on hwnd={hwnd} failed: {e}"
					)
		
		if not callHwnd:
			log.debug("LINE: no incoming call window found")
		
		return callHwnd
	
	def _ocrWindowArea(self, hwnd, region=None, sync=False, timeout=3.0):
		"""OCR a window (or part of it) and return the recognized text.
		
		Args:
			hwnd: Window handle to capture.
			region: Optional (left, top, width, height) tuple in screen
				coordinates.  If None, uses the full window rect.
			sync: If True, block until OCR completes (up to timeout).
			timeout: Max seconds to wait when sync=True.
		
		Returns:
			The OCR text string, or empty string on failure.
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
			return ""
		
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
				return ""
			
			recognizer = uwpOcr.UwpOcr(language=ocrLang)
			resizeFactor = recognizer.getResizeFactor(width, height)
			
			if resizeFactor > 1:
				sb2 = screenBitmap.ScreenBitmap(
					width * resizeFactor, height * resizeFactor
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
					return ""
				text = getattr(result, 'text', '') or ''
				return _removeCJKSpaces(text.strip())
			else:
				# Async — not used for incoming call detection
				return ""
		except Exception as e:
			log.debug(f"LINE: _ocrWindowArea failed: {e}", exc_info=True)
			return ""
	
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
			f"LINE: call window rect=({rect.left},{rect.top},"
			f"{rect.right},{rect.bottom}), size={winW}x{winH}"
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
						elRectStr = (
							f"({elRect.left},{elRect.top},"
							f"{elRect.right},{elRect.bottom})"
						)
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
						f"rect={elRectStr}, invoke={hasInvoke}"
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
				if (elCX < rect.left or elCX > rect.right or
						elCY < rect.top or elCY > rect.bottom):
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
				f"LINE: filtered out {outsideCount} border elements "
				f"outside window rect"
			)
		
		if not candidates:
			log.info("LINE: no button candidates INSIDE window")
			return None
		
		log.info(
			f"LINE: {len(candidates)} button candidates found, "
			f"looking for '{side}' button"
		)
		
		if side == "right":
			rightCandidates = [
				c for c in candidates if c[2] > winCenterX
			]
			if rightCandidates:
				rightCandidates.sort(key=lambda c: c[2], reverse=True)
				best = rightCandidates[0]
				log.info(
					f"LINE: selected right button at "
					f"({best[2]},{best[3]}), size={best[4]}x{best[5]}"
				)
				return (best[0], best[2], best[3])
		else:
			leftCandidates = [
				c for c in candidates if c[2] < winCenterX
			]
			if leftCandidates:
				leftCandidates.sort(key=lambda c: c[2])
				best = leftCandidates[0]
				log.info(
					f"LINE: selected left button at "
					f"({best[2]},{best[3]}), size={best[4]}x{best[5]}"
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
				f"LINE: fallback button at "
				f"({best[2]},{best[3]}), size={best[4]}x{best[5]}"
			)
			return (best[0], best[2], best[3])
		
		return None
	
	def _ocrFindButtonKeyword(self, hwnd, keywords):
		"""Use OCR to check if any keyword appears in the call window.
		
		Returns True if a keyword is found, False otherwise.
		Note: NVDA's uwpOcr result only provides flat text,
		not per-word positions, so we just confirm presence.
		"""
		try:
			ocrText = self._ocrWindowArea(hwnd, sync=True, timeout=3.0)
			if not ocrText:
				log.info("LINE: OCR returned no text for call window")
				return False
			
			ocrTextLower = ocrText.lower()
			log.info(
				f"LINE: OCR call window text: '{ocrText}'"
			)
			for kw in keywords:
				if kw.lower() in ocrTextLower:
					log.info(f"LINE: OCR found keyword '{kw}'")
					return True
			
			log.info("LINE: OCR no keyword match")
			return False
		except Exception as e:
			log.debug(
				f"LINE: _ocrFindButtonKeyword failed: {e}",
				exc_info=True
			)
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
		
		allElements, handler, rootEl = self._getCallButtonElements(hwnd)
		
		# Strategy 1: UIA keyword search
		if allElements:
			answerEl = self._findButtonByKeywords(
				allElements,
				["接聽", "accept", "answer", "応答", "รับสาย", "接受"],
				excludeKeywords=["拒絕", "decline", "reject"]
			)
			if answerEl:
				if self._invokeElement(answerEl, "已接聽"):
					return True
		
		# Strategy 2: UIA bounding-rect analysis (inside window only)
		# In LINE's call popup, answer (green) is on the LEFT
		if allElements:
			result = self._findCallButtonByRect(
				hwnd, allElements, side="left"
			)
			if result:
				el, cx, cy = result
				invoked = False
				try:
					pat = el.GetCurrentPattern(10000)
					if pat:
						import comtypes
						pat.QueryInterface(
							comtypes.gen.UIAutomationClient
							.IUIAutomationInvokePattern
						).Invoke()
						invoked = True
						log.info("LINE: answered via InvokePattern")
				except Exception as e:
					log.debug(f"LINE: InvokePattern failed: {e}")
				
				if not invoked:
					log.info(
						f"LINE: clicking answer button at ({cx}, {cy})"
					)
					self._clickAtPosition(cx, cy)
				
				ui.message("已接聽")
				return True
		
		# Strategy 3: OCR confirms call window, then click at position
		log.info("LINE: trying OCR to confirm call window")
		ocrConfirmed = self._ocrFindButtonKeyword(
			hwnd,
			["接聽", "accept", "answer", "応答", "รับสาย",
			 "拒絕", "decline", "reject", "來電"]
		)
		
		# Strategy 4: Click at position inside the window
		# Screenshot layout: [avatar][caller text][red reject ~80%][green answer ~92%]
		# Answer (green phone) is the RIGHTMOST button
		if winH > 200:
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
			f"({answerX}, {answerY}), winRect=({rect.left},"
			f"{rect.top},{rect.right},{rect.bottom})"
		)
		self._clickAtPosition(answerX, answerY)
		ui.message("已接聽")
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
		
		allElements, handler, rootEl = self._getCallButtonElements(hwnd)
		
		# Strategy 1: UIA keyword search
		if allElements:
			rejectEl = self._findButtonByKeywords(
				allElements,
				["拒絕", "decline", "reject", "拒否", "ปฏิเสธ"],
				excludeKeywords=["接聽", "accept", "answer"]
			)
			if rejectEl:
				if self._invokeElement(rejectEl, "已拒絕"):
					return True
		
		# Strategy 2: UIA bounding-rect analysis (inside window only)
		# In LINE's call popup, reject (red) is on the RIGHT
		if allElements:
			result = self._findCallButtonByRect(
				hwnd, allElements, side="right"
			)
			if result:
				el, cx, cy = result
				invoked = False
				try:
					pat = el.GetCurrentPattern(10000)
					if pat:
						import comtypes
						pat.QueryInterface(
							comtypes.gen.UIAutomationClient
							.IUIAutomationInvokePattern
						).Invoke()
						invoked = True
						log.info("LINE: rejected via InvokePattern")
				except Exception as e:
					log.debug(f"LINE: InvokePattern failed: {e}")
				
				if not invoked:
					log.info(
						f"LINE: clicking reject button at ({cx}, {cy})"
					)
					self._clickAtPosition(cx, cy)
				
				ui.message("已拒絕")
				return True
		
		# Strategy 3: OCR confirms call window, then click at position
		log.info("LINE: trying OCR to confirm call window")
		ocrConfirmed = self._ocrFindButtonKeyword(
			hwnd,
			["拒絕", "decline", "reject", "拒否", "ปฏิเสธ",
			 "接聽", "accept", "answer", "來電"]
		)
		
		# Strategy 4: Click at position inside the window
		# Screenshot layout: [avatar][caller text][red reject ~80%][green answer ~92%]
		# Reject (red phone) is second from right
		if winH > 200:
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
			f"({rejectX}, {rejectY}), winRect=({rect.left},"
			f"{rect.top},{rect.right},{rect.bottom})"
		)
		self._clickAtPosition(rejectX, rejectY)
		ui.message("已拒絕")
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
			for removeKw in ["來電", "着信", "ringing", "incoming call",
							 "สายเรียกเข้า"]:
				callerName = callerName.replace(removeKw, "")
			callerName = callerName.strip()
			if callerName:
				ui.message(f"來電：{callerName}")
			else:
				ui.message(f"來電（OCR: {ocrText}）")
			log.info(f"LINE: caller info OCR: {ocrText!r}")
		else:
			ui.message("無法辨識來電者")
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
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(f"接聽功能錯誤: {e}")
	
	def script_rejectCall(self, gesture):
		"""Reject an incoming LINE call."""
		try:
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._rejectIncomingCall(hwnd)
			else:
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(f"拒絕功能錯誤: {e}")
	
	def script_checkCaller(self, gesture):
		"""Announce who is calling."""
		try:
			hwnd = self._findIncomingCallWindow()
			if hwnd:
				self._getCallerInfo(hwnd)
			else:
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(f"來電查看功能錯誤: {e}")

	def script_focusCallWindow(self, gesture):
		"""Find the LineCall window, bring it to foreground, and OCR its content."""
		import ctypes
		import ctypes.wintypes

		hwnd = self._findIncomingCallWindow()
		if not hwnd:
			ui.message("未偵測到通話視窗")
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
					ui.message("通話視窗（無法辨識內容）")
			except Exception as e:
				log.warning(f"LINE: call window OCR error: {e}", exc_info=True)
				ui.message("通話視窗")

		core.callLater(300, _announceCallWindow)

	# ── Outgoing call scripts ──────────────────────────────────────────

	@script(
		description="撥打語音通話",
		gesture="kb:NVDA+shift+c",
		category="LINE Desktop",
	)
	def script_makeCall(self, gesture):
		"""Click the phone icon, then auto-select voice call from the popup menu."""
		try:
			if not self._makeCallByType("voice"):
				ui.message("找不到 LINE 視窗，請先開啟聊天室")
		except Exception as e:
			log.warning(f"LINE makeCall error: {e}", exc_info=True)
			ui.message(f"通話功能錯誤: {e}")
	
	@script(
		description="撥打視訊通話",
		gesture="kb:NVDA+shift+v",
		category="LINE Desktop",
	)
	def script_makeVideoCall(self, gesture):
		"""Click the phone icon, then auto-select video call from the popup menu."""
		try:
			if not self._makeCallByType("video"):
				ui.message("找不到 LINE 視窗，請先開啟聊天室")
		except Exception as e:
			log.warning(f"LINE makeVideoCall error: {e}", exc_info=True)
			ui.message(f"視訊通話功能錯誤: {e}")

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
					hwnd, buf, 256
				)
				if buf.value == "#32770":
					# Check if this dialog belongs to LINE's process
					pid = ctypes.wintypes.DWORD()
					ctypes.windll.user32.GetWindowThreadProcessId(
						hwnd, ctypes.byref(pid)
					)
					if pid.value == lineProcessId:
						foundOurDialog = True
						return False  # stop enumeration
				return True  # continue enumeration

			ctypes.windll.user32.EnumWindows(
				WNDENUMPROC(_enumCallback), 0
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

	def script_openFileDialog(self, gesture):
		"""Pass Ctrl+O to LINE, suppress addon while file dialog is open."""
		global _suppressAddon
		_suppressAddon = True
		log.info("LINE: Ctrl+O pressed, addon suppressed, waiting for file dialog...")
		gesture.send()
		# Start polling for file dialog to close after a short delay
		core.callLater(1000, self._pollFileDialog)

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
		global _lastAnnouncedUIAElement
		# Reset tracking so we always announce after navigation
		_lastAnnouncedUIAElement = None
		# Pass the gesture through to LINE
		gesture.send()
		# After a delay, query the UIA focused element
		core.callLater(100, _queryAndSpeakUIAFocus)

	__gestures = {
		"kb:tab": "navigateAndTrack",
		"kb:shift+tab": "navigateAndTrack",
		"kb:upArrow": "navigateAndTrack",
		"kb:downArrow": "navigateAndTrack",
		"kb:leftArrow": "navigateAndTrack",
		"kb:rightArrow": "navigateAndTrack",
		"kb:control+o": "openFileDialog",
	}
