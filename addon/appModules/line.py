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

	def script_navigateAndTrack(self, gesture):
		"""Pass navigation key to LINE, then poll UIA focused element.
		
		LINE's Qt6 framework does not fire UIA focus change events when
		navigating with Tab/arrows. This script sends the key through,
		waits briefly for LINE to process it, then queries the UIA
		focused element directly and announces it.
		"""
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
	}
