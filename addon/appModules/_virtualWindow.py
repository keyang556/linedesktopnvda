from ._utils import message

import api
from inputCore import decide_executeGesture, InputGesture
import mouseHandler
import winUser

import ctypes
import pkgutil
from importlib import import_module

# Key sets used for gesture matching
_PREVIOUS_KEYS = {"kb:uparrow", "kb:shift+tab"}
_NEXT_KEYS = {"kb:downarrow", "kb:tab"}
_CLICK_KEYS = {"kb:enter", "kb:space"}
_ESCAPE_KEYS = {"kb:escape"}
_HANDLED_KEYS = _PREVIOUS_KEYS | _NEXT_KEYS | _CLICK_KEYS | _ESCAPE_KEYS


class VirtualWindow:
	"""
	VirtualWindow is a base class for creating virtual windows for different screens in the Line App. It allows users to navigate and interact with elements on the screen using keyboard gestures.
	To create a virtual window for a specific screen, subclass VirtualWindow and implement the following methods:
	- isMatchLineScreen(cls, obj): A class method that determines if the current Line App screen matches the virtual window. It should return True if it matches, False otherwise.
	- makeElements(self): An instance method that populates the elements list based on the current Line App screen. Each element should be a dictionary with at least 'name' and optionally 'role' and 'clickPoint' keys.
	The virtual window will be activated when the user focuses on a screen that matches the virtual window's criteria. Once activated, users can use the following keyboard gestures to navigate and interact with the elements:
	- Previous Element: kb:up or kb:shift+tab
	- Next Element: kb:down or kb:tab
	- Click Element: kb:enter or kb:space
	- Dismiss: kb:escape
	"""

	title = None

	windowClasses = tuple()
	currentWindow = None
	_initialized = False
	# Foreground window handle captured when the virtual window opened;
	# used by isStillValid() to notice the popup closing behind our back.
	_popupHwnd = None

	@classmethod
	def initialize(cls):
		"""
		Initializes the VirtualWindow system by dynamically importing all virtual window classes and registering the gesture handler.
		Should be called once during the initialization of the Line AppModule.
		"""
		if cls._initialized:
			# AppModule.__init__ runs again every time LINE starts; the
			# decide_executeGesture handler must only be registered once.
			return
		# Dynamically import all virtual window classes from the virtualWindows package.
		assert __package__
		pkg = import_module(__package__ + "._virtualWindows")
		for importer, modname, ispkg in pkgutil.iter_modules(pkg.__path__):
			import_module(f"{__package__}._virtualWindows.{modname}")

		# Deduplicate by class name: after a plugin reload, classes from the
		# previous import can linger in __subclasses__() until GC runs.
		seen = set()
		classes = []
		for windowClass in VirtualWindow.__subclasses__():
			if windowClass.__name__ in seen:
				continue
			seen.add(windowClass.__name__)
			classes.append(windowClass)
		cls.windowClasses = tuple(classes)

		decide_executeGesture.register(cls.handleGesture)
		cls._initialized = True

	@classmethod
	def handleGesture(cls, gesture: InputGesture):
		"""
		This method is called by inputCore.decide_executeGesture.
		When a virtual window is active, navigation and action keys are
		consumed (return False) so that the AppModule scripts do not
		also process them. Other keys pass through normally.
		"""
		if not cls.currentWindow:
			return True

		try:
			foreground = api.getForegroundObject()
			if foreground.appModule.appName != "line":
				return True
		except Exception:
			return True

		window = cls.currentWindow
		try:
			stillValid = window.isStillValid()
		except Exception:
			stillValid = False
		if not stillValid:
			# The underlying LINE popup is gone (e.g. closed with the mouse);
			# stop hijacking the navigation keys.
			if cls.currentWindow is window:
				cls.currentWindow = None
			return True

		ids = gesture.normalizedIdentifiers
		if _HANDLED_KEYS.intersection(ids):
			import core

			core.callLater(1, cls.processKey, gesture)
			return False

		return True

	@classmethod
	def processKey(cls, gesture: InputGesture):
		"""
		Processes a keyboard gesture for the current virtual window.

		"""
		if not cls.currentWindow:
			return

		try:
			foreground = api.getForegroundObject()
			if foreground.appModule.appName != "line":
				return
		except Exception:
			return

		# Process the gesture for the current virtual window.
		ids = gesture.normalizedIdentifiers
		if _PREVIOUS_KEYS.intersection(ids):
			cls.currentWindow.previous()
		elif _NEXT_KEYS.intersection(ids):
			cls.currentWindow.next()
		elif _CLICK_KEYS.intersection(ids):
			cls.currentWindow.click()
		elif _ESCAPE_KEYS.intersection(ids):
			cls.currentWindow.dismiss()

		return

	@classmethod
	def onFocusChanged(cls, obj):
		"""
		Called when the focus changes in the Line App. It checks if the new focused screen matches any virtual window and activates it if it does.
		"""
		window = cls.getWindowClass(obj)
		if getattr(cls.currentWindow, "__class__", None) is window:
			return

		if window is None:
			cls.currentWindow = None
			return
		try:
			cls.currentWindow = window(obj)
		except Exception:
			from logHandler import log

			log.debugWarning(
				f"VirtualWindow {window.__name__} failed to initialize",
				exc_info=True,
			)
			cls.currentWindow = None

	@classmethod
	def getWindowClass(cls, obj):
		for windowClass in cls.windowClasses:
			if windowClass.isMatchLineScreen(obj):
				return windowClass

	@staticmethod
	def isMatchLineScreen(obj):
		# This method should be overridden by subclasses to determine if current Line App screen is matches the virtual window.
		raise NotImplementedError()

	def __init__(self, obj):
		self.obj = obj
		self.elements = []
		self.pos = -1
		self.captureForegroundHwnd()
		self.makeElements()
		message(self.title) if self.title else None

	def captureForegroundHwnd(self):
		"""Record the current foreground window for later liveness checks."""
		try:
			self._popupHwnd = ctypes.windll.user32.GetForegroundWindow()
		except Exception:
			self._popupHwnd = None

	def _isPopupForegroundAlive(self):
		"""True while the window that was foreground at creation still is."""
		hwnd = self._popupHwnd
		if not hwnd:
			return True
		try:
			user32 = ctypes.windll.user32
			return bool(
				user32.IsWindow(hwnd)
				and user32.IsWindowVisible(hwnd)
				and user32.GetForegroundWindow() == hwnd,
			)
		except Exception:
			return True

	def isStillValid(self):
		"""Return False once the matched LINE screen/popup no longer exists.

		handleGesture calls this before consuming navigation keys so a popup
		closed behind our back (e.g. with the mouse) releases the keyboard.
		"""
		if not self._isPopupForegroundAlive():
			return False
		try:
			return bool(type(self).isMatchLineScreen(self.obj))
		except Exception:
			return False

	def makeElements(self):
		"""
		This method should be overridden by subclasses to populate the elements list based on the current Line App screen.

		elements should be a list of dictionaries with at least 'name' and optionally 'role' and 'clickPoint' keys, for example:
		[
			{'name': 'Button 1', 'role': roleObject, 'clickPoint': (x, y)},
			{'name': 'Button 2', 'role': roleObject, 'clickPoint': (x, y)},
			...
		]
		"""
		raise NotImplementedError()

	def rectGetCenterPoint(self, rect):
		return rect.left + (rect.width // 2), rect.top + (rect.height // 2)

	def previous(self):
		if not self.elements:
			return

		if self.pos > 0:
			self.pos -= 1
		else:
			self.pos = len(self.elements) - 1

		self.show()

	def next(self):
		if not self.elements:
			return

		if self.pos < len(self.elements) - 1:
			self.pos += 1
		else:
			self.pos = 0

		self.show()

	def show(self):
		element = self.element
		if not element:
			return

		role = element.get("role")
		roleName = role.displayString if role else ""
		displayText = f"{element['name']}" + (f" ({roleName})" if roleName else "")
		message(displayText)

	@property
	def element(self):
		if not self.elements:
			return None

		return self.elements[self.pos]

	def click(self):
		"""
		Simulates a click on the current element by executing mouse events at the element's click point.
		"""
		element = self.element
		if not element or not element.get("clickPoint"):
			return

		originalPos = winUser.getCursorPos()
		winUser.setCursorPos(*element.get("clickPoint"))
		mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_LEFTDOWN, 0, 0)
		mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_LEFTUP, 0, 0)
		winUser.setCursorPos(*originalPos)

	def dismiss(self):
		"""
		Dismisses the virtual window and sends Escape to close any popup.
		Subclasses can override this for custom dismiss behavior.
		"""
		VirtualWindow.currentWindow = None
		from keyboardHandler import KeyboardInputGesture

		KeyboardInputGesture.fromName("escape").send()
