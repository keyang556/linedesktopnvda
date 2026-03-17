from ._utils import message

from logHandler import log
import api
from inputCore import decide_executeGesture, InputGesture
import mouseHandler
import winUser

import pkgutil
from importlib import import_module

class VirtualWindow:
	'''
	VirtualWindow is a base class for creating virtual windows for different screens in the Line App. It allows users to navigate and interact with elements on the screen using keyboard gestures.
	To create a virtual window for a specific screen, subclass VirtualWindow and implement the following methods:
	- isMatchLineScreen(cls, obj): A class method that determines if the current Line App screen matches the virtual window. It should return True if it matches, False otherwise.
	- makeElements(self): An instance method that populates the elements list based on the current Line App screen. Each element should be a dictionary with at least 'name' and optionally 'role' and 'clickPoint' keys.
	The virtual window will be activated when the user focuses on a screen that matches the virtual window's criteria. Once activated, users can use the following keyboard gestures to navigate and interact with the elements:
	- Previous Element: kb:up or kb:shift+tab
	- Next Element: kb:down or kb:tab
	- Click Element: kb:enter or kb:space
	'''
	title = None
	
	windowClasses = tuple()
	currentWindow = None
	
	@classmethod
	def initialize(cls):
		'''
		Initializes the VirtualWindow system by dynamically importing all virtual window classes and registering the gesture handler.
		Should be called once during the initialization of the Line AppModule.
		'''
		# Dynamically import all virtual window classes from the virtualWindows package.
		assert __package__
		pkg = import_module(__package__ + '._virtualWindows')
		for importer, modname, ispkg in pkgutil.iter_modules(pkg.__path__):
			import_module(f'{__package__}._virtualWindows.{modname}')
		
		cls.windowClasses = tuple(VirtualWindow.__subclasses__())
		
		decide_executeGesture.register(cls.handleGesture)
	
	@classmethod
	def handleGesture(cls, gesture: InputGesture):
		'''
		This method is called by inputCore.decide_executeGesture. 
		This method must return True, Because it is only responsible for processing gestures when a virtual window is active, and it should not block any other gesture processing.
		'''
		import core
		core.callLater(1, cls.processKey, gesture)
		return True
	
	@classmethod
	def processKey(cls, gesture: InputGesture):
		'''
		Processes a keyboard gesture for the current virtual window.
		
		'''
		if not cls.currentWindow:
			return
		
		foreground = api.getForegroundObject()
		if foreground.appModule.appName != 'line':
			return
		
		# Process the gesture for the current virtual window.
		previousKeys = {'kb:uparrow', 'kb:shift+tab'}
		nextKeys = {'kb:downarrow', 'kb:tab'}
		clickKeys = {'kb:enter', 'kb:space'}
		ids = gesture.normalizedIdentifiers
		if previousKeys.intersection(ids):
			cls.currentWindow.previous()
		elif nextKeys.intersection(ids):
			cls.currentWindow.next()
		elif clickKeys.intersection(ids):
			cls.currentWindow.click()
		
		return
	
	@classmethod
	def onFocusChanged(cls, obj):
		'''
		Called when the focus changes in the Line App. It checks if the new focused screen matches any virtual window and activates it if it does.
		'''
		window = cls.getWindowClass(obj)
		if getattr(cls.currentWindow, '__class__', None) is window:
			return
		
		cls.currentWindow = window(obj) if window else None
	
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
		self.makeElements()
		message(self.title) if self.title else None
	
	def makeElements(self):
		'''
		This method should be overridden by subclasses to populate the elements list based on the current Line App screen.
		
		elements should be a list of dictionaries with at least 'name' and optionally 'role' and 'clickPoint' keys, for example:
		[
			{'name': 'Button 1', 'role': roleObject, 'clickPoint': (x, y)},
			{'name': 'Button 2', 'role': roleObject, 'clickPoint': (x, y)},
			...
		]
		'''
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
		
		role = element.get('role')
		roleName = role.displayString if role else ''
		displayText = f'{element["name"]}' + (f' ({roleName})' if roleName else '')
		message(displayText)
	
	@property
	def element(self):
		if not self.elements:
			return None
		
		return self.elements[self.pos]
	
	def click(self):
		'''
		Simulates a click on the current element by executing mouse events at the element's click point.
		'''
		element = self.element
		if not element or not element.get('clickPoint'):
			return
		
		originalPos = winUser.getCursorPos()
		winUser.setCursorPos(*element.get('clickPoint'))
		mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_LEFTDOWN, 0, 0)
		mouseHandler.executeMouseEvent(winUser.MOUSEEVENTF_LEFTUP, 0, 0)
		winUser.setCursorPos(*originalPos)
	
