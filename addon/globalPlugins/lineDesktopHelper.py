# LINE Desktop Global Plugin for NVDA
# Maps alternative executable names to the LINE appModule.
# Also adds a "LINE Desktop" submenu under NVDA's Tools menu.

import appModuleHandler
import globalPluginHandler
from scriptHandler import script
from logHandler import log
import gui
import wx
import addonHandler

addonHandler.initTranslation()


def _getLineAppModule():
	"""Find and return the LINE appModule instance, or None."""
	for app in appModuleHandler.runningTable.values():
		if app and getattr(app, 'appName', '').lower() in (
			'line', 'line_app', 'linecall',
		):
			return app
	return None


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Global plugin to ensure the LINE appModule is loaded
	for all known LINE desktop executable variants.

	LINE desktop may run as:
	- LINE.exe (standard installer)
	- LINE_APP.exe (Microsoft Store version, older)
	- LineLauncher.exe (launcher process)
	"""

	# Alternative executable names that should use the line appModule.
	# NVDA lowercases exe names, so "LINE.exe" becomes "line" automatically.
	# We register additional variants here for safety.
	_LINE_EXECUTABLES = [
		"LINE",
		"Line",
		"LINE_APP",
		"LineCall",
		"linecall",
	]

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		for exe in self._LINE_EXECUTABLES:
			try:
				appModuleHandler.registerExecutableWithAppModule(exe, "line")
				log.debug(f"Registered {exe} with line appModule")
			except Exception:
				log.debugWarning(
					f"Failed to register {exe} with LINE appModule",
					exc_info=True,
				)
		self._toolsMenu = None
		self._lineSubMenu = None
		self._createToolsMenu()

	def _createToolsMenu(self):
		"""Create the LINE Desktop submenu under NVDA Tools menu."""
		try:
			self._lineSubMenu = wx.Menu()

			# ── Chat room tab navigation ──
			self._allChatsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to All chats tab
				_("全部聊天室(&1)") + "\tNVDA+Windows+1",
			)
			self._friendsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Friends tab
				_("好友(&2)") + "\tNVDA+Windows+2",
			)
			self._groupsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Groups tab
				_("群組(&3)") + "\tNVDA+Windows+3",
			)
			self._communitiesItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Communities tab
				_("社群(&4)") + "\tNVDA+Windows+4",
			)
			self._officialItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Official accounts tab
				_("官方帳號(&5)") + "\tNVDA+Windows+5",
			)

			self._lineSubMenu.AppendSeparator()

			# ── Chat functions ──
			self._voiceCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for making a voice call
				_("語音通話(&C)") + "\tNVDA+Windows+C",
			)
			self._videoCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for making a video call
				_("視訊通話(&V)") + "\tNVDA+Windows+V",
			)
			self._moreOptionsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for clicking more options button
				_("更多選項(&O)") + "\tNVDA+Windows+O",
			)
			self._messageReaderItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for opening the message reader
				_("訊息閱讀器(&J)") + "\tNVDA+Windows+J",
			)
			self._readChatNameItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for reading chat room name
				_("讀出聊天室名稱(&T)") + "\tNVDA+Windows+T",
			)

			self._lineSubMenu.AppendSeparator()

			# ── Incoming call functions ──
			self._answerCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for answering a call
				_("接聽來電(&A)") + "\tNVDA+Windows+A",
			)
			self._rejectCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for rejecting a call
				_("拒絕來電(&D)") + "\tNVDA+Windows+D",
			)
			self._checkCallerItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for checking who is calling
				_("查看來電者(&S)") + "\tNVDA+Windows+S",
			)
			self._focusCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for focusing the call window
				_("跳到通話視窗(&F)") + "\tNVDA+Windows+F",
			)

			# Bind events
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onAllChats, self._allChatsItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onFriends, self._friendsItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onGroups, self._groupsItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onCommunities, self._communitiesItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onOfficial, self._officialItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onVoiceCall, self._voiceCallItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onVideoCall, self._videoCallItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onMoreOptions, self._moreOptionsItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onMessageReader, self._messageReaderItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onReadChatName, self._readChatNameItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onAnswerCall, self._answerCallItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onRejectCall, self._rejectCallItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onCheckCaller, self._checkCallerItem
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU, self._onFocusCallWindow, self._focusCallItem
			)

			# Add the submenu to NVDA's Tools menu
			self._toolsMenu = gui.mainFrame.sysTrayIcon.toolsMenu
			# Translators: The label for the LINE Desktop submenu in NVDA's Tools menu
			self._lineMenuItem = self._toolsMenu.AppendSubMenu(
				self._lineSubMenu,
				"LINE Desktop",
			)
			log.info("LINE Desktop: tools menu created")
		except Exception:
			log.debugWarning(
				"Failed to create LINE Desktop tools menu",
				exc_info=True,
			)

	def _removeToolsMenu(self):
		"""Remove the LINE Desktop submenu from NVDA Tools menu."""
		try:
			if self._toolsMenu and self._lineMenuItem:
				self._toolsMenu.Remove(self._lineMenuItem)
				self._lineMenuItem = None
				log.info("LINE Desktop: tools menu removed")
		except Exception:
			log.debugWarning(
				"Failed to remove LINE Desktop tools menu",
				exc_info=True,
			)
		self._lineSubMenu = None
		self._toolsMenu = None

	# ── Menu event handlers ──────────────────────────────────────────

	def _onAllChats(self, evt):
		wx.CallAfter(self._doNavigateTab, "全部")

	def _onFriends(self, evt):
		wx.CallAfter(self._doNavigateTab, "好友")

	def _onGroups(self, evt):
		wx.CallAfter(self._doNavigateTab, "群組")

	def _onCommunities(self, evt):
		wx.CallAfter(self._doNavigateTab, "社群")

	def _onOfficial(self, evt):
		wx.CallAfter(self._doNavigateTab, "官方帳號")

	def _doNavigateTab(self, tabName):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab(tabName):
					ui.message(tabName)
				else:
					ui.message(_("無法切換到{tabName}").format(tabName=tabName))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	def _onVoiceCall(self, evt):
		# Defer execution so the NVDA menu closes first and LINE regains focus
		wx.CallAfter(self._doVoiceCall)

	def _doVoiceCall(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._makeCallByType("voice"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeCall error: {e}", exc_info=True)
			ui.message(_("通話功能錯誤: {error}").format(error=e))

	def _onVideoCall(self, evt):
		wx.CallAfter(self._doVideoCall)

	def _doVideoCall(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._makeCallByType("video"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeVideoCall error: {e}", exc_info=True)
			ui.message(_("視訊通話功能錯誤: {error}").format(error=e))

	def _onMessageReader(self, evt):
		wx.CallAfter(self._doMessageReader)

	def _doMessageReader(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_openMessageReader(None)
		except Exception as e:
			log.warning(f"LINE openMessageReader error: {e}", exc_info=True)
			ui.message(_("訊息閱讀器功能錯誤: {error}").format(error=e))

	def _onMoreOptions(self, evt):
		wx.CallAfter(self._doMoreOptions)

	def _doMoreOptions(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._clickMoreOptionsButton():
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE clickMoreOptions error: {e}", exc_info=True)
			ui.message(_("更多選項功能錯誤: {error}").format(error=e))

	def _onReadChatName(self, evt):
		wx.CallAfter(self._doReadChatName)

	def _doReadChatName(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp._readChatRoomName()
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	def _onAnswerCall(self, evt):
		wx.CallAfter(self._doAnswerCall)

	def _doAnswerCall(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._answerIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(_("接聽功能錯誤: {error}").format(error=e))

	def _onRejectCall(self, evt):
		wx.CallAfter(self._doRejectCall)

	def _doRejectCall(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._rejectIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(_("拒絕功能錯誤: {error}").format(error=e))

	def _onCheckCaller(self, evt):
		wx.CallAfter(self._doCheckCaller)

	def _doCheckCaller(self):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._getCallerInfo(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(_("來電查看功能錯誤: {error}").format(error=e))

	def _onFocusCallWindow(self, evt):
		wx.CallAfter(self._doFocusCallWindow)

	def _doFocusCallWindow(self):
		import ui
		import core
		import speech
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			import ctypes
			import ctypes.wintypes
			hwnd = lineApp._findIncomingCallWindow()
			if not hwnd:
				ui.message(_("未偵測到通話視窗"))
				return
			try:
				ctypes.windll.user32.SetForegroundWindow(hwnd)
			except Exception:
				pass

			def _announceCallWindow():
				try:
					ocrText = lineApp._ocrWindowArea(hwnd, sync=True, timeout=3.0)
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
		except Exception as e:
			log.warning(f"LINE focusCallWindow error: {e}", exc_info=True)
			ui.message(_("跳到通話視窗功能錯誤: {error}").format(error=e))

	def terminate(self, *args, **kwargs):
		self._removeToolsMenu()
		super().terminate(*args, **kwargs)
		for exe in self._LINE_EXECUTABLES:
			try:
				appModuleHandler.unregisterExecutable(exe)
			except Exception:
				pass

	@script(
		# Translators: Description of a debug script to report focused object info
		description=_("Debug: Report focused object's appModule and executable"),
		gesture="kb:NVDA+shift+j"
	)
	def script_reportFocusInfo(self, gesture):
		import api
		import ui
		obj = api.getFocusObject()
		app = obj.appModule
		appName = app.appName if app else "None"
		processID = app.processID if app else "None"
		moduleName = app.__class__.__module__ if app else "None"
		className = app.__class__.__name__ if app else "None"
		
		msg = (
			f"App: {appName}, PID: {processID}, "
			f"Module: {moduleName}.{className}"
		)
		log.info(f"LINE Debug Focus Info: {msg}")
		ui.message(msg)

	# ── Incoming call global shortcuts ─────────────────────────────

	@script(
		# Translators: Description of a script to answer an incoming LINE call
		description=_("LINE: 接聽來電"),
		gesture="kb:NVDA+windows+a",
		category="LINE Desktop",
	)
	def script_answerCall(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._answerIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(_("接聽功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to reject an incoming LINE call
		description=_("LINE: 拒絕來電"),
		gesture="kb:NVDA+windows+d",
		category="LINE Desktop",
	)
	def script_rejectCall(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._rejectIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(_("拒絕功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to check who is calling
		description=_("LINE: 查看來電者"),
		gesture="kb:NVDA+windows+s",
		category="LINE Desktop",
	)
	def script_checkCaller(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._getCallerInfo(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(_("來電查看功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to focus the call window
		description=_("LINE: 跳到通話視窗"),
		gesture="kb:NVDA+windows+f",
		category="LINE Desktop",
	)
	def script_focusCallWindow(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_focusCallWindow(gesture)
		except Exception as e:
			log.warning(f"LINE focusCallWindow error: {e}", exc_info=True)
			ui.message(_("跳到通話視窗功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to read the current chat room name
		description=_("LINE: 讀出目前聊天室名稱"),
		gesture="kb:NVDA+windows+t",
		category="LINE Desktop",
	)
	def script_readChatRoomName(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_readChatRoomName(gesture)
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to open the message reader
		description=_("LINE: 開啟訊息閱讀器"),
		gesture="kb:NVDA+windows+j",
		category="LINE Desktop",
	)
	def script_openMessageReader(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_openMessageReader(gesture)
		except Exception as e:
			log.warning(f"LINE openMessageReader error: {e}", exc_info=True)
			ui.message(_("訊息閱讀器功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to click the more options button
		description=_("LINE: 點擊更多選項按鈕"),
		gesture="kb:NVDA+windows+o",
		category="LINE Desktop",
	)
	def script_clickMoreOptions(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_clickMoreOptions(gesture)
		except Exception as e:
			log.warning(f"LINE clickMoreOptions error: {e}", exc_info=True)
			ui.message(_("更多選項功能錯誤: {error}").format(error=e))

	# ── Chat room tab navigation shortcuts ─────────────────────────

	@script(
		# Translators: Description of a script to navigate to all chats tab
		description=_("LINE: 跳到全部聊天室"),
		gesture="kb:NVDA+windows+1",
		category="LINE Desktop",
	)
	def script_navigateAllChats(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab("全部"):
					ui.message(_("全部"))
				else:
					# Translators: Shown when unable to switch to a chat tab
					ui.message(_("無法切換到全部"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to friends tab
		description=_("LINE: 跳到好友"),
		gesture="kb:NVDA+windows+2",
		category="LINE Desktop",
	)
	def script_navigateFriends(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab("好友"):
					ui.message(_("好友"))
				else:
					ui.message(_("無法切換到好友"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to groups tab
		description=_("LINE: 跳到群組"),
		gesture="kb:NVDA+windows+3",
		category="LINE Desktop",
	)
	def script_navigateGroups(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab("群組"):
					ui.message(_("群組"))
				else:
					ui.message(_("無法切換到群組"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to communities tab
		description=_("LINE: 跳到社群"),
		gesture="kb:NVDA+windows+4",
		category="LINE Desktop",
	)
	def script_navigateCommunities(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab("社群"):
					ui.message(_("社群"))
				else:
					ui.message(_("無法切換到社群"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to official accounts tab
		description=_("LINE: 跳到官方帳號"),
		gesture="kb:NVDA+windows+5",
		category="LINE Desktop",
	)
	def script_navigateOfficial(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, '_navigateToChatTab'):
				if lineApp._navigateToChatTab("官方帳號"):
					ui.message(_("官方帳號"))
				else:
					ui.message(_("無法切換到官方帳號"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))
