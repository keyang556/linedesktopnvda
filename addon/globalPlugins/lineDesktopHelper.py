# LINE Desktop Global Plugin for NVDA
# Maps alternative executable names to the LINE appModule.

import appModuleHandler
import globalPluginHandler
from scriptHandler import script
from logHandler import log


def _getLineAppModule():
	"""Find and return the LINE appModule instance, or None."""
	for app in appModuleHandler.runningTable.values():
		if app and getattr(app, 'appName', '').lower() in ('line', 'line_app'):
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

	def terminate(self, *args, **kwargs):
		super().terminate(*args, **kwargs)
		for exe in self._LINE_EXECUTABLES:
			try:
				appModuleHandler.unregisterExecutable(exe)
			except Exception:
				pass

	@script(
		description="Debug: Report focused object's appModule and executable",
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
		description="LINE: 接聽來電",
		gesture="kb:NVDA+windows+a",
		category="LINE Desktop",
	)
	def script_answerCall(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message("LINE 未執行")
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._answerIncomingCall(hwnd)
			else:
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(f"接聽功能錯誤: {e}")

	@script(
		description="LINE: 拒絕來電",
		gesture="kb:NVDA+windows+d",
		category="LINE Desktop",
	)
	def script_rejectCall(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message("LINE 未執行")
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._rejectIncomingCall(hwnd)
			else:
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(f"拒絕功能錯誤: {e}")

	@script(
		description="LINE: 查看來電者",
		gesture="kb:NVDA+windows+s",
		category="LINE Desktop",
	)
	def script_checkCaller(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message("LINE 未執行")
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._getCallerInfo(hwnd)
			else:
				ui.message("未偵測到來電")
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(f"來電查看功能錯誤: {e}")

	@script(
		description="LINE: 跳到通話視窗",
		gesture="kb:NVDA+windows+f",
		category="LINE Desktop",
	)
	def script_focusCallWindow(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message("LINE 未執行")
			return
		try:
			lineApp.script_focusCallWindow(gesture)
		except Exception as e:
			log.warning(f"LINE focusCallWindow error: {e}", exc_info=True)
			ui.message(f"跳到通話視窗功能錯誤: {e}")

	@script(
		description="LINE: 讀出目前聊天室名稱",
		gesture="kb:NVDA+windows+t",
		category="LINE Desktop",
	)
	def script_readChatRoomName(self, gesture):
		import ui
		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message("LINE 未執行")
			return
		try:
			lineApp.script_readChatRoomName(gesture)
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(f"讀取聊天室名稱錯誤: {e}")
