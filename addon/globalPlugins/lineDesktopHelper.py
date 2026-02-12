# LINE Desktop Global Plugin for NVDA
# Maps alternative executable names to the LINE appModule.

import appModuleHandler
import globalPluginHandler
from scriptHandler import script
from logHandler import log


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
