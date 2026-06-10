from .._virtualWindow import VirtualWindow
from controlTypes.role import Role
from controlTypes.state import State
from logHandler import log
import addonHandler

addonHandler.initTranslation()


def _walk(obj, *steps):
	"""Follow attribute steps, returning None as soon as any link is missing.

	The tray menu has no UIA names, so items are located purely by their
	position relative to each other; any link can be None or go stale
	(COMError) while LINE redraws the menu.
	"""
	for step in steps:
		if obj is None:
			return None
		try:
			obj = getattr(obj, step)
		except Exception:
			return None
	return obj


class Tray(VirtualWindow):
	title = f"Line {Role.MENU.displayString}"

	@staticmethod
	def isMatchLineScreen(obj):
		try:
			return "LcContextMenu" in obj.UIAElement.CurrentClassName
		except AttributeError:
			return False

	def makeElements(self):
		try:
			self._buildElements()
		except Exception:
			log.debugWarning("LINE tray menu: failed to build elements", exc_info=True)
		finally:
			self.elements.reverse()

	def _appendClickable(self, name, obj):
		"""Append a positional menu item; returns False when obj is unusable."""
		location = getattr(obj, "location", None) if obj is not None else None
		if not location:
			return False
		self.elements.append(
			{
				"name": name,
				"role": None,
				"clickPoint": self.rectGetCenterPoint(location),
			},
		)
		return True

	def _buildElements(self):
		obj = _walk(self.obj, "firstChild", "next", "lastChild")
		if obj is None:
			return
		# Translators: Tray menu item that quits the LINE application.
		if not self._appendClickable(_("結束應用程式"), obj):
			return
		parentFirstStates = getattr(_walk(obj, "parent", "firstChild"), "states", None) or set()
		stateText = State.UNAVAILABLE.displayString if State.UNAVAILABLE in parentFirstStates else ""
		previousStates = getattr(_walk(obj, "previous"), "states", None) or set()
		isLoggedIn = not stateText or State.UNAVAILABLE in previousStates
		if isLoggedIn:
			obj = _walk(obj, "previous")
			# Translators: Tray menu item that logs out of LINE.
			if not self._appendClickable(_("登出") + (f" ({stateText})" if stateText else ""), obj):
				return

		obj = _walk(obj, "previous", "previous")
		# Translators: Tray menu item that checks for LINE updates.
		if not self._appendClickable(_("確認有無最新版本"), obj):
			return
		obj = _walk(obj, "previous")
		# Translators: Tray menu item that shows the About LINE dialog.
		if not self._appendClickable(_("關於LINE"), obj):
			return
		obj = _walk(obj, "previous")
		# Translators: Tray menu item that opens Keep notes.
		if not self._appendClickable(_("Keep筆記") + (f" ({stateText})" if stateText else ""), obj):
			return
		obj = _walk(obj, "previous")
		# Translators: Tray menu item that opens LINE settings.
		if not self._appendClickable(_("設定"), obj):
			return
		if not isLoggedIn:
			obj = _walk(obj, "previous", "previous")
			# Translators: Tray menu item that logs into LINE.
			if not self._appendClickable(_("登入"), obj):
				return

		obj = _walk(obj, "previous", "previous")
		# Translators: Tray menu item that opens the friends list.
		self._appendClickable(_("好友名單") + (f" ({stateText})" if stateText else ""), obj)
