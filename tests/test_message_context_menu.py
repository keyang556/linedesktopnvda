from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path


def _load_message_context_menu_module():
	module_name = "addon.appModules._virtualWindows.messageContextMenu"
	module_path = (
		Path(__file__).resolve().parents[1]
		/ "addon"
		/ "appModules"
		/ "_virtualWindows"
		/ "messageContextMenu.py"
	)

	for name in (
		"addon",
		"addon.appModules",
		"addon.appModules._virtualWindows",
	):
		pkg = types.ModuleType(name)
		pkg.__path__ = []  # type: ignore[attr-defined]
		sys.modules[name] = pkg

	virtual_window_mod = types.ModuleType("addon.appModules._virtualWindow")

	class VirtualWindow:
		currentWindow = None

		@property
		def element(self):
			if not getattr(self, "elements", None):
				return None
			return self.elements[self.pos]

		def click(self):
			return None

	virtual_window_mod.VirtualWindow = VirtualWindow
	sys.modules["addon.appModules._virtualWindow"] = virtual_window_mod

	utils_mod = types.ModuleType("addon.appModules._utils")
	utils_mod.ocrGetText = lambda *args, **kwargs: None
	utils_mod.message = lambda *args, **kwargs: None
	sys.modules["addon.appModules._utils"] = utils_mod

	log_handler_mod = types.ModuleType("logHandler")

	class _Log:
		def debug(self, *args, **kwargs):
			pass

		def info(self, *args, **kwargs):
			pass

	log_handler_mod.log = _Log()
	sys.modules["logHandler"] = log_handler_mod

	spec = importlib.util.spec_from_file_location(module_name, module_path)
	assert spec and spec.loader
	module = importlib.util.module_from_spec(spec)
	sys.modules[module_name] = module
	spec.loader.exec_module(module)
	return module


message_context_menu = _load_message_context_menu_module()


def test_message_context_menu_click_invokes_action_callback_and_closes_window():
	calls = []
	window = object.__new__(message_context_menu.MessageContextMenu)
	window.elements = [{"name": "收回", "clickPoint": (1108, 518)}]
	window.pos = 0
	window.onAction = calls.append

	message_context_menu.VirtualWindow.currentWindow = window
	message_context_menu.MessageContextMenu.click(window)

	assert calls == ["收回"]
	assert message_context_menu.VirtualWindow.currentWindow is None


def test_build_menu_elements_ignores_noise_lines_before_copy_row():
	lines = [
		{"text": "50", "rect": (700, 430, 732, 450)},
		{"text": "回覆", "rect": (706, 449, 742, 469)},
		{"text": "複製", "rect": (706, 489, 742, 511)},
		{"text": "分享", "rect": (706, 530, 742, 550)},
	]
	row_rects = [
		(639, 439, 837, 479),
		(639, 480, 837, 520),
		(639, 520, 837, 560),
	]

	elements = message_context_menu._buildMenuElements(
		lines,
		(624, 415, 852, 860),
		rowRects=row_rects,
	)

	assert [element["name"] for element in elements] == ["回覆", "複製", "分享"]
	assert [element["clickPoint"] for element in elements] == [
		(738, 459),
		(738, 500),
		(738, 540),
	]
