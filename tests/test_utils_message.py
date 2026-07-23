"""Tests for _utils: braille output goes through the public API, and the
pending-OCR bookkeeping cannot leak buffers forever."""

from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path


def _load_utils_module():
	module_name = "addon.appModules._utils_under_test"
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "_utils.py"

	saved = {name: sys.modules.get(name) for name in ("logHandler", module_name)}
	try:
		log_handler_mod = types.ModuleType("logHandler")

		class _Log:
			def debug(self, *args, **kwargs):
				pass

			def warning(self, *args, **kwargs):
				pass

		log_handler_mod.log = _Log()
		sys.modules["logHandler"] = log_handler_mod

		spec = importlib.util.spec_from_file_location(module_name, module_path)
		assert spec and spec.loader
		module = importlib.util.module_from_spec(spec)
		sys.modules[module_name] = module
		spec.loader.exec_module(module)
		return module
	finally:
		for name, mod in saved.items():
			if mod is None:
				sys.modules.pop(name, None)
			else:
				sys.modules[name] = mod


class _SysModules:
	"""Temporarily registers fake modules for the function-local imports."""

	def __init__(self, fakes):
		self._fakes = fakes
		self._saved = {}

	def __enter__(self):
		for name, mod in self._fakes.items():
			self._saved[name] = sys.modules.get(name)
			sys.modules[name] = mod
		return self

	def __exit__(self, *exc_info):
		for name, mod in self._saved.items():
			if mod is None:
				sys.modules.pop(name, None)
			else:
				sys.modules[name] = mod
		return False


def _speech_module(spoken):
	mod = types.ModuleType("speech")
	mod.speakMessage = spoken.append
	return mod


def test_message_speaks_and_routes_braille_through_handler_message():
	utils = _load_utils_module()
	spoken = []
	braille_messages = []

	braille_mod = types.ModuleType("braille")
	braille_mod.handler = types.SimpleNamespace(message=braille_messages.append)

	with _SysModules({"speech": _speech_module(spoken), "braille": braille_mod}):
		utils.message("hello")

	assert spoken == ["hello"]
	assert braille_messages == ["hello"]


def test_message_survives_uninitialized_braille():
	utils = _load_utils_module()
	spoken = []

	braille_mod = types.ModuleType("braille")
	braille_mod.handler = None

	with _SysModules({"speech": _speech_module(spoken), "braille": braille_mod}):
		utils.message("hello")

	assert spoken == ["hello"]


def test_message_survives_braille_handler_errors():
	utils = _load_utils_module()
	spoken = []

	def _raise(_text):
		raise RuntimeError("braille broke")

	braille_mod = types.ModuleType("braille")
	braille_mod.handler = types.SimpleNamespace(message=_raise)

	with _SysModules({"speech": _speech_module(spoken), "braille": braille_mod}):
		utils.message("hello")

	assert spoken == ["hello"]


def test_prune_pending_ocr_drops_only_stale_entries():
	import time

	utils = _load_utils_module()
	now = time.monotonic()
	utils._pendingOcr.clear()
	utils._pendingOcr[1] = (now - 120, "recognizer", "pixels", "info", "bitmap")
	utils._pendingOcr[2] = (now, "recognizer", "pixels", "info", "bitmap")

	utils._prunePendingOcr()

	assert 1 not in utils._pendingOcr, "entries older than the max age must be released"
	assert 2 in utils._pendingOcr, "fresh (possibly in-flight) entries must be kept"
	utils._pendingOcr.clear()


def test_schedule_pending_ocr_cleanup_defers_via_call_later():
	utils = _load_utils_module()
	scheduled = []

	core_mod = types.ModuleType("core")
	core_mod.callLater = lambda delay, func, *args: scheduled.append((delay, func, args))

	with _SysModules({"core": core_mod}):
		utils.schedulePendingOcrCleanup()

	assert len(scheduled) == 1
	delay, func, args = scheduled[0]
	# Deferred long enough for any in-flight recognition to finish first.
	assert delay >= 5000
	assert func is utils._prunePendingOcr
