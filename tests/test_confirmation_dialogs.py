"""Behavioral tests for the recall / photo-consent message dialogs.

These dialogs replaced transient single-key modes that bound Y/N/P and A/D while
LINE's own dialog was up. The replacement must:
- offer exactly the actions LINE itself offers (no stealth recall button when
  LINE does not show one);
- keep the old letters working, now as button accelerators;
- treat closing the dialog (escape, alt+f4, NVDA shutting down) as the
  non-destructive answer the ten-second timeout used to pick;
- hand the chosen action back only once the dialog is off the screen, because
  the caller OCRs and clicks LINE's dialog underneath it.
"""

from __future__ import annotations

import importlib.util
import inspect
import sys
import types
from pathlib import Path


class _FakeButton:
	"""Stands in for gui.message.Button, whose fields we assert on."""

	def __init__(
		self,
		id,
		label,
		callback=None,
		defaultFocus=False,
		fallbackAction=False,
		closesDialog=True,
		returnCode=None,
	):
		self.id = id
		self.label = label
		self.callback = callback
		self.defaultFocus = defaultFocus
		self.fallbackAction = fallbackAction
		self.closesDialog = closesDialog
		self.returnCode = returnCode


class _FakeReturnCode:
	YES = "YES"
	NO = "NO"
	CANCEL = "CANCEL"
	CUSTOM_1 = "CUSTOM_1"


class _FakeMessageDialog:
	"""Records what the add-on asked for instead of building a real dialog."""

	shown: list["_FakeMessageDialog"] = []

	def __init__(self, parent, message, title, dialogType=None, *, buttons=()):
		self.parent = parent
		self.message = message
		self.title = title
		self.dialogType = dialogType
		self.buttons = list(buttons)
		self.isShown = False

	def Show(self):
		self.isShown = True
		type(self).shown.append(self)

	def Raise(self):
		pass

	def SetFocus(self):
		pass

	def button(self, label):
		return next(button for button in self.buttons if button.label == label)


def _load_confirmation_dialogs_module():
	module_name = "addon.appModules._confirmationDialogs"
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "_confirmationDialogs.py"

	fakedNames = (
		"addon",
		"addon.appModules",
		"wx",
		"gui",
		"gui.message",
		"logHandler",
		"addonHandler",
		module_name,
	)
	saved = {name: sys.modules.get(name) for name in fakedNames}
	try:
		for name in ("addon", "addon.appModules"):
			pkg = types.ModuleType(name)
			pkg.__path__ = []  # type: ignore[attr-defined]
			sys.modules[name] = pkg

		wx_mod = types.ModuleType("wx")
		# Run deferred callables straight away; the tests care about what was
		# scheduled, not about wx's event loop.
		wx_mod.CallAfter = lambda func, *args, **kwargs: func(*args, **kwargs)
		sys.modules["wx"] = wx_mod

		message_mod = types.ModuleType("gui.message")
		message_mod.Button = _FakeButton
		message_mod.ReturnCode = _FakeReturnCode
		message_mod.MessageDialog = _FakeMessageDialog
		message_mod.DialogType = types.SimpleNamespace(STANDARD="standard", WARNING="warning")

		gui_mod = types.ModuleType("gui")
		gui_mod.mainFrame = object()
		gui_mod.message = message_mod
		sys.modules["gui"] = gui_mod
		sys.modules["gui.message"] = message_mod

		log_handler_mod = types.ModuleType("logHandler")
		log_handler_mod.log = types.SimpleNamespace(
			debug=lambda *args, **kwargs: None,
			info=lambda *args, **kwargs: None,
			warning=lambda *args, **kwargs: None,
		)
		sys.modules["logHandler"] = log_handler_mod

		def _initTranslation():
			"""Install a passthrough gettext into the caller, as NVDA's does."""
			inspect.currentframe().f_back.f_globals["_"] = lambda text: text

		addon_handler_mod = types.ModuleType("addonHandler")
		addon_handler_mod.initTranslation = _initTranslation
		sys.modules["addonHandler"] = addon_handler_mod

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


confirmation_dialogs = _load_confirmation_dialogs_module()


def _openRecall(availableActions, chosen):
	_FakeMessageDialog.shown.clear()
	confirmation_dialogs._recallDialog = None
	confirmation_dialogs.openRecallConfirmationDialog(
		"確認要收回嗎？",
		availableActions,
		chosen.append,
	)
	return _FakeMessageDialog.shown[-1]


def _openPhotoConsent(chosen):
	_FakeMessageDialog.shown.clear()
	confirmation_dialogs._photoConsentDialog = None
	confirmation_dialogs.openPhotoTextConsentDialog(
		"轉為文字會將照片上傳到 LINE 伺服器處理。是否同意？",
		chosen.append,
	)
	return _FakeMessageDialog.shown[-1]


def test_recall_dialog_keeps_the_old_y_n_p_keys_as_button_accelerators():
	"""Muscle memory from the single-key mode has to survive the move."""
	dlg = _openRecall({"收回", "取消", "無痕收回"}, [])

	assert [button.label for button in dlg.buttons] == ["收回(&Y)", "無痕收回(&P)", "取消(&N)"]


def test_recall_dialog_omits_stealth_recall_when_line_does_not_offer_it():
	"""Offering a button LINE has not drawn would leave nothing to click."""
	dlg = _openRecall({"收回", "取消"}, [])

	assert [button.label for button in dlg.buttons] == ["收回(&Y)", "取消(&N)"]


def test_recall_dialog_buttons_report_the_line_action_names():
	for label, expected in (("收回(&Y)", "收回"), ("無痕收回(&P)", "無痕收回"), ("取消(&N)", "取消")):
		chosen = []
		dlg = _openRecall({"收回", "取消", "無痕收回"}, chosen)
		dlg.button(label).callback(None)
		assert chosen == [expected]


def test_recall_dialog_falls_back_to_cancel():
	"""Escape, alt+f4 or NVDA shutting down must pick the harmless action, the
	way the ten-second auto-cancel used to."""
	dlg = _openRecall({"收回", "取消", "無痕收回"}, [])

	assert dlg.button("取消(&N)").fallbackAction is True
	assert [button.label for button in dlg.buttons if button.defaultFocus] == ["收回(&Y)"]


def test_photo_consent_dialog_keeps_the_old_a_d_keys_and_declines_on_escape():
	chosen = []
	dlg = _openPhotoConsent(chosen)

	assert [button.label for button in dlg.buttons] == ["同意(&A)", "不同意(&D)"]
	assert dlg.button("不同意(&D)").fallbackAction is True
	assert [button.label for button in dlg.buttons if button.defaultFocus] == ["同意(&A)"]

	dlg.button("同意(&A)").callback(None)
	assert chosen == ["同意"]


def test_photo_consent_dialog_reports_the_decline_action():
	chosen = []
	dlg = _openPhotoConsent(chosen)
	dlg.button("不同意(&D)").callback(None)

	assert chosen == ["不同意"]


def test_choosing_an_action_releases_the_singleton_so_the_next_prompt_opens():
	"""A stale reference would silently swallow every later confirmation."""
	dlg = _openRecall({"收回", "取消"}, [])
	dlg.button("取消(&N)").callback(None)

	assert confirmation_dialogs._recallDialog is None

	dlg = _openPhotoConsent([])
	dlg.button("不同意(&D)").callback(None)

	assert confirmation_dialogs._photoConsentDialog is None


def test_a_second_detection_refocuses_the_open_dialog_instead_of_stacking_one():
	chosen = []
	first = _openRecall({"收回", "取消"}, chosen)
	confirmation_dialogs.openRecallConfirmationDialog("確認要收回嗎？", {"收回", "取消"}, chosen.append)

	assert _FakeMessageDialog.shown == [first]


def test_a_dialog_that_cannot_be_shown_reports_back_instead_of_wedging_the_feature():
	"""The caller marks the confirmation pending before the dialog appears. If
	building it throws and nobody says so, that flag stays set and every later
	recall is silently ignored."""
	failures = []

	def _explode(*args, **kwargs):
		raise RuntimeError("no GUI")

	original = confirmation_dialogs.MessageDialog
	confirmation_dialogs.MessageDialog = _explode
	try:
		confirmation_dialogs._recallDialog = None
		confirmation_dialogs.openRecallConfirmationDialog(
			"確認要收回嗎？",
			{"收回", "取消"},
			lambda action: failures.append(("chose", action)),
			onFailed=lambda: failures.append("recall failed"),
		)
		confirmation_dialogs._photoConsentDialog = None
		confirmation_dialogs.openPhotoTextConsentDialog(
			"是否同意？",
			lambda action: failures.append(("chose", action)),
			onFailed=lambda: failures.append("consent failed"),
		)
	finally:
		confirmation_dialogs.MessageDialog = original

	assert failures == ["recall failed", "consent failed"]


def test_dialogs_warn_rather_than_showing_a_plain_message():
	"""Both dialogs gate something irreversible (an unsend) or privacy-relevant
	(uploading a photo), so they get the warning icon and sound."""
	assert _openRecall({"收回", "取消"}, []).dialogType == "warning"
	assert _openPhotoConsent([]).dialogType == "warning"
