"""Behavioral tests for the one-time AI image-description consent.

The consent gate must:
- never ask users whose active provider uses their own API key;
- ask once, persist acceptance, and never ask again;
- block the upload (without persisting anything) when declined;
- refuse on the worker-thread chokepoint (_callImageDescriptionApi) without
  showing UI when consent is somehow missing there.
"""

from __future__ import annotations

import ast
import sys
import types
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"

NEEDED_ASSIGNMENTS = {
	"_AI_CONSENT_FILENAME",
	"_aiConsentGrantedInSession",
}
NEEDED_FUNCTIONS = {
	"_getAiConsentStorePath",
	"_hasAiConsent",
	"_recordAiConsent",
	"_isImageDescriptionUsingUserKey",
	"_ensureImageDescriptionConsent",
	"_callImageDescriptionApi",
}


class _Log:
	def debug(self, *args, **kwargs):
		pass

	def warning(self, *args, **kwargs):
		pass

	def info(self, *args, **kwargs):
		pass


def _load_consent_helpers(extra_namespace):
	import os

	source = MODULE_PATH.read_text(encoding="utf-8")
	module = ast.parse(source)
	namespace = {"os": os, "log": _Log(), "_": lambda text: text}
	namespace.update(extra_namespace)
	for node in module.body:
		is_needed_assignment = isinstance(node, ast.Assign) and any(
			isinstance(target, ast.Name) and target.id in NEEDED_ASSIGNMENTS for target in node.targets
		)
		is_needed_function = isinstance(node, ast.FunctionDef) and node.name in NEEDED_FUNCTIONS
		if is_needed_assignment or is_needed_function:
			exec(
				compile(ast.Module(body=[node], type_ignores=[]), str(MODULE_PATH), "exec"),
				namespace,
			)
	return namespace


class _FakeGuiWx:
	"""Registers gui/wx/globalVars stubs in sys.modules for the lazy imports."""

	def __init__(self, config_path, answers):
		self._config_path = config_path
		self.message_boxes = []
		self._answers = list(answers)
		self._saved = {}

	def __enter__(self):
		wx_mod = types.ModuleType("wx")
		wx_mod.YES_NO = 0x0A
		wx_mod.ICON_WARNING = 0x100
		wx_mod.YES = 2
		wx_mod.NO = 8

		gui_mod = types.ModuleType("gui")

		def _message_box(message, caption, style):
			self.message_boxes.append((message, caption, style))
			return self._answers.pop(0)

		gui_mod.messageBox = _message_box

		global_vars_mod = types.ModuleType("globalVars")
		global_vars_mod.appArgs = types.SimpleNamespace(configPath=str(self._config_path))

		for name, mod in (("wx", wx_mod), ("gui", gui_mod), ("globalVars", global_vars_mod)):
			self._saved[name] = sys.modules.get(name)
			sys.modules[name] = mod
		self.wx = wx_mod
		return self

	def __exit__(self, *exc_info):
		for name, mod in self._saved.items():
			if mod is None:
				sys.modules.pop(name, None)
			else:
				sys.modules[name] = mod
		return False


def _make_namespace(user_key=None, provider="google"):
	return _load_consent_helpers(
		{
			"_getEffectiveImageProvider": lambda: provider,
			"_IMAGE_DESCRIPTION_PROVIDER_LABELS": {"google": "Google AI"},
			"_IMAGE_DESCRIPTION_PROVIDER_GOOGLE": "google",
			"_IMAGE_DESCRIPTION_PROVIDER_OLLAMA": "ollama",
			"_IMAGE_DESCRIPTION_PROVIDER_NVIDIA": "nvidia",
			"_IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS": "pollinations",
			"_IMAGE_DESCRIPTION_PROVIDER_OPENAI": "openai",
			"_IMAGE_DESCRIPTION_PROVIDER_MISTRAL": "mistral",
			"getUserImageApiKey": lambda: user_key,
			"getUserOllamaApiKey": lambda: None,
			"getUserNvidiaApiKey": lambda: None,
			"getUserPollinationsApiKey": lambda: None,
			"getUserOpenaiApiKey": lambda: None,
			"getUserMistralApiKey": lambda: None,
		},
	)


def test_consent_dialog_accept_persists_and_never_asks_again(tmp_path):
	ns = _make_namespace()
	with _FakeGuiWx(tmp_path, answers=[2, 2]) as fake:  # wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1
		# Recorded: second call must not show another dialog.
		assert ns["_hasAiConsent"]() is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1


def test_consent_dialog_decline_blocks_and_persists_nothing(tmp_path):
	ns = _make_namespace()
	with _FakeGuiWx(tmp_path, answers=[8, 8]) as fake:  # wx.NO
		assert ns["_ensureImageDescriptionConsent"]() is False
		assert ns["_hasAiConsent"]() is False
		# Declining is not remembered: the next use asks again.
		assert ns["_ensureImageDescriptionConsent"]() is False
		assert len(fake.message_boxes) == 2


def test_consent_grant_falls_back_to_in_session_cache_when_disk_write_fails(tmp_path):
	"""If the config directory can't be written to (read-only, full disk,
	etc.), a granted "Yes" must still be honored for the rest of this NVDA
	session instead of every subsequent call failing with "not consented"
	right after the user agreed."""
	ns = _make_namespace()
	unwritable_config_path = tmp_path / "does-not-exist"  # never created
	with _FakeGuiWx(unwritable_config_path, answers=[2, 2]) as fake:  # wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1
		assert ns["_hasAiConsent"]() is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1


def test_consent_not_asked_when_user_key_configured(tmp_path):
	ns = _make_namespace(user_key="sk-user-supplied")
	with _FakeGuiWx(tmp_path, answers=[]) as fake:
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert fake.message_boxes == []
		assert ns["_hasAiConsent"]() is False


def test_worker_chokepoint_refuses_without_consent(tmp_path):
	provider_calls = []
	ns = _make_namespace()
	for name in (
		"_callGoogleImageDescriptionApi",
		"_callOllamaImageDescriptionApi",
		"_callNvidiaImageDescriptionApi",
		"_callPollinationsImageDescriptionApi",
		"_callOpenaiImageDescriptionApi",
		"_callMistralImageDescriptionApi",
	):
		ns[name] = lambda *args, **kwargs: provider_calls.append("called") or ("text", None)

	with _FakeGuiWx(tmp_path, answers=[]):
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
	assert text is None
	assert error
	assert provider_calls == []


def test_worker_chokepoint_dispatches_after_consent(tmp_path):
	provider_calls = []
	ns = _make_namespace()
	for name in (
		"_callGoogleImageDescriptionApi",
		"_callOllamaImageDescriptionApi",
		"_callNvidiaImageDescriptionApi",
		"_callPollinationsImageDescriptionApi",
		"_callOpenaiImageDescriptionApi",
		"_callMistralImageDescriptionApi",
	):
		ns[name] = lambda *args, **kwargs: provider_calls.append("called") or ("text", None)

	with _FakeGuiWx(tmp_path, answers=[]):
		assert ns["_recordAiConsent"]() is True
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
	assert text == "text"
	assert error is None
	assert provider_calls == ["called"]
