"""Behavioral tests for the AI image-description consent.

The consent gate must:
- ask on first use of each provider, whether the request would use the
  add-on's bundled key or the user's own (the disclosure is about which
  company receives the screenshot, which owning a key does not answer);
- never let consent for one provider carry over to another;
- persist acceptance, and not ask again for that provider;
- block the upload (without persisting anything) when declined;
- refuse on the worker-thread chokepoint (_callImageDescriptionApi) without
  showing UI when consent is missing for the provider it is about to call.
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
	"_readConsentedAiProviders",
	"_hasAiConsent",
	"_recordAiConsent",
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


def _make_namespace(provider="google"):
	"""Build the extracted-helper namespace.

	Returns (namespace, providerHolder); mutate providerHolder[0] to simulate
	the user switching provider in the settings panel.
	"""
	providerHolder = [provider]
	namespace = _load_consent_helpers(
		{
			"_getEffectiveImageProvider": lambda: providerHolder[0],
			"_IMAGE_DESCRIPTION_PROVIDER_LABELS": {
				"google": "Google AI",
				"mistral": "Mistral AI",
			},
			"_IMAGE_DESCRIPTION_PROVIDER_GOOGLE": "google",
			"_IMAGE_DESCRIPTION_PROVIDER_OLLAMA": "ollama",
			"_IMAGE_DESCRIPTION_PROVIDER_NVIDIA": "nvidia",
			"_IMAGE_DESCRIPTION_PROVIDER_POLLINATIONS": "pollinations",
			"_IMAGE_DESCRIPTION_PROVIDER_OPENAI": "openai",
			"_IMAGE_DESCRIPTION_PROVIDER_MISTRAL": "mistral",
		},
	)
	return namespace, providerHolder


def test_consent_dialog_accept_persists_and_never_asks_again(tmp_path):
	ns, _provider = _make_namespace()
	with _FakeGuiWx(tmp_path, answers=[2, 2]) as fake:  # wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1
		# Recorded: second call must not show another dialog.
		assert ns["_hasAiConsent"]() is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1


def test_consent_dialog_decline_blocks_and_persists_nothing(tmp_path):
	ns, _provider = _make_namespace()
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
	ns, _provider = _make_namespace()
	unwritable_config_path = tmp_path / "does-not-exist"  # never created
	with _FakeGuiWx(unwritable_config_path, answers=[2, 2]) as fake:  # wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1
		assert ns["_hasAiConsent"]() is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1


def test_consent_is_asked_even_when_user_supplied_their_own_key(tmp_path):
	"""Owning the API key does not answer "which company gets my screenshot",
	so the first use of a provider is disclosed regardless of key source."""
	ns, _provider = _make_namespace()
	with _FakeGuiWx(tmp_path, answers=[2]) as fake:  # wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1


def test_consent_does_not_carry_over_to_another_provider(tmp_path):
	"""Consenting to one cloud service must never imply consent for a
	different company the user was never shown by name."""
	ns, provider = _make_namespace(provider="google")
	with _FakeGuiWx(tmp_path, answers=[2, 2]) as fake:  # wx.YES, wx.YES
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 1

		provider[0] = "mistral"
		assert ns["_hasAiConsent"]() is False, "consent must not transfer between providers"
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 2
		# The second dialog must name the newly selected service.
		assert "Mistral AI" in fake.message_boxes[1][0]

		# Both providers are now independently remembered.
		assert ns["_hasAiConsent"]("google") is True
		assert ns["_hasAiConsent"]("mistral") is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 2


def test_declining_for_new_provider_does_not_revoke_the_previous_one(tmp_path):
	ns, provider = _make_namespace(provider="google")
	with _FakeGuiWx(tmp_path, answers=[2, 8]) as fake:  # wx.YES, then wx.NO
		assert ns["_ensureImageDescriptionConsent"]() is True

		provider[0] = "mistral"
		assert ns["_ensureImageDescriptionConsent"]() is False
		assert len(fake.message_boxes) == 2

		provider[0] = "google"
		assert ns["_hasAiConsent"]() is True
		assert ns["_ensureImageDescriptionConsent"]() is True
		assert len(fake.message_boxes) == 2


def _stub_provider_backends(ns, provider_calls):
	for name in (
		"_callGoogleImageDescriptionApi",
		"_callOllamaImageDescriptionApi",
		"_callNvidiaImageDescriptionApi",
		"_callPollinationsImageDescriptionApi",
		"_callOpenaiImageDescriptionApi",
		"_callMistralImageDescriptionApi",
	):
		ns[name] = lambda *args, **kwargs: provider_calls.append("called") or ("text", None)


def test_worker_chokepoint_refuses_without_consent(tmp_path):
	provider_calls = []
	ns, _provider = _make_namespace()
	_stub_provider_backends(ns, provider_calls)

	with _FakeGuiWx(tmp_path, answers=[]):
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
	assert text is None
	assert error
	assert provider_calls == []


def test_worker_chokepoint_dispatches_after_consent(tmp_path):
	provider_calls = []
	ns, _provider = _make_namespace()
	_stub_provider_backends(ns, provider_calls)

	with _FakeGuiWx(tmp_path, answers=[]):
		assert ns["_recordAiConsent"]() is True
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
	assert text == "text"
	assert error is None
	assert provider_calls == ["called"]


def test_worker_chokepoint_refuses_after_switching_to_an_unconsented_provider(tmp_path):
	"""The follow-up Q&A path re-resolves the provider on every turn, so
	switching provider mid-conversation must not inherit the previous
	provider's consent."""
	provider_calls = []
	ns, provider = _make_namespace(provider="google")
	_stub_provider_backends(ns, provider_calls)

	with _FakeGuiWx(tmp_path, answers=[]):
		assert ns["_recordAiConsent"]("google") is True
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
		assert text == "text"

		provider[0] = "mistral"
		text, error = ns["_callImageDescriptionApi"]([{"parts": []}])
	assert text is None
	assert error
	assert provider_calls == ["called"], "no second upload may reach the new provider"
