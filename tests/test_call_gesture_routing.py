"""Source guards for the incoming-call gesture routing in the global plugin.

The NVDA+Windows+A/D/S/F handlers (and their Tools-menu equivalents) must
delegate to the appModule's worker-thread scripts instead of inlining
`_findIncomingCallWindow` / `_answerIncomingCall` / `_getCallerInfo` /
`_ocrWindowArea`, which block NVDA's main thread for seconds (fixed sync
OCR + sleeps). Reintroducing the inline pattern would resurface the freeze.
"""

from __future__ import annotations

import ast
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parents[1] / "addon" / "globalPlugins" / "lineDesktopHelper.py"

# handler name -> appModule script it must delegate to
DELEGATING_HANDLERS = {
	"script_answerCall": "script_answerCall",
	"script_rejectCall": "script_rejectCall",
	"script_checkCaller": "script_checkCaller",
	"script_focusCallWindow": "script_focusCallWindow",
	"_doAnswerCall": "script_answerCall",
	"_doRejectCall": "script_rejectCall",
	"_doCheckCaller": "script_checkCaller",
	"_doFocusCallWindow": "script_focusCallWindow",
}

# Names whose presence in a handler means blocking work happens on the main thread.
FORBIDDEN_ATTRIBUTES = {
	"_findIncomingCallWindow",
	"_answerIncomingCall",
	"_rejectIncomingCall",
	"_getCallerInfo",
	"_ocrWindowArea",
}


def _get_global_plugin_methods():
	source = MODULE_PATH.read_text(encoding="utf-8")
	module = ast.parse(source)
	plugin = next(
		node
		for node in module.body
		if isinstance(node, ast.ClassDef) and node.name == "GlobalPlugin"
	)
	return {
		node.name: node for node in plugin.body if isinstance(node, ast.FunctionDef)
	}


def test_call_handlers_delegate_to_app_module_scripts():
	methods = _get_global_plugin_methods()
	for handler_name, delegate in DELEGATING_HANDLERS.items():
		method = methods[handler_name]
		delegating_calls = [
			node
			for node in ast.walk(method)
			if isinstance(node, ast.Call)
			and isinstance(node.func, ast.Attribute)
			and node.func.attr == delegate
			and isinstance(node.func.value, ast.Name)
			and node.func.value.id == "lineApp"
		]
		assert delegating_calls, f"{handler_name} must delegate to lineApp.{delegate}"


def test_call_handlers_do_not_inline_blocking_call_work():
	methods = _get_global_plugin_methods()
	for handler_name in DELEGATING_HANDLERS:
		method = methods[handler_name]
		used = {
			node.attr
			for node in ast.walk(method)
			if isinstance(node, ast.Attribute)
		}
		inlined = used & FORBIDDEN_ATTRIBUTES
		assert not inlined, (
			f"{handler_name} references {sorted(inlined)}; blocking call work "
			"must stay on the appModule's worker thread"
		)
