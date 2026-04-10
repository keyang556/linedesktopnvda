from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path

from addon.appModules._chatParser import parseChatFile


def _load_message_reader_module():
	module_name = "addon.appModules._messageReader"
	module_path = (
		Path(__file__).resolve().parents[1]
		/ "addon"
		/ "appModules"
		/ "_messageReader.py"
	)

	gui_mod = types.ModuleType("gui")
	gui_mod.mainFrame = object()
	sys.modules["gui"] = gui_mod

	wx_mod = types.ModuleType("wx")
	wx_mod.Dialog = type("Dialog", (), {})
	sys.modules["wx"] = wx_mod

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
	module._ = lambda text: text
	sys.modules[module_name] = module
	spec.loader.exec_module(module)
	return module


message_reader = _load_message_reader_module()


def test_parse_chat_file_keeps_date_rows_in_original_positions(tmp_path):
	chat_file = tmp_path / "chat.txt"
	chat_file.write_text(
		"\n".join(
			[
				"2026.04.09 星期四",
				"09:00 Alice 早安",
				"2026.04.10 星期五",
				"10:30 Bob 已收回訊息",
			]
		),
		encoding="utf-8",
	)

	assert parseChatFile(chat_file) == [
		{"type": "date", "content": "2026.04.09 星期四"},
		{"type": "message", "time": "09:00", "name": "Alice", "content": "早安"},
		{"type": "date", "content": "2026.04.10 星期五"},
		{"type": "message", "time": "10:30", "name": "Bob", "content": "已收回訊息"},
	]


def test_parse_chat_file_appends_continuation_lines_only_to_messages(tmp_path):
	chat_file = tmp_path / "chat.txt"
	chat_file.write_text(
		"\n".join(
			[
				"2026.04.09 星期四",
				"09:00 Alice 第一行",
				"第二行",
				"2026.04.10 星期五",
				"日期後面的孤立文字",
				"10:00 Bob 新的一天",
			]
		),
		encoding="utf-8",
	)

	assert parseChatFile(chat_file) == [
		{"type": "date", "content": "2026.04.09 星期四"},
		{
			"type": "message",
			"time": "09:00",
			"name": "Alice",
			"content": "第一行\n第二行",
		},
		{"type": "date", "content": "2026.04.10 星期五"},
		{"type": "message", "time": "10:00", "name": "Bob", "content": "新的一天"},
	]


def test_message_reader_formats_date_rows_without_removing_original_text():
	dialog = object.__new__(message_reader.MessageReaderDialog)

	assert dialog._formatMessage({"type": "date", "content": "2026.04.09 星期四"}) == (
		"2026.04.09 星期四"
	)
	assert dialog._formatMessage(
		{"type": "message", "name": "Alice", "content": "早安", "time": "09:00"}
	) == "Alice 早安 09:00"


def test_message_reader_progress_counts_only_real_messages():
	dialog = object.__new__(message_reader.MessageReaderDialog)
	dialog._messages = [
		{"type": "date", "content": "2026.04.09 星期四"},
		{"type": "message", "name": "Alice", "content": "早安", "time": "09:00"},
		{"type": "date", "content": "2026.04.10 星期五"},
		{"type": "message", "name": "Bob", "content": "晚安", "time": "21:00"},
	]
	dialog._messageCount = 2
	dialog._messageIndexMap = [0, 1, 1, 2]

	dialog._pos = 0
	assert dialog._getProgressLabel() == "1 / 2"

	dialog._pos = 1
	assert dialog._getProgressLabel() == "1 / 2"

	dialog._pos = 2
	assert dialog._getProgressLabel() == "2 / 2"

	dialog._pos = 3
	assert dialog._getProgressLabel() == "2 / 2"


def test_message_reader_hides_progress_on_trailing_date_row():
	dialog = object.__new__(message_reader.MessageReaderDialog)
	dialog._messages = [
		{"type": "message", "name": "Alice", "content": "早安", "time": "09:00"},
		{"type": "date", "content": "2026.04.10 星期五"},
	]
	dialog._messageCount = 1
	dialog._messageIndexMap = [1, 1]
	dialog._pos = 1

	assert dialog._getProgressLabel() == ""


def test_message_reader_boundary_prompts_use_generic_item_wording():
	spoken = []
	dialog = object.__new__(message_reader.MessageReaderDialog)
	dialog._messages = [{"type": "date", "content": "2026.04.09 星期四"}]
	dialog._pos = 0
	dialog._speakMessage = spoken.append
	dialog._updateDisplay = lambda: None

	dialog._movePrevious()
	dialog._moveNext()

	assert spoken == ["已經是第一項", "已經是最後一項"]
