from __future__ import annotations

import ast
import re
from pathlib import Path


def _load_call_helpers():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	namespace = {"re": re}
	needed_assignments = {
		"_CJK_CHAR",
		"_CJK_SPACE_RE",
		"_CALL_OCR_LOG_NOISE_MARKERS",
		"_CALL_CHAT_CLOCK_RE",
		"_CALL_DURATION_RE",
		"_CALL_DURATION_NOISE_LINES",
	}
	needed_functions = {
		"_removeCJKSpaces",
		"_looksLikeOcrLogNoise",
		"_normalizeCallOcrLine",
		"_isChatClockTimeLine",
		"_isCallDurationFallbackNoiseLine",
		"_extractCallDuration",
		"_getCallAnnouncementFromOcr",
	}

	for node in module.body:
		if isinstance(node, ast.Assign):
			names = {target.id for target in node.targets if isinstance(target, ast.Name)}
			if names & needed_assignments:
				exec(
					compile(
						ast.Module(body=[node], type_ignores=[]),
						str(module_path),
						"exec",
					),
					namespace,
				)
		elif isinstance(node, ast.FunctionDef) and node.name in needed_functions:
			exec(
				compile(
					ast.Module(body=[node], type_ignores=[]),
					str(module_path),
					"exec",
				),
				namespace,
			)
	return namespace


helpers = _load_call_helpers()


def test_call_duration_ignores_chat_timestamp_only_lines():
	assert helpers["_extractCallDuration"]("午 11 : 40") is None
	assert helpers["_extractCallDuration"]("下午 11 : 46") is None
	assert helpers["_extractCallDuration"]("上牛 12 : 17") is None


def test_call_duration_handles_ocr_punctuation_variants():
	assert helpers["_extractCallDuration"]("00•.04\n下午 11 : 46") == "00:04"
	assert helpers["_extractCallDuration"]("00 : 1 3\n下午 11 : 47") == "00:13"
	assert helpers["_extractCallDuration"]("全選\n00 : 31\n下午 3 : 10") == "00:31"
	assert helpers["_extractCallDuration"]("已\n下午 12 : 45\n01 : 59") == "01:59"


def test_call_duration_ignores_log_ocr_noise():
	assert helpers["_extractCallDuration"](
		"- conng.conngManager._loaaconTlg 1 : zy : 45\n"
		"Loading config: C:\\lJsers\\chang\\AppData\\Roaming\\\n"
		"INFO - config.ConfigManager._loadConfig ( 1 1 : 29 : 45",
	) is None


def test_call_duration_ignores_message_body_with_ocr_clock_suffix():
	assert helpers["_extractCallDuration"]("已讀\n關於江同學的事情\n上牛 12 : 17") is None


def test_call_announcement_preserves_special_statuses():
	assert helpers["_getCallAnnouncementFromOcr"]("取消\n下午 3:14") == "取消的通話"
	assert helpers["_getCallAnnouncementFromOcr"]("無應答\n上午 9:41") == "無應答"
	assert helpers["_getCallAnnouncementFromOcr"]("未接來電\n下午 11:40") == "未接來電"


def test_call_announcement_formats_regular_calls_with_duration():
	assert helpers["_getCallAnnouncementFromOcr"]("00•.04\n下午 11 : 46") == "通話時間：00:04"
	assert helpers["_getCallAnnouncementFromOcr"]("00 : 12") == "通話時間：00:12"
	assert helpers["_getCallAnnouncementFromOcr"]("1:02:03\n下午 1:20") == "通話時間：01:02:03"
