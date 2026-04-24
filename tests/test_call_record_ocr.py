import ast
import re
import unittest
from pathlib import Path


def _load_call_helpers():
	source = Path("addon/appModules/line.py").read_text(encoding="utf-8-sig")
	module = ast.parse(source, filename="addon/appModules/line.py")
	needed = {
		"_CJK_CHAR",
		"_CJK_SPACE_RE",
		"_CALL_OCR_LOG_NOISE_MARKERS",
		"_CALL_CHAT_CLOCK_RE",
		"_CALL_DURATION_RE",
		"_CALL_DURATION_NOISE_LINES",
		"_removeCJKSpaces",
		"_looksLikeOcrLogNoise",
		"_normalizeCallOcrLine",
		"_isChatClockTimeLine",
		"_isCallDurationFallbackNoiseLine",
		"_extractCallDuration",
		"_getCallAnnouncementFromOcr",
	}
	selected = []
	for node in module.body:
		if isinstance(node, ast.Assign):
			for target in node.targets:
				if isinstance(target, ast.Name) and target.id in needed:
					selected.append(node)
					break
		elif isinstance(node, ast.FunctionDef) and node.name in needed:
			selected.append(node)
	namespace = {"re": re}
	exec(compile(ast.Module(body=selected, type_ignores=[]), "<call_helpers>", "exec"), namespace)
	return namespace


HELPERS = _load_call_helpers()


class CallRecordOcrTests(unittest.TestCase):
	def test_duration_call_record_is_normalized(self):
		text = "已讀\n下午 IO : 25\n00 : 52"
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"](text),
			"通話時間：00:52",
		)

	def test_duration_call_record_with_ocr_noise_is_normalized(self):
		text = "已謴\n下午 12 : 27\n00 : 08"
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"](text),
			"通話時間：00:08",
		)

	def test_missed_call_state_is_detected(self):
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"]("未接來電"),
			"未接來電",
		)

	def test_no_answer_call_state_is_detected(self):
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"]("上午 3:10\n無應答"),
			"無應答",
		)

	def test_cancelled_call_state_accepts_optional_de(self):
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"]("取消通話"),
			"取消的通話",
		)
		self.assertEqual(
			HELPERS["_getCallAnnouncementFromOcr"]("取消的通話"),
			"取消的通話",
		)

	def test_plain_cancel_text_is_not_treated_as_call_record(self):
		self.assertIsNone(HELPERS["_getCallAnnouncementFromOcr"]("取消"))

	def test_log_ocr_noise_is_not_treated_as_call_record(self):
		text = (
			"- conng.conngManager._loaaconTlg 1 : zy : 45\n"
			"Loading config: C:\\lJsers\\chang\\AppData\\Roaming\\\n"
			"INFO - config.ConfigManager._loadConfig ( 1 1 : 29 : 45\n"
			"Confiq loaded (after upqrade, and in the state itwill"
		)
		self.assertIsNone(HELPERS["_getCallAnnouncementFromOcr"](text))

	def test_message_body_with_ocr_clock_suffix_is_not_treated_as_call_record(self):
		text = "已讀\n關於江同學的事情\n上牛 12 : 17"
		self.assertIsNone(HELPERS["_getCallAnnouncementFromOcr"](text))


if __name__ == "__main__":
	unittest.main()
