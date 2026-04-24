from __future__ import annotations

import ast
import re
import sys
import time
from pathlib import Path
from types import SimpleNamespace


def _load_line_symbols(*, assignment_names=(), function_names=(), namespace=None):
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	ns = {} if namespace is None else dict(namespace)

	for node in module.body:
		if isinstance(node, ast.Assign):
			names = {target.id for target in node.targets if isinstance(target, ast.Name)}
			if names & set(assignment_names):
				exec(
					compile(
						ast.Module(body=[node], type_ignores=[]),
						str(module_path),
						"exec",
					),
					ns,
				)
		elif isinstance(node, ast.FunctionDef) and node.name in set(function_names):
			exec(
				compile(
					ast.Module(body=[node], type_ignores=[]),
					str(module_path),
					"exec",
				),
				ns,
			)
	return ns


class _Log:
	def debug(self, *args, **kwargs):
		pass

	def info(self, *args, **kwargs):
		pass

	def debugWarning(self, *args, **kwargs):
		pass


def test_notes_window_context_uses_only_fresh_cache():
	ns = _load_line_symbols(
		assignment_names={
			"_NOTES_WINDOW_KEYWORDS",
			"_NOTES_OCR_KEYWORDS",
			"_NOTES_OCR_CACHE_TTL",
			"_notesWindowDetectionCache",
		},
		function_names={"_isNotesWindowContext"},
		namespace={"log": _Log(), "time": time},
	)
	walker = SimpleNamespace(GetParentElement=lambda _element: None)
	ns["_getForegroundWindowInfo"] = lambda: (101, "line", (0, 0, 1200, 800))

	ns["_notesWindowDetectionCache"] = {
		"key": (101, "line", (0, 0, 1200, 800)),
		"expiresAt": time.monotonic() + 5.0,
		"isNotesWindow": True,
	}
	assert ns["_isNotesWindowContext"](object(), walker) == (True, "line")

	ns["_notesWindowDetectionCache"] = {
		"key": (101, "line", (0, 0, 1200, 800)),
		"expiresAt": time.monotonic() - 1.0,
		"isNotesWindow": True,
	}
	assert ns["_isNotesWindowContext"](object(), walker, allowOcr=False) == (False, "line")


def test_rect_visibility_checks_overlap_with_foreground_window():
	ns = _load_line_symbols(
		function_names={"_rectsIntersect", "_isRectVisibleInForegroundWindow"},
	)
	ns["_getForegroundWindowInfo"] = lambda: (101, "line", (100, 100, 400, 400))

	assert ns["_isRectVisibleInForegroundWindow"](150, 150, 300, 300) is True
	assert ns["_isRectVisibleInForegroundWindow"](450, 450, 500, 500) is False


def test_should_dismiss_copy_read_menu_requires_same_process_foreground():
	ns = _load_line_symbols(
		function_names={"_shouldDismissCopyReadMenu"},
		namespace={
			"ctypes": SimpleNamespace(
				windll=SimpleNamespace(
					user32=SimpleNamespace(GetForegroundWindow=lambda: 202),
				),
			),
		},
	)
	ns["_getWindowProcessId"] = lambda hwnd: {
		101: 7001,
		202: 7001,
		303: 7002,
	}.get(hwnd, 0)

	assert ns["_shouldDismissCopyReadMenu"](101) is True
	ns["ctypes"].windll.user32.GetForegroundWindow = lambda: 303
	assert ns["_shouldDismissCopyReadMenu"](101) is False
	ns["ctypes"].windll.user32.GetForegroundWindow = lambda: 0
	assert ns["_shouldDismissCopyReadMenu"](101) is False


def test_message_probe_hit_test_rejects_edit_fields_but_allows_qt_hit_test_drift():
	class _Element:
		def __init__(self, runtime_id, control_type=50007):
			self.runtime_id = runtime_id
			self.CurrentControlType = control_type

		def GetRuntimeId(self):
			return self.runtime_id

	class _Walker:
		def __init__(self, parents):
			self.parents = parents

		def GetParentElement(self, element):
			return self.parents.get(element)

	class _Client:
		def __init__(self, hit_element, walker):
			self.hit_element = hit_element
			self.RawViewWalker = walker

		def ElementFromPoint(self, _point):
			return self.hit_element

	target = _Element((1, 2, 3))
	target_child = _Element((1, 2, 4))
	other = _Element((9, 9, 9))
	edit = _Element((8, 8, 8), control_type=50004)
	walker = _Walker({target_child: target})
	client = _Client(target_child, walker)
	handler = SimpleNamespace(clientObject=client)

	ns = _load_line_symbols(
		function_names={
			"_normalizeRuntimeId",
			"_getElementRuntimeId",
			"_messageProbePointHitsTargetElement",
		},
		namespace={
			"ctypes": SimpleNamespace(
				wintypes=SimpleNamespace(
					POINT=lambda x, y: SimpleNamespace(x=x, y=y),
				),
			),
			"log": _Log(),
		},
	)
	target_id = ns["_getElementRuntimeId"](target)

	assert ns["_messageProbePointHitsTargetElement"](handler, target, 10, 20, target_id) is True
	client.hit_element = other
	assert ns["_messageProbePointHitsTargetElement"](handler, target, 10, 20, target_id) is True
	client.hit_element = edit
	assert ns["_messageProbePointHitsTargetElement"](handler, target, 10, 20, target_id) is False


def test_message_probe_hit_test_preserves_existing_behavior_when_uia_hit_test_fails():
	class _Client:
		RawViewWalker = None

		def ElementFromPoint(self, _point):
			raise RuntimeError("UIA unavailable")

	ns = _load_line_symbols(
		function_names={"_messageProbePointHitsTargetElement"},
		namespace={
			"ctypes": SimpleNamespace(
				wintypes=SimpleNamespace(
					POINT=lambda x, y: SimpleNamespace(x=x, y=y),
				),
			),
			"log": _Log(),
		},
	)

	assert ns["_messageProbePointHitsTargetElement"](
		SimpleNamespace(clientObject=_Client()),
		object(),
		10,
		20,
		(1, 2, 3),
	) is True


def test_extract_matched_message_context_menu_labels_ignores_message_body_text():
	known_labels = {"回覆", "複製", "分享", "刪除", "收回"}
	ns = _load_line_symbols(
		function_names={"_extractMatchedMessageContextMenuLabels"},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"_matchMessageContextMenuLabel": lambda text: text if text in known_labels else None,
		},
	)

	popup_lines, line_matches, matched_labels = ns["_extractMatchedMessageContextMenuLabels"](
		"本次更新新增了訊息右鍵選單\n天、回覆、複製、收回等重要站點\n日期：2026 年 3 月 27 日",
	)

	assert popup_lines[1] == "天、回覆、複製、收回等重要站點"
	assert line_matches[1] == ("天、回覆、複製、收回等重要站點", None)
	assert matched_labels == []


def test_extract_matched_message_context_menu_labels_accepts_real_menu_rows():
	known_labels = {"回覆", "複製", "分享", "刪除", "收回"}
	ns = _load_line_symbols(
		function_names={"_extractMatchedMessageContextMenuLabels"},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"_matchMessageContextMenuLabel": lambda text: text if text in known_labels else None,
		},
	)

	_popup_lines, line_matches, matched_labels = ns["_extractMatchedMessageContextMenuLabels"](
		"回覆\n複製\n分享",
	)

	assert line_matches == [("回覆", "回覆"), ("複製", "複製"), ("分享", "分享")]
	assert matched_labels == ["回覆", "複製", "分享"]


def test_resolve_popup_menu_label_click_point_returns_aligned_click_target():
	ns = _load_line_symbols(
		function_names={"_resolvePopupMenuLabelClickPoint"},
		namespace={"log": _Log()},
	)

	row_rects = [
		(639, 439, 837, 479),
		(639, 480, 837, 520),
		(639, 520, 837, 560),
	]
	calls = {}

	def _fake_build_menu_elements(lines, popup_rect, rowRects=None):
		calls["lines"] = lines
		calls["popup_rect"] = popup_rect
		calls["row_rects"] = rowRects
		return [
			{"name": "回覆", "clickPoint": (738, 459)},
			{"name": "複製", "clickPoint": (738, 500)},
			{"name": "分享", "clickPoint": (738, 540)},
		]

	target = ns["_resolvePopupMenuLabelClickPoint"](
		"複製",
		[
			{"text": "50", "rect": (700, 430, 732, 450)},
			{"text": "回覆", "rect": (706, 449, 742, 469)},
			{"text": "複製", "rect": (706, 489, 742, 511)},
			{"text": "分享", "rect": (706, 530, 742, 550)},
		],
		(624, 415, 852, 860),
		popupRowRects=row_rects,
		buildMenuElements=_fake_build_menu_elements,
	)

	assert calls["row_rects"] == row_rects
	assert target == {
		"clickPoint": (738, 500),
		"index": 1,
		"count": 3,
	}


def test_build_message_bubble_click_positions_preserves_existing_priority_before_new_left_fallbacks():
	ns = _load_line_symbols(
		function_names={"_buildMessageBubbleClickPositions"},
	)

	positions = ns["_buildMessageBubbleClickPositions"](
		(547, 274, 1102, 397),
		0,
		1000,
	)

	assert [label for _x, _y, label in positions[:6]] == [
		"1/6-left",
		"5/6-right",
		"1/4-left",
		"3/4-right",
		"9/10-right",
		"7/8-right",
	]
	assert [label for _x, _y, label in positions[6:8]] == [
		"1/10-left",
		"1/8-left",
	]
	assert [label for _x, _y, label in positions[-1:]] == [
		"center",
	]


def test_build_message_bubble_click_positions_adds_vertical_probes_only_when_requested():
	ns = _load_line_symbols(
		function_names={"_buildMessageBubbleClickPositions"},
	)

	basePositions = ns["_buildMessageBubbleClickPositions"](
		(547, 274, 1102, 397),
		0,
		1000,
	)
	verticalPositions = ns["_buildMessageBubbleClickPositions"](
		(547, 274, 1102, 397),
		0,
		1000,
		includeVerticalOffsets=True,
	)

	assert [label for _x, _y, label in basePositions] == [
		"1/6-left",
		"5/6-right",
		"1/4-left",
		"3/4-right",
		"9/10-right",
		"7/8-right",
		"1/10-left",
		"1/8-left",
		"center",
	]
	assert [label for _x, _y, label in verticalPositions] == [
		"1/6-left",
		"5/6-right",
		"1/4-left",
		"3/4-right",
		"9/10-right",
		"7/8-right",
		"1/10-left",
		"1/8-left",
		"1/10-top",
		"1/8-top",
		"1/6-top",
		"1/10-bottom",
		"1/8-bottom",
		"5/6-bottom",
		"center",
	]


def test_build_message_bubble_click_positions_prioritizes_right_edge_vertical_probes_for_lower_bubbles():
	ns = _load_line_symbols(
		function_names={"_buildMessageBubbleClickPositions"},
	)

	verticalPositions = ns["_buildMessageBubbleClickPositions"](
		(547, 274, 1102, 397),
		0,
		735,
		includeVerticalOffsets=True,
	)

	assert [label for _x, _y, label in verticalPositions] == [
		"1/6-left",
		"5/6-right",
		"1/4-left",
		"3/4-right",
		"9/10-right",
		"7/8-right",
		"1/10-left",
		"1/8-left",
		"5/6-top",
		"5/6-bottom",
		"1/10-top",
		"1/8-top",
		"1/6-top",
		"1/10-bottom",
		"1/8-bottom",
		"center",
	]


def test_is_message_bubble_metadata_ocr_line_filters_timestamps_and_read_receipts():
	ns = _load_line_symbols(
		function_names={
			"_normalizeMessageBubbleOcrLine",
			"_isMessageBubbleMetadataOcrLine",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"re": re,
		},
	)

	assert ns["_isMessageBubbleMetadataOcrLine"]("下午 12 : 38") is True
	assert ns["_isMessageBubbleMetadataOcrLine"](", 下午 12 : 38") is True
	assert ns["_isMessageBubbleMetadataOcrLine"]("已讀") is True
	assert ns["_isMessageBubbleMetadataOcrLine"]("已讀 下午 12 ℃ 5") is True
	assert ns["_isMessageBubbleMetadataOcrLine"]("謝謝 !") is False


def test_looks_like_line_date_separator_text_matches_line_style_dates():
	ns = _load_line_symbols(
		assignment_names={"_LINE_DATE_SEPARATOR_RE"},
		function_names={
			"_normalizeLineDateSeparatorOcrText",
			"_getSpokenLineDateSeparatorText",
			"_looksLikeLineDateSeparatorText",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"re": re,
		},
	)

	assert ns["_looksLikeLineDateSeparatorText"]("今天") is True
	assert ns["_looksLikeLineDateSeparatorText"]("昨天") is True
	assert ns["_looksLikeLineDateSeparatorText"]("阼天") is True
	assert ns["_looksLikeLineDateSeparatorText"]("4 月 8 日 ( 三 )") is True
	assert ns["_looksLikeLineDateSeparatorText"]("2026 年 4 月 8 日 ( 三 )") is True
	assert ns["_looksLikeLineDateSeparatorText"]("今天下午 5 : 28") is False
	assert ns["_looksLikeLineDateSeparatorText"]("下午 5 : 28") is False
	assert ns["_looksLikeLineDateSeparatorText"]("4 月 8 日 ( 三 ) 下午 5 : 28") is False
	assert ns["_getSpokenLineDateSeparatorText"]("阼天") == "昨天"
	assert ns["_getSpokenLineDateSeparatorText"]("昨天") == "昨天"
	assert ns["_getSpokenLineDateSeparatorText"]("4 月 8 日 ( 三 )") == "4 月 8 日 ( 三 )"


def test_is_centered_line_date_separator_ocr_requires_compact_centered_geometry():
	ns = _load_line_symbols(
		assignment_names={"_LINE_DATE_SEPARATOR_RE"},
		function_names={
			"_normalizeLineDateSeparatorOcrText",
			"_looksLikeLineDateSeparatorText",
			"_isCenteredLineDateSeparatorOcr",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"re": re,
		},
	)

	assert (
		ns["_isCenteredLineDateSeparatorOcr"](
			"今天",
			[{"text": "今天", "rect": (803, 164, 847, 186)}],
			(547, 142, 1102, 211),
		)
		is True
	)
	assert (
		ns["_isCenteredLineDateSeparatorOcr"](
			"阼天",
			[{"text": "阼天", "rect": (803, 164, 847, 186)}],
			(547, 142, 1102, 211),
		)
		is True
	)
	assert (
		ns["_isCenteredLineDateSeparatorOcr"](
			"4 月 8 日 ( 三 )",
			[{"text": "4 月 8 日 ( 三 )", "rect": (782, 164, 867, 186)}],
			(547, 142, 1102, 211),
		)
		is True
	)
	assert (
		ns["_isCenteredLineDateSeparatorOcr"](
			"4 月 8 日 ( 三 )",
			[{"text": "4 月 8 日 ( 三 )", "rect": (980, 164, 1065, 186)}],
			(547, 142, 1102, 211),
		)
		is False
	)
	assert (
		ns["_isCenteredLineDateSeparatorOcr"](
			"OK",
			[{"text": "OK", "rect": (982, 336, 1038, 359)}],
			(547, 274, 1102, 397),
		)
		is False
	)


def test_build_message_bubble_ocr_click_positions_targets_side_padding_for_right_aligned_content():
	ns = _load_line_symbols(
		function_names={
			"_normalizeMessageBubbleOcrLine",
			"_isMessageBubbleMetadataOcrLine",
			"_buildMessageBubbleOcrClickPositions",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
			"re": re,
		},
	)

	positions = ns["_buildMessageBubbleOcrClickPositions"](
		[
			{"text": "謝謝 !", "rect": (996, 306, 1038, 333)},
			{"text": ", 下午 12 : 38", "rect": (982, 336, 1060, 359)},
		],
		(547, 274, 1102, 397),
		0,
		735,
	)

	assert [label for _x, _y, label in positions] == [
		"ocr-right-center",
		"ocr-right-upper",
		"ocr-right-lower",
	]
	assert all(x > 1038 for x, _y, _label in positions)
	assert all(x < 1102 for x, _y, _label in positions)


def test_has_exhausted_message_bubble_fallback_probes_ignores_trailing_center():
	ns = _load_line_symbols(
		function_names={
			"_buildMessageBubbleClickPositions",
			"_hasExhaustedMessageBubbleFallbackProbes",
		},
	)

	positions = ns["_buildMessageBubbleClickPositions"](
		(547, 274, 1102, 397),
		0,
		1000,
		includeVerticalOffsets=True,
	)

	assert ns["_hasExhaustedMessageBubbleFallbackProbes"](5, positions) is False
	assert ns["_hasExhaustedMessageBubbleFallbackProbes"](7, positions) is False
	assert ns["_hasExhaustedMessageBubbleFallbackProbes"](12, positions) is False
	assert ns["_hasExhaustedMessageBubbleFallbackProbes"](13, positions) is True
	assert ns["_hasExhaustedMessageBubbleFallbackProbes"](14, positions) is True


def test_extract_recall_dialog_action_labels_handles_modern_dialog_without_matching_body_text():
	ns = _load_line_symbols(
		function_names={
			"_normalizeRecallDialogLine",
			"_matchRecallDialogActionLabel",
			"_extractRecallDialogActionLabels",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
		},
	)

	labels = ns["_extractRecallDialogActionLabels"](
		"確定要收回訊息嗎？\n"
		"您可無痕收回此則未讀訊息，對方不會收到任何提醒。\n"
		"依對方使用的LINE版本而定，有可能無法收回訊息。\n"
		"無痕收回 Premium\n"
		"收回\n"
		"關閉",
	)

	assert labels == ["無痕收回", "收回", "取消"]


def test_get_recall_confirmation_prompt_marks_stealth_option_as_premium():
	ns = _load_line_symbols(
		function_names={"_getRecallConfirmationPrompt"},
		namespace={"_": lambda text: text},
	)

	assert ns["_getRecallConfirmationPrompt"]({"收回", "取消"}) == "確認要收回嗎？按 Y 收回，按 N 取消"
	assert (
		ns["_getRecallConfirmationPrompt"](
			{"收回", "取消"},
			isModernDialog=True,
		)
		== "確認要收回嗎？按 Y 收回，按 N 取消"
	)
	assert (
		ns["_getRecallConfirmationPrompt"](
			{"無痕收回", "收回", "取消"},
			isModernDialog=True,
		)
		== "確認要收回嗎？按 Y 收回，按 N 取消，按 P 無痕收回，需要 Premium"
	)


def test_extract_photo_text_consent_action_labels_ignores_title_and_body_mentions_of_agree():
	ns = _load_line_symbols(
		function_names={
			"_normalizePhotoTextConsentDialogLine",
			"_matchPhotoTextConsentActionLabel",
			"_extractPhotoTextConsentActionLabels",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
		},
	)

	labels = ns["_extractPhotoTextConsentActionLabels"](
		"同意提供照片\n"
		"使用本功能會將照片傳至本公司伺服器進行處理\n"
		"您要同意本功能的服務規定並開始使用嗎？\n"
		"同意\n"
		"不同意",
	)

	assert labels == ["同意", "不同意"]


def test_is_photo_text_consent_dialog_text_requires_upload_notice_and_buttons():
	ns = _load_line_symbols(
		function_names={"_isPhotoTextConsentDialogText"},
		namespace={"_removeCJKSpaces": lambda text: text.replace(" ", "")},
	)

	assert (
		ns["_isPhotoTextConsentDialogText"](
			"同意提供照片\n"
			"使用本功能會將照片傳至本公司伺服器進行處理\n"
			"您要同意本功能的服務規定並開始使用嗎？\n"
			"同意\n"
			"不同意",
			["同意", "不同意"],
		)
		is True
	)
	assert (
		ns["_isPhotoTextConsentDialogText"](
			"同意提供照片\n同意\n不同意",
			["同意", "不同意"],
		)
		is False
	)


def test_get_photo_text_consent_prompt_mentions_upload_notice_and_a_d_shortcuts():
	ns = _load_line_symbols(
		function_names={"_getPhotoTextConsentPrompt"},
		namespace={"_": lambda text: text},
	)

	assert (
		ns["_getPhotoTextConsentPrompt"]() == "轉為文字會將照片上傳到 LINE 伺服器處理。按 A 同意，按 D 不同意"
	)


def test_extract_photo_text_consent_action_click_points_use_button_rows_not_dialog_title():
	ns = _load_line_symbols(
		function_names={
			"_normalizePhotoTextConsentDialogLine",
			"_matchPhotoTextConsentActionLabel",
			"_rectsIntersect",
			"_extractPhotoTextConsentActionClickPoints",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
		},
	)

	points = ns["_extractPhotoTextConsentActionClickPoints"](
		[
			{"text": "同意提供照片", "rect": (360, 240, 560, 292)},
			{"text": "同意", "rect": (432, 600, 516, 636)},
			{"text": "不同意", "rect": (544, 600, 652, 636)},
		],
		(320, 220, 760, 760),
	)

	assert points == {
		"同意": {
			"clickPoint": (474, 618),
			"rect": (432, 600, 516, 636),
		},
		"不同意": {
			"clickPoint": (598, 618),
			"rect": (544, 600, 652, 636),
		},
	}


def test_is_modern_recall_dialog_text_accepts_compact_two_button_modern_layout():
	ns = _load_line_symbols(
		function_names={"_isModernRecallDialogText"},
		namespace={"_removeCJKSpaces": lambda text: text.replace(" ", "")},
	)

	assert (
		ns["_isModernRecallDialogText"](
			"確定要收回此訊息嗎 ?\n"
			"收回已讀訊息時, 對方將會收到通知。\n"
			"依對方使用的LINE版本而定, 有可能無法收回訊息。\n"
			"收回\n"
			"關閉",
			["收回", "取消"],
		)
		is True
	)


def test_is_compact_modern_recall_dialog_requires_two_button_modern_state():
	ns = _load_line_symbols(
		function_names={"_isCompactModernRecallDialog"},
	)

	assert ns["_isCompactModernRecallDialog"](["收回", "取消"], isModernDialog=True) is True
	assert ns["_isCompactModernRecallDialog"](["無痕收回", "收回", "取消"], isModernDialog=True) is False
	assert ns["_isCompactModernRecallDialog"](["收回", "取消"], isModernDialog=False) is False


def test_try_invoke_uia_element_prefers_direct_invoke_without_generated_stubs():
	ns = _load_line_symbols(
		function_names={"_invokeUIAInvokePattern", "_tryInvokeUIAElement"},
	)

	class _Pattern:
		def __init__(self):
			self.invoked = 0

		def Invoke(self):
			self.invoked += 1

	class _Element:
		def __init__(self):
			self.pattern_ids = []
			self.pattern = _Pattern()

		def GetCurrentPattern(self, patternId):
			self.pattern_ids.append(patternId)
			return self.pattern

	element = _Element()
	assert ns["_tryInvokeUIAElement"](element) is True
	assert element.pattern.invoked == 1
	assert element.pattern_ids == [10000]


def test_invoke_uia_invoke_pattern_falls_back_to_query_interface_without_comtypes_gen():
	class _IUnknown:
		pass

	class _Comtypes:
		IUnknown = _IUnknown

		@staticmethod
		def GUID(value):
			return value

		@staticmethod
		def COMMETHOD(*args):
			return ("COMMETHOD", args)

	ns = _load_line_symbols(
		function_names={"_invokeUIAInvokePattern"},
		namespace={
			"comtypes": _Comtypes,
			"ctypes": SimpleNamespace(c_long=object()),
		},
	)

	class _Invokable:
		def __init__(self):
			self.invoked = 0

		def Invoke(self):
			self.invoked += 1

	class _Pattern:
		def __init__(self):
			self.requestedInterface = None
			self.invokable = _Invokable()

		def QueryInterface(self, interfaceType):
			self.requestedInterface = interfaceType
			return self.invokable

	pattern = _Pattern()
	assert ns["_invokeUIAInvokePattern"](pattern) is True
	assert pattern.invokable.invoked == 1
	assert pattern.requestedInterface._iid_ == "{FB377FBE-8EA6-46D5-9C73-6499642D3059}"


def test_extract_ocr_rect_like_unions_word_rects_when_line_rect_is_missing():
	ns = _load_line_symbols(function_names={"_extractOcrRectLike"})

	class _WordRect:
		def __init__(self, left, top, right, bottom):
			self.left = left
			self.top = top
			self.right = right
			self.bottom = bottom

	class _Word:
		def __init__(self, rect):
			self.boundingRect = rect

	class _Line:
		def __init__(self):
			self.words = [
				_Word(_WordRect(480, 620, 520, 650)),
				_Word(_WordRect(522, 618, 560, 648)),
			]

	assert ns["_extractOcrRectLike"](_Line()) == (480, 618, 560, 650)


def test_extract_ocr_rect_like_supports_uppercase_xywh_rects():
	ns = _load_line_symbols(function_names={"_extractOcrRectLike"})

	class _Rect:
		X = 996
		Y = 306
		Width = 42
		Height = 27

	class _Line:
		boundingRect = _Rect()

	assert ns["_extractOcrRectLike"](_Line()) == (996, 306, 1038, 333)


def test_extract_ocr_rect_like_supports_tuple_location_rects():
	ns = _load_line_symbols(function_names={"_extractOcrRectLike"})

	class _Line:
		location = (996, 306, 42, 27)

	assert ns["_extractOcrRectLike"](_Line()) == (996, 306, 1038, 333)


def test_extract_ocr_rect_like_unions_capitalized_word_polygons():
	ns = _load_line_symbols(function_names={"_extractOcrRectLike"})

	class _Word:
		def __init__(self, polygon):
			self.polygon = polygon

	class _Line:
		def __init__(self):
			self.Words = [
				_Word([480, 620, 520, 620, 520, 650, 480, 650]),
				_Word([522, 618, 560, 618, 560, 648, 522, 648]),
			]

	assert ns["_extractOcrRectLike"](_Line()) == (480, 618, 560, 650)


def test_extract_ocr_lines_supports_nvda_lines_words_result_shape():
	ns = _load_line_symbols(
		function_names={"_extractOcrRectLike", "_extractOcrLines"},
		namespace={"_removeCJKSpaces": lambda text: text.replace(" ", "")},
	)

	class _Word:
		def __init__(self, offset, left, top, width, height):
			self.offset = offset
			self.left = left
			self.top = top
			self.width = width
			self.height = height

	text = "謝謝 !\n, 下午 12 : 38\n"
	result = SimpleNamespace(
		text=text,
		lines=[len("謝謝 !\n"), len(text)],
		words=[
			_Word(0, 996, 306, 42, 27),
			_Word(len("謝謝 !\n"), 982, 336, 78, 23),
		],
	)

	assert ns["_extractOcrLines"](result) == [
		{"text": "謝謝 !", "rect": (996, 306, 1038, 333)},
		{"text": ", 下午 12 : 38", "rect": (982, 336, 1060, 359)},
	]


def test_infer_recall_dialog_targets_by_geometry_recovers_modern_unlabeled_buttons():
	ns = _load_line_symbols(
		function_names={
			"_rectIntersectionArea",
			"_rectIoU",
			"_isCompactModernRecallDialog",
			"_inferRecallDialogTargetsByGeometry",
		},
	)

	candidates = [
		{
			"element": "junk-top",
			"rect": (150, 160, 260, 190),
			"controlType": 50003,
			"hasInvoke": False,
		},
		{
			"element": "stealth",
			"rect": (150, 290, 450, 332),
			"controlType": 50000,
			"hasInvoke": True,
		},
		{
			"element": "recall",
			"rect": (150, 346, 450, 388),
			"controlType": 50000,
			"hasInvoke": True,
		},
		{
			"element": "cancel-text",
			"rect": (238, 420, 362, 446),
			"controlType": 50003,
			"hasInvoke": False,
		},
	]

	inferred = ns["_inferRecallDialogTargetsByGeometry"](
		candidates,
		(100, 120, 500, 500),
		["無痕收回", "收回"],
		isModernDialog=True,
	)

	assert inferred["無痕收回"]["element"] == "stealth"
	assert inferred["收回"]["element"] == "recall"


def test_infer_recall_dialog_targets_by_geometry_recovers_legacy_primary_button():
	ns = _load_line_symbols(
		function_names={
			"_rectIntersectionArea",
			"_rectIoU",
			"_isCompactModernRecallDialog",
			"_inferRecallDialogTargetsByGeometry",
		},
	)

	candidates = [
		{
			"element": "body",
			"rect": (140, 180, 430, 220),
			"controlType": 50003,
			"hasInvoke": False,
		},
		{
			"element": "recall",
			"rect": (150, 320, 450, 364),
			"controlType": 50000,
			"hasInvoke": True,
		},
	]

	inferred = ns["_inferRecallDialogTargetsByGeometry"](
		candidates,
		(100, 120, 500, 500),
		["收回", "取消"],
		isModernDialog=False,
	)

	assert inferred["收回"]["element"] == "recall"


def test_infer_recall_dialog_targets_by_geometry_prefers_compact_modern_button_zone():
	ns = _load_line_symbols(
		function_names={
			"_rectIntersectionArea",
			"_rectIoU",
			"_isCompactModernRecallDialog",
			"_inferRecallDialogTargetsByGeometry",
		},
	)

	candidates = [
		{
			"element": "topCard",
			"rect": (150, 260, 450, 304),
			"controlType": 50000,
			"hasInvoke": True,
		},
		{
			"element": "recallButton",
			"rect": (150, 352, 450, 404),
			"controlType": 50000,
			"hasInvoke": True,
		},
	]

	inferred = ns["_inferRecallDialogTargetsByGeometry"](
		candidates,
		(100, 120, 500, 560),
		["收回", "取消"],
		isModernDialog=True,
	)

	assert inferred["收回"]["element"] == "recallButton"


def test_get_recall_dialog_fallback_click_point_covers_legacy_and_modern_layouts():
	ns = _load_line_symbols(
		function_names={
			"_isCompactModernRecallDialog",
			"_getRecallDialogFallbackClickPoint",
		},
	)

	assert ns["_getRecallDialogFallbackClickPoint"]("收回", (100, 200, 500, 600), False) == (300, 432)
	assert ns["_getRecallDialogFallbackClickPoint"]("取消", (100, 200, 500, 600), False) is None
	assert ns["_getRecallDialogFallbackClickPoint"]("無痕收回", (100, 200, 500, 600), True) == (300, 396)
	assert ns["_getRecallDialogFallbackClickPoint"](
		"收回",
		(100, 200, 500, 600),
		True,
		["收回", "取消"],
	) == (300, 456)
	assert ns["_getRecallDialogFallbackClickPoint"]("取消", (100, 200, 500, 600), True) is None


def test_extract_recall_dialog_action_click_points_uses_ocr_label_centers():
	ns = _load_line_symbols(
		function_names={
			"_normalizeRecallDialogLine",
			"_matchRecallDialogActionLabel",
			"_rectsIntersect",
			"_extractRecallDialogActionClickPoints",
		},
		namespace={
			"_removeCJKSpaces": lambda text: text.replace(" ", ""),
		},
	)

	points = ns["_extractRecallDialogActionClickPoints"](
		[
			{"text": "收回訊息。", "rect": (180, 260, 300, 292)},
			{"text": "收回", "rect": (470, 618, 550, 648)},
			{"text": "關閉", "rect": (472, 676, 548, 706)},
		],
		(320, 220, 700, 760),
	)

	assert points == {
		"收回": {
			"clickPoint": (510, 633),
			"rect": (470, 618, 550, 648),
		},
		"取消": {
			"clickPoint": (510, 691),
			"rect": (472, 676, 548, 706),
		},
	}


def test_begin_recall_confirmation_binds_y_n_p_shortcuts():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	begin_method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_beginRecallConfirmation"
	)

	bind_calls = set()
	for node in ast.walk(begin_method):
		if not (
			isinstance(node, ast.Call)
			and isinstance(node.func, ast.Attribute)
			and node.func.attr == "bindGesture"
			and len(node.args) >= 2
		):
			continue
		first_arg, second_arg = node.args[:2]
		if all(
			isinstance(arg, ast.Constant) and isinstance(arg.value, str) for arg in (first_arg, second_arg)
		):
			bind_calls.add((first_arg.value, second_arg.value))

	assert ("kb:y", "confirmRecall") in bind_calls
	assert ("kb:n", "cancelRecall") in bind_calls
	assert ("kb:p", "stealthRecall") in bind_calls


def test_begin_photo_text_consent_binds_a_d_shortcuts():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	begin_method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_beginPhotoTextConsent"
	)

	bind_calls = set()
	for node in ast.walk(begin_method):
		if not (
			isinstance(node, ast.Call)
			and isinstance(node.func, ast.Attribute)
			and node.func.attr == "bindGesture"
			and len(node.args) >= 2
		):
			continue
		first_arg, second_arg = node.args[:2]
		if all(
			isinstance(arg, ast.Constant) and isinstance(arg.value, str) for arg in (first_arg, second_arg)
		):
			bind_calls.add((first_arg.value, second_arg.value))

	assert ("kb:a", "acceptPhotoTextConsent") in bind_calls
	assert ("kb:d", "declinePhotoTextConsent") in bind_calls


def test_end_recall_confirmation_defers_user_feedback_until_post_click_verification():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	end_method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_endRecallConfirmation"
	)

	calls = set()
	for node in ast.walk(end_method):
		if not (
			isinstance(node, ast.Call)
			and isinstance(node.func, ast.Attribute)
			and isinstance(node.func.value, ast.Name)
			and node.func.value.id == "self"
		):
			continue
		calls.add(node.func.attr)

	assert "_performRecallConfirmationAction" in calls
	assert "_scheduleRecallCompletionAnnouncement" in calls
	assert "_clearRecallConfirmationBindings" not in calls


def test_perform_recall_confirmation_action_prefers_ocr_click_point_for_legacy_recall():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_performRecallConfirmationAction"
	)
	compact_helper = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_isCompactModernRecallDialog"
	)
	ns = {
		"log": _Log(),
		"_getRecallDialogFallbackClickPoint": lambda *args, **kwargs: None,
		"_": lambda text: text,
	}
	exec(
		compile(
			ast.Module(body=[compact_helper, method], type_ignores=[]),
			str(module_path),
			"exec",
		),
		ns,
	)
	perform = ns["_performRecallConfirmationAction"]
	clicks = []

	class _Self:
		def _refreshRecallConfirmationState(self):
			return {
				"targets": {
					"收回": {
						"element": object(),
						"rect": (100, 200, 300, 320),
						"clickPoint": (210, 278),
					},
				},
				"isModernDialog": False,
				"hwnd": 123,
				"dialogRect": (50, 50, 350, 350),
			}

		def _invokeElement(self, *args, **kwargs):
			raise AssertionError("legacy recall should click the OCR hit before invoking UIA")

		def _clickAtPosition(self, *args, **kwargs):
			clicks.append((args, kwargs))

	assert perform(_Self(), "收回") is True
	assert clicks == [((210, 278), {"hwnd": 123})]


def test_perform_recall_confirmation_action_prefers_ocr_click_point_for_modern_recall():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_performRecallConfirmationAction"
	)
	compact_helper = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_isCompactModernRecallDialog"
	)
	ns = {
		"log": _Log(),
		"_getRecallDialogFallbackClickPoint": lambda *args, **kwargs: None,
		"_": lambda text: text,
	}
	exec(
		compile(
			ast.Module(body=[compact_helper, method], type_ignores=[]),
			str(module_path),
			"exec",
		),
		ns,
	)
	perform = ns["_performRecallConfirmationAction"]
	clicks = []

	class _Self:
		def _refreshRecallConfirmationState(self):
			return {
				"targets": {
					"收回": {
						"element": object(),
						"rect": (100, 200, 300, 320),
						"clickPoint": (215, 284),
					},
				},
				"isModernDialog": True,
				"hwnd": 456,
				"dialogRect": (50, 50, 350, 350),
			}

		def _invokeElement(self, *args, **kwargs):
			raise AssertionError("OCR click point should win before UIA even for modern dialog")

		def _clickAtPosition(self, *args, **kwargs):
			clicks.append((args, kwargs))

	assert perform(_Self(), "收回") is True
	assert clicks == [((215, 284), {"hwnd": 456})]


def test_perform_recall_confirmation_action_prefers_compact_modern_fallback_before_uia_when_ocr_missing():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_performRecallConfirmationAction"
	)
	compact_helper = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_isCompactModernRecallDialog"
	)
	fallback_calls = []
	ns = {
		"log": _Log(),
		"_getRecallDialogFallbackClickPoint": lambda action, rect, **kwargs: (
			fallback_calls.append((action, rect, kwargs)) or (222, 333)
		),
		"_": lambda text: text,
	}
	exec(
		compile(
			ast.Module(body=[compact_helper, method], type_ignores=[]),
			str(module_path),
			"exec",
		),
		ns,
	)
	perform = ns["_performRecallConfirmationAction"]
	clicks = []

	class _Self:
		def _refreshRecallConfirmationState(self):
			return {
				"targets": {
					"收回": {
						"element": object(),
						"rect": (100, 200, 300, 320),
						"clickPoint": None,
					},
				},
				"actionLabels": ["收回", "取消"],
				"isModernDialog": True,
				"hwnd": 789,
				"dialogRect": (50, 50, 350, 350),
			}

		def _invokeElement(self, *args, **kwargs):
			raise AssertionError("compact modern recall should fallback-click before invoking UIA")

		def _clickAtPosition(self, *args, **kwargs):
			clicks.append((args, kwargs))

	assert perform(_Self(), "收回") is True
	assert fallback_calls == [
		("收回", (50, 50, 350, 350), {"isModernDialog": True, "availableActions": ["收回", "取消"]}),
	]
	assert clicks == [((222, 333), {"hwnd": 789})]


def test_handle_message_context_menu_action_starts_photo_consent_watch_for_convert_to_text():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_handleMessageContextMenuAction"
	)
	ns = {}
	exec(
		compile(
			ast.Module(body=[method], type_ignores=[]),
			str(module_path),
			"exec",
		),
		ns,
	)
	handle = ns["_handleMessageContextMenuAction"]
	calls = []

	class _Self:
		def _watchForRecallConfirmationDialog(self):
			calls.append("recall")

		def _watchForPhotoTextConsentDialog(self):
			calls.append("photo")

	handle(_Self(), "轉為文字")
	assert calls == ["photo"]
	calls.clear()

	handle(_Self(), "收回")
	assert calls == ["recall"]


def test_perform_photo_text_consent_action_prefers_ocr_click_point():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_performPhotoTextConsentAction"
	)
	ns = {
		"log": _Log(),
		"_getPhotoTextConsentDialogFallbackClickPoint": lambda *args, **kwargs: None,
		"_": lambda text: text,
	}
	exec(
		compile(
			ast.Module(body=[method], type_ignores=[]),
			str(module_path),
			"exec",
		),
		ns,
	)
	perform = ns["_performPhotoTextConsentAction"]
	clicks = []

	class _Self:
		def _refreshPhotoTextConsentState(self):
			return {
				"targets": {
					"同意": {
						"element": object(),
						"rect": (100, 200, 300, 320),
						"clickPoint": (184, 266),
					},
				},
				"hwnd": 321,
				"dialogRect": (50, 50, 350, 350),
			}

		def _invokeElement(self, *args, **kwargs):
			raise AssertionError("photo consent should click the OCR hit before invoking UIA")

		def _clickAtPosition(self, *args, **kwargs):
			clicks.append((args, kwargs))

	assert perform(_Self(), "同意") is True
	assert clicks == [((184, 266), {"hwnd": 321})]


def test_activate_message_context_menu_supports_keyboard_fallback_hooks():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_activateMessageContextMenu"
	)

	arg_names = [arg.arg for arg in method.args.args]
	assert arg_names[:5] == ["self", "retriesLeft", "onAction", "onFailure", "shouldAbort"]

	message_context_menu_calls = [
		node
		for node in ast.walk(method)
		if isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "MessageContextMenu"
	]
	assert message_context_menu_calls, "expected MessageContextMenu construction"
	assert any(
		any(keyword.arg == "onAction" for keyword in call.keywords) for call in message_context_menu_calls
	)
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_popupLooksLikeMessageContextMenu"
		for node in ast.walk(method)
	), "keyboard popup activation should validate popup OCR before opening the virtual window"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_popupLooksLikeMessageContextMenu"
		and len(node.args) == 3
		for node in ast.walk(method)
	), "keyboard popup activation should pass app module, window handle, and popup rect into OCR validation"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_sendGestureWithAddonSuppressed"
		and node.args
		and isinstance(node.args[0], ast.Constant)
		and node.args[0].value == "escape"
		for node in ast.walk(method)
	), "keyboard popup activation should dismiss invalid popups before failing"


def test_message_context_menu_script_tries_keyboard_before_mouse_probes():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	method = [
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "script_messageContextMenu"
	][-1]

	keyboard_fallback = next(
		node
		for node in method.body
		if isinstance(node, ast.FunctionDef) and node.name == "_tryKeyboardFallback"
	)
	start_mouse_probes = next(
		node
		for node in method.body
		if isinstance(node, ast.FunctionDef) and node.name == "_startMouseContextMenuProbes"
	)

	assert any(
		isinstance(node, ast.Expr)
		and isinstance(node.value, ast.Call)
		and isinstance(node.value.func, ast.Name)
		and node.value.func.id == "_tryKeyboardFallback"
		for node in method.body
	), "message context menu should try the native applications key before mouse probes"

	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_sendGestureWithAddonSuppressed"
		and node.args
		and isinstance(node.args[0], ast.Constant)
		and node.args[0].value == "applications"
		for node in ast.walk(keyboard_fallback)
	), "keyboard-first path should send the applications key directly to LINE"

	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Attribute)
		and isinstance(node.func.value, ast.Name)
		and node.func.value.id == "self"
		and node.func.attr == "_activateMessageContextMenu"
		for node in ast.walk(keyboard_fallback)
	), "keyboard-first path should activate popup detection after sending the gesture"

	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_startMouseContextMenuProbes"
		for node in ast.walk(keyboard_fallback)
	), "keyboard failure should hand off to mouse probe fallback"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_attemptAtOffset"
		for node in ast.walk(start_mouse_probes)
	), "mouse probe fallback should use the multi-coordinate probe path"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_popupLooksLikeMessageContextMenu"
		and len(node.args) == 3
		for node in ast.walk(method)
	), "mouse probe path should pass popup validation the app module, window handle, and popup rect"


def test_copy_read_and_context_menu_actions_use_popup_label_click_point_resolution():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	copy_read = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_copyAndReadMessage"
	)
	context_menu_action = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_contextMenuAction"
	)

	for function_node in (copy_read, context_menu_action):
		assert any(
			isinstance(node, ast.Call)
			and isinstance(node.func, ast.Name)
			and node.func.id == "_resolvePopupMenuLabelClickPoint"
			for node in ast.walk(function_node)
		), "popup OCR flows should resolve labels through the shared click-point helper"
		assert any(
			isinstance(node, ast.Call)
			and isinstance(node.func, ast.Name)
			and node.func.id == "_collectPopupMenuRowRects"
			for node in ast.walk(function_node)
		), "popup OCR flows should collect row rects before clicking popup menu items"


def test_copy_read_uses_spoken_date_separator_text_helper():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	copy_read = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_copyAndReadMessage"
	)

	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_getSpokenLineDateSeparatorText"
		for node in ast.walk(copy_read)
	), "copy-read should normalize the spoken text for detected date separator OCR"


def test_popup_validation_and_native_gesture_helpers_exist_for_message_context_menu():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)

	popup_helper = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_popupLooksLikeMessageContextMenu"
	)
	send_helper = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_sendGestureWithAddonSuppressed"
	)

	assert [arg.arg for arg in popup_helper.args.args] == ["appMod", "hwnd", "popupRect"]
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_extractMatchedMessageContextMenuLabels"
		for node in ast.walk(popup_helper)
	), "popup validation helper should reuse message-menu label matching"

	assert [arg.arg for arg in send_helper.args.args] == ["gestureName"]
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Attribute)
		and isinstance(node.func.value, ast.Name)
		and node.func.value.id == "KeyboardInputGesture"
		and node.func.attr == "fromName"
		for node in ast.walk(send_helper)
	), "native gesture helper should synthesize keys by name"


def test_detect_edit_field_label_message_hint_skips_notes_detection():
	class _Rect:
		left = 700
		top = 520
		bottom = 640

	class _Element:
		CurrentName = "輸入訊息"
		CurrentBoundingRectangle = _Rect()

		def GetCurrentPropertyValue(self, _prop):
			return ""

	class _EditElements:
		Length = 1

	class _Parent:
		def FindAll(self, _scope, _condition):
			return _EditElements()

	class _Walker:
		def GetParentElement(self, _element):
			return _Parent()

	class _Client:
		RawViewWalker = _Walker()

		def CreatePropertyCondition(self, *_args):
			return object()

	ns = _load_line_symbols(
		function_names={"_getEditPlaceholder", "_detectEditFieldLabel"},
		namespace={
			"_": lambda text: text,
			"UIAHandler": SimpleNamespace(TreeScope_Children=1),
			"ctypes": SimpleNamespace(
				windll=SimpleNamespace(
					user32=SimpleNamespace(
						GetForegroundWindow=lambda: 101,
						GetWindowRect=lambda _hwnd, rect: None,
					),
				),
				wintypes=SimpleNamespace(RECT=lambda: SimpleNamespace(left=0, top=0, right=1200, bottom=800)),
				byref=lambda value: value,
			),
			"log": _Log(),
		},
	)
	ns["_isNotesWindowContext"] = lambda *_args, **_kwargs: (_ for _ in ()).throw(
		AssertionError("notes detection should not run for message fast-path"),
	)

	handler = SimpleNamespace(clientObject=_Client())
	assert ns["_detectEditFieldLabel"](_Element(), handler) == "Message input"


def test_detect_edit_field_label_query_text_skips_notes_ocr():
	seen_allow_ocr = []

	class _Rect:
		left = 120
		top = 130
		bottom = 180

	class _Element:
		CurrentName = ""
		CurrentBoundingRectangle = _Rect()

		def GetCurrentPropertyValue(self, prop_id):
			if prop_id == 30045:
				return "黃淑"
			return ""

	class _EditElements:
		Length = 1

	class _Parent:
		def FindAll(self, _scope, _condition):
			return _EditElements()

	class _Walker:
		def GetParentElement(self, _element):
			return _Parent()

	class _Client:
		RawViewWalker = _Walker()

		def CreatePropertyCondition(self, *_args):
			return object()

	ns = _load_line_symbols(
		function_names={"_getEditPlaceholder", "_detectEditFieldLabel"},
		namespace={
			"_": lambda text: text,
			"UIAHandler": SimpleNamespace(TreeScope_Children=1),
			"ctypes": SimpleNamespace(
				windll=SimpleNamespace(
					user32=SimpleNamespace(
						GetForegroundWindow=lambda: 101,
						GetWindowRect=lambda _hwnd, rect: None,
					),
				),
				wintypes=SimpleNamespace(RECT=lambda: SimpleNamespace(left=0, top=0, right=1200, bottom=800)),
				byref=lambda value: value,
			),
			"log": _Log(),
		},
	)

	def _fake_notes_context(*_args, **kwargs):
		seen_allow_ocr.append(kwargs["allowOcr"])
		return False, "line"

	ns["_isNotesWindowContext"] = _fake_notes_context

	handler = SimpleNamespace(clientObject=_Client())
	assert ns["_detectEditFieldLabel"](_Element(), handler) == "Search chat rooms"
	assert seen_allow_ocr == [False]


def test_schedule_query_invalidates_active_copy_read():
	scheduled = []
	focus_calls = []

	ns = _load_line_symbols(
		assignment_names={"_copyReadRequestId", "_focusQueryRequestId", "_messageContextMenuRequestId"},
		function_names={
			"_invalidateActiveCopyRead",
			"_invalidateActiveMessageContextMenu",
			"_scheduleQueryAndSpeakUIAFocus",
		},
		namespace={
			"core": SimpleNamespace(callLater=lambda _delay, callback: scheduled.append(callback)),
		},
	)
	ns["_queryAndSpeakUIAFocus"] = lambda: focus_calls.append("focus")
	ns["_copyReadRequestId"] = 7
	ns["_messageContextMenuRequestId"] = 11
	ns["_focusQueryRequestId"] = 0

	ns["_scheduleQueryAndSpeakUIAFocus"](50)

	assert ns["_copyReadRequestId"] == 8
	assert ns["_messageContextMenuRequestId"] == 12
	assert ns["_focusQueryRequestId"] == 1
	assert len(scheduled) == 1

	scheduled[0]()
	assert focus_calls == ["focus"]


def test_copy_read_stale_request_restores_clipboard_without_dismissing_other_windows():
	copy_calls = []
	escape_calls = []
	scheduled = []
	fallback_calls = []

	class _Rect:
		left = 100
		top = 200
		right = 400
		bottom = 260

	class _Target:
		CurrentBoundingRectangle = _Rect()

		def GetRuntimeId(self):
			return (1, 2, 3)

	user32 = SimpleNamespace(
		GetForegroundWindow=lambda: 101,
		SetForegroundWindow=lambda _hwnd: None,
		SetCursorPos=lambda _x, _y: None,
		mouse_event=lambda *_args: None,
	)
	previousKeyboardHandler = sys.modules.get("keyboardHandler")
	sys.modules["keyboardHandler"] = SimpleNamespace(
		KeyboardInputGesture=SimpleNamespace(
			fromName=lambda name: SimpleNamespace(
				send=lambda: escape_calls.append(name),
			),
		),
	)

	try:
		ns = _load_line_symbols(
			assignment_names={"_copyReadRequestId", "_copyReadClipboardOwnerId"},
			function_names={"_buildMessageBubbleClickPositions", "_copyAndReadMessage"},
			namespace={
				"api": SimpleNamespace(
					getClipData=lambda: "orig",
					copyToClip=lambda value: copy_calls.append(value),
				),
				"core": SimpleNamespace(
					callLater=lambda _delay, callback: scheduled.append(callback),
				),
				"ctypes": SimpleNamespace(windll=SimpleNamespace(user32=user32)),
				"time": SimpleNamespace(sleep=lambda _seconds: None),
				"log": _Log(),
				"_getElementRuntimeId": lambda _element: (1, 2, 3),
				"_getFocusedElementRuntimeId": lambda: (1, 2, 3),
				"_messageProbePointHitsTargetElement": lambda *_args: True,
				"_ocrReadMessageFallback": lambda _element: fallback_calls.append("fallback"),
				"_restoreFocusToElement": lambda *_args: None,
				"_shouldDismissCopyReadMenu": lambda _hwnd: False,
				"UIAHandler": SimpleNamespace(handler=object()),
			},
		)

		ns["_copyAndReadMessage"](_Target())
		assert copy_calls == [""]
		assert len(scheduled) == 1

		ns["_copyReadRequestId"] += 1
		scheduled[0]()

		assert copy_calls == ["", "orig"]
		assert escape_calls == []
		assert fallback_calls == []
		assert ns["_copyReadClipboardOwnerId"] == 0
	finally:
		if previousKeyboardHandler is None:
			sys.modules.pop("keyboardHandler", None)
		else:
			sys.modules["keyboardHandler"] = previousKeyboardHandler


def test_copy_read_does_not_restore_focus_when_line_is_not_foreground():
	copy_calls = []
	focus_restores = []
	scheduled = []
	current_runtime_id = [(1, 2, 3)]

	class _Rect:
		left = 100
		top = 200
		right = 400
		bottom = 260

	class _Target:
		CurrentBoundingRectangle = _Rect()

		def GetRuntimeId(self):
			return (1, 2, 3)

	user32 = SimpleNamespace(
		GetForegroundWindow=lambda: 101,
		SetForegroundWindow=lambda _hwnd: None,
		SetCursorPos=lambda _x, _y: None,
		mouse_event=lambda *_args: None,
	)

	ns = _load_line_symbols(
		assignment_names={"_copyReadRequestId", "_copyReadClipboardOwnerId"},
		function_names={"_buildMessageBubbleClickPositions", "_copyAndReadMessage"},
		namespace={
			"api": SimpleNamespace(
				getClipData=lambda: "orig",
				copyToClip=lambda value: copy_calls.append(value),
			),
			"core": SimpleNamespace(
				callLater=lambda _delay, callback: scheduled.append(callback),
			),
			"ctypes": SimpleNamespace(windll=SimpleNamespace(user32=user32)),
			"time": SimpleNamespace(sleep=lambda _seconds: None),
			"log": _Log(),
			"_getElementRuntimeId": lambda _element: (1, 2, 3),
			"_getFocusedElementRuntimeId": lambda: current_runtime_id[0],
			"_messageProbePointHitsTargetElement": lambda *_args: True,
			"_ocrReadMessageFallback": lambda _element: None,
			"_restoreFocusToElement": lambda *_args: focus_restores.append("restore"),
			"_shouldDismissCopyReadMenu": lambda _hwnd: False,
			"UIAHandler": SimpleNamespace(handler=object()),
		},
	)

	ns["_copyAndReadMessage"](_Target())
	assert len(scheduled) == 1

	current_runtime_id[0] = (9, 9, 9)
	scheduled[0]()

	assert copy_calls == ["", "orig"]
	assert focus_restores == []


def test_copy_read_skips_probe_when_hit_test_finds_different_element():
	copy_calls = []
	cursor_moves = []
	mouse_events = []
	scheduled = []

	class _Rect:
		left = 100
		top = 200
		right = 400
		bottom = 260

	class _Target:
		CurrentBoundingRectangle = _Rect()

		def GetRuntimeId(self):
			return (1, 2, 3)

	def _get_window_rect(_hwnd, rect):
		rect.left = 0
		rect.top = 0
		rect.right = 1200
		rect.bottom = 900
		return True

	user32 = SimpleNamespace(
		GetForegroundWindow=lambda: 101,
		GetWindowRect=_get_window_rect,
		SetForegroundWindow=lambda _hwnd: None,
		SetCursorPos=lambda x, y: cursor_moves.append((x, y)),
		mouse_event=lambda *args: mouse_events.append(args),
	)

	ns = _load_line_symbols(
		assignment_names={"_copyReadRequestId", "_copyReadClipboardOwnerId"},
		function_names={"_buildMessageBubbleClickPositions", "_copyAndReadMessage"},
		namespace={
			"api": SimpleNamespace(
				getClipData=lambda: "orig",
				copyToClip=lambda value: copy_calls.append(value),
			),
			"core": SimpleNamespace(
				callLater=lambda _delay, callback: scheduled.append(callback),
			),
			"ctypes": SimpleNamespace(
				windll=SimpleNamespace(user32=user32),
				byref=lambda value: value,
			),
			"time": SimpleNamespace(sleep=lambda _seconds: None),
			"log": _Log(),
			"_getElementRuntimeId": lambda _element: (1, 2, 3),
			"_getFocusedElementRuntimeId": lambda: (1, 2, 3),
			"_messageProbePointHitsTargetElement": lambda *_args: False,
			"_ocrReadMessageFallback": lambda _element: None,
			"_restoreFocusToElement": lambda *_args: None,
			"_shouldDismissCopyReadMenu": lambda _hwnd: False,
			"UIAHandler": SimpleNamespace(handler=object()),
		},
	)

	ns["_copyAndReadMessage"](_Target())

	assert copy_calls == [""]
	assert cursor_moves == []
	assert mouse_events == []
	assert len(scheduled) == 1


def test_copy_read_tries_keyboard_menu_before_mouse_probes():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	copy_read = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_copyAndReadMessage"
	)
	keyboard_fallback = next(
		node
		for node in copy_read.body
		if isinstance(node, ast.FunctionDef) and node.name == "_tryKeyboardCopyFallback"
	)
	start_mouse_probes = next(
		node
		for node in copy_read.body
		if isinstance(node, ast.FunctionDef) and node.name == "_startMouseCopyProbes"
	)
	attempt_copy = next(
		node
		for node in copy_read.body
		if isinstance(node, ast.FunctionDef) and node.name == "_attemptCopyAtOffset"
	)

	assert any(
		isinstance(node, ast.Expr)
		and isinstance(node.value, ast.Call)
		and isinstance(node.value.func, ast.Name)
		and node.value.func.id == "_tryKeyboardCopyFallback"
		for node in copy_read.body
	), "copy-read should try the native applications key before mouse probes"

	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_sendGestureWithAddonSuppressed"
		and node.args
		and isinstance(node.args[0], ast.Constant)
		and node.args[0].value == "applications"
		for node in ast.walk(keyboard_fallback)
	), "copy-read keyboard-first path should send the native applications key"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_findCopyMenuItem"
		for node in ast.walk(keyboard_fallback)
	), "copy-read keyboard-first path should reuse copy menu detection"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_attemptCopyAtOffset"
		for node in ast.walk(start_mouse_probes)
	), "copy-read keyboard failure should hand off to mouse probe fallback"
	assert any(
		isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_tryKeyboardCopyFallback"
		for node in ast.walk(attempt_copy)
	), "copy-read should fall back to OCR after mouse probes exhaust the already-tried keyboard path"


def test_message_focus_skips_off_window_elements_before_copy_read():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	query_focus = next(
		node
		for node in module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_queryAndSpeakUIAFocus"
	)

	visible_checks = [
		node
		for node in ast.walk(query_focus)
		if isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_isElementVisibleInForegroundWindow"
	]
	copy_read_calls = [
		node
		for node in ast.walk(query_focus)
		if isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "_copyAndReadMessage"
	]

	assert visible_checks, "message focus should guard against virtualized off-window UIA rects"
	assert copy_read_calls, "message focus should still use copy-read for visible message items"
