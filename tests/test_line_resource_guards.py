from __future__ import annotations

import ast
import importlib.util
import re
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


def _load_voice_call_confirmation_symbols():
	return _load_line_symbols(
		assignment_names={
			"_CJK_CHAR",
			"_CJK_SPACE_RE",
			"_VOICE_CALL_CONFIRMATION_START_LABELS",
			"_VOICE_CALL_CONFIRMATION_JOIN_LABELS",
			"_GROUP_VOICE_CALL_CONFIRMATION_HINTS",
		},
		function_names={
			"_removeCJKSpaces",
			"_normalizeVoiceCallConfirmationLine",
			"_matchVoiceCallConfirmationActionLabel",
			"_extractVoiceCallConfirmationActionLabels",
			"_isGroupVoiceCallConfirmationText",
			"_getVoiceCallConfirmationState",
			"_rectsIntersect",
			"_extractVoiceCallConfirmationActionClickPoints",
			"_selectVoiceCallConfirmationActionTarget",
		},
		namespace={"re": re},
	)


def test_voice_call_confirmation_state_keeps_personal_start_layout():
	ns = _load_voice_call_confirmation_symbols()

	assert ns["_getVoiceCallConfirmationState"](
		"確定要與小明進行語音通話？\n開始\n取消",
	) == {
		"action": "start",
		"isGroup": False,
	}


def test_voice_call_confirmation_state_detects_group_start_layout():
	ns = _load_voice_call_confirmation_symbols()

	assert ns["_getVoiceCallConfirmationState"](
		"確定要開始群組語音通話？\n開始\n取消",
	) == {
		"action": "start",
		"isGroup": True,
	}


def test_voice_call_confirmation_state_detects_group_join_layout():
	ns = _load_voice_call_confirmation_symbols()

	assert ns["_getVoiceCallConfirmationState"](
		"群 組 語 音 通 話\n2 人 已 加 入\n加 入\n取消",
	) == {
		"action": "join",
		"isGroup": True,
	}


def test_extract_voice_call_confirmation_action_click_points_uses_join_label_center():
	ns = _load_voice_call_confirmation_symbols()

	points = ns["_extractVoiceCallConfirmationActionClickPoints"](
		[
			{"text": "2 人已加入", "rect": (430, 340, 540, 370)},
			{"text": "加入", "rect": (462, 442, 510, 472)},
			{"text": "取消", "rect": (548, 442, 596, 472)},
		],
		(350, 250, 650, 520),
	)

	assert points == {
		"join": {
			"clickPoint": (486, 457),
			"rect": (462, 442, 510, 472),
		},
	}


def test_voice_call_confirmation_join_prompt_prefers_start_button_target_when_both_are_visible():
	ns = _load_voice_call_confirmation_symbols()
	targets = {
		"start": {
			"clickPoint": (470, 352),
			"rect": (420, 330, 520, 374),
		},
		"join": {
			"clickPoint": (758, 514),
			"rect": (732, 500, 784, 528),
		},
	}

	assert ns["_selectVoiceCallConfirmationActionTarget"](targets, "join") == targets["start"]


def test_window_client_screen_rect_converts_client_area_to_screen_coordinates():
	class _Rect:
		left = 0
		top = 0
		right = 1280
		bottom = 720

	class _Point:
		def __init__(self):
			self.x = 0
			self.y = 0

	def _get_client_rect(_hwnd, rect):
		rect.left = 0
		rect.top = 0
		rect.right = 1280
		rect.bottom = 720
		return 1

	def _client_to_screen(_hwnd, point):
		point.x += 320
		point.y += 180
		return 1

	user32 = SimpleNamespace(
		GetClientRect=_get_client_rect,
		ClientToScreen=_client_to_screen,
	)
	ctypes_module = SimpleNamespace(
		windll=SimpleNamespace(user32=user32),
		wintypes=SimpleNamespace(RECT=_Rect, POINT=_Point),
		byref=lambda value: value,
	)
	ns = _load_line_symbols(function_names={"_getWindowClientScreenRect"})

	assert ns["_getWindowClientScreenRect"](101, ctypesModule=ctypes_module) == (320, 180, 1600, 900)


def test_window_screen_rect_prefers_dwm_extended_frame_bounds():
	class _Rect:
		def __init__(self):
			self.left = 0
			self.top = 0
			self.right = 0
			self.bottom = 0

	def _get_window_rect(_hwnd, rect):
		rect.left = 10
		rect.top = 20
		rect.right = 1300
		rect.bottom = 760
		return 1

	def _dwm_get_window_attribute(_hwnd, _attr, rect, _size):
		rect.left = 18
		rect.top = 28
		rect.right = 1292
		rect.bottom = 752
		return 0

	ctypes_module = SimpleNamespace(
		windll=SimpleNamespace(
			user32=SimpleNamespace(GetWindowRect=_get_window_rect),
			dwmapi=SimpleNamespace(DwmGetWindowAttribute=_dwm_get_window_attribute),
		),
		wintypes=SimpleNamespace(RECT=_Rect),
		byref=lambda value: value,
		sizeof=lambda _value: 16,
	)
	ns = _load_line_symbols(function_names={"_rectTupleFromWinRect", "_getWindowScreenRect"})

	assert ns["_getWindowScreenRect"](101, ctypesModule=ctypes_module) == (18, 28, 1292, 752)


def test_chat_header_icon_point_scales_from_client_rect_for_common_dpi_values():
	ns = _load_line_symbols(
		assignment_names={
			"_CHAT_HEADER_ICON_Y_BASE",
			"_CHAT_HEADER_ICON_SPACING_BASE",
			"_CHAT_HEADER_FIRST_ICON_OFFSET_BASE",
		},
		function_names={"_scaleLineUiPixels", "_getChatHeaderIconPointFromRect"},
	)
	rect = (100, 50, 2020, 1130)

	assert ns["_getChatHeaderIconPointFromRect"](rect, 1.0, iconIndex=0) == (2005, 105)
	assert ns["_getChatHeaderIconPointFromRect"](rect, 1.5, iconIndex=0) == (1998, 132)
	assert ns["_getChatHeaderIconPointFromRect"](rect, 2.0, iconIndex=0) == (1990, 160)
	assert ns["_getChatHeaderIconPointFromRect"](rect, 1.5, iconIndex=2) == (1918, 132)


def test_chat_header_icon_point_prefers_client_rect_over_outer_frame():
	ns = _load_line_symbols(
		assignment_names={
			"_CHAT_HEADER_ICON_Y_BASE",
			"_CHAT_HEADER_ICON_SPACING_BASE",
			"_CHAT_HEADER_FIRST_ICON_OFFSET_BASE",
		},
		function_names={
			"_scaleLineUiPixels",
			"_getChatHeaderIconPointFromRect",
			"_getChatHeaderIconPointInfo",
			"_getChatHeaderIconPoint",
		},
		namespace={"log": _Log()},
	)
	ns["_getDpiScale"] = lambda _hwnd: 1.5
	ns["_getWindowClientScreenRect"] = lambda _hwnd: (12, 34, 1912, 1074)
	ns["_getWindowScreenRect"] = lambda _hwnd: (0, 0, 1920, 1080)

	assert ns["_getChatHeaderIconPointInfo"](101, iconIndex=0) == ((1890, 116), (12, 34, 1912, 1074))
	assert ns["_getChatHeaderIconPoint"](101, iconIndex=0) == (1890, 116)


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


def _load_message_input_classifier(*, name, edit_count=1, rect=(700, 520, 1100, 640)):
	"""Load _isMessageInputElement against a fake 1200x800 LINE window."""
	left, top, right, bottom = rect

	class _Element:
		CurrentName = name
		CurrentBoundingRectangle = SimpleNamespace(left=left, top=top, right=right, bottom=bottom)

		def GetCurrentPropertyValue(self, _prop):
			return ""

	class _Parent:
		def FindAll(self, _scope, _condition):
			return SimpleNamespace(Length=edit_count)

	class _Client:
		RawViewWalker = SimpleNamespace(GetParentElement=lambda _element: _Parent())

		def CreatePropertyCondition(self, *_args):
			return object()

	def _get_window_rect(_hwnd, rect):
		rect.left, rect.top, rect.right, rect.bottom = 0, 0, 1200, 800

	ns = _load_line_symbols(
		function_names={"_getEditPlaceholder", "_isMessageInputElement"},
		namespace={
			"UIAHandler": SimpleNamespace(TreeScope_Children=1),
			"ctypes": SimpleNamespace(
				windll=SimpleNamespace(
					user32=SimpleNamespace(
						GetForegroundWindow=lambda: 101,
						GetWindowRect=_get_window_rect,
					),
				),
				wintypes=SimpleNamespace(RECT=lambda: SimpleNamespace(left=0, top=0, right=0, bottom=0)),
				byref=lambda value: value,
			),
			"log": _Log(),
		},
	)
	handler = SimpleNamespace(clientObject=_Client())
	return lambda: ns["_isMessageInputElement"](_Element(), handler)


def test_message_input_classifier_trusts_native_line_names():
	"""LINE 26.6.0.4004+ names its edit fields, so the name decides first."""
	assert _load_message_input_classifier(name="輸入訊息")() is True
	assert _load_message_input_classifier(name="メッセージを入力", rect=(100, 60, 400, 90))() is True
	assert _load_message_input_classifier(name="搜尋聊天和訊息", rect=(100, 60, 400, 90))() is False


def test_message_input_classifier_lets_search_hint_win_over_message_hint():
	"""An English search box may be named "Search chats and messages"; Enter
	there must not play the send sound even though the name mentions messages
	and the field could sit low in a short window."""
	assert (
		_load_message_input_classifier(name="Search chats and messages", rect=(700, 600, 1100, 640))()
		is False
	)


def test_message_input_classifier_falls_back_to_position_and_skips_login_fields():
	assert _load_message_input_classifier(name="")() is True
	assert _load_message_input_classifier(name="", rect=(100, 60, 400, 90))() is False
	assert _load_message_input_classifier(name="輸入訊息", edit_count=2)() is False


def _app_module_bound_keys():
	"""Every key the LINE AppModule binds, via __gestures or @script(gesture=...)."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	module = ast.parse(module_path.read_text(encoding="utf-8"))
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	bound_keys = set()
	for node in app_module.body:
		if isinstance(node, ast.Assign) and any(
			isinstance(target, ast.Name) and target.id == "__gestures" for target in node.targets
		):
			bound_keys.update(key.value.lower() for key in node.value.keys)
		if isinstance(node, ast.FunctionDef):
			for decorator in node.decorator_list:
				if not isinstance(decorator, ast.Call):
					continue
				for keyword in decorator.keywords:
					if keyword.arg == "gesture":
						bound_keys.add(keyword.value.value.lower())
					elif keyword.arg == "gestures":
						bound_keys.update(element.value.lower() for element in keyword.value.elts)
	return bound_keys


def test_navigation_keys_are_left_to_native_uia_focus():
	"""LINE 26.6.0.4004 fires real UIA focus events, so NVDA announces Tab and
	arrow navigation itself. Re-binding those keys would bring back the old
	focus-polling path (double speech, synthetic Tabs in the chat list)."""
	bound_keys = _app_module_bound_keys()
	for key in ("kb:tab", "kb:shift+tab", "kb:uparrow", "kb:downarrow", "kb:leftarrow", "kb:rightarrow"):
		assert key not in bound_keys, f"{key} must stay with NVDA's native focus handling"


def test_native_menu_and_tab_keys_are_left_to_line():
	"""LINE 26.6.0.4004 opens its own readable context menu on Shift+F10 /
	Applications, and Control+1/2/3 move the focus to a tab that NVDA already
	announces. Intercepting them again would bring back the OCR menu window and
	speak a second, hard-coded tab name over NVDA's own announcement. The
	message shortcuts also send Shift+F10 themselves and rely on it reaching
	LINE."""
	bound_keys = _app_module_bound_keys()
	for key in ("kb:shift+f10", "kb:applications", "kb:control+1", "kb:control+2", "kb:control+3"):
		assert key not in bound_keys, f"{key} must reach LINE unchanged"


def _load_line_labels():
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "_lineLabels.py"
	spec = importlib.util.spec_from_file_location("line_labels_under_test", module_path)
	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module


def test_label_key_of_matches_whole_uia_names_only():
	labels = _load_line_labels()
	table = labels._MESSAGE_MENU_LABELS
	assert labels.labelKeyOf("回覆", table) == "回覆"
	assert labels.labelKeyOf(" copy ", table) == "複製"
	assert labels.labelKeyOf("送信取り消し", table) == "收回"
	assert labels.labelKeyOf("Copy link", table) is None
	assert labels.labelKeyOf("", table) is None
	assert labels.labelKeyOf(None, table) is None


class _FakeUIAElement:
	def __init__(self, name="", hwnd=0, parent=None, buttons=(), clickable=True):
		self.CurrentName = name
		self.CurrentNativeWindowHandle = hwnd
		self.parent = parent
		self.buttons = list(buttons)
		self.clickable = clickable
		self.searched = False

	def FindAll(self, scope, condition):
		self.searched = True
		buttons = self.buttons
		return SimpleNamespace(Length=len(buttons), GetElement=lambda index: buttons[index])


def _fake_menu_handler():
	walker = SimpleNamespace(GetParentElement=lambda element: element.parent)
	return SimpleNamespace(
		clientObject=SimpleNamespace(
			RawViewWalker=walker,
			CreatePropertyCondition=lambda *args: args,
		),
	)


def _load_native_menu_symbols(namespace=None):
	return _load_line_symbols(
		assignment_names={
			"_UIA_CONTROL_TYPE_PROPERTY_ID",
			"_UIA_BUTTON_CONTROL_TYPE_ID",
			"_MENU_ITEM_CLICKED",
			"_MENU_ITEM_NOT_FOUND",
			"_MENU_NOT_OPENED",
		},
		function_names={"_findMenuItem", "_chooseNativeMenuItem", "_waitForModifierRelease"},
		namespace={
			"_lineLabels": _load_line_labels(),
			"UIAHandler": SimpleNamespace(TreeScope_Descendants=4),
			"log": _Log(),
			**(namespace or {}),
		},
	)


def test_find_menu_item_searches_only_the_popup_window():
	"""LINE opens its menus in windows of their own; a same-named button in the
	main LINE window must never be picked when the menu lacks the item."""
	ns = _load_native_menu_symbols()
	labels = ns["_lineLabels"]
	mainWindow = _FakeUIAElement("LINE", hwnd=1, buttons=[_FakeUIAElement("收回")])
	reply = _FakeUIAElement("回覆")
	copy = _FakeUIAElement("複製")
	popup = _FakeUIAElement("LINE", hwnd=2, parent=mainWindow, buttons=[reply, copy])
	row = _FakeUIAElement(parent=popup)
	reply.parent = row
	handler = _fake_menu_handler()

	assert ns["_findMenuItem"](handler, reply, labels._MESSAGE_MENU_LABELS, "複製") == (copy, True)
	assert ns["_findMenuItem"](handler, reply, labels._MESSAGE_MENU_LABELS, "收回") == (None, True)
	assert popup.searched and not mainWindow.searched
	assert ns["_findMenuItem"](handler, reply, labels._MORE_OPTIONS_LABELS, "儲存聊天") == (None, False)


def test_choose_native_menu_item_reports_each_outcome():
	ns = _load_native_menu_symbols()
	labels = ns["_lineLabels"]
	save = _FakeUIAElement("Save chat")
	failing = _FakeUIAElement("儲存聊天", clickable=False)
	focused = _FakeUIAElement("關閉提醒")
	ns["_clickElementCenter"] = lambda element: element.clickable

	def choose(menu, key="儲存聊天"):
		ns["_waitForFocusedMenuItem"] = (
			lambda handler, previousRuntimeId: focused if menu is not None else None
		)
		ns["_findMenuItem"] = lambda handler, item, table, wanted: menu
		return ns["_chooseNativeMenuItem"](None, None, labels._MORE_OPTIONS_LABELS, key)

	assert choose(None) == ("noMenu", "")
	assert choose((None, True)) == ("notFound", "")
	assert choose((None, False)) == ("noMenu", "")
	assert choose((save, True)) == ("clicked", "Save chat")
	assert choose((failing, True)) == ("notFound", "")


def test_modifier_wait_includes_the_windows_key():
	"""NVDA+Windows+R etc. would otherwise send Windows+Shift+F10 to LINE."""
	held = {"polls": 0}

	def getAsyncKeyState(vk):
		if vk == 0x5B and held["polls"] < 3:
			held["polls"] += 1
			return 0x8000
		return 0

	clock = {"now": 0.0}

	def sleep(seconds):
		clock["now"] += seconds

	fakeTime = SimpleNamespace(monotonic=lambda: clock["now"], sleep=sleep)
	fakeCtypes = SimpleNamespace(
		windll=SimpleNamespace(user32=SimpleNamespace(GetAsyncKeyState=getAsyncKeyState)),
	)
	ns = _load_native_menu_symbols({"ctypes": fakeCtypes, "time": fakeTime})
	assert ns["_waitForModifierRelease"]() is True
	assert held["polls"] == 3

	fakeCtypes.windll.user32.GetAsyncKeyState = lambda vk: 0x8000 if vk == 0x11 else 0
	assert ns["_waitForModifierRelease"](timeoutSec=0.2) is False


def test_message_shortcuts_use_lines_own_menu():
	"""Reply, copy, save-as and recall pick the item from LINE's Shift+F10 menu
	by name; the right-click + OCR probing they used to do is gone."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	module = ast.parse(source)
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	methods = {node.name: node for node in app_module.body if isinstance(node, ast.FunctionDef)}
	expected = {
		"script_replyMessage": "回覆",
		"script_copyMessage": "複製",
		"script_saveAsMessage": "另存新檔",
		"script_recallMessage": "收回",
	}
	for script_name, label in expected.items():
		calls = [
			node
			for node in ast.walk(methods[script_name])
			if isinstance(node, ast.Call)
			and isinstance(node.func, ast.Attribute)
			and node.func.attr == "_nativeMessageMenuAction"
		]
		assert calls, f"{script_name} must use LINE's own message menu"
		assert calls[0].args[0].value == label
	menu_action = methods["_nativeMessageMenuAction"]
	called = {
		node.func.id
		for node in ast.walk(menu_action)
		if isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
	}
	assert "_waitForModifierRelease" in called, "wait for the shortcut's keys before sending Shift+F10"
	assert "_contextMenuAction" not in source
	assert "_virtualWindows.messageContextMenu" not in source
	assert "_virtualWindows.chatMoreOptions" not in source


def test_no_transient_letter_gesture_bindings():
	"""The add-on never binds bare letters at runtime. Such bindings swallowed
	those letters inside LINE's chat until a timeout released them, and mutating
	the private _gestureMap to undo them was worse still. LINE's own
	confirmation dialogs are readable, so there is nothing left to confirm with
	letter keys."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	source = module_path.read_text(encoding="utf-8")
	assert "._gestureMap[" not in source, "never mutate the private _gestureMap"
	assert "bindGesture" not in source, "never bind transient letter keys at runtime"


def test_no_deprecated_message_box():
	"""gui.messageBox is a deprecated alias as of NVDA 2025.1; every dialog the
	add-on shows goes through gui.message.MessageDialog instead."""
	addon_dir = Path(__file__).resolve().parents[1] / "addon"
	offenders = [
		str(path) for path in addon_dir.rglob("*.py") if "messageBox" in path.read_text(encoding="utf-8")
	]
	assert not offenders, f"gui.messageBox used in: {offenders}"


def test_no_deprecated_braille_text_region():
	"""braille.TextRegion is a deprecated moved alias as of NVDA 2026.3; the
	add-on shows braille via the public braille.handler.message instead."""
	addon_dir = Path(__file__).resolve().parents[1] / "addon"
	offenders = [
		str(path)
		for path in addon_dir.rglob("*.py")
		if "braille.TextRegion" in path.read_text(encoding="utf-8")
	]
	assert not offenders, f"braille.TextRegion used in: {offenders}"


def test_script_categories_use_shared_constant():
	"""Every @script category must reference SCRIPT_CATEGORY so all gestures
	stay grouped under one translated category in the Input Gestures dialog."""
	for relative in (
		("addon", "appModules", "line.py"),
		("addon", "globalPlugins", "lineDesktopHelper.py"),
	):
		module_path = Path(__file__).resolve().parents[1].joinpath(*relative)
		module = ast.parse(module_path.read_text(encoding="utf-8"))
		for node in ast.walk(module):
			if not isinstance(node, ast.Call):
				continue
			for keyword in node.keywords:
				if keyword.arg != "category":
					continue
				assert isinstance(keyword.value, ast.Name) and keyword.value.id == "SCRIPT_CATEGORY", (
					f"{module_path.name}:{keyword.value.lineno}: category= must use SCRIPT_CATEGORY"
				)


def test_terminate_schedules_pending_ocr_cleanup():
	"""AppModule.terminate must schedule the pending-OCR cleanup so buffers
	from recognitions whose callback never fires are released after LINE exits."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	module = ast.parse(module_path.read_text(encoding="utf-8"))
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	terminate = next(
		node for node in app_module.body if isinstance(node, ast.FunctionDef) and node.name == "terminate"
	)
	cleanup_calls = [
		node
		for node in ast.walk(terminate)
		if isinstance(node, ast.Call)
		and isinstance(node.func, ast.Attribute)
		and node.func.attr == "schedulePendingOcrCleanup"
	]
	assert cleanup_calls, "terminate must call _utils.schedulePendingOcrCleanup"


def test_describe_image_defers_consent_dialog_off_the_hook_stack():
	"""script_describeImage must not show its (potentially modal) consent
	dialog synchronously: scripts run on the keyboard hook's call stack, and
	every dialog-opening script in NVDA's own globalCommands.py defers via
	wx.CallAfter for exactly this reason. Calling _ensureImageDescriptionConsent
	directly from the script crashed on the first NVDA+Windows+I use."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	module = ast.parse(module_path.read_text(encoding="utf-8"))
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	script_describe_image = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "script_describeImage"
	)
	after_consent = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_describeImageAfterConsent"
	)

	assert not any(
		isinstance(call, ast.Call)
		and isinstance(call.func, ast.Name)
		and call.func.id == "_ensureImageDescriptionConsent"
		for call in ast.walk(script_describe_image)
	), "script_describeImage must not call the consent check synchronously"

	assert any(
		isinstance(call, ast.Call)
		and isinstance(call.func, ast.Attribute)
		and call.func.attr == "CallAfter"
		and call.args
		and isinstance(call.args[0], ast.Attribute)
		and call.args[0].attr == "_describeImageAfterConsent"
		for call in ast.walk(script_describe_image)
	), "script_describeImage must defer to _describeImageAfterConsent via wx.CallAfter"

	assert any(
		isinstance(call, ast.Call)
		and isinstance(call.func, ast.Name)
		and call.func.id == "_ensureImageDescriptionConsent"
		for call in ast.walk(after_consent)
	), "_describeImageAfterConsent must perform the consent check"


def test_file_dialog_wait_requires_seeing_dialog_before_declaring_it_gone():
	"""_suppressAddonForFileDialog's watcher must only resume the add-on once
	the dialog has actually been observed open and then closed. A prior
	two-phase version (wait up to 2s for it to appear, then wait for "not
	exists") would time out on a slow-to-appear dialog and immediately treat
	"never seen" as "gone", re-enabling navigation/copy interception while
	the real file dialog was still about to open."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	module = ast.parse(module_path.read_text(encoding="utf-8"))
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	suppress_method = next(
		node
		for node in app_module.body
		if isinstance(node, ast.FunctionDef) and node.name == "_suppressAddonForFileDialog"
	)

	block_calls = [
		node
		for node in ast.walk(suppress_method)
		if isinstance(node, ast.Call)
		and isinstance(node.func, ast.Name)
		and node.func.id == "blockUntilConditionMet"
	]
	assert len(block_calls) == 1, (
		"the wait must be a single poll, not a separate appear-then-disappear pair "
		"(the pair is what let a slow-to-appear dialog be missed entirely)"
	)

	wait_for_gone = next(
		node
		for node in ast.walk(suppress_method)
		if isinstance(node, ast.FunctionDef) and node.name == "_waitForDialogGone"
	)
	seen_flag_writes = [
		node
		for node in ast.walk(wait_for_gone)
		if isinstance(node, ast.Assign)
		and any(isinstance(target, ast.Name) and target.id == "seenDialog" for target in node.targets)
	]
	assert seen_flag_writes, "the poll must track whether the dialog was ever observed open"

	# ...but a watch whose token was invalidated (superseded, or LINE exited)
	# must stop at the next poll rather than run out the full give-up window,
	# which would keep the thread and the old AppModule alive into the next
	# LINE session, polling a process ID Windows may have recycled.
	evaluator = next(
		node
		for node in ast.walk(wait_for_gone)
		if isinstance(node, ast.FunctionDef) and node.name == "_dialogSeenThenClosed"
	)
	assert any(
		isinstance(node, ast.Compare)
		and any(isinstance(op, ast.NotEq) for op in node.ops)
		and any(isinstance(name, ast.Name) and name.id == "token" for name in ast.walk(node))
		and any(
			isinstance(attr, ast.Attribute) and attr.attr == "_fileDialogWaitToken" for attr in ast.walk(node)
		)
		for node in ast.walk(evaluator)
	), "the poll must bail out once its watch token has been invalidated"


def test_terminate_invalidates_file_dialog_watch_token():
	"""terminate() must bump the watch token so any in-flight file-dialog
	watcher stops instead of outliving this AppModule."""
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "line.py"
	module = ast.parse(module_path.read_text(encoding="utf-8"))
	app_module = next(
		node for node in module.body if isinstance(node, ast.ClassDef) and node.name == "AppModule"
	)
	terminate = next(
		node for node in app_module.body if isinstance(node, ast.FunctionDef) and node.name == "terminate"
	)
	token_bumps = [
		node
		for node in ast.walk(terminate)
		if isinstance(node, ast.AugAssign)
		and isinstance(node.target, ast.Attribute)
		and node.target.attr == "_fileDialogWaitToken"
	]
	assert token_bumps, "terminate must invalidate the file-dialog watch token"
