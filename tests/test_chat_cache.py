from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path


def _load_chat_cache():
	module_name = "addon.appModules._chatCache"
	module_path = Path(__file__).resolve().parents[1] / "addon" / "appModules" / "_chatCache.py"

	log_handler_mod = types.ModuleType("logHandler")

	class _Log:
		def debug(self, *args, **kwargs):
			pass

		def info(self, *args, **kwargs):
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


chat_cache = _load_chat_cache()


def _reset_cache(messages, room="Test Room"):
	"""Set internal state as if setCache() was called, rebuilding index maps."""
	chat_cache.setCache(messages, None, room)
	# setCache calls clearCache first which sets path; keep path None
	chat_cache._tempPath = None


def test_lookup_matches_message_by_content_and_time():
	_reset_cache(
		[
			{"type": "date", "content": "2026.04.09 星期四"},
			{"type": "message", "name": "Alice", "content": "午安", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "好的", "time": "09:01"},
		],
	)

	formatted, idx = chat_cache.lookupMessage("Alice 午安 09:00")
	assert idx == 1
	assert formatted == "Alice 午安 09:00"


def test_lookup_matches_fullwidth_content_via_halfwidth_ocr():
	# LINE stores full-width punctuation; OCR returns half-width
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "謝謝！", "time": "12:38"},
		],
	)

	formatted, idx = chat_cache.lookupMessage("謝謝 !\n, 下午 12 : 38")
	assert idx == 0
	assert "謝謝！" in formatted


def test_extract_times_handles_ampm_and_spaces():
	assert chat_cache._extractTimes("下午 12 : 38") == ["12:38"]
	assert chat_cache._extractTimes("下午 3:05") == ["15:05"]
	assert chat_cache._extractTimes("上午 12:00") == ["00:00"]
	assert chat_cache._extractTimes("上午 9:30") == ["09:30"]
	assert chat_cache._extractTimes("09:00") == ["09:00"]


def test_normalize_converts_fullwidth_to_halfwidth():
	assert chat_cache._normalize("謝謝！") == "謝謝!"
	assert chat_cache._normalize("，你好") == ",你好"
	assert chat_cache._normalize("  spaces  ") == "spaces"


def test_message_index_map_excludes_dates_counts_all_message_types():
	"""_messageIndexMap follows the same rules as MessageReaderDialog."""
	_reset_cache(
		[
			{"type": "date", "content": "2026.04.09 星期四"},  # idx 0 → msgIdx 0
			{"type": "message", "name": "A", "content": "早安", "time": "09:00"},  # idx 1 → msgIdx 1
			{"type": "message", "name": "B", "content": "已收回訊息", "time": "09:01"},  # idx 2 → msgIdx 2
			{"type": "date", "content": "2026.04.10 星期五"},  # idx 3 → msgIdx 0
			{"type": "message", "name": "A", "content": "晚安", "time": "21:00"},  # idx 4 → msgIdx 3
		],
	)
	assert chat_cache._messageIndexMap == [0, 1, 2, 0, 3]
	assert chat_cache._messageDateGroups == [0, 0, 0, 3, 3]


def test_date_group_bias_picks_correct_day_when_duplicate_content():
	"""When the same message text appears on two days, cursor date group wins."""
	_reset_cache(
		[
			{"type": "date", "content": "2026.04.09 星期四"},
			{"type": "message", "name": "Alice", "content": "好", "time": "09:00"},
			{"type": "date", "content": "2026.04.10 星期五"},
			{"type": "message", "name": "Alice", "content": "好", "time": "21:05"},
		],
	)

	# Anchor to the second date group first
	chat_cache.lookupMessage("2026.04.10 星期五")

	# Now '好 21:05' should win because cursor is in the 2026.04.10 date group.
	# "下午 9 : 05" → 9+12=21 → 21:05, which matches the cached time.
	formatted, idx = chat_cache.lookupMessage("Alice 好 下午 9 : 05")
	assert idx == 3


def test_recalled_message_counted_and_matchable():
	_reset_cache(
		[
			{"type": "date", "content": "2026.04.09 星期四"},
			{"type": "message", "name": "Alice", "content": "已收回訊息", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "早安", "time": "09:01"},
		],
	)
	# Recalled message should appear in the index and be matchable
	assert chat_cache._messageIndexMap[1] == 1  # counts as message 1
	formatted, idx = chat_cache.lookupMessage("Alice已收回訊息 09:00")
	assert idx == 1


def test_lookup_disambiguates_duplicates_by_chat_order():
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "好", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "好", "time": "09:05"},
			{"type": "message", "name": "Carol", "content": "好", "time": "09:10"},
		],
	)

	# Sweep top-down — each lookup should pick the next occurrence even though
	# content is identical, because the cursor advances after each match.
	first, idx1 = chat_cache.lookupMessage("Alice 好 09:00")
	assert idx1 == 0
	second, idx2 = chat_cache.lookupMessage("Bob 好 09:05")
	assert idx2 == 1
	third, idx3 = chat_cache.lookupMessage("Carol 好 09:10")
	assert idx3 == 2


def test_lookup_matches_date_separator_by_date_fragment():
	_reset_cache(
		[
			{"type": "date", "content": "2026.04.09 星期四"},
			{"type": "message", "name": "Alice", "content": "早安", "time": "09:00"},
		],
	)

	formatted, idx = chat_cache.lookupMessage("2026.04.09 星期四")
	assert idx == 0
	assert formatted == "2026.04.09 星期四"


def test_lookup_tolerates_single_char_drop_via_lcs():
	"""OCR sometimes drops a single character; a long contiguous substring
	on either side should still match."""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "陳圻囷",
				"content": "變化還很多，都不知道明年會不會打仗",
				"time": "15:25",
			},
		],
	)

	# OCR dropped the "還" character but the rest is intact
	formatted, idx = chat_cache.lookupMessage(
		"已讀\n變化很多 , 都不知道明年會不會打仗\n下午 3 : 25",
	)
	assert idx == 0
	assert "變化還很多" in formatted


def test_lookup_tolerates_truncated_time_minutes():
	"""When OCR cuts off the trailing minute digit ('3 : 1' for '15:10'), the
	match should still succeed using the hour:minute prefix."""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "陳圻囷",
				"content": "如果我還在餐廳可以有招待哈哈哈",
				"time": "15:10",
			},
		],
	)

	formatted, idx = chat_cache.lookupMessage(
		"如果我還在餐可以有招待哈哈哈\n下午 3 : 1",
	)
	assert idx == 0
	assert "餐廳" in formatted


def test_lookup_finds_match_in_multi_bubble_ocr():
	"""OCR sometimes captures fragments from neighbouring bubbles. A long
	contiguous slice of the right message should still win."""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "可揚",
				"content": "今年的目標是有機會去宜蘭玩，去你們的餐廳吃",
				"time": "22:25",
			},
		],
	)

	formatted, idx = chat_cache.lookupMessage(
		"今年的目\n2 月 1 6 日 ( - )\n已讀\n廳吃\n下午 10 : 25\n蘭玩 , 去你們的餐",
	)
	assert idx == 0
	assert "宜蘭玩" in formatted


def test_short_content_requires_time_match_to_avoid_false_positives():
	"""A cached message whose content is < 4 chars must NOT match unless time also fits."""
	_reset_cache(
		[
			{"type": "message", "name": "可揚", "content": "有", "time": "11:29"},
		],
	)

	# OCR contains "有" but time is completely different — should NOT match
	formatted, idx = chat_cache.lookupMessage("你有吃午餐 ?\n下午 11 : 38")
	assert formatted is None
	assert idx is None

	# OCR contains "有" AND matching time "11:29" — should match
	formatted2, idx2 = chat_cache.lookupMessage("有\n上午 11 : 29")
	assert idx2 == 0
	assert "有" in formatted2


def test_lookup_returns_none_when_ocr_text_unrelated():
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "午安", "time": "09:00"},
		],
	)

	formatted, idx = chat_cache.lookupMessage("完全不同的內容沒有時間")
	assert formatted is None
	assert idx is None


def test_lookup_returns_none_when_cache_inactive():
	chat_cache._messages = []
	chat_cache._tempPath = None
	chat_cache._chatRoomName = None
	chat_cache._lastMatchedIdx = None

	formatted, idx = chat_cache.lookupMessage("Alice 午安 09:00")
	assert formatted is None
	assert idx is None


def _make_temp_file():
	import os
	import tempfile

	fd, path = tempfile.mkstemp(prefix="lineDesktop_test_", suffix=".txt")
	os.close(fd)
	with open(path, "w", encoding="utf-8") as f:
		f.write("hi")
	return path


def test_on_chat_room_changed_clears_cache_when_room_differs():
	import os

	temp_file = _make_temp_file()
	try:
		chat_cache.setCache(
			[{"type": "message", "name": "Alice", "content": "x", "time": "09:00"}],
			temp_file,
			"Room A",
		)

		chat_cache.onChatRoomChanged("Room B")

		assert not chat_cache.isActive()
		assert not os.path.exists(temp_file)
	finally:
		if os.path.exists(temp_file):
			os.remove(temp_file)


def test_on_chat_room_changed_adopts_first_observed_name():
	import os

	temp_file = _make_temp_file()
	try:
		chat_cache.setCache(
			[{"type": "message", "name": "Alice", "content": "x", "time": "09:00"}],
			temp_file,
			None,
		)

		chat_cache.onChatRoomChanged("Room A")

		# Cache must remain active and the room name must now equal "Room A".
		assert chat_cache.isActive()
		assert chat_cache.getChatRoomName() == "Room A"

		chat_cache.onChatRoomChanged("Room B")
		assert not chat_cache.isActive()
	finally:
		if os.path.exists(temp_file):
			os.remove(temp_file)


def test_clear_cache_removes_temp_file():
	import os

	temp_file = _make_temp_file()
	try:
		chat_cache.setCache(
			[{"type": "message", "name": "Alice", "content": "x", "time": "09:00"}],
			temp_file,
			"Room A",
		)
		chat_cache.clearCache()

		assert not chat_cache.isActive()
		assert not os.path.exists(temp_file)
	finally:
		if os.path.exists(temp_file):
			os.remove(temp_file)
