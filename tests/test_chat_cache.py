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


def test_reply_bubble_matches_actual_reply_not_quoted_preview():
	"""When a bubble shows a reply preview (sender + quoted message),
	OCR captures both names plus the quoted text. The lookup must match
	the actual reply content, not the longer quoted preview."""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "王昱涵",
				"content": "然後認領想要的工作(可揚你可以休息\n1. 科系有兩個表格爆掉了要改一下\n2. 格子要合併儲存格",
				"time": "22:20",
			},
			{"type": "message", "name": "陳禹安", "content": "那我用標題", "time": "10:26"},
			{"type": "message", "name": "莊忠諺", "content": "我修結論", "time": "10:27"},
		],
	)

	# 莊忠諺 replies to 王昱涵's earlier message. OCR captures both names
	# plus a chunk of the quoted preview, then the actual reply "我修結論".
	# Without the reply-sender filter, the long quoted overlap would win.
	formatted, idx = chat_cache.lookupMessage(
		"莊忠諺\n0 王昱涵\n然後認領想要的工作\n( 可揚你可以休息 \n我修結論",
	)
	assert idx == 2
	assert "我修結論" in formatted

	# Same shape for the 陳禹安 reply.
	formatted2, idx2 = chat_cache.lookupMessage(
		"陳禹安\n0 王昱涵\n然後認領想要的工作\n( 可揚你可以休息 \n那我用標題\n上午 10:26",
	)
	assert idx2 == 1
	assert "那我用標題" in formatted2


def test_reply_lookup_exposes_original_message_for_left_arrow():
	"""After matching a reply, getLastReplyInfo() returns the quoted
	original located upward in the cache, including its content for
	the left-arrow read-aloud handler."""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "王昱涵",
				"content": "然後認領想要的工作(可揚你可以休息\n1. 科系有兩個表格爆掉了",
				"time": "22:20",
			},
			{"type": "message", "name": "陳禹安", "content": "那我用標題", "time": "10:26"},
			{"type": "message", "name": "莊忠諺", "content": "我修結論", "time": "10:27"},
		],
	)

	formatted, idx = chat_cache.lookupMessage(
		"莊忠諺\n0 王昱涵\n然後認領想要的工作\n( 可揚你可以休息 \n我修結論",
	)
	assert idx == 2
	assert "我修結論" in formatted

	info = chat_cache.getLastReplyInfo()
	assert info is not None
	assert info["replySender"] == "莊忠諺"
	assert info["replyContent"] == "我修結論"
	assert info["originalName"] == "王昱涵"
	assert info["originalIdx"] == 0
	assert "然後認領想要的工作" in info["originalContent"]


def test_reply_lookup_clears_reply_info_for_non_reply_message():
	"""Non-reply OCR (single name) must clear stale reply info so the
	left-arrow handler doesn't read an old original."""
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "舊訊息", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "回覆內容", "time": "09:05"},
			{"type": "message", "name": "Bob", "content": "後續訊息", "time": "09:10"},
		],
	)

	# First lookup: a reply (2 names) — populates reply info.
	chat_cache.lookupMessage("Bob\n0 Alice\n舊訊息\n回覆內容\n上午 9 : 05")
	assert chat_cache.getLastReplyInfo() is not None

	# Next lookup: regular non-reply message — reply info must clear.
	chat_cache.lookupMessage("Bob 後續訊息 上午 9 : 10")
	assert chat_cache.getLastReplyInfo() is None


def test_reply_lookup_clears_reply_info_when_no_match():
	"""When the cache can't match anything, stale reply info must clear."""
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "舊訊息", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "回覆", "time": "09:05"},
		],
	)

	chat_cache.lookupMessage("Bob\n0 Alice\n舊訊息\n回覆\n上午 9 : 05")
	assert chat_cache.getLastReplyInfo() is not None

	# Unrelated OCR — no match, reply info must clear.
	formatted, idx = chat_cache.lookupMessage("完全不同的內容沒有時間")
	assert formatted is None
	assert chat_cache.getLastReplyInfo() is None


def test_find_quoted_original_only_searches_upward():
	"""The original message is always BEFORE the reply in chat order;
	never match a later message even if it has the same content."""
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "說了某句話", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "回應", "time": "09:05"},
			{"type": "message", "name": "Alice", "content": "說了某句話", "time": "10:00"},
		],
	)

	# Bob (idx=1) replies to Alice's earlier message (idx=0). Even though
	# the later Alice message (idx=2) has identical content, the search
	# upward must pick idx=0.
	chat_cache.lookupMessage("Bob\n0 Alice\n說了某句話\n回應\n上午 9 : 05")
	info = chat_cache.getLastReplyInfo()
	assert info is not None
	assert info["originalIdx"] == 0


def test_reply_filter_does_not_engage_when_only_one_name_in_ocr():
	"""Single-name OCR is not a reply preview — the filter must not engage,
	otherwise messages from anyone else become unmatchable."""
	_reset_cache(
		[
			{"type": "message", "name": "Alice", "content": "早安", "time": "09:00"},
			{"type": "message", "name": "Bob", "content": "回覆內容", "time": "09:05"},
		],
	)

	# Only Bob appears in OCR; reply filter must not exclude Bob's own
	# message. (If it did engage incorrectly using "Alice", Bob's message
	# would be filtered out and we'd return None.)
	formatted, idx = chat_cache.lookupMessage("Bob\n回覆內容\n上午 9 : 05")
	assert idx == 1
	assert "回覆內容" in formatted


def test_reply_filter_not_triggered_by_name_in_message_body():
	"""A name mentioned inside message text must not trigger reply detection.

	If Alice sends '感謝Bob你的幫助' and both Alice and Bob are in the cache,
	the old find() approach would detect two names and wrongly enter reply
	mode.  The line-anchored regex must not match 'Bob' mid-sentence.
	"""
	_reset_cache(
		[
			{"type": "message", "name": "Bob", "content": "沒問題", "time": "09:00"},
			{
				"type": "message",
				"name": "Alice",
				"content": "感謝Bob你的幫助",
				"time": "09:05",
			},
		],
	)

	# OCR for Alice's message — 'Bob' is inside the content line, not a
	# standalone line.  lookupMessage must match Alice's message normally
	# (idx=1) without entering reply mode.
	formatted, idx = chat_cache.lookupMessage("Alice\n感謝Bob你的幫助\n上午 9 : 05")
	assert idx == 1
	assert chat_cache.getLastReplyInfo() is None


def test_reply_detection_handles_quote_indicator_prefix():
	"""Real LINE OCR renders a quote-indicator glyph (often "0 ") before
	the quoted user name in reply bubbles, so the quoted name does NOT
	occupy a standalone line.  The regex must still detect it.

	Regression for an over-strict ``^name$`` regex that broke the
	original reply-bubble fix on actual LINE OCR.
	"""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "王昱涵",
				"content": "然後認領想要的工作(可揚你可以休息",
				"time": "22:20",
			},
			{"type": "message", "name": "莊忠諺", "content": "我修結論", "time": "10:23"},
		],
	)

	# Exact OCR shape from the LINE log: actual sender on its own line,
	# quoted user preceded by the "0 " quote-indicator glyph.
	formatted, idx = chat_cache.lookupMessage(
		"莊忠諺\n0 王昱涵\n然後認領想要的工作\n( 可揚你可以休息 \n我修結論",
	)
	assert idx == 1
	assert "我修結論" in formatted
	info = chat_cache.getLastReplyInfo()
	assert info is not None
	assert info["replySender"] == "莊忠諺"
	assert info["originalName"] == "王昱涵"


def test_reply_filter_not_triggered_by_substring_name():
	"""A short name that is a substring of a longer name must not match.

	If '王昱' and '王昱涵' are both in the cache and the OCR line is
	'王昱涵', the regex must not match '王昱' against that line.
	"""
	_reset_cache(
		[
			{
				"type": "message",
				"name": "王昱",
				"content": "你好",
				"time": "09:00",
			},
			{
				"type": "message",
				"name": "王昱涵",
				"content": "收到",
				"time": "09:05",
			},
		],
	)

	# OCR for 王昱涵's message — only '王昱涵' occupies a standalone line.
	# '王昱' must NOT match, so only one name is found → no reply filter.
	formatted, idx = chat_cache.lookupMessage("王昱涵\n收到\n上午 9 : 05")
	assert idx == 1
	assert chat_cache.getLastReplyInfo() is None


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
