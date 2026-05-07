"""Background chat-text cache used by NVDA+Windows+U.

Stores parsed messages from a saved chat export so the app module can read
each focused message bubble directly from text instead of running the
copy-read flow. OCR'd text from the focused bubble is matched against the
cache by content overlap, with time/name/date hints and chat-order proximity
as tie-breakers.

Counting rules (same as MessageReaderDialog):
  - Date separator entries do NOT increment the message index.
  - Every other entry (text, recalled "已收回訊息", sticker, etc.) counts as
    one message.

Dates also serve as positional anchors: when a date separator is matched,
the cursor is set to that exact position, so subsequent lookups are biased
toward messages in the same date group.
"""

import os
import re

from logHandler import log


_messages = []
_tempPath = None
_chatRoomName = None
_lastMatchedIdx = None
# Reply context for the most recent successful lookup, or None when the
# matched bubble was not a reply. Populated by lookupMessage(); consumed
# by the left-arrow handler so users can hear the quoted original.
_lastReplyInfo = None

# _messageIndexMap[i] = 1-based message number for position i (0 for date rows)
_messageIndexMap = []
# _messageDateGroups[i] = list index of the most recent date row before position i
# (-1 means no date has appeared yet)
_messageDateGroups = []


_DATE_FRAGMENT_RE = re.compile(r"\d{4}\.\d{2}\.\d{2}")
# Matches "HH:MM" or "HH : MM" (OCR sometimes inserts spaces around the colon)
_TIME_RE = re.compile(r"\b(\d{1,2})\s*[:：]\s*(\d{2})\b")
# Matches Chinese 12-hour notation like "下午 3:38" or "上午 10 : 05".
# Minutes are 1–2 digits because OCR sometimes truncates the trailing digit
# (e.g. captures "下午 3 : 1" when the real time is "15:10").
_AMPM_TIME_RE = re.compile(r"(上午|下午)\s*(\d{1,2})\s*[:：]\s*(\d{1,2})")
# Minimum contiguous-character overlap (after normalization) to consider a
# cached message a content match. Below this, we treat the overlap as noise.
_MIN_FUZZY_OVERLAP = 4


def setCache(messages, tempPath, chatRoomName):
	"""Activate the background chat cache.

	Clears any prior cache (and deletes the previous temp file) first so the
	cache always reflects a single chat room.  Builds _messageIndexMap and
	_messageDateGroups using the same counting rules as MessageReaderDialog.
	"""
	global _messages, _tempPath, _chatRoomName, _lastMatchedIdx, _lastReplyInfo
	global _messageIndexMap, _messageDateGroups
	clearCache()
	_messages = list(messages or [])
	_tempPath = tempPath
	_chatRoomName = chatRoomName
	_lastMatchedIdx = None
	_lastReplyInfo = None

	msgIdx = 0
	lastDateIdx = -1
	_messageIndexMap = []
	_messageDateGroups = []
	for i, msg in enumerate(_messages):
		if msg.get("type") == "date":
			lastDateIdx = i
		else:
			msgIdx += 1
		_messageIndexMap.append(msgIdx if msg.get("type") != "date" else 0)
		_messageDateGroups.append(lastDateIdx)

	log.info(
		f"LINE chat cache: stored {msgIdx} messages"
		f" ({len(_messages)} total entries) for room {chatRoomName!r}",
	)


def clearCache():
	"""Reset cache state and remove the temp export file if present."""
	global _messages, _tempPath, _chatRoomName, _lastMatchedIdx, _lastReplyInfo
	global _messageIndexMap, _messageDateGroups
	path = _tempPath
	_messages = []
	_tempPath = None
	_chatRoomName = None
	_lastMatchedIdx = None
	_lastReplyInfo = None
	_messageIndexMap = []
	_messageDateGroups = []
	if path:
		try:
			if os.path.isfile(path):
				os.remove(path)
				log.debug(f"LINE chat cache: removed temp file {path}")
		except Exception as e:
			log.warning(f"LINE chat cache: failed to remove temp file: {e}")


def isActive():
	return bool(_messages)


def getMessageCount():
	return len(_messages)


def getChatRoomName():
	return _chatRoomName


def getTempPath():
	return _tempPath


def onChatRoomChanged(newName):
	"""Drop the cache (and temp file) when the active chat room differs.

	Called from the chat-room-name detection paths in line.py so leaving the
	cached room invalidates the cache automatically. If the cache was set
	before any chat room name was recorded, adopt the first observed name
	instead of clearing on the first detection.
	"""
	global _chatRoomName
	if not isActive():
		return
	if not newName:
		return
	if _chatRoomName is None:
		_chatRoomName = newName
		log.debug(f"LINE chat cache: adopted chat room name {newName!r}")
		return
	if newName != _chatRoomName:
		log.info(
			f"LINE chat cache: chat room changed ({_chatRoomName!r} -> {newName!r}); clearing cache",
		)
		clearCache()


def _toHalfWidth(text):
	"""Convert full-width ASCII variants (U+FF01–U+FF5E) to half-width equivalents.

	OCR returns half-width punctuation even when LINE stores full-width.
	Normalizing both sides prevents false misses like '謝謝！' vs '謝謝!'.
	"""
	result = []
	for ch in text:
		cp = ord(ch)
		if 0xFF01 <= cp <= 0xFF5E:
			result.append(chr(cp - 0xFEE0))
		else:
			result.append(ch)
	return "".join(result)


def _normalize(text):
	if not text:
		return ""
	return re.sub(r"\s+", "", _toHalfWidth(text))


def _extractTimes(text):
	"""Extract HH:MM strings, handling Chinese AM/PM notation and OCR spacing.

	When OCR captures a single-digit minutes value (e.g. "下午 3 : 1" because the
	trailing digit fell outside the bubble), the result is the hour prefix only
	(e.g. "15:1") so the caller can do a startswith comparison against the
	cached HH:MM strings.
	"""
	times = []
	if not text:
		return times

	# Prefer the AM/PM form because LINE bubbles show it; this lets us recover
	# the 24-hour time that the chat export records.
	for m in _AMPM_TIME_RE.finditer(text):
		ampm = m.group(1)
		hh = int(m.group(2))
		mmStr = m.group(3)
		mm = int(mmStr)
		if not (0 <= hh <= 12):
			continue
		if ampm == "下午" and hh < 12:
			hh += 12
		elif ampm == "上午" and hh == 12:
			hh = 0
		if len(mmStr) == 2 and 0 <= mm <= 59:
			times.append(f"{hh:02d}:{mmStr}")
		elif len(mmStr) == 1:
			# OCR truncation — keep the hour + partial minutes prefix.
			times.append(f"{hh:02d}:{mmStr}")
	if times:
		return times

	# Fall back to bare HH:MM (24-hour from the export, or OCR without label)
	for m in _TIME_RE.finditer(text):
		hh = int(m.group(1))
		mm = int(m.group(2))
		if 0 <= hh <= 23 and 0 <= mm <= 59:
			times.append(f"{hh:02d}:{m.group(2)}")
	return times


def _longestCommonSubstring(a, b):
	"""Return the length of the longest common contiguous substring of a and b.

	OCR often drops single characters (e.g. "變化還很多" → "變化很多") or
	captures fragments from neighbouring bubbles. A contiguous-substring metric
	tolerates both: minor drops still leave a long block on each side, and
	multi-bubble OCR still contains an intact slice of the right message.
	"""
	if not a or not b:
		return 0
	if len(a) > len(b):
		a, b = b, a
	m, n = len(a), len(b)
	prev = [0] * (n + 1)
	longest = 0
	for i in range(1, m + 1):
		curr = [0] * (n + 1)
		ai = a[i - 1]
		for j in range(1, n + 1):
			if ai == b[j - 1]:
				curr[j] = prev[j - 1] + 1
				if curr[j] > longest:
					longest = curr[j]
		prev = curr
	return longest


def _timeMatches(msgTime, ocrTimes):
	"""True if msgTime equals any OCR time, allowing single-digit-minute prefix."""
	if not msgTime or not ocrTimes:
		return False
	for t in ocrTimes:
		if msgTime == t:
			return True
		# OCR may have truncated minutes (e.g. "15:1" matches "15:10"–"15:19")
		if len(t) < len(msgTime) and msgTime.startswith(t):
			return True
	return False


def _formatMessage(msg):
	if msg.get("type") == "date":
		return msg.get("content", "")
	name = msg.get("name", "")
	content = msg.get("content", "")
	timeStr = msg.get("time", "")
	return f"{name} {content} {timeStr}".strip()


def _detectReplyNames(ocrText):
	"""Identify both names when the OCR shows a reply preview.

	Reply bubbles show the actual sender's name first (above the bubble),
	then a quote block containing the quoted user's name and message
	preview, then the actual reply content. OCR captures both names, and
	the longer quoted preview otherwise wins on content overlap alone.

	Returns (replySender, quotedSender) when 2+ distinct cached sender
	names appear in the OCR text (reply pattern). Names are ordered by
	their position in the OCR text — the earliest is the actual sender,
	the next is the quoted user being replied to. Returns (None, None)
	when the OCR doesn't look like a reply.
	"""
	if not ocrText:
		return None, None
	occurrences = []
	seen = set()
	for msg in _messages:
		if msg.get("type") != "message":
			continue
		name = msg.get("name", "")
		if not name or name in seen:
			continue
		# Require the name at the start of a line, optionally preceded by
		# non-letter glyphs (e.g. the "0 " quote-indicator icon LINE
		# renders before quoted-user names in reply bubbles), and not
		# immediately followed by another CJK/Latin letter.  This:
		#   - lets quoted names like "0 王昱涵" still match (real LINE OCR);
		#   - blocks names mentioned mid-sentence (e.g. "感謝Bob你的幫助"
		#     — 感謝 is CJK so the non-letter prefix can't consume it);
		#   - blocks short names matching inside longer names (e.g.
		#     "王昱" must not match a line "王昱涵" — the lookahead fails
		#     because 涵 is a CJK letter).
		m = re.search(
			r"(?m)^[^一-鿿぀-ヿa-zA-Z]*" + re.escape(name) + r"(?![一-鿿぀-ヿa-zA-Z])",
			ocrText,
		)
		if m:
			occurrences.append((m.start(), name))
			seen.add(name)
	if len(occurrences) < 2:
		return None, None
	occurrences.sort(key=lambda t: t[0])
	return occurrences[0][1], occurrences[1][1]


def _findQuotedOriginal(replyIdx, quotedSender, ocrText):
	"""Locate the original message a reply is quoting.

	Replies always come AFTER the original in chat order, so we scan
	upward from ``replyIdx`` for messages by ``quotedSender`` and pick
	the one whose content has the largest overlap with the OCR text
	(which contains a fragment of the quoted preview). The OCR preview
	is shorter than the full message — substring containment in either
	direction or longest-common-substring (≥ ``_MIN_FUZZY_OVERLAP``)
	count as a match.

	Returns ``(msg, idx)`` on success, ``(None, None)`` otherwise.
	"""
	if not quotedSender or replyIdx <= 0:
		return None, None
	ocrNorm = _normalize(ocrText)
	if not ocrNorm:
		return None, None
	bestOverlap = 0
	bestIdx = -1
	for i in range(replyIdx - 1, -1, -1):
		msg = _messages[i]
		if msg.get("type") != "message":
			continue
		if msg.get("name") != quotedSender:
			continue
		contentNorm = _normalize(msg.get("content", ""))
		if not contentNorm:
			continue
		if contentNorm in ocrNorm:
			overlap = len(contentNorm)
		elif ocrNorm in contentNorm:
			overlap = len(ocrNorm)
		else:
			overlap = _longestCommonSubstring(contentNorm, ocrNorm)
		if overlap >= _MIN_FUZZY_OVERLAP and overlap > bestOverlap:
			bestOverlap = overlap
			bestIdx = i
	if bestIdx < 0:
		return None, None
	return _messages[bestIdx], bestIdx


def getLastReplyInfo():
	"""Return reply context for the most recent successful lookup.

	Returns a dict with ``replySender``, ``replyContent``, ``replyTime``,
	``originalName``, ``originalContent``, ``originalTime`` and
	``originalIdx`` when the matched bubble was a reply, otherwise
	``None``. ``originalContent``/``originalTime``/``originalIdx`` may
	be ``None`` when the original message couldn't be located in the
	cache (e.g. it was outside the exported window).
	"""
	return _lastReplyInfo


def clearLastReplyInfo():
	"""Clear cached reply context so the left-arrow handler fires only once."""
	global _lastReplyInfo
	_lastReplyInfo = None


def lookupMessage(ocrText):
	"""Return the cached message that best matches the OCR snippet.

	Scoring (higher = better match):
	  - Content substring overlap × 10
	  - Time match (including AM/PM conversion)  +50
	  - Name match                               +10
	  - Same date group as cursor                +20
	  - Chat-order proximity to cursor           −0.05 per position

	Date separators act as definitive anchors: matching one sets the cursor
	precisely, giving a strong date-group bias to the next lookups.

	Returns:
		(formattedText, index) on match, otherwise (None, None).
	"""
	global _lastMatchedIdx, _lastReplyInfo
	# Default to no reply context; populated below when the matched bubble
	# turns out to be a reply.
	_lastReplyInfo = None
	if not isActive() or not ocrText:
		return None, None

	# Date separator: unique enough to be decisive; anchor the cursor here.
	dateFragMatch = _DATE_FRAGMENT_RE.search(ocrText)
	if dateFragMatch:
		datePrefix = dateFragMatch.group(0)
		for i, msg in enumerate(_messages):
			if msg.get("type") != "date":
				continue
			if datePrefix in msg.get("content", ""):
				_lastMatchedIdx = i
				log.debug(
					f"LINE chat cache: date anchor matched at [{i}]: {msg['content']!r}",
				)
				return _formatMessage(msg), i

	ocrNorm = _normalize(ocrText)
	if not ocrNorm:
		return None, None

	ocrTimes = _extractTimes(ocrText)
	cursor = _lastMatchedIdx if _lastMatchedIdx is not None else 0

	# Date group of the cursor position (−1 if no date seen yet)
	cursorDateGroup = _messageDateGroups[cursor] if cursor < len(_messageDateGroups) else -1

	# Reply pattern: when the bubble shows a quote preview, OCR contains both
	# the actual sender's name and the quoted user's name. The quoted preview
	# usually has a longer content overlap than the actual reply, so without
	# this restriction the cache would return the quoted message instead of
	# the reply that the user just navigated to.
	replySender, quotedSender = _detectReplyNames(ocrText)

	bestScore = 0.0
	bestIdx = -1

	for i, msg in enumerate(_messages):
		if msg.get("type") != "message":
			continue
		msgContentNorm = _normalize(msg.get("content", ""))
		msgName = msg.get("name", "")
		msgTime = msg.get("time", "")
		if not msgContentNorm:
			continue
		if replySender and msgName != replySender:
			continue

		# Try exact substring containment first — that's the strongest signal.
		contentOverlap = 0
		if msgContentNorm in ocrNorm:
			contentOverlap = len(msgContentNorm)
		elif ocrNorm in msgContentNorm:
			contentOverlap = len(ocrNorm)
		else:
			# Fallback: longest contiguous substring. Tolerates OCR character
			# drops ("變化還很多" → "變化很多") and multi-bubble OCR captures.
			lcs = _longestCommonSubstring(msgContentNorm, ocrNorm)
			if lcs >= _MIN_FUZZY_OVERLAP:
				contentOverlap = lcs

		timeMatched = _timeMatches(msgTime, ocrTimes)
		nameMatched = bool(msgName and msgName in ocrText)

		# Short cached content (< 4 chars) is too ambiguous to use without a time
		# match — single characters like "有" appear in almost any OCR text.
		if not timeMatched and (contentOverlap == 0 or len(msgContentNorm) < 4):
			continue

		score = contentOverlap * 10
		if timeMatched:
			score += 50
		if nameMatched:
			score += 10
		# Same date group as the cursor → same day, strong signal.
		if cursorDateGroup >= 0 and _messageDateGroups[i] == cursorDateGroup:
			score += 20
		# Chat-order proximity to last confirmed position.
		score -= abs(i - cursor) * 0.05

		if score > bestScore:
			bestScore = score
			bestIdx = i

	if bestIdx < 0 or bestScore <= 0:
		return None, None

	_lastMatchedIdx = bestIdx
	log.debug(
		f"LINE chat cache: matched [{bestIdx}]"
		f" msgIdx={_messageIndexMap[bestIdx]}"
		f" dateGroup=[{_messageDateGroups[bestIdx]}]: {_formatMessage(_messages[bestIdx])!r}",
	)

	# Reply context: when the OCR pattern indicated a reply, locate the
	# original (quoted) message upward in the cache so the left-arrow
	# handler can read it on demand.
	if replySender and quotedSender:
		matched = _messages[bestIdx]
		originalMsg, originalIdx = _findQuotedOriginal(bestIdx, quotedSender, ocrText)
		_lastReplyInfo = {
			"replySender": replySender,
			"replyContent": matched.get("content", ""),
			"replyTime": matched.get("time", ""),
			"originalName": quotedSender,
			"originalContent": originalMsg.get("content", "") if originalMsg else None,
			"originalTime": originalMsg.get("time", "") if originalMsg else None,
			"originalIdx": originalIdx if originalMsg else None,
		}

	return _formatMessage(_messages[bestIdx]), bestIdx
