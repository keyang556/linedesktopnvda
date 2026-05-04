from .._virtualWindow import VirtualWindow
from .._utils import ocrGetText, message
from logHandler import log

import difflib
import re
from typing import Any

_CJK_CHAR = (
	r"[\u2E80-\u9FFF\uF900-\uFAFF"
	r"\U00020000-\U0002A6DF\U0002A700-\U0002EBEF\U00030000-\U000323AF]"
)
_CJK_SPACE_RE = re.compile(
	r"(?<=" + _CJK_CHAR + r") (?=" + _CJK_CHAR + r")",
)


def _removeCJKSpaces(text):
	return _CJK_SPACE_RE.sub("", text)


_KNOWN_MENU_LABELS = (
	"開啟提醒",
	"關閉提醒",
	"邀請",
	"相簿",
	"照片・影片",
	"檔案",
	"連結",
	"投票",
	"儲存聊天",
	"背景設定",
	"檢舉",
	"封鎖",
)

_MENU_LABEL_ALIASES = {
	"開啟提醒": ("開啟提醒",),
	"關閉提醒": ("關閉提醒",),
	"邀請": ("邀請",),
	"相簿": ("相簿",),
	"照片・影片": (
		"照片影片",
		"照片影⽚",
		"照片影像",
		"照片•影片",
		"照片‧影片",
		"眧片影片",
		"照片 影片",
	),
	"檔案": ("檔案",),
	"連結": ("連結",),
	"投票": ("投票",),
	"儲存聊天": ("儲存聊天",),
	"背景設定": ("背景設定", "冃景言殳定", "背景设定"),
	"檢舉": ("檢舉",),
	"封鎖": ("封鎖",),
}

_NOISE_LINE_RE = re.compile(r"^[\W_]*[\d０-９]+[\W_]*$|^[A-Za-z]{4,}$")


def _normalizeLineText(text: str) -> str:
	text = _removeCJKSpaces((text or "").strip())
	text = text.replace("•", "・").replace("‧", "・").replace("·", "・")
	text = text.replace("・", "")
	text = text.replace(" ", "")
	return text


def _matchMenuLabel(text: str) -> str | None:
	normalized = _normalizeLineText(text)
	if not normalized:
		return None

	for canonical, aliases in _MENU_LABEL_ALIASES.items():
		for alias in aliases:
			if alias in normalized:
				return canonical

	bestLabel = None
	bestRatio = 0.0
	for canonical in _KNOWN_MENU_LABELS:
		ratio = difflib.SequenceMatcher(None, normalized, canonical).ratio()
		if ratio > bestRatio:
			bestRatio = ratio
			bestLabel = canonical

	if bestLabel and bestRatio >= 0.62:
		return bestLabel
	return None


def _getObjectValue(obj: Any, *names: str) -> Any:
	if obj is None:
		return None
	if isinstance(obj, dict):
		for name in names:
			value = obj.get(name)
			if value is not None:
				return value
	for name in names:
		value = getattr(obj, name, None)
		if value is not None:
			return value
	return None


def _coerceRectNumber(value: Any) -> float | None:
	if value is None:
		return None
	for attr in ("value", "Value"):
		nested = getattr(value, attr, None)
		if nested is not None and nested is not value:
			value = nested
			break
	try:
		return float(value)
	except Exception:
		return None


def _coerceRectTuple(*values: Any) -> tuple[int, ...] | None:
	coerced = []
	for value in values:
		number = _coerceRectNumber(value)
		if number is None:
			return None
		coerced.append(int(round(number)))
	return tuple(coerced)


def _rectFromSequence(source: Any, preferXYWH: bool = False) -> tuple[int, int, int, int] | None:
	if not isinstance(source, (list, tuple)) or len(source) < 4:
		return None
	flat = _coerceRectTuple(*source[:4])
	if not flat:
		return None
	left, top, third, fourth = flat
	edgeRect = (left, top, third, fourth) if third > left and fourth > top else None
	sizeRect = (left, top, left + third, top + fourth) if third > 0 and fourth > 0 else None
	for rect in (sizeRect, edgeRect) if preferXYWH else (edgeRect, sizeRect):
		if rect and rect[2] > rect[0] and rect[3] > rect[1]:
			return rect
	return None


def _rectFromObject(source: Any, preferXYWH: bool = False) -> tuple[int, int, int, int] | None:
	if source is None:
		return None
	rect = _coerceRectTuple(
		_getObjectValue(source, "left", "Left", "minX", "MinX", "x1", "X1"),
		_getObjectValue(source, "top", "Top", "minY", "MinY", "y1", "Y1"),
		_getObjectValue(source, "right", "Right", "maxX", "MaxX", "x2", "X2"),
		_getObjectValue(source, "bottom", "Bottom", "maxY", "MaxY", "y2", "Y2"),
	)
	if rect and rect[2] > rect[0] and rect[3] > rect[1]:
		return rect

	for leftName, topName in (("x", "y"), ("X", "Y"), ("left", "top"), ("Left", "Top")):
		rect = _coerceRectTuple(
			_getObjectValue(source, leftName),
			_getObjectValue(source, topName),
			_getObjectValue(source, "width", "Width", "w", "W"),
			_getObjectValue(source, "height", "Height", "h", "H"),
		)
		if rect and rect[2] > 0 and rect[3] > 0:
			left, top, width, height = rect
			return (left, top, left + width, top + height)

	return _rectFromSequence(source, preferXYWH=preferXYWH)


def _extractRectLike(obj: Any) -> tuple[int, int, int, int] | None:
	for attr in (
		"boundingRect",
		"boundingRectangle",
		"bounding_rect",
		"bounding_rectangle",
		"rect",
		"location",
		"bounds",
		"box",
	):
		rect = _rectFromObject(_getObjectValue(obj, attr), preferXYWH=True)
		if rect:
			return rect

	rect = _rectFromObject(obj)
	if rect:
		return rect

	words = _getObjectValue(obj, "words", "Words")
	try:
		wordIterator = iter(words or ())
	except Exception:
		wordIterator = ()
	wordRects = []
	for word in wordIterator:
		rect = _extractRectLike(word)
		if rect:
			wordRects.append(rect)
	if wordRects:
		return (
			min(rect[0] for rect in wordRects),
			min(rect[1] for rect in wordRects),
			max(rect[2] for rect in wordRects),
			max(rect[3] for rect in wordRects),
		)

	return None


def _extractOcrLines(result: Any) -> list[dict[str, Any]]:
	rawLines = getattr(result, "lines", None) or []
	extracted: list[dict[str, Any]] = []
	for rawLine in rawLines:
		text = getattr(rawLine, "text", "") or ""
		text = text.strip()
		if not text:
			continue
		extracted.append(
			{
				"text": text,
				"rect": _extractRectLike(rawLine),
			},
		)
	return extracted


def _normalizeMenuRowRects(
	rowRects: list[tuple[int, int, int, int]] | None,
	popupRect: tuple[int, int, int, int],
) -> list[tuple[int, int, int, int]]:
	if not rowRects:
		return []

	left, top, right, bottom = popupRect
	normalized: list[tuple[int, int, int, int]] = []
	seen = set()
	for rect in rowRects:
		if not rect or len(rect) != 4:
			continue
		rowLeft, rowTop, rowRight, rowBottom = [int(value) for value in rect]
		rowLeft = max(left, rowLeft)
		rowTop = max(top, rowTop)
		rowRight = min(right, rowRight)
		rowBottom = min(bottom, rowBottom)
		if rowRight <= rowLeft or rowBottom <= rowTop:
			continue
		key = (rowLeft, rowTop, rowRight, rowBottom)
		if key in seen:
			continue
		seen.add(key)
		normalized.append(key)

	normalized.sort(key=lambda rect: (((rect[1] + rect[3]) / 2), rect[0]))
	return normalized


_KNOWN_100_PERCENT_ROW_LAYOUT = (
	("關閉提醒", 0),
	("邀請", 1),
	("相簿", 2),
	("照片・影片", 3),
	("檔案", 4),
	("連結", 5),
	("投票", 6),
	("儲存聊天", 7),
	("背景設定", 8),
	("檢舉", 9),
	("封鎖", 10),
)
_KNOWN_100_PERCENT_ROW_INDEX_BY_LABEL = dict(_KNOWN_100_PERCENT_ROW_LAYOUT)
_REMINDER_TOGGLE_LABELS = {"開啟提醒", "關閉提醒"}
_KNOWN_100_PERCENT_ROW_INDEX_BY_LABEL.update((label, 0) for label in _REMINDER_TOGGLE_LABELS)
_KNOWN_100_PERCENT_ANCHOR_LABELS = {
	"投票",
	"儲存聊天",
	"背景設定",
	"檢舉",
	"封鎖",
}


def _inferKnownMenuRowIndex(
	label: str | None,
	rowRects: list[tuple[int, int, int, int]],
) -> int | None:
	# At 100%/125% LINE exposes 12 row rects; the final bottom info row is non-actionable.
	if len(rowRects) != 12:
		return None
	return _KNOWN_100_PERCENT_ROW_INDEX_BY_LABEL.get(label or "")


def _buildKnown100PercentLayoutElements(
	rowRects: list[tuple[int, int, int, int]],
	detectedElements: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
	if len(rowRects) != 12:
		return []

	detectedReminderLabel = next(
		(
			element.get("name")
			for element in detectedElements or []
			if element.get("name") in _REMINDER_TOGGLE_LABELS
		),
		None,
	)
	elements: list[dict[str, Any]] = []
	for label, rowIndex in _KNOWN_100_PERCENT_ROW_LAYOUT:
		if rowIndex >= len(rowRects):
			continue
		if rowIndex == 0 and detectedReminderLabel:
			label = detectedReminderLabel
		rowLeft, rowTop, rowRight, rowBottom = rowRects[rowIndex]
		elements.append(
			{
				"name": label,
				"role": None,
				"clickPoint": (
					int((rowLeft + rowRight) / 2),
					int((rowTop + rowBottom) / 2),
				),
			},
		)
	return elements


def _shouldUseKnown100PercentLayout(
	elements: list[dict[str, Any]],
	rowRects: list[tuple[int, int, int, int]],
) -> bool:
	if len(rowRects) != 12 or len(elements) >= 8:
		return False
	labels = {element.get("name") for element in elements}
	return len(labels & _KNOWN_100_PERCENT_ANCHOR_LABELS) >= 2


def _assignRowRectsToElements(
	elements: list[dict[str, Any]],
	rowRects: list[tuple[int, int, int, int]],
) -> None:
	if not elements or not rowRects:
		return

	currentRow = 0
	totalRows = len(rowRects)
	for elementIndex, element in enumerate(elements):
		remainingElements = len(elements) - elementIndex
		maxRowIndex = totalRows - remainingElements
		if maxRowIndex < currentRow:
			break

		targetY = element.get("_lineCenterY")
		chosenRowIndex = currentRow
		if targetY is not None:
			bestDistance = None
			for rowIndex in range(currentRow, maxRowIndex + 1):
				rowLeft, rowTop, rowRight, rowBottom = rowRects[rowIndex]
				rowCenterY = (rowTop + rowBottom) / 2
				distance = abs(rowCenterY - targetY)
				if bestDistance is None or distance < bestDistance:
					bestDistance = distance
					chosenRowIndex = rowIndex
		else:
			knownRowIndex = _inferKnownMenuRowIndex(element.get("name"), rowRects)
			if knownRowIndex is not None and currentRow <= knownRowIndex <= maxRowIndex:
				chosenRowIndex = knownRowIndex

		rowLeft, rowTop, rowRight, rowBottom = rowRects[chosenRowIndex]
		element["clickPoint"] = (
			int((rowLeft + rowRight) / 2),
			int((rowTop + rowBottom) / 2),
		)
		currentRow = chosenRowIndex + 1


def _buildMenuElements(
	lines: list[dict[str, Any]],
	popupRect: tuple[int, int, int, int],
	rowRects: list[tuple[int, int, int, int]] | None = None,
) -> list[dict[str, Any]]:
	left, top, right, bottom = popupRect
	centerX = (left + right) // 2
	rowRects = _normalizeMenuRowRects(rowRects, popupRect)
	elements: list[dict[str, Any]] = []

	for line in lines:
		rawText = line["text"]
		menuLabel = _matchMenuLabel(rawText)
		if not menuLabel:
			normalized = _normalizeLineText(rawText)
			if normalized and not _NOISE_LINE_RE.fullmatch(normalized):
				log.debug(
					f"LINE: ChatMoreOptions skipping non-menu OCR line: {rawText!r}",
				)
			continue

		rect = line.get("rect")
		lineCenterY = None
		if rect:
			lineLeft, lineTop, lineRight, lineBottom = rect
			if lineRight <= left or lineLeft >= right or lineBottom <= top or lineTop >= bottom:
				rect = None
			else:
				clickY = int((lineTop + lineBottom) / 2)
				clickX = int((lineLeft + lineRight) / 2)
				lineCenterY = clickY
		if not rect:
			clickY = None
			clickX = centerX

		elements.append(
			{
				"name": menuLabel,
				"role": None,
				"clickPoint": (clickX, clickY) if clickY is not None else None,
				"_lineCenterY": lineCenterY,
			},
		)

	if elements:
		knownLayoutElements = []
		if _shouldUseKnown100PercentLayout(elements, rowRects):
			knownLayoutElements = _buildKnown100PercentLayoutElements(rowRects, elements)
		if knownLayoutElements:
			log.debug("LINE: ChatMoreOptions using known 100% popup row layout")
			return knownLayoutElements

		_assignRowRectsToElements(elements, rowRects)
		itemHeight = (bottom - top) / len(elements)
		for index, element in enumerate(elements):
			if element["clickPoint"] is None:
				itemCenterY = int(top + itemHeight * index + itemHeight / 2)
				element["clickPoint"] = (centerX, itemCenterY)
			element.pop("_lineCenterY", None)
		return elements

	textLines = [line["text"].strip() for line in lines if line["text"].strip()]
	if not textLines:
		return []

	itemHeight = (bottom - top) / len(textLines)
	for index, text in enumerate(textLines):
		normalized = _normalizeLineText(text)
		if _NOISE_LINE_RE.fullmatch(normalized):
			continue
		itemCenterY = int(top + itemHeight * index + itemHeight / 2)
		elements.append(
			{
				"name": text,
				"role": None,
				"clickPoint": (centerX, itemCenterY),
				"_lineCenterY": None,
			},
		)
	_assignRowRectsToElements(elements, rowRects)
	for element in elements:
		element.pop("_lineCenterY", None)
	return elements


class ChatMoreOptions(VirtualWindow):
	title = "更多選項"

	@staticmethod
	def isMatchLineScreen(obj):
		return False

	def __init__(self, popupRect, rowRects=None, onAction=None):
		self.elements = []
		self.pos = -1
		self.popupRect = popupRect
		self.rowRects = rowRects or []
		self.onAction = onAction
		left, top, right, bottom = popupRect
		width = right - left
		height = bottom - top
		if width > 0 and height > 0:
			ocrGetText(left, top, width, height, self._onOcrResult)
		message(self.title)

	def makeElements(self):
		pass

	def _onOcrResult(self, result):
		if not result or isinstance(result, Exception):
			log.debug(f"LINE: ChatMoreOptions OCR error: {result}")
			return

		lineInfos = _extractOcrLines(result)
		if not lineInfos:
			text = getattr(result, "text", "") or ""
			text = _removeCJKSpaces(text.strip())
			lineInfos = [{"text": line.strip(), "rect": None} for line in text.split("\n") if line.strip()]

		if not lineInfos:
			log.debug("LINE: ChatMoreOptions OCR returned no lines")
			return

		self.elements = _buildMenuElements(
			lineInfos,
			self.popupRect,
			rowRects=self.rowRects,
		)
		log.debug(
			f"LINE: ChatMoreOptions click points: "
			f"{[(e['name'], e.get('clickPoint')) for e in self.elements]}",
		)

		log.info(
			f"LINE: ChatMoreOptions found {len(self.elements)} items: {[e['name'] for e in self.elements]}",
		)

		if self.elements:
			self.pos = 0
			self.show()

	def click(self):
		element = self.element
		actionName = element.get("name") if element else None
		hasClickPoint = bool(element and element.get("clickPoint"))
		super().click()
		VirtualWindow.currentWindow = None
		if hasClickPoint and callable(self.onAction):
			try:
				self.onAction(actionName)
			except Exception:
				log.debug("LINE: ChatMoreOptions action callback failed", exc_info=True)
