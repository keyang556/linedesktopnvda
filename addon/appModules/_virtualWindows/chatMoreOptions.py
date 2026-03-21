from .._virtualWindow import VirtualWindow
from .._utils import ocrGetText, message
from logHandler import log

import re

_CJK_CHAR = (
	r'[\u2E80-\u9FFF\uF900-\uFAFF'
	r'\U00020000-\U0002A6DF\U0002A700-\U0002EBEF\U00030000-\U000323AF]'
)
_CJK_SPACE_RE = re.compile(
	r'(?<=' + _CJK_CHAR + r') (?=' + _CJK_CHAR + r')'
)

# OCR 常見錯字修正表（原文 → 正確文字）
_OCR_CORRECTIONS = {
	'眧': '照',
}

def _removeCJKSpaces(text):
	return _CJK_SPACE_RE.sub('', text)

def _fixOcrErrors(text):
	for wrong, correct in _OCR_CORRECTIONS.items():
		text = text.replace(wrong, correct)
	return text


class ChatMoreOptions(VirtualWindow):
	title = '更多選項'

	@staticmethod
	def isMatchLineScreen(obj):
		return False

	def __init__(self, popupRect):
		self.elements = []
		self.pos = -1
		self.popupRect = popupRect
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

		text = getattr(result, 'text', '') or ''
		text = _fixOcrErrors(_removeCJKSpaces(text.strip()))
		lines = [l.strip() for l in text.split('\n') if l.strip()]

		if not lines:
			log.debug("LINE: ChatMoreOptions OCR returned no lines")
			return

		left, top, right, bottom = self.popupRect
		height = bottom - top
		itemHeight = height / len(lines)
		centerX = (left + right) // 2

		for i, line in enumerate(lines):
			itemCenterY = int(top + itemHeight * i + itemHeight / 2)
			self.elements.append({
				'name': line,
				'role': None,
				'clickPoint': (centerX, itemCenterY)
			})

		log.info(f"LINE: ChatMoreOptions found {len(self.elements)} items: {[e['name'] for e in self.elements]}")

		if self.elements:
			self.pos = 0
			self.show()

	def click(self):
		super().click()
		VirtualWindow.currentWindow = None
