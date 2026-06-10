from .._virtualWindow import VirtualWindow
from .._utils import ocrGetText
import addonHandler

addonHandler.initTranslation()


class PinCode(VirtualWindow):
	title = "電腦版認證"

	@staticmethod
	def isMatchLineScreen(obj):
		try:
			return obj.UIAElement.CurrentClassName == "PinCodeInputWindow"
		except AttributeError:
			return False

	def makeElements(self):
		self.elements.extend(
			[
				# Translators: Title line of LINE's PC login verification screen.
				{"name": _("電腦版認證")},
				# Translators: Description line of LINE's PC login verification screen.
				{"name": _("為了確保帳號安全性,您必須在首次登入電腦版時認證您的帳號。")},
				# Translators: Instruction line of LINE's PC login verification screen.
				{"name": _("請在智慧手機上輸入以下代碼。")},
			],
		)
		location = getattr(self.obj, "location", None)
		if location:
			ocrGetText(*location, self.onOcrResult)

	def onOcrResult(self, result):
		# UWP OCR invokes this on a background thread; hop to the main thread
		# before touching the elements list.
		import core

		core.callLater(0, self._handleOcrResult, result)

	def _handleOcrResult(self, result):
		import re

		if VirtualWindow.currentWindow is not self:
			# The verification screen went away while OCR was still running.
			return
		text = getattr(result, "text", "")
		if not text:
			return

		matchPinCode = re.search(r"\d+", text)
		if not matchPinCode:
			return

		self.elements.extend(
			[
				{"name": matchPinCode.group()},
				# Translators: Instruction line of LINE's PC login verification screen.
				{"name": _("輸入完成後,請點選智慧手機上的確認按鈕。")},
				# Translators: Instruction line of LINE's PC login verification screen.
				{"name": _("若無法在智慧手機上使用LINE,請執行移動LINE帳號的操作。")},
			],
		)
