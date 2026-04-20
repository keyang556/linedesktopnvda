from .._virtualWindow import VirtualWindow
from .._utils import ocrGetText


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
				{"name": "電腦版認證"},
				{"name": "為了確保帳號安全性,您必須在首次登入電腦版時認證您的帳號。"},
				{"name": "請在智慧手機上輸入以下代碼。"},
			],
		)
		ocrGetText(*self.obj.location, self.onOcrResult)

	def onOcrResult(self, result):
		import re

		text = getattr(result, "text", "")
		if not text:
			return

		matchPinCode = re.search(r"\d+", text)
		if not matchPinCode:
			return

		self.elements.extend(
			[
				{"name": matchPinCode.group()},
				{"name": "輸入完成後,請點選智慧手機上的確認按鈕。"},
				{"name": "若無法在智慧手機上使用LINE,請執行移動LINE帳號的操作。"},
			],
		)
