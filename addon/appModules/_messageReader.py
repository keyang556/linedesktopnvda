import wx
import gui
from logHandler import log


class MessageReaderDialog(wx.Dialog):
	"""A dialog for reading LINE chat messages with up/down arrow navigation.

	Each message is displayed as: name content time
	Up arrow moves to the previous message, down arrow moves to the next.
	"""

	def __init__(self, messages, title="訊息閱讀器", cleanupPath=None):
		super().__init__(
			gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._messages = messages
		self._pos = len(messages) - 1 if messages else -1
		self._cleanupPath = cleanupPath

		panel = wx.Panel(self)
		sizer = wx.BoxSizer(wx.VERTICAL)

		self._totalLabel = wx.StaticText(panel, label="")
		sizer.Add(self._totalLabel, 0, wx.ALL, 5)

		self._textCtrl = wx.TextCtrl(
			panel,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
			size=(500, 100),
		)
		sizer.Add(self._textCtrl, 1, wx.EXPAND | wx.ALL, 5)

		closeBtn = wx.Button(panel, wx.ID_CLOSE, "關閉(&C)")
		sizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.ALL, 5)

		panel.SetSizer(sizer)
		sizer.Fit(self)

		self._textCtrl.Bind(wx.EVT_KEY_DOWN, self._onKeyDown)
		closeBtn.Bind(wx.EVT_BUTTON, self._onClose)
		self.Bind(wx.EVT_CLOSE, self._onClose)
		self.Bind(wx.EVT_CHAR_HOOK, self._onCharHook)

		self._updateDisplay()
		self._textCtrl.SetFocus()

	def _formatMessage(self, msg):
		return f"{msg['name']} {msg['content']} {msg['time']}"

	def _updateDisplay(self):
		if not self._messages or self._pos < 0:
			self._textCtrl.SetValue("沒有訊息")
			self._totalLabel.SetLabel("")
			return
		msg = self._messages[self._pos]
		text = self._formatMessage(msg)
		self._textCtrl.SetValue(text)
		self._totalLabel.SetLabel(
			f"{self._pos + 1} / {len(self._messages)}"
		)
		self._speakMessage(text)

	def _speakMessage(self, text):
		try:
			import speech
			speech.cancelSpeech()
			speech.speakMessage(text)
		except Exception:
			pass

	def _onKeyDown(self, evt):
		keyCode = evt.GetKeyCode()
		if keyCode == wx.WXK_UP:
			self._movePrevious()
		elif keyCode == wx.WXK_DOWN:
			self._moveNext()
		elif keyCode == wx.WXK_ESCAPE:
			self.Close()
		else:
			evt.Skip()

	def _onCharHook(self, evt):
		keyCode = evt.GetKeyCode()
		if keyCode == wx.WXK_ESCAPE:
			self.Close()
			return
		evt.Skip()

	def _movePrevious(self):
		if not self._messages:
			return
		if self._pos > 0:
			self._pos -= 1
			self._updateDisplay()
		else:
			self._speakMessage("已經是第一則訊息")

	def _moveNext(self):
		if not self._messages:
			return
		if self._pos < len(self._messages) - 1:
			self._pos += 1
			self._updateDisplay()
		else:
			self._speakMessage("已經是最後一則訊息")

	def _onClose(self, evt):
		# Clean up temp file if specified
		if self._cleanupPath:
			try:
				import os
				if os.path.isfile(self._cleanupPath):
					os.remove(self._cleanupPath)
					log.debug(f"Deleted temp chat export: {self._cleanupPath}")
			except Exception as e:
				log.warning(f"Failed to delete temp chat export: {e}")
		self.Destroy()


def openMessageReader(messages, title="訊息閱讀器", cleanupPath=None):
	"""Open the message reader dialog on the main GUI thread.

	Args:
		messages: List of parsed message dicts
		title: Dialog window title
		cleanupPath: Optional file path to delete when dialog closes
	"""
	def _show():
		dlg = MessageReaderDialog(messages, title=title, cleanupPath=cleanupPath)
		dlg.Show()
		dlg.Raise()
	wx.CallAfter(_show)
