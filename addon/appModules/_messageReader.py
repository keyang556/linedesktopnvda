import wx
import gui
from logHandler import log


class MessageReaderDialog(wx.Dialog):
	"""A dialog for reading LINE chat messages with up/down arrow navigation.

	Messages are displayed as: name content time
	Date separators are displayed in their original positions.
	Up arrow moves to the previous message, down arrow moves to the next.
	"""

	def __init__(self, messages, title=None, cleanupPath=None):
		if title is None:
			title = _("訊息閱讀器")
		super().__init__(
			gui.mainFrame,
			title=title,
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._messages = messages
		self._pos = 0 if messages else -1
		self._cleanupPath = cleanupPath
		self._messageCount = sum(1 for msg in messages if msg.get('type') != 'date')
		self._messageIndexMap = []
		messageIndex = 0
		for msg in messages:
			if msg.get('type') != 'date':
				messageIndex += 1
			self._messageIndexMap.append(messageIndex)

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

		closeBtn = wx.Button(panel, wx.ID_CLOSE, _("關閉(&C)"))
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
		if msg.get('type') == 'date':
			return msg.get('content', '')
		return f"{msg['name']} {msg['content']} {msg['time']}"

	def _getProgressLabel(self):
		"""Return progress text counting only actual messages."""
		if self._messageCount <= 0 or self._pos < 0:
			return ""
		currentMessageIndex = self._messageIndexMap[self._pos]
		if self._messages[self._pos].get('type') == 'date':
			if currentMessageIndex < self._messageCount:
				currentMessageIndex += 1
			else:
				return ""
		return f"{currentMessageIndex} / {self._messageCount}"

	def _updateDisplay(self):
		if not self._messages or self._pos < 0:
			self._textCtrl.SetValue(_("沒有訊息"))
			self._totalLabel.SetLabel("")
			return
		msg = self._messages[self._pos]
		text = self._formatMessage(msg)
		self._textCtrl.SetValue(text)
		self._totalLabel.SetLabel(self._getProgressLabel())
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
			self._speakMessage(_("已經是第一項"))

	def _moveNext(self):
		if not self._messages:
			return
		if self._pos < len(self._messages) - 1:
			self._pos += 1
			self._updateDisplay()
		else:
			self._speakMessage(_("已經是最後一項"))

	def _onClose(self, evt):
		global _readerDlg
		_readerDlg = None  # Allow future invocations to create a new instance
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


_readerDlg = None  # module-level singleton sentinel


def openMessageReader(messages, title=None, cleanupPath=None):
	"""Open the message reader dialog on the main GUI thread.

	Args:
		messages: List of parsed message dicts
		title: Dialog window title
		cleanupPath: Optional file path to delete when dialog closes
	"""
	def _show():
		global _readerDlg
		if _readerDlg and _readerDlg.IsShown():
			_readerDlg.Raise()
			return
		_readerDlg = MessageReaderDialog(messages, title=title, cleanupPath=cleanupPath)
		_readerDlg.Show()
		_readerDlg.Raise()
	wx.CallAfter(_show)
