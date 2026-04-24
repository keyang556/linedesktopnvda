import threading

import wx
import gui
from logHandler import log


class ImageDescriptionDialog(wx.Dialog):
	"""Image description dialog with a read-only transcript and multi-turn follow-up input."""

	def __init__(self, apiCaller, initialContents, initialUserPrompt, initialDescription):
		super().__init__(
			gui.mainFrame,
			title=_("圖片描述"),
			style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER,
		)
		self._apiCaller = apiCaller
		self._contents = list(initialContents)
		self._contents.append(
			{"role": "model", "parts": [{"text": initialDescription}]},
		)
		self._transcript = ""
		self._pending = False
		self._closed = False

		panel = wx.Panel(self)
		sizer = wx.BoxSizer(wx.VERTICAL)

		# Translators: Label above the conversation transcript.
		transcriptLabel = wx.StaticText(panel, label=_("對話內容(&T)："))
		sizer.Add(transcriptLabel, 0, wx.LEFT | wx.RIGHT | wx.TOP, 5)

		self._transcriptCtrl = wx.TextCtrl(
			panel,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
			size=(600, 300),
		)
		sizer.Add(self._transcriptCtrl, 1, wx.EXPAND | wx.ALL, 5)

		self._statusLabel = wx.StaticText(panel, label="")
		sizer.Add(self._statusLabel, 0, wx.LEFT | wx.RIGHT, 5)

		# Translators: Label above the follow-up question input box.
		inputLabel = wx.StaticText(
			panel,
			label=_("後續提問(&Q)（按 Ctrl+Enter 送出）："),
		)
		sizer.Add(inputLabel, 0, wx.LEFT | wx.RIGHT | wx.TOP, 5)

		self._inputCtrl = wx.TextCtrl(
			panel,
			style=wx.TE_MULTILINE,
			size=(600, 80),
		)
		sizer.Add(self._inputCtrl, 0, wx.EXPAND | wx.ALL, 5)

		btnRow = wx.BoxSizer(wx.HORIZONTAL)
		# Translators: Button that sends the follow-up question.
		self._sendBtn = wx.Button(panel, label=_("送出(&S)"))
		# Translators: Button that closes the image description dialog.
		closeBtn = wx.Button(panel, wx.ID_CLOSE, _("關閉(&C)"))
		btnRow.Add(self._sendBtn, 0, wx.RIGHT, 10)
		btnRow.Add(closeBtn, 0)
		sizer.Add(btnRow, 0, wx.ALIGN_CENTER | wx.ALL, 5)

		panel.SetSizer(sizer)
		sizer.Fit(self)

		# Translators: Transcript label for the original prompt turn.
		self._appendTurn(_("提示"), initialUserPrompt)
		# Translators: Transcript label for the model's initial description.
		self._appendTurn(_("描述"), initialDescription)

		self._sendBtn.Bind(wx.EVT_BUTTON, self._onSend)
		closeBtn.Bind(wx.EVT_BUTTON, self._onClose)
		self.Bind(wx.EVT_CLOSE, self._onClose)
		self.Bind(wx.EVT_CHAR_HOOK, self._onCharHook)

		self.SetEscapeId(wx.ID_CLOSE)
		# wx.CallAfter defers speech until after Show() so NVDA's focus announcement doesn't cut it off.
		self._transcriptCtrl.SetInsertionPoint(0)
		self._transcriptCtrl.SetFocus()
		wx.CallAfter(self._speak, initialDescription)

	def _appendTurn(self, speaker, text):
		if self._transcript:
			self._transcript += "\n\n"
		self._transcript += f"【{speaker}】\n{text}"
		self._transcriptCtrl.SetValue(self._transcript)
		self._transcriptCtrl.ShowPosition(self._transcriptCtrl.GetLastPosition())

	def _speak(self, text):
		"""Speak ``text`` via NVDA, cancelling any prior speech."""
		if not text:
			return
		try:
			import speech

			speech.cancelSpeech()
			speech.speakMessage(text)
		except Exception:
			log.debug("LINE: image description dialog speech failed", exc_info=True)

	def _onCharHook(self, evt):
		keyCode = evt.GetKeyCode()
		if keyCode == wx.WXK_ESCAPE:
			self.Close()
			return
		if (
			keyCode in (wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER)
			and evt.ControlDown()
			and self.FindFocus() is self._inputCtrl
		):
			self._onSend(evt)
			return
		evt.Skip()

	def _onSend(self, evt):
		if self._pending:
			return
		question = self._inputCtrl.GetValue().strip()
		if not question:
			return
		self._inputCtrl.SetValue("")
		# Translators: Transcript label for the user's follow-up question.
		self._appendTurn(_("提問"), question)

		followup = list(self._contents)
		followup.append({"role": "user", "parts": [{"text": question}]})

		self._pending = True
		self._sendBtn.Disable()
		self._inputCtrl.Disable()
		# Translators: Status text shown while waiting for the follow-up reply.
		waitMsg = _("回答中，請稍候…")
		self._statusLabel.SetLabel(waitMsg)
		self._speak(waitMsg)

		def worker():
			try:
				answer, err = self._apiCaller(followup)
			except Exception as e:
				log.warning(
					f"LINE: image description follow-up failed: {e}",
					exc_info=True,
				)
				answer, err = None, _("圖片描述失敗")
			wx.CallAfter(self._onApiResult, question, answer, err)

		threading.Thread(target=worker, daemon=True).start()

	def _onApiResult(self, question, answer, err):
		if self._closed:
			return
		self._pending = False
		self._sendBtn.Enable()
		self._inputCtrl.Enable()
		self._statusLabel.SetLabel("")
		if answer:
			self._contents.append(
				{"role": "user", "parts": [{"text": question}]},
			)
			self._contents.append(
				{"role": "model", "parts": [{"text": answer}]},
			)
			# Translators: Transcript label for the model's follow-up reply.
			self._appendTurn(_("回答"), answer)
			self._transcriptCtrl.SetInsertionPoint(
				self._transcriptCtrl.GetLastPosition(),
			)
			self._transcriptCtrl.SetFocus()
			self._speak(answer)
		else:
			msg = err or _("回答失敗")
			# Translators: Transcript label for an error reply.
			self._appendTurn(_("錯誤"), msg)
			self._inputCtrl.SetFocus()
			self._speak(msg)

	def _onClose(self, evt):
		global _dlg
		self._closed = True  # prevent in-flight wx.CallAfter callbacks from touching destroyed widgets
		_dlg = None
		self.Destroy()


_dlg = None


def openImageDescriptionDialog(
	apiCaller,
	initialContents,
	initialUserPrompt,
	initialDescription,
):
	"""Show the image-description dialog on the GUI thread (singleton)."""

	def _show():
		global _dlg
		if _dlg and _dlg.IsShown():
			_dlg.Close()
		_dlg = ImageDescriptionDialog(
			apiCaller,
			initialContents,
			initialUserPrompt,
			initialDescription,
		)
		_dlg.Show()
		_dlg.Raise()

	wx.CallAfter(_show)
