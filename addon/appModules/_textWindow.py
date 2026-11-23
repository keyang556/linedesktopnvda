"""Read-only multi-line text window, as UnigramPlus uses for NVDA+Alt+V.

Arrow keys read the text line by line; Escape or Alt+F4 closes the window.
"""

import wx
import gui


class TextWindow(wx.Frame):
	"""A top-level window holding one read-only multi-line text field."""

	def __init__(self, text, title):
		super().__init__(gui.mainFrame, title=title)
		sizer = wx.BoxSizer(wx.VERTICAL)
		self.textCtrl = wx.TextCtrl(
			self,
			style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2,
			size=(500, 200),
		)
		self.textCtrl.Bind(wx.EVT_KEY_DOWN, self._onKeyDown)
		sizer.Add(self.textCtrl, proportion=1, flag=wx.EXPAND)
		self.SetSizer(sizer)
		sizer.Fit(self)
		self.textCtrl.SetValue(text)
		self.textCtrl.SetInsertionPoint(0)
		self.CentreOnScreen()

	def _onKeyDown(self, evt):
		if evt.GetKeyCode() == wx.WXK_ESCAPE:
			self.Close()
			return
		evt.Skip()


_window = None


def showTextWindow(text, title):
	"""Show ``text`` in a read-only window on the GUI thread, replacing any open one."""

	def _show():
		global _window
		# A destroyed wx window is falsy, so a window the user closed is skipped.
		if _window:
			_window.Destroy()
		_window = TextWindow(text, title)
		_window.Show()
		_window.Raise()
		_window.textCtrl.SetFocus()

	wx.CallAfter(_show)
