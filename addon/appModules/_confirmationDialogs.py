"""Standard NVDA message dialogs for LINE's recall and photo-to-text confirmations.

These replace the transient single-key confirmation modes (Y/N/P and A/D) the
add-on used to bind while LINE's own dialog was on screen. Those bindings had to
swallow plain letter keys inside LINE and needed a timeout to release them again;
a real dialog owns the focus instead, so the letters stay usable in LINE and no
timer is needed.

Built on ``gui.message.MessageDialog`` as documented in the NVDA developer guide.
The original letters survive as button accelerators, so pressing Y/N/P (or A/D)
with the dialog focused still picks the same action.
"""

import wx
import gui
from gui.message import Button, DialogType, MessageDialog, ReturnCode
from logHandler import log
import addonHandler

addonHandler.initTranslation()


_recallDialog = None
_photoConsentDialog = None


def _showOnGuiThread(show, dialogName, onFailed):
	"""Run ``show`` on the GUI thread; MessageDialogs cannot be built anywhere else.

	``onFailed`` runs when the dialog could not be shown, so the caller can
	release whatever it marked as pending. Without it a single failure would
	leave the feature wedged until NVDA restarts.
	"""

	def _guarded():
		try:
			show()
		except Exception:
			log.warning(f"LINE: showing the {dialogName} dialog failed", exc_info=True)
			if onFailed is not None:
				onFailed()

	wx.CallAfter(_guarded)


def _focusExistingDialog(dlg):
	"""Raise and focus an already-open dialog; return False if it is gone.

	A closed dialog leaves a stale Python wrapper behind, and touching the
	destroyed wx object raises RuntimeError; treat that as "no dialog".
	"""
	if dlg is None:
		return False
	try:
		dlg.Raise()
		dlg.SetFocus()
	except RuntimeError:
		return False
	return True


def _makeCallback(action, onAction, clearDialog):
	"""Build a button callback that reports ``action`` back to the caller.

	``MessageDialog`` hides the dialog before running the callback but only
	closes it afterwards, so hand control back to the event loop first: the
	caller acts on LINE's own dialog by OCRing and clicking it, which must not
	race our own dialog's teardown.
	"""

	def _callback(payload):
		clearDialog()
		wx.CallAfter(onAction, action)

	return _callback


def openRecallConfirmationDialog(message, availableActions, onAction, onFailed=None):
	"""Show LINE's recall confirmation as a standard NVDA message dialog.

	``availableActions`` are the canonical action names detected on LINE's own
	dialog, so stealth recall is only offered when LINE offers it. ``onAction``
	is called on the GUI thread with the chosen action name ("收回", "無痕收回"
	or "取消") once this dialog has closed. Closing the dialog with escape picks
	"取消", matching what the N key used to do. ``onFailed`` runs instead when
	the dialog could not be shown at all.
	"""
	offersStealthRecall = "無痕收回" in set(availableActions or ())

	def _show():
		global _recallDialog
		if _focusExistingDialog(_recallDialog):
			return

		def _clear():
			global _recallDialog
			_recallDialog = None

		buttons = [
			Button(
				id=ReturnCode.YES,
				# Translators: Button confirming a message recall. Its Y accelerator
				# is the key this action used to be bound to.
				label=_("收回(&Y)"),
				callback=_makeCallback("收回", onAction, _clear),
				defaultFocus=True,
			),
		]
		if offersStealthRecall:
			buttons.append(
				Button(
					id=ReturnCode.CUSTOM_1,
					# Translators: Button choosing LINE's stealth recall. Its P
					# accelerator is the key this action used to be bound to.
					label=_("無痕收回(&P)"),
					callback=_makeCallback("無痕收回", onAction, _clear),
				),
			)
		buttons.append(
			Button(
				id=ReturnCode.CANCEL,
				# Translators: Button cancelling a message recall. Its N accelerator
				# is the key this action used to be bound to.
				label=_("取消(&N)"),
				callback=_makeCallback("取消", onAction, _clear),
				fallbackAction=True,
			),
		)

		_recallDialog = MessageDialog(
			gui.mainFrame,
			message,
			# Translators: Title of the dialog asking whether to recall a message.
			_("LINE - 收回訊息"),
			DialogType.WARNING,
			buttons=buttons,
		)
		_recallDialog.Show()

	_showOnGuiThread(_show, "recall confirmation", onFailed)


def openPhotoTextConsentDialog(message, onAction, onFailed=None):
	"""Show LINE's first-run photo upload notice as a standard NVDA message dialog.

	``onAction`` is called on the GUI thread with "同意" or "不同意" once this
	dialog has closed. Closing the dialog with escape picks "不同意", matching
	what the D key (and the old ten-second timeout) used to do. ``onFailed``
	runs instead when the dialog could not be shown at all.
	"""

	def _show():
		global _photoConsentDialog
		if _focusExistingDialog(_photoConsentDialog):
			return

		def _clear():
			global _photoConsentDialog
			_photoConsentDialog = None

		_photoConsentDialog = MessageDialog(
			gui.mainFrame,
			message,
			# Translators: Title of the dialog asking whether to let LINE upload a
			# photo for its "Convert to text" feature.
			_("LINE - 同意提供照片"),
			DialogType.WARNING,
			buttons=(
				Button(
					id=ReturnCode.YES,
					# Translators: Button agreeing to LINE's photo upload notice. Its A
					# accelerator is the key this action used to be bound to.
					label=_("同意(&A)"),
					callback=_makeCallback("同意", onAction, _clear),
					defaultFocus=True,
				),
				Button(
					id=ReturnCode.NO,
					# Translators: Button declining LINE's photo upload notice. Its D
					# accelerator is the key this action used to be bound to.
					label=_("不同意(&D)"),
					callback=_makeCallback("不同意", onAction, _clear),
					fallbackAction=True,
				),
			),
		)
		_photoConsentDialog.Show()

	_showOnGuiThread(_show, "photo consent", onFailed)
