# LINE Desktop Global Plugin for NVDA
# Maps alternative executable names to the LINE appModule.
# Also adds a "LINE Desktop" submenu under NVDA's Tools menu.

import appModuleHandler
import globalPluginHandler
from scriptHandler import script
from logHandler import log
import gui
from gui import guiHelper
from gui.settingsDialogs import SettingsPanel
import wx
import winreg
import ctypes
import addonHandler

addonHandler.initTranslation()


# ---------------------------------------------------------------------------
# Qt accessibility environment variable helpers (duplicated from line.py
# so the global plugin can toggle the setting even when LINE is not running)
# ---------------------------------------------------------------------------

_QT_ACCESSIBILITY_ENV_NAME = "QT_ACCESSIBILITY"
_HWND_BROADCAST = 0xFFFF
_WM_SETTINGCHANGE = 0x001A
_SMTO_ABORTIFHUNG = 0x0002


def _isQtAccessibleSet():
	"""Check if QT_ACCESSIBILITY=1 is set in user environment variables."""
	try:
		with winreg.OpenKey(
			winreg.HKEY_CURRENT_USER,
			"Environment",
			0,
			winreg.KEY_READ,
		) as key:
			value, _ = winreg.QueryValueEx(key, _QT_ACCESSIBILITY_ENV_NAME)
			return str(value) == "1"
	except FileNotFoundError:
		return False
	except Exception:
		return False


def _setQtAccessible(enable=True):
	"""Set or remove QT_ACCESSIBILITY in user environment variables."""
	try:
		with winreg.OpenKey(
			winreg.HKEY_CURRENT_USER,
			"Environment",
			0,
			winreg.KEY_SET_VALUE | winreg.KEY_READ,
		) as key:
			if enable:
				winreg.SetValueEx(
					key,
					_QT_ACCESSIBILITY_ENV_NAME,
					0,
					winreg.REG_SZ,
					"1",
				)
			else:
				try:
					winreg.DeleteValue(key, _QT_ACCESSIBILITY_ENV_NAME)
				except FileNotFoundError:
					pass
		ctypes.windll.user32.SendMessageTimeoutW(
			_HWND_BROADCAST,
			_WM_SETTINGCHANGE,
			0,
			"Environment",
			_SMTO_ABORTIFHUNG,
			5000,
			None,
		)
		return True
	except Exception:
		log.warning("Failed to set QT_ACCESSIBILITY in registry", exc_info=True)
		return False


def _getLineAppModule():
	"""Find and return the LINE appModule instance, or None."""
	for app in appModuleHandler.runningTable.values():
		if app and getattr(app, "appName", "").lower() in (
			"line",
			"line_app",
			"linecall",
		):
			return app
	return None


class LineDesktopSettingsPanel(SettingsPanel):
	"""Settings panel shown under NVDA Preferences → Settings → LINE Desktop."""

	# Translators: Title of the LINE Desktop settings panel in NVDA Preferences
	title = _("LINE Desktop")

	def makeSettings(self, settingsSizer):
		sHelper = guiHelper.BoxSizerHelper(self, sizer=settingsSizer)

		# Translators: Checkbox label in LINE Desktop settings panel
		qtLabel = _("啟用 Qt 無障礙環境變數 QT_ACCESSIBILITY=1 (&Q)")
		self._qtCheck = sHelper.addItem(wx.CheckBox(self, label=qtLabel))
		self._qtCheck.SetValue(_isQtAccessibleSet())

		providerOptions = self._loadProviderOptions()
		self._providerIds = tuple(opt[0] for opt in providerOptions)
		providerLabels = [opt[1] for opt in providerOptions]
		# Translators: Dropdown label for selecting the image-description backend provider.
		providerLabel = _("圖片描述服務 (&D)")
		self._providerChoice = sHelper.addLabeledControl(
			providerLabel,
			wx.Choice,
			choices=providerLabels,
		)
		currentProvider = self._loadCurrentProvider()
		try:
			providerIndex = self._providerIds.index(currentProvider)
		except ValueError:
			providerIndex = 0
		if self._providerIds:
			self._providerChoice.SetSelection(providerIndex)

		# Pending edits per-provider so switching the dropdown doesn't lose
		# in-flight changes the user has made for the other backend.
		self._pendingApiKey = {pid: self._loadStoredApiKey(pid) for pid in self._providerIds}
		self._pendingModel = {pid: self._loadStoredModel(pid) for pid in self._providerIds}

		# Translators: Text field label for the image-description API key
		apiLabel = _("圖片描述 API Key，留空則使用預設金鑰 (&I)")
		self._apiKeyText = sHelper.addLabeledControl(apiLabel, wx.TextCtrl)

		# Translators: Dropdown label for selecting the image-description model
		modelLabel = _("圖片描述模型 (&M)")
		self._modelChoice = sHelper.addLabeledControl(
			modelLabel,
			wx.Choice,
			choices=[],
		)

		# Translators: Text field label for the image-description prompt
		promptLabel = _("圖片描述提示詞，留空則使用預設 (&P)")
		self._promptText = sHelper.addLabeledControl(
			promptLabel,
			wx.TextCtrl,
			style=wx.TE_MULTILINE,
			size=(-1, 80),
		)
		self._promptText.SetValue(self._loadCurrentPrompt())

		self._activeProviderId = self._currentSelectedProviderId() or _safeDefaultProvider()
		self._refreshProviderUI(self._activeProviderId)

		self._providerChoice.Bind(wx.EVT_CHOICE, self._onProviderChange)

	def _currentSelectedProviderId(self):
		"""Return the provider ID matching the dropdown selection, or None."""
		idx = self._providerChoice.GetSelection() if hasattr(self, "_providerChoice") else wx.NOT_FOUND
		if idx == wx.NOT_FOUND or not self._providerIds:
			return None
		if 0 <= idx < len(self._providerIds):
			return self._providerIds[idx]
		return None

	def _onProviderChange(self, evt):
		# Stash whatever the user has typed/selected for the previously active
		# provider before we repaint the controls for the newly selected one.
		previous = getattr(self, "_activeProviderId", None)
		if previous and previous in self._pendingApiKey:
			self._pendingApiKey[previous] = self._apiKeyText.GetValue().strip()
		if previous and previous in self._pendingModel:
			selectedModel = self._currentModelChoiceValue()
			if selectedModel is not None:
				self._pendingModel[previous] = selectedModel

		newProvider = self._currentSelectedProviderId() or previous
		self._activeProviderId = newProvider
		self._refreshProviderUI(newProvider)

	def _currentModelChoiceValue(self):
		choices = getattr(self, "_modelChoices", ())
		idx = self._modelChoice.GetSelection()
		if idx == wx.NOT_FOUND or not choices:
			return None
		if 0 <= idx < len(choices):
			return choices[idx]
		return None

	def _refreshProviderUI(self, provider):
		"""Repopulate the API-key text field and the model dropdown for ``provider``."""
		choices, defaultModel = self._modelOptionsFor(provider)
		self._modelChoices = choices
		self._modelChoice.Clear()
		if choices:
			self._modelChoice.AppendItems(list(choices))
		desiredModel = self._pendingModel.get(provider) or defaultModel
		try:
			selectIndex = choices.index(desiredModel)
		except ValueError:
			try:
				selectIndex = choices.index(defaultModel)
			except ValueError:
				selectIndex = 0
		if choices:
			self._modelChoice.SetSelection(selectIndex)

		self._apiKeyText.ChangeValue(self._pendingApiKey.get(provider, ""))

	def _loadProviderOptions(self):
		"""Return [(provider_id, display_label), …] in the order shown to the user."""
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS,
				_IMAGE_DESCRIPTION_PROVIDER_LABELS,
			)

			return [
				(pid, _IMAGE_DESCRIPTION_PROVIDER_LABELS.get(pid, pid))
				for pid in _IMAGE_DESCRIPTION_AVAILABLE_PROVIDERS
			]
		except Exception:
			log.debug("LINE: cannot load provider options", exc_info=True)
			return []

	def _loadCurrentProvider(self):
		try:
			from appModules.line import getUserImageProvider

			return getUserImageProvider() or _safeDefaultProvider()
		except Exception:
			log.debug("LINE: cannot load current image provider", exc_info=True)
			return _safeDefaultProvider()

	def _loadStoredApiKey(self, provider):
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
				getUserImageApiKey,
				getUserOllamaApiKey,
			)

			if provider == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
				return getUserOllamaApiKey() or ""
			return getUserImageApiKey() or ""
		except Exception:
			log.debug(
				f"LINE: cannot load API key for provider {provider!r}",
				exc_info=True,
			)
			return ""

	def _loadStoredModel(self, provider):
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
				getUserImageModel,
				getUserOllamaModel,
			)

			if provider == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
				return getUserOllamaModel() or _IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL
			return getUserImageModel() or _IMAGE_DESCRIPTION_DEFAULT_MODEL
		except Exception:
			log.debug(
				f"LINE: cannot load model for provider {provider!r}",
				exc_info=True,
			)
			return ""

	def _modelOptionsFor(self, provider):
		"""Return (choices_tuple, default_model_id) for the given provider."""
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_AVAILABLE_MODELS,
				_IMAGE_DESCRIPTION_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS,
				_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
			)

			if provider == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
				return (
					_IMAGE_DESCRIPTION_OLLAMA_AVAILABLE_MODELS,
					_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL,
				)
			return (
				_IMAGE_DESCRIPTION_AVAILABLE_MODELS,
				_IMAGE_DESCRIPTION_DEFAULT_MODEL,
			)
		except Exception:
			log.debug("LINE: cannot load image model options", exc_info=True)
			return ((), "")

	def _loadCurrentPrompt(self):
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_DEFAULT_PROMPT,
				getUserImagePrompt,
			)

			return getUserImagePrompt() or _IMAGE_DESCRIPTION_DEFAULT_PROMPT
		except Exception:
			log.debug("LINE: cannot load image prompt for settings panel", exc_info=True)
			return ""

	def onSave(self):
		# Qt accessibility env var
		wantSet = bool(self._qtCheck.GetValue())
		if wantSet != _isQtAccessibleSet():
			if not _setQtAccessible(wantSet):
				gui.messageBox(
					# Translators: Error shown when writing Qt accessibility env var fails
					_("設定 Qt 無障礙環境變數失敗，請確認系統權限。"),
					# Translators: Title of the settings error dialog
					_("LINE Desktop - 設定錯誤"),
					wx.OK | wx.ICON_ERROR,
					self,
				)

		# Capture pending edits for the currently visible provider before saving.
		activeProvider = self._currentSelectedProviderId() or _safeDefaultProvider()
		if activeProvider in self._pendingApiKey:
			self._pendingApiKey[activeProvider] = self._apiKeyText.GetValue().strip()
		if activeProvider in self._pendingModel:
			selectedModel = self._currentModelChoiceValue()
			if selectedModel is not None:
				self._pendingModel[activeProvider] = selectedModel

		# Active provider selection
		try:
			from appModules.line import getUserImageProvider, setUserImageProvider

			currentProvider = getUserImageProvider() or _safeDefaultProvider()
			if activeProvider != currentProvider:
				if not setUserImageProvider(activeProvider):
					gui.messageBox(
						# Translators: Error shown when saving the provider fails
						_("儲存圖片描述服務失敗，請重試。"),
						_("LINE Desktop - 設定錯誤"),
						wx.OK | wx.ICON_ERROR,
						self,
					)
		except Exception:
			log.warning(
				"LINE: cannot load image provider helpers from settings panel",
				exc_info=True,
			)

		# Image description API keys (one per provider)
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
				getUserImageApiKey,
				getUserOllamaApiKey,
				setUserImageApiKey,
				setUserOllamaApiKey,
			)

			for providerId, pendingKey in self._pendingApiKey.items():
				if providerId == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
					currentKey = getUserOllamaApiKey() or ""
					setter = setUserOllamaApiKey
				else:
					currentKey = getUserImageApiKey() or ""
					setter = setUserImageApiKey
				if pendingKey != currentKey:
					if not setter(pendingKey):
						gui.messageBox(
							# Translators: Error shown when saving the image API key fails
							_("儲存圖片描述 API Key 失敗，請重試。"),
							_("LINE Desktop - 設定錯誤"),
							wx.OK | wx.ICON_ERROR,
							self,
						)
		except Exception:
			log.warning(
				"LINE: cannot load image API key helpers from settings panel",
				exc_info=True,
			)
			gui.messageBox(
				# Translators: Error shown when the API key module cannot be loaded
				_("無法載入 API Key 設定，請確認附加元件完整性。"),
				_("LINE Desktop - 設定錯誤"),
				wx.OK | wx.ICON_ERROR,
				self,
			)

		# Image description models (one per provider)
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL,
				_IMAGE_DESCRIPTION_PROVIDER_OLLAMA,
				getUserImageModel,
				getUserOllamaModel,
				setUserImageModel,
				setUserOllamaModel,
			)

			for providerId, pendingModel in self._pendingModel.items():
				if not pendingModel:
					continue
				if providerId == _IMAGE_DESCRIPTION_PROVIDER_OLLAMA:
					currentModel = getUserOllamaModel() or _IMAGE_DESCRIPTION_OLLAMA_DEFAULT_MODEL
					setter = setUserOllamaModel
				else:
					currentModel = getUserImageModel() or _IMAGE_DESCRIPTION_DEFAULT_MODEL
					setter = setUserImageModel
				if pendingModel != currentModel:
					if not setter(pendingModel):
						gui.messageBox(
							# Translators: Error shown when saving the image model fails
							_("儲存圖片描述模型失敗，請重試。"),
							_("LINE Desktop - 設定錯誤"),
							wx.OK | wx.ICON_ERROR,
							self,
						)
		except Exception:
			log.warning(
				"LINE: cannot load image model helpers from settings panel",
				exc_info=True,
			)

		# Image description prompt
		try:
			from appModules.line import (
				_IMAGE_DESCRIPTION_DEFAULT_PROMPT,
				getUserImagePrompt,
				setUserImagePrompt,
			)

			newPrompt = self._promptText.GetValue().strip()
			currentPrompt = getUserImagePrompt() or _IMAGE_DESCRIPTION_DEFAULT_PROMPT
			if newPrompt != currentPrompt:
				if not setUserImagePrompt(newPrompt):
					gui.messageBox(
						# Translators: Error shown when saving the image prompt fails
						_("儲存圖片描述提示詞失敗，請重試。"),
						_("LINE Desktop - 設定錯誤"),
						wx.OK | wx.ICON_ERROR,
						self,
					)
		except Exception:
			log.warning(
				"LINE: cannot load image prompt helpers from settings panel",
				exc_info=True,
			)


def _safeDefaultProvider():
	"""Return the default provider ID, falling back to a literal if line.py is unavailable."""
	try:
		from appModules.line import _IMAGE_DESCRIPTION_DEFAULT_PROVIDER

		return _IMAGE_DESCRIPTION_DEFAULT_PROVIDER
	except Exception:
		return "google"


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Global plugin to ensure the LINE appModule is loaded
	for all known LINE desktop executable variants.

	LINE desktop may run as:
	- LINE.exe (standard installer)
	- LINE_APP.exe (Microsoft Store version, older)
	- LineLauncher.exe (launcher process)
	"""

	# Alternative executable names that should use the line appModule.
	# NVDA lowercases exe names, so "LINE.exe" becomes "line" automatically.
	# We register additional variants here for safety.
	_LINE_EXECUTABLES = [
		"LINE",
		"Line",
		"LINE_APP",
		"LineCall",
		"linecall",
	]

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		for exe in self._LINE_EXECUTABLES:
			try:
				appModuleHandler.registerExecutableWithAppModule(exe, "line")
				log.debug(f"Registered {exe} with line appModule")
			except Exception:
				log.debugWarning(
					f"Failed to register {exe} with LINE appModule",
					exc_info=True,
				)
		self._toolsMenu = None
		self._lineSubMenu = None
		self._createToolsMenu()
		self._registerSettingsPanel()

	def _registerSettingsPanel(self):
		try:
			if LineDesktopSettingsPanel not in gui.NVDASettingsDialog.categoryClasses:
				gui.NVDASettingsDialog.categoryClasses.append(LineDesktopSettingsPanel)
		except Exception:
			log.debugWarning(
				"Failed to register LINE Desktop settings panel",
				exc_info=True,
			)

	def _unregisterSettingsPanel(self):
		try:
			gui.NVDASettingsDialog.categoryClasses.remove(LineDesktopSettingsPanel)
		except (ValueError, AttributeError):
			pass
		except Exception:
			log.debugWarning(
				"Failed to unregister LINE Desktop settings panel",
				exc_info=True,
			)

	def _createToolsMenu(self):
		"""Create the LINE Desktop submenu under NVDA Tools menu."""
		try:
			self._lineSubMenu = wx.Menu()

			# ── Chat room tab navigation ──
			self._allChatsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to All chats tab
				_("全部聊天室(&1)") + "\tNVDA+Windows+1",
			)
			self._friendsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Friends tab
				_("好友(&2)") + "\tNVDA+Windows+2",
			)
			self._groupsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Groups tab
				_("群組(&3)") + "\tNVDA+Windows+3",
			)
			self._communitiesItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Communities tab
				_("社群(&4)") + "\tNVDA+Windows+4",
			)
			self._officialItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for navigating to Official accounts tab
				_("官方帳號(&5)") + "\tNVDA+Windows+5",
			)

			self._lineSubMenu.AppendSeparator()

			# ── Chat functions ──
			self._voiceCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for making a voice call
				_("語音通話(&C)") + "\tNVDA+Windows+C",
			)
			self._videoCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for making a video call
				_("視訊通話(&V)") + "\tNVDA+Windows+V",
			)
			self._moreOptionsItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for clicking more options button
				_("更多選項(&O)") + "\tNVDA+Windows+O",
			)
			self._messageReaderItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for opening the message reader
				_("訊息閱讀器(&J)") + "\tNVDA+Windows+J",
			)
			self._readChatNameItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for reading chat room name
				_("讀出聊天室名稱(&T)") + "\tNVDA+Windows+T",
			)
			self._describeImageItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for describing the current image message
				_("圖片描述(&I)") + "\tNVDA+Windows+I",
			)

			self._lineSubMenu.AppendSeparator()

			# ── Incoming call functions ──
			self._answerCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for answering a call
				_("接聽來電(&A)") + "\tNVDA+Windows+A",
			)
			self._rejectCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for rejecting a call
				_("拒絕來電(&D)") + "\tNVDA+Windows+D",
			)
			self._checkCallerItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for checking who is calling
				_("查看來電者(&S)") + "\tNVDA+Windows+S",
			)
			self._focusCallItem = self._lineSubMenu.Append(
				wx.ID_ANY,
				# Translators: Menu item for focusing the call window
				_("跳到通話視窗(&F)") + "\tNVDA+Windows+F",
			)

			# Bind events
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onAllChats,
				self._allChatsItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onFriends,
				self._friendsItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onGroups,
				self._groupsItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onCommunities,
				self._communitiesItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onOfficial,
				self._officialItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onVoiceCall,
				self._voiceCallItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onVideoCall,
				self._videoCallItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onMoreOptions,
				self._moreOptionsItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onMessageReader,
				self._messageReaderItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onReadChatName,
				self._readChatNameItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onDescribeImage,
				self._describeImageItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onAnswerCall,
				self._answerCallItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onRejectCall,
				self._rejectCallItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onCheckCaller,
				self._checkCallerItem,
			)
			gui.mainFrame.sysTrayIcon.Bind(
				wx.EVT_MENU,
				self._onFocusCallWindow,
				self._focusCallItem,
			)

			# Add the submenu to NVDA's Tools menu
			self._toolsMenu = gui.mainFrame.sysTrayIcon.toolsMenu
			# Translators: The label for the LINE Desktop submenu in NVDA's Tools menu
			self._lineMenuItem = self._toolsMenu.AppendSubMenu(
				self._lineSubMenu,
				"LINE Desktop",
			)
			log.info("LINE Desktop: tools menu created")
		except Exception:
			log.debugWarning(
				"Failed to create LINE Desktop tools menu",
				exc_info=True,
			)

	def _removeToolsMenu(self):
		"""Remove the LINE Desktop submenu from NVDA Tools menu."""
		try:
			if self._toolsMenu and self._lineMenuItem:
				self._toolsMenu.Remove(self._lineMenuItem)
				self._lineMenuItem = None
				log.info("LINE Desktop: tools menu removed")
		except Exception:
			log.debugWarning(
				"Failed to remove LINE Desktop tools menu",
				exc_info=True,
			)
		self._lineSubMenu = None
		self._toolsMenu = None

	# ── Menu event handlers ──────────────────────────────────────────

	def _onAllChats(self, evt):
		wx.CallAfter(self._doNavigateTab, "全部")

	def _onFriends(self, evt):
		wx.CallAfter(self._doNavigateTab, "好友")

	def _onGroups(self, evt):
		wx.CallAfter(self._doNavigateTab, "群組")

	def _onCommunities(self, evt):
		wx.CallAfter(self._doNavigateTab, "社群")

	def _onOfficial(self, evt):
		wx.CallAfter(self._doNavigateTab, "官方帳號")

	def _doNavigateTab(self, tabName):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab(tabName):
					ui.message(tabName)
				else:
					ui.message(_("無法切換到{tabName}").format(tabName=tabName))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	def _onVoiceCall(self, evt):
		# Defer execution so the NVDA menu closes first and LINE regains focus
		wx.CallAfter(self._doVoiceCall)

	def _doVoiceCall(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._makeCallByType("voice"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeCall error: {e}", exc_info=True)
			ui.message(_("通話功能錯誤: {error}").format(error=e))

	def _onVideoCall(self, evt):
		wx.CallAfter(self._doVideoCall)

	def _doVideoCall(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._makeCallByType("video"):
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE makeVideoCall error: {e}", exc_info=True)
			ui.message(_("視訊通話功能錯誤: {error}").format(error=e))

	def _onMessageReader(self, evt):
		wx.CallAfter(self._doMessageReader)

	def _doMessageReader(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_openMessageReader(None)
		except Exception as e:
			log.warning(f"LINE openMessageReader error: {e}", exc_info=True)
			ui.message(_("訊息閱讀器功能錯誤: {error}").format(error=e))

	def _onMoreOptions(self, evt):
		wx.CallAfter(self._doMoreOptions)

	def _doMoreOptions(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if not lineApp._clickMoreOptionsButton():
				ui.message(_("找不到 LINE 視窗，請先開啟聊天室"))
		except Exception as e:
			log.warning(f"LINE clickMoreOptions error: {e}", exc_info=True)
			ui.message(_("更多選項功能錯誤: {error}").format(error=e))

	def _onReadChatName(self, evt):
		wx.CallAfter(self._doReadChatName)

	def _doReadChatName(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp._readChatRoomName()
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	def _onDescribeImage(self, evt):
		wx.CallAfter(self._doDescribeImage)

	def _doDescribeImage(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_describeImage(None)
		except Exception as e:
			log.warning(f"LINE describeImage error: {e}", exc_info=True)
			ui.message(_("圖片描述功能錯誤: {error}").format(error=e))

	def _onAnswerCall(self, evt):
		wx.CallAfter(self._doAnswerCall)

	def _doAnswerCall(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._answerIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(_("接聽功能錯誤: {error}").format(error=e))

	def _onRejectCall(self, evt):
		wx.CallAfter(self._doRejectCall)

	def _doRejectCall(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._rejectIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(_("拒絕功能錯誤: {error}").format(error=e))

	def _onCheckCaller(self, evt):
		wx.CallAfter(self._doCheckCaller)

	def _doCheckCaller(self):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._getCallerInfo(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(_("來電查看功能錯誤: {error}").format(error=e))

	def _onFocusCallWindow(self, evt):
		wx.CallAfter(self._doFocusCallWindow)

	def _doFocusCallWindow(self):
		import ui
		import core
		import speech

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			import ctypes
			import ctypes.wintypes

			hwnd = lineApp._findIncomingCallWindow()
			if not hwnd:
				ui.message(_("未偵測到通話視窗"))
				return
			try:
				ctypes.windll.user32.SetForegroundWindow(hwnd)
			except Exception:
				pass

			def _announceCallWindow():
				try:
					ocrText = lineApp._ocrWindowArea(hwnd, sync=True, timeout=3.0)
					if ocrText:
						speech.cancelSpeech()
						ui.message(ocrText)
						log.info(f"LINE: call window OCR: {ocrText!r}")
					else:
						ui.message(_("通話視窗（無法辨識內容）"))
				except Exception as e:
					log.warning(f"LINE: call window OCR error: {e}", exc_info=True)
					ui.message(_("通話視窗"))

			core.callLater(300, _announceCallWindow)
		except Exception as e:
			log.warning(f"LINE focusCallWindow error: {e}", exc_info=True)
			ui.message(_("跳到通話視窗功能錯誤: {error}").format(error=e))

	def terminate(self, *args, **kwargs):
		self._removeToolsMenu()
		self._unregisterSettingsPanel()
		super().terminate(*args, **kwargs)
		for exe in self._LINE_EXECUTABLES:
			try:
				appModuleHandler.unregisterExecutable(exe)
			except Exception:
				pass

	@script(
		# Translators: Description of a debug script to report focused object info
		description=_("Debug: Report focused object's appModule and executable"),
		gesture="kb:NVDA+shift+j",
	)
	def script_reportFocusInfo(self, gesture):
		import api
		import ui

		obj = api.getFocusObject()
		app = obj.appModule
		appName = app.appName if app else "None"
		processID = app.processID if app else "None"
		moduleName = app.__class__.__module__ if app else "None"
		className = app.__class__.__name__ if app else "None"

		msg = f"App: {appName}, PID: {processID}, Module: {moduleName}.{className}"
		log.info(f"LINE Debug Focus Info: {msg}")
		ui.message(msg)

	# ── Incoming call global shortcuts ─────────────────────────────

	@script(
		# Translators: Description of a script to answer an incoming LINE call
		description=_("LINE: 接聽來電"),
		gesture="kb:NVDA+windows+a",
		category="LINE Desktop",
	)
	def script_answerCall(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._answerIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE answerCall error: {e}", exc_info=True)
			ui.message(_("接聽功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to reject an incoming LINE call
		description=_("LINE: 拒絕來電"),
		gesture="kb:NVDA+windows+d",
		category="LINE Desktop",
	)
	def script_rejectCall(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._rejectIncomingCall(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE rejectCall error: {e}", exc_info=True)
			ui.message(_("拒絕功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to check who is calling
		description=_("LINE: 查看來電者"),
		gesture="kb:NVDA+windows+s",
		category="LINE Desktop",
	)
	def script_checkCaller(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			hwnd = lineApp._findIncomingCallWindow()
			if hwnd:
				lineApp._getCallerInfo(hwnd)
			else:
				ui.message(_("未偵測到來電"))
		except Exception as e:
			log.warning(f"LINE checkCaller error: {e}", exc_info=True)
			ui.message(_("來電查看功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to focus the call window
		description=_("LINE: 跳到通話視窗"),
		gesture="kb:NVDA+windows+f",
		category="LINE Desktop",
	)
	def script_focusCallWindow(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_focusCallWindow(gesture)
		except Exception as e:
			log.warning(f"LINE focusCallWindow error: {e}", exc_info=True)
			ui.message(_("跳到通話視窗功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to read the current chat room name
		description=_("LINE: 讀出目前聊天室名稱"),
		gesture="kb:NVDA+windows+t",
		category="LINE Desktop",
	)
	def script_readChatRoomName(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_readChatRoomName(gesture)
		except Exception as e:
			log.warning(f"LINE readChatRoomName error: {e}", exc_info=True)
			ui.message(_("讀取聊天室名稱錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to open the message reader
		description=_("LINE: 開啟訊息閱讀器"),
		gesture="kb:NVDA+windows+j",
		category="LINE Desktop",
	)
	def script_openMessageReader(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_openMessageReader(gesture)
		except Exception as e:
			log.warning(f"LINE openMessageReader error: {e}", exc_info=True)
			ui.message(_("訊息閱讀器功能錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to click the more options button
		description=_("LINE: 點擊更多選項按鈕"),
		gesture="kb:NVDA+windows+o",
		category="LINE Desktop",
	)
	def script_clickMoreOptions(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			lineApp.script_clickMoreOptions(gesture)
		except Exception as e:
			log.warning(f"LINE clickMoreOptions error: {e}", exc_info=True)
			ui.message(_("更多選項功能錯誤: {error}").format(error=e))

	# ── Chat room tab navigation shortcuts ─────────────────────────

	@script(
		# Translators: Description of a script to navigate to all chats tab
		description=_("LINE: 跳到全部聊天室"),
		gesture="kb:NVDA+windows+1",
		category="LINE Desktop",
	)
	def script_navigateAllChats(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab("全部"):
					ui.message(_("全部"))
				else:
					# Translators: Shown when unable to switch to a chat tab
					ui.message(_("無法切換到全部"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to friends tab
		description=_("LINE: 跳到好友"),
		gesture="kb:NVDA+windows+2",
		category="LINE Desktop",
	)
	def script_navigateFriends(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab("好友"):
					ui.message(_("好友"))
				else:
					ui.message(_("無法切換到好友"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to groups tab
		description=_("LINE: 跳到群組"),
		gesture="kb:NVDA+windows+3",
		category="LINE Desktop",
	)
	def script_navigateGroups(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab("群組"):
					ui.message(_("群組"))
				else:
					ui.message(_("無法切換到群組"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to communities tab
		description=_("LINE: 跳到社群"),
		gesture="kb:NVDA+windows+4",
		category="LINE Desktop",
	)
	def script_navigateCommunities(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab("社群"):
					ui.message(_("社群"))
				else:
					ui.message(_("無法切換到社群"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))

	@script(
		# Translators: Description of a script to navigate to official accounts tab
		description=_("LINE: 跳到官方帳號"),
		gesture="kb:NVDA+windows+5",
		category="LINE Desktop",
	)
	def script_navigateOfficial(self, gesture):
		import ui

		lineApp = _getLineAppModule()
		if not lineApp:
			ui.message(_("LINE 未執行"))
			return
		try:
			if hasattr(lineApp, "_navigateToChatTab"):
				if lineApp._navigateToChatTab("官方帳號"):
					ui.message(_("官方帳號"))
				else:
					ui.message(_("無法切換到官方帳號"))
			else:
				ui.message(_("此功能需要更新 LINE 模組"))
		except Exception as e:
			log.warning(f"LINE navigateTab error: {e}", exc_info=True)
			ui.message(_("切換分頁錯誤: {error}").format(error=e))
