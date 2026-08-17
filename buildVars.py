# Build customizations
# Change this file instead of sconstruct or manifest files, whenever possible.

from site_scons.site_tools.NVDATool.typings import (
	AddonInfo,
	BrailleTables,
	SpeechDictionaries,
	SymbolDictionaries,
)

# Since some strings in `addon_info` are translatable,
# we need to include them in the .po files.
# Gettext recognizes only strings given as parameters to the `_` function.
# To avoid initializing translations in this module we simply import a "fake" `_` function
# which returns whatever is given to it as an argument.
from site_scons.site_tools.NVDATool.utils import _


# Add-on information variables
addon_info = AddonInfo(
	# add-on Name/identifier, internal for NVDA
	addon_name="lineDesktop",
	# Add-on summary/title, usually the user visible name of the add-on
	# Translators: Summary/title for this add-on
	# to be shown on installation and add-on information found in add-on store
	addon_summary=_("LINE Desktop Accessibility"),
	# Add-on description
	# Translators: Long description to be shown for this add-on on add-on information from add-on store
	addon_description=_("""Enhances NVDA accessibility support for LINE Desktop on Windows.
Provides improved navigation for chat lists, messages, contacts, and message input.
Supports calls, incoming call handling, OCR-assisted reading, message export reading, and AI image description with follow-up questions."""),
	# version
	addon_version="1.3.0-beta4",
	# Brief changelog for this version
	# Translators: what's new content for the add-on version to be shown in the add-on store
	addon_changelog=_("""Message recall and the "Convert to text" photo notice now ask in a standard NVDA dialog instead of temporarily binding the Y/N/P and A/D keys.
Those letters are no longer swallowed inside LINE, and the confirmation no longer times out after ten seconds; Y/N/P and A/D still work as button shortcuts in the dialog.
All dialogs now use NVDA's message dialog API, replacing the deprecated gui.messageBox.
NVDA 2025.1 or later is now required."""),
	# Author(s)
	addon_author="張可揚 <lindsay714322@gmail.com>; 洪鳳恩 <kittyhong0208@gmail.com>; 蔡頭 <tommytsaitou>",
	# URL for the add-on documentation support
	addon_url="https://keyang556.github.io/linedesktopnvda/",
	# URL for the add-on repository where the source code can be found
	addon_sourceURL="https://github.com/keyang556/linedesktopnvda",
	# Documentation file name
	addon_docFileName="readme.html",
	# Minimum NVDA version supported (e.g. "2019.3.0", minor version is optional)
	# 2025.1 is the floor: every dialog the add-on shows is built on the
	# gui.message message dialog API (added in 2025.1), on top of the
	# speakOnDemand script parameter (2024.1) and the controlTypes.Role/State
	# enums (2021.2) used throughout.
	addon_minimumNVDAVersion="2025.1",
	# Last NVDA version supported/tested (e.g. "2024.4.0", ideally more recent than minimum version)
	addon_lastTestedNVDAVersion="2026.2",
	# Add-on update channel (default is None, denoting stable releases,
	# and for development releases, use "dev".)
	# Do not change unless you know what you are doing!
	addon_updateChannel=None,
	# Add-on license such as GPL 2
	addon_license="GPL 2",
	# URL for the license document the ad-on is licensed under
	addon_licenseURL="https://www.gnu.org/licenses/old-licenses/gpl-2.0.html",
)

# Define the python files that are the sources of your add-on.
pythonSources: list[str] = ["addon/appModules/*.py", "addon/globalPlugins/*.py"]

# Files that contain strings for translation. Usually your python sources
i18nSources: list[str] = pythonSources + ["buildVars.py"]

# Files that will be ignored when building the nvda-addon file
# *.pyc excludes __pycache__ contents left behind by running tests or
# py_compile against the addon sources (e.g. in CI, where the test step runs
# before the build step and imports addon.appModules.* as real packages).
excludedFiles: list[str] = ["*.pyc"]

# Base language for the NVDA add-on
baseLanguage: str = "en"

# Markdown extensions for add-on documentation
# The README files use pipe tables for shortcut references.
markdownExtensions: list[str] = ["markdown.extensions.tables"]

# Custom braille translation tables
brailleTables: BrailleTables = {}

# Custom speech symbol dictionaries
symbolDictionaries: SymbolDictionaries = {}

# Custom speech pronunciation dictionaries
speechDictionaries: SpeechDictionaries = {}
