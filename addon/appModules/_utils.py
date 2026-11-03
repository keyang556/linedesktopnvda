# Optional hint (LINE's UI language code, e.g. 'zh-TW', 'ja', 'th') so OCR can
# prefer the language LINE is actually displaying. Set by the app module at
# startup; None keeps the historical Traditional-Chinese-first default.
_preferredOcrLanguage = None


def setPreferredOcrLanguage(lang):
	"""Record LINE's UI language so OCR can prefer it (see pickOcrLanguage)."""
	global _preferredOcrLanguage
	_preferredOcrLanguage = lang


def pickOcrLanguage(langs, preferLang=None):
	"""Choose the best available UWP OCR language tag.

	``langs`` is the list of installed OCR language tags (e.g. 'zh-Hant-TW',
	'ja', 'en-US'). ``preferLang`` is an optional hint from LINE's UI language
	(e.g. 'zh-TW', 'ja', 'th'); when omitted the module-level hint is used.

	Falls back to Traditional Chinese, then any Chinese, then the first
	installed language — the add-on's historical default — so behaviour is
	unchanged when no hint is given or LINE is Chinese.
	"""
	if not langs:
		return None
	hint = (preferLang if preferLang is not None else _preferredOcrLanguage) or ""
	hint = hint.lower()
	prefixes = []
	if hint.startswith("ja"):
		prefixes = ["ja"]
	elif hint.startswith("ko"):
		prefixes = ["ko"]
	elif hint.startswith("th"):
		prefixes = ["th"]
	elif hint.startswith(("zh-cn", "zh-hans", "zh_cn")):
		prefixes = ["zh-hans", "zh-cn", "zh"]
	elif hint.startswith("en"):
		prefixes = ["en"]

	lowerLangs = [(lang, lang.lower()) for lang in langs]
	for prefix in prefixes:
		for orig, low in lowerLangs:
			if low.startswith(prefix):
				return orig
	# Historical default: Traditional Chinese, then any Chinese.
	for cand in ("zh-hant-tw", "zh-tw", "zh-hant"):
		for orig, low in lowerLangs:
			if low == cand:
				return orig
	for orig, low in lowerLangs:
		if low.startswith("zh"):
			return orig
	return langs[0]
