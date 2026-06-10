import itertools

from logHandler import log


# Keeps recognizer/pixel buffers referenced until each async recognition
# finishes; without this they can be garbage-collected mid-recognition and the
# callback never fires (or crashes). Keyed by a monotonic token.
_pendingOcr = {}
_ocrTokenCounter = itertools.count()

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


def ocrGetText(left, top, width, height, onResult):
	"""
	Performs OCR on the specified screen region and calls the onResult callback with the OCR result.
	onResult should be a function that takes a single argument, which will be the OCR result object.

	Note: the underlying UWP OCR invokes ``onResult`` on a background thread.
	Callers that touch NVDA speech/braille or UI must marshal back to the main
	thread themselves (e.g. via ``core.callLater``).
	"""
	try:
		import screenBitmap
		from contentRecog import uwpOcr

		langs = uwpOcr.getLanguages()
		if not langs:
			return

		# Pick language: follow LINE's UI language when known, else prefer
		# Traditional Chinese (the historical default).
		ocrLang = pickOcrLanguage(langs)
		if not ocrLang:
			return

		recognizer = uwpOcr.UwpOcr(language=ocrLang)
		resizeFactor = recognizer.getResizeFactor(width, height)
		# Use higher resize factor for small elements to improve OCR accuracy
		minFactor = 2
		if width < 100 or height < 100:
			minFactor = max(3, int(200 / max(min(width, height), 1)))
		if resizeFactor < minFactor:
			resizeFactor = minFactor

		class _ImgInfo:
			def __init__(self, w, h, factor, sLeft, sTop):
				self.recogWidth = w * factor
				self.recogHeight = h * factor
				self.resizeFactor = factor
				self._screenLeft = sLeft
				self._screenTop = sTop

			def convertXToScreen(self, x):
				return self._screenLeft + int(x / self.resizeFactor)

			def convertYToScreen(self, y):
				return self._screenTop + int(y / self.resizeFactor)

			def convertWidthToScreen(self, w):
				return int(w / self.resizeFactor)

			def convertHeightToScreen(self, h):
				return int(h / self.resizeFactor)

		imgInfo = _ImgInfo(width, height, resizeFactor, left, top)

		if resizeFactor > 1:
			sb = screenBitmap.ScreenBitmap(
				width * resizeFactor,
				height * resizeFactor,
			)
			ocrPixels = sb.captureImage(left, top, width, height)
		else:
			sb = screenBitmap.ScreenBitmap(width, height)
			ocrPixels = sb.captureImage(left, top, width, height)

		# Keep the recognizer and pixel buffer alive until recognition
		# completes, then release them in the wrapper.
		token = next(_ocrTokenCounter)
		_pendingOcr[token] = (recognizer, ocrPixels, imgInfo, sb)

		def _wrappedOnResult(result):
			_pendingOcr.pop(token, None)
			try:
				onResult(result)
			except Exception:
				log.debug("OCR onResult callback failed", exc_info=True)

		recognizer.recognize(ocrPixels, imgInfo, _wrappedOnResult)
	except Exception:
		log.debug("OCR fallback failed", exc_info=True)


def message(text):
	"""
	Same as ui.message, but without the time limit for braille display.
	"""
	import speech
	import braille

	speech.speakMessage(text)

	handler = braille.handler
	if handler is None:
		# Braille may be uninitialized (secure screens, minimal setups);
		# speech already happened, so just skip braille output.
		return
	try:
		if handler.buffer is handler.messageBuffer:
			handler.buffer.clear()
		else:
			handler.buffer = handler.messageBuffer

		region = braille.TextRegion(text)
		region.update()
		handler.buffer.regions.append(region)
		handler.buffer.update()
		handler.update()
	except Exception:
		log.debug("Braille output failed for text, skipping braille output", exc_info=True)
