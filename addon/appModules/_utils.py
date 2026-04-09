from logHandler import log

def ocrGetText(left, top, width, height, onResult):
	'''
	Performs OCR on the specified screen region and calls the onResult callback with the OCR result.
	onResult should be a function that takes a single argument, which will be the OCR result object.
	'''
	try:
		import screenBitmap
		from contentRecog import uwpOcr

		langs = uwpOcr.getLanguages()
		if not langs:
			return

		# Pick language: prefer Traditional Chinese
		ocrLang = None
		for candidate in ["zh-Hant-TW", "zh-TW", "zh-Hant"]:
			if candidate in langs:
				ocrLang = candidate
				break
		if not ocrLang:
			for lang in langs:
				if lang.startswith("zh"):
					ocrLang = lang
					break
		if not ocrLang:
			ocrLang = langs[0]

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
				height * resizeFactor
			)
			ocrPixels = sb.captureImage(left, top, width, height)
		else:
			sb = screenBitmap.ScreenBitmap(width, height)
			ocrPixels = sb.captureImage(left, top, width, height)

		recognizer.recognize(ocrPixels, imgInfo, onResult)
	except Exception:
		log.debug("OCR fallback failed", exc_info=True)
	

def message(text):
	'''
	Same as ui.message, but without the time limit for braille display.
	'''
	import speech
	import braille
	
	speech.speakMessage(text)
	
	handler = braille.handler
	assert handler
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
	except RuntimeError:
		log.debug("Braille translation failed for text, skipping braille output", exc_info=True)
