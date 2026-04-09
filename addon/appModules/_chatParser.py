import re

_DATE_RE = re.compile(r'^\d{4}\.\d{2}\.\d{2} .+$')
_MSG_RE = re.compile(r'^(\d{2}:\d{2}) (\S+?) (.+)$')
_RECALL_RE = re.compile(r'^(\d{2}:\d{2}) (\S+?) 已收回訊息$')


def parseChatFile(filePath):
	"""Parse a LINE chat export text file.

	Returns a list of dicts with keys: name, content, time.
	Display format: name content time
	Continuation lines (Shift+Enter multi-line messages) are appended to the
	previous message's content.
	"""
	messages = []
	with open(filePath, 'r', encoding='utf-8') as f:
		for line in f:
			line = line.rstrip('\r\n')
			if not line:
				continue
			if _DATE_RE.match(line):
				continue
			m = _RECALL_RE.match(line)
			if m:
				messages.append({
					'time': m.group(1),
					'name': m.group(2),
					'content': '已收回訊息',
				})
				continue
			m = _MSG_RE.match(line)
			if m:
				messages.append({
					'time': m.group(1),
					'name': m.group(2),
					'content': m.group(3),
				})
				continue
			# Continuation line (Shift+Enter multi-line message)
			if messages:
				messages[-1]['content'] += '\n' + line
	return messages
