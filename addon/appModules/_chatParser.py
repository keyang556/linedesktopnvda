import re

_DATE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2} .+$")
_MSG_RE = re.compile(r"^(\d{2}:\d{2}) (\S+?) (.+)$")
_RECALL_RE = re.compile(r"^(\d{2}:\d{2}) (\S+?)已收回訊息$")


def parseChatFile(filePath):
	"""Parse a LINE chat export text file.

	Returns a list of dicts for message-reader navigation.
	Message entries use keys: type, name, content, time.
	Date separator entries use keys: type, content.
	Display format keeps date separators in their original positions.
	Continuation lines (Shift+Enter multi-line messages) are appended to the
	previous message's content.
	"""
	messages = []
	with open(filePath, "r", encoding="utf-8") as f:
		for line in f:
			line = line.rstrip("\r\n")
			if not line:
				continue
			if _DATE_RE.match(line):
				messages.append(
					{
						"type": "date",
						"content": line,
					},
				)
				continue
			m = _RECALL_RE.match(line)
			if m:
				messages.append(
					{
						"type": "message",
						"time": m.group(1),
						"name": m.group(2),
						"content": "已收回訊息",
					},
				)
				continue
			m = _MSG_RE.match(line)
			if m:
				messages.append(
					{
						"type": "message",
						"time": m.group(1),
						"name": m.group(2),
						"content": m.group(3),
					},
				)
				continue
			# Continuation line (Shift+Enter multi-line message)
			if messages and messages[-1].get("type") != "date":
				messages[-1]["content"] += "\n" + line
	return messages
