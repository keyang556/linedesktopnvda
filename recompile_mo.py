"""Recompile all .po files to .mo files using Python's msgfmt module."""

import os
import sys
import subprocess
import ast
import struct


def _po_string(token):
	"""Return the decoded value of a quoted PO string token."""
	return ast.literal_eval(token)


def _compile_po_to_mo(po_path, mo_path):
	"""Compile a UTF-8 .po file to .mo without external msgfmt tools."""
	messages = {}
	msgid = None
	msgstr = None
	active = None

	with open(po_path, "r", encoding="utf-8") as f:
		for rawLine in f:
			line = rawLine.strip()
			if not line or line.startswith("#"):
				continue
			if line.startswith("msgid "):
				if msgid is not None:
					messages[msgid] = msgstr or ""
				msgid = _po_string(line[6:].strip())
				msgstr = None
				active = "msgid"
			elif line.startswith("msgstr "):
				msgstr = _po_string(line[7:].strip())
				active = "msgstr"
			elif line.startswith('"') and line.endswith('"'):
				value = _po_string(line)
				if active == "msgid":
					msgid = (msgid or "") + value
				elif active == "msgstr":
					msgstr = (msgstr or "") + value

	if msgid is not None:
		messages[msgid] = msgstr or ""

	keys = sorted(messages)
	ids = b""
	strs = b""
	offsets = []
	for key in keys:
		idBytes = key.encode("utf-8")
		strBytes = messages[key].encode("utf-8")
		offsets.append((len(idBytes), len(ids), len(strBytes), len(strs)))
		ids += idBytes + b"\0"
		strs += strBytes + b"\0"

	n = len(keys)
	keystart = 7 * 4 + n * 8 * 2
	valuestart = keystart + len(ids)
	output = struct.pack("Iiiiiii", 0x950412DE, 0, n, 7 * 4, 7 * 4 + n * 8, 0, 0)
	for idLen, idOffset, _strLen, _strOffset in offsets:
		output += struct.pack("ii", idLen, keystart + idOffset)
	for _idLen, _idOffset, strLen, strOffset in offsets:
		output += struct.pack("ii", strLen, valuestart + strOffset)
	output += ids
	output += strs

	with open(mo_path, "wb") as f:
		f.write(output)


locale_dir = os.path.join("addon", "locale")

for lang in os.listdir(locale_dir):
	po_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.po")
	mo_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.mo")
	if os.path.exists(po_path):
		print(f"Compiling {po_path} -> {mo_path}")
		# Use Python's built-in msgfmt tool
		result = subprocess.run(
			[sys.executable, "-m", "tools.i18n.msgfmt", "-o", mo_path, po_path],
			capture_output=True,
			text=True,
		)
		if result.returncode != 0:
			# Try the standard library msgfmt
			result2 = subprocess.run(
				[
					sys.executable,
					os.path.join(os.path.dirname(os.__file__), "Tools", "i18n", "msgfmt.py"),
					"-o",
					mo_path,
					po_path,
				],
				capture_output=True,
				text=True,
			)
			if result2.returncode != 0:
				# Last resort: use our own implementation
				print("  Standard tools failed, using custom compiler")
				_compile_po_to_mo(po_path, mo_path)
			else:
				print("  Done (standard msgfmt)")
		else:
			print("  Done (tools.i18n.msgfmt)")

print("All done!")
