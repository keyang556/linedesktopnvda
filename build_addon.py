"""Build script to create .nvda-addon package without requiring scons.

This creates a minimal .nvda-addon package (which is a .zip file)
containing the required manifest.ini and addon files.
"""

import zipfile
import os
import subprocess
import markdown

ADDON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "addon")
OUTPUT_NAME = "lineDesktop-1.2.4-beta2.nvda-addon"
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTPUT_NAME)

# Manifest content matching the format of working NVDA add-ons
MANIFEST_CONTENT = """\
name = lineDesktop
summary = "LINE Desktop Accessibility"
description = \"\"\"Enhances NVDA accessibility support for the LINE desktop application on Windows.
Provides improved navigation for chat lists, messages, contacts, and message input.\"\"\"
author = "張可揚 <lindsay714322@gmail.com>; 洪鳳恩 <kittyhong0208@gmail.com>; 蔡頭<tommytsaitou>"
url = https://keyang556.github.io/linedesktopnvda/
version = 1.2.4-beta2
changelog = \"\"\"Initial release with LINE desktop accessibility support.\"\"\"
docFileName = readme.html
minimumNVDAVersion = 2019.3
lastTestedNVDAVersion = 2026.1
updateChannel = None
"""


def compile_po_files():
	"""Compile all .po files to .mo files using msgfmt."""
	locale_dir = os.path.join(ADDON_DIR, "locale")
	if not os.path.isdir(locale_dir):
		return
	for lang in os.listdir(locale_dir):
		lc_dir = os.path.join(locale_dir, lang, "LC_MESSAGES")
		if not os.path.isdir(lc_dir):
			continue
		for po_file in os.listdir(lc_dir):
			if not po_file.endswith(".po"):
				continue
			po_path = os.path.join(lc_dir, po_file)
			mo_path = po_path[:-3] + ".mo"
			try:
				# Use Python's msgfmt tool
				from tools.i18n import msgfmt

				msgfmt.make(po_path, mo_path)
				print(f"  Compiled {lang}/{po_file} -> .mo")
			except ImportError:
				# Fallback: try system msgfmt command
				try:
					subprocess.run(
						["msgfmt", "-o", mo_path, po_path],
						check=True,
						capture_output=True,
					)
					print(f"  Compiled {lang}/{po_file} -> .mo (system msgfmt)")
				except (FileNotFoundError, subprocess.CalledProcessError):
					# Fallback: use inline minimal msgfmt implementation
					_compile_po_to_mo(po_path, mo_path)
					print(f"  Compiled {lang}/{po_file} -> .mo (inline)")


def _compile_po_to_mo(po_path, mo_path):
	"""Minimal .po to .mo compiler for when system tools are unavailable."""
	import struct

	messages = {}
	msgid = None
	msgstr = None
	in_msgid = False
	in_msgstr = False

	with open(po_path, "r", encoding="utf-8") as f:
		for line in f:
			line = line.strip()
			if line.startswith("msgid "):
				# Save previous entry (including empty msgid for metadata)
				if msgid is not None:
					messages[msgid] = msgstr or ""
				in_msgid = True
				in_msgstr = False
				msgid = line[6:].strip('"')
			elif line.startswith("msgstr "):
				in_msgid = False
				in_msgstr = True
				msgstr = line[7:].strip('"')
			elif line.startswith('"') and line.endswith('"'):
				s = line[1:-1]
				if in_msgid:
					msgid = (msgid or "") + s
				elif in_msgstr:
					msgstr = (msgstr or "") + s
			elif not line or line.startswith("#"):
				if msgid is not None:
					messages[msgid] = msgstr or ""
				msgid = None
				msgstr = None
				in_msgid = False
				in_msgstr = False

	# Don't forget the last entry
	if msgid is not None:
		messages[msgid] = msgstr or ""

	# Process escape sequences
	def unescape(s):
		return s.replace("\\n", "\n").replace("\\t", "\t").replace('\\"', '"').replace("\\\\", "\\")

	# Build .mo file
	keys = sorted(messages.keys())
	offsets = []
	ids = b""
	strs = b""
	for key in keys:
		id_bytes = unescape(key).encode("utf-8")
		str_bytes = unescape(messages[key]).encode("utf-8")
		offsets.append((len(ids), len(id_bytes), len(strs), len(str_bytes)))
		ids += id_bytes + b"\x00"
		strs += str_bytes + b"\x00"

	n = len(keys)
	# Header: magic, revision, nstrings, offset_orig, offset_trans, size_hash, offset_hash
	keystart = 7 * 4 + n * 8 + n * 8
	valuestart = keystart + len(ids)

	koffsets = []
	voffsets = []
	for o in offsets:
		koffsets.append((o[1], o[0] + keystart))
		voffsets.append((o[3], o[2] + valuestart))

	output = struct.pack(
		"Iiiiiii",
		0x950412DE,  # magic
		0,  # revision
		n,  # nstrings
		7 * 4,  # offset orig table
		7 * 4 + n * 8,  # offset trans table
		0,  # size of hash
		0,  # offset of hash
	)
	for length, offset in koffsets:
		output += struct.pack("ii", length, offset)
	for length, offset in voffsets:
		output += struct.pack("ii", length, offset)
	output += ids
	output += strs

	with open(mo_path, "wb") as f:
		f.write(output)


def convert_md_to_html(md_path):
	"""Convert a markdown file to HTML string."""
	try:
		with open(md_path, "r", encoding="utf-8") as f:
			md_content = f.read()
		html = markdown.markdown(md_content, extensions=["tables"])
		return f"<!DOCTYPE html>\n<html><head><meta charset='utf-8'></head><body>\n{html}\n</body></html>"
	except ImportError:
		# If markdown module is not available, create minimal HTML
		with open(md_path, "r", encoding="utf-8") as f:
			content = f.read()
		return f"<!DOCTYPE html>\n<html><head><meta charset='utf-8'></head><body><pre>\n{content}\n</pre></body></html>"
	except Exception:
		return None


def build():
	print(f"Building {OUTPUT_NAME}...")

	# Compile .po files to .mo files
	print("Compiling translations...")
	compile_po_files()

	with zipfile.ZipFile(OUTPUT_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
		# Write manifest.ini at the root of the archive
		zf.writestr("manifest.ini", MANIFEST_CONTENT)
		print("  Added manifest.ini")

		# Walk the addon directory and add all files
		for root, dirs, files in os.walk(ADDON_DIR):
			# Skip __pycache__ directories
			dirs[:] = [d for d in dirs if d != "__pycache__"]
			for filename in files:
				filepath = os.path.join(root, filename)
				# Archive path relative to addon directory (so appModules/line.py, etc.)
				arcname = os.path.relpath(filepath, ADDON_DIR)
				zf.write(filepath, arcname)
				print(f"  Added {arcname}")

				# If this is a .md file in doc/, also generate .html version
				if filename.endswith(".md") and "doc" in root:
					html_content = convert_md_to_html(filepath)
					if html_content:
						html_arcname = arcname.replace(".md", ".html")
						zf.writestr(html_arcname, html_content)
						print(f"  Added {html_arcname} (generated from {filename})")

	print(f"\nBuild complete: {OUTPUT_PATH}")
	print(f"File size: {os.path.getsize(OUTPUT_PATH)} bytes")
	print("\nTo install: double-click the .nvda-addon file")


if __name__ == "__main__":
	build()
