from __future__ import annotations

from pathlib import Path

import buildVars
from site_scons.site_tools.NVDATool.docs import md2html


def test_documentation_build_renders_markdown_tables(tmp_path: Path):
	source = tmp_path / "readme.md"
	destination = tmp_path / "readme.html"
	source.write_text(
		"| Shortcut | Action |\n"
		"| --- | --- |\n"
		"| `NVDA+Windows+J` | Open the message reader |\n",
		encoding="utf-8",
	)

	md2html(
		source,
		destination,
		moFile=None,
		mdExtensions=buildVars.markdownExtensions,
		addon_info=buildVars.addon_info,
	)

	html = destination.read_text(encoding="utf-8")
	assert "<table>" in html
	assert "<thead>" in html
	assert "<td><code>NVDA+Windows+J</code></td>" in html
