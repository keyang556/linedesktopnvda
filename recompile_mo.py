"""Recompile all .po files to .mo files using Python's msgfmt module."""
import os
import sys
import subprocess

locale_dir = os.path.join("addon", "locale")

for lang in os.listdir(locale_dir):
    po_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.po")
    mo_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.mo")
    if os.path.exists(po_path):
        print(f"Compiling {po_path} -> {mo_path}")
        # Use Python's built-in msgfmt tool
        result = subprocess.run(
            [sys.executable, "-m", "tools.i18n.msgfmt", "-o", mo_path, po_path],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            # Try the standard library msgfmt
            result2 = subprocess.run(
                [sys.executable, os.path.join(os.path.dirname(os.__file__), "Tools", "i18n", "msgfmt.py"), "-o", mo_path, po_path],
                capture_output=True, text=True
            )
            if result2.returncode != 0:
                # Last resort: use our own implementation
                print(f"  Standard tools failed, using custom compiler")
                from compile_mo_custom import compile_po_to_mo
                compile_po_to_mo(po_path, mo_path)
            else:
                print(f"  Done (standard msgfmt)")
        else:
            print(f"  Done (tools.i18n.msgfmt)")

print("All done!")
