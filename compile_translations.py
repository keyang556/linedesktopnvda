#!/usr/bin/env python3
"""Simple script to compile .po files to .mo files."""
import os
import struct
import array

def compile_po_to_mo(po_file, mo_file):
    """Compile a .po file to .mo format."""
    # Read the .po file
    with open(po_file, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Parse messages
    messages = {}
    msgid = None
    msgstr = None
    in_msgid = False
    in_msgstr = False
    
    for line in lines:
        line = line.strip()
        
        if line.startswith('msgid '):
            if msgid is not None and msgstr is not None:
                messages[msgid] = msgstr
            msgid = line[6:].strip('"')
            msgstr = None
            in_msgid = True
            in_msgstr = False
        elif line.startswith('msgstr '):
            msgstr = line[7:].strip('"')
            in_msgid = False
            in_msgstr = True
        elif line.startswith('"') and (in_msgid or in_msgstr):
            text = line.strip('"')
            if in_msgid:
                msgid += text
            elif in_msgstr:
                msgstr += text
        elif line == '':
            if msgid is not None and msgstr is not None:
                messages[msgid] = msgstr
            msgid = None
            msgstr = None
            in_msgid = False
            in_msgstr = False
    
    # Add last message
    if msgid is not None and msgstr is not None:
        messages[msgid] = msgstr
    
    # Remove empty msgid (header)
    if '' in messages:
        del messages['']
    
    # Build .mo file
    keys = sorted(messages.keys())
    offsets = []
    ids = b''
    strs = b''
    
    for key in keys:
        offsets.append((len(ids), len(key), len(strs), len(messages[key])))
        ids += key.encode('utf-8') + b'\x00'
        strs += messages[key].encode('utf-8') + b'\x00'
    
    # Write .mo file
    with open(mo_file, 'wb') as f:
        # Magic number
        f.write(struct.pack('I', 0x950412de))
        # Version
        f.write(struct.pack('I', 0))
        # Number of entries
        f.write(struct.pack('I', len(keys)))
        # Offset of table with original strings
        f.write(struct.pack('I', 28))
        # Offset of table with translation strings
        f.write(struct.pack('I', 28 + len(keys) * 8))
        # Size of hashing table
        f.write(struct.pack('I', 0))
        # Offset of hashing table
        f.write(struct.pack('I', 0))
        
        # Write original strings table
        for o in offsets:
            f.write(struct.pack('II', o[1], 28 + len(keys) * 16 + o[0]))
        
        # Write translation strings table
        for o in offsets:
            f.write(struct.pack('II', o[3], 28 + len(keys) * 16 + len(ids) + o[2]))
        
        # Write strings
        f.write(ids)
        f.write(strs)

# Compile all translations
locale_dir = os.path.join("addon", "locale")
for lang in os.listdir(locale_dir):
    po_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.po")
    mo_path = os.path.join(locale_dir, lang, "LC_MESSAGES", "nvda.mo")
    if os.path.exists(po_path):
        print(f"Compiling {po_path} -> {mo_path}")
        try:
            compile_po_to_mo(po_path, mo_path)
            print(f"  Done!")
        except Exception as e:
            print(f"  Error: {e}")

print("All translations compiled!")
