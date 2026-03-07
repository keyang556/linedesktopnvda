"""Compile .po files to .mo files using a simple custom implementation."""
import os
import struct
import array

def compile_po_to_mo(po_path, mo_path):
    messages = {}
    msgid = None
    msgstr = None
    in_msgid = False
    in_msgstr = False

    with open(po_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('#') or not line:
                if msgid is not None and msgstr is not None:
                    messages[msgid] = msgstr
                msgid = None
                msgstr = None
                in_msgid = False
                in_msgstr = False
                continue
            if line.startswith('msgid '):
                if msgid is not None and msgstr is not None:
                    messages[msgid] = msgstr
                in_msgid = True
                in_msgstr = False
                msgid = line[6:].strip('"')
            elif line.startswith('msgstr '):
                in_msgid = False
                in_msgstr = True
                msgstr = line[7:].strip('"')
            elif line.startswith('"'):
                val = line.strip('"')
                if in_msgid:
                    msgid = (msgid or '') + val
                elif in_msgstr:
                    msgstr = (msgstr or '') + val
        # Don't forget last entry
        if msgid is not None and msgstr is not None:
            messages[msgid] = msgstr

    # Build .mo binary
    # Keep header (empty msgid)
    header = messages.pop('', '')

    keys = sorted(messages.keys())
    # Prepend header
    all_keys = [''] + keys
    all_vals = [header] + [messages[k] for k in keys]

    offsets = []
    ids = b''
    strs = b''
    for key, val in zip(all_keys, all_vals):
        encoded_key = key.encode('utf-8')
        encoded_val = val.encode('utf-8')
        offsets.append((len(encoded_key), len(ids), len(encoded_val), len(strs)))
        ids += encoded_key + b'\x00'
        strs += encoded_val + b'\x00'

    n = len(all_keys)
    keystart = 7 * 4 + n * 8 * 2
    valuestart = keystart + len(ids)
    koffsets = []
    voffsets = []
    for l1, o1, l2, o2 in offsets:
        koffsets += [l1, o1 + keystart]
        voffsets += [l2, o2 + valuestart]

    output = struct.pack('Iiiiiii',
        0x950412de,       # magic
        0,                # version
        n,                # nstrings
        7 * 4,            # offset of table with original strings
        7 * 4 + n * 8,    # offset of table with translation strings
        0, 0              # size / offset of hash table
    )
    output += array.array('i', koffsets).tobytes()
    output += array.array('i', voffsets).tobytes()
    output += ids
    output += strs

    with open(mo_path, 'wb') as f:
        f.write(output)

locale_dir = os.path.join(os.path.dirname(__file__), 'addon', 'locale')
for lang in os.listdir(locale_dir):
    po = os.path.join(locale_dir, lang, 'LC_MESSAGES', 'nvda.po')
    mo = os.path.join(locale_dir, lang, 'LC_MESSAGES', 'nvda.mo')
    if os.path.exists(po):
        compile_po_to_mo(po, mo)
        print(f'Compiled {lang}: {po} -> {mo}')
print('Done!')
