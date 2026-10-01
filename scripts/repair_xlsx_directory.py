"""Recover an XLSX whose ZIP directory is truncated, without changing its parts.

Only ordinary ZIP files with complete local headers/payloads are supported.
CRC, uncompressed size, XML syntax and internal relationships are verified.
An incomplete payload, data descriptor or ZIP64 file is rejected. This repairs
the container only; it does not validate or recalculate inventory economics.
The source is never modified and an existing output is never overwritten.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import struct
import zipfile
import zlib
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET


def recover(raw: bytes) -> tuple[bytes, dict]:
    if zipfile.is_zipfile(io.BytesIO(raw)):
        raise ValueError('The input has a readable ZIP directory; this repair is not applicable.')
    offset = 0
    parts = {}
    central = []
    records = []
    while raw[offset:offset + 4] == b'PK\x03\x04':
        if len(raw) - offset < 30:
            raise ValueError('Truncated local header.')
        _, version, flags, method, tm, dt, crc, compressed, size, nl, el = struct.unpack_from(
            '<IHHHHHIIIHH', raw, offset)
        if flags & 9 or method not in (0, 8) or max(size, compressed) == 0xffffffff:
            raise ValueError('Encrypted, data-descriptor, ZIP64 or unsupported compression input.')
        name_raw = raw[offset + 30:offset + 30 + nl]
        name = name_raw.decode('utf-8' if flags & 0x800 else 'cp437')
        if name in parts or name.startswith('/') or '..' in name.split('/'):
            raise ValueError(f'Duplicate or unsafe part name: {name}')
        start = offset + 30 + nl + el
        end = start + compressed
        if not name or end > len(raw):
            raise ValueError(f'Truncated part: {name}')
        payload = raw[start:end]
        if method == 8:
            decoder = zlib.decompressobj(-15)
            decoded = decoder.decompress(payload) + decoder.flush()
            if not decoder.eof or decoder.unused_data:
                raise ValueError(f'Incomplete or trailing compressed data: {name}')
        else:
            decoded = payload
        if len(decoded) != size or zlib.crc32(decoded) != crc:
            raise ValueError(f'Part checksum/size mismatch: {name}')
        if name.endswith(('.xml', '.rels')):
            ET.fromstring(decoded)
        parts[name] = decoded
        extra = raw[offset + 30 + nl:start]
        central.append(struct.pack('<IHHHHHHIIIHHHHHII',
            0x02014b50, 20, version, flags, method, tm, dt, crc, compressed,
            size, nl, el, 0, 0, 0, 0, offset) + name_raw + extra)
        records.append({'name': name, 'sha256': hashlib.sha256(decoded).hexdigest(),
                        'crc32': f'{crc:08x}', 'size': size})
        offset = end
    if not parts or raw[offset:offset + 4] != b'PK\x01\x02':
        raise ValueError('Complete payload sequence is not followed by a central-directory header.')
    for required in ('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
                     'xl/_rels/workbook.xml.rels'):
        if required not in parts:
            raise ValueError(f'Missing required workbook part: {required}')
    for name, decoded in parts.items():
        if not name.endswith('.rels'):
            continue
        # A .rels part lives in a _rels subdirectory beside its owning part.
        owner_dir = posixpath.dirname(posixpath.dirname(name))
        for rel in ET.fromstring(decoded):
            if rel.attrib.get('TargetMode') == 'External':
                continue
            target = unquote(rel.attrib['Target']).split('#')[0]
            resolved = (target.lstrip('/') if target.startswith('/') else
                        posixpath.normpath(posixpath.join(owner_dir, target)))
            if resolved not in parts:
                raise ValueError(f'Missing relationship target: {name} -> {resolved}')
    directory = b''.join(central)
    if len(parts) >= 65535 or offset + len(directory) >= 0xffffffff:
        raise ValueError('ZIP64 output is unsupported.')
    repaired = raw[:offset] + directory + struct.pack(
        '<IHHHHIIH', 0x06054b50, 0, 0, len(parts), len(parts), len(directory), offset, 0)
    with zipfile.ZipFile(io.BytesIO(repaired)) as result:
        if result.testzip() is not None or set(result.namelist()) != set(parts):
            raise ValueError('Recovered archive verification failed.')
        for name, decoded in parts.items():
            if result.read(name) != decoded:
                raise ValueError(f'Payload changed during recovery: {name}')
    return repaired, {'method': 'Rebuild central directory; original local headers and compressed payloads unchanged.',
                      'source_sha256': hashlib.sha256(raw).hexdigest(),
                      'recovered_sha256': hashlib.sha256(repaired).hexdigest(),
                      'source_size': len(raw), 'recovered_size': len(repaired),
                      'preserved_prefix_bytes': offset, 'verified_parts': records}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('source', type=Path)
    ap.add_argument('output', type=Path)
    args = ap.parse_args()
    if args.source.resolve() == args.output.resolve():
        ap.error('Use a separate output path.')
    report_path = args.output.with_suffix('.recovery.json')
    if args.output.exists() or report_path.exists():
        ap.error('Output or recovery report already exists; choose a new destination.')
    repaired, report = recover(args.source.read_bytes())
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with args.output.open('xb') as f:
        f.write(repaired)
    report_path.write_text(json.dumps(report, indent=2) + '\n')
    print(f'Recovered {len(report["verified_parts"])} unchanged parts to {args.output}')


if __name__ == '__main__':
    main()
