# scripts/hwp5/ole.py
# -*- coding: utf-8 -*-
"""
Lightweight OLE (CFB) Reader with full FAT / MiniFAT / MiniStream support.

- Supports:
  * Header parsing (sector sizes, cutoffs, sector ids)
  * DIFAT → FAT sector discovery
  * FAT chains (regular sectors)
  * Directory stream parsing (names/types/tree)
  * MiniFAT chains + MiniStream (64-byte mini sectors)
  * Stream read by full path name, e.g. "BodyText/Section0"
- No external dependencies; pure Python and safe for sandbox use.

Notes:
- This is a minimal, read-only implementation tailored for HWP 5.x reading.
- Directory tree is traversed to materialize full paths like "Storage/Child/Stream".
"""

import io
import os
import struct
from typing import Dict, List, Optional, Tuple

# Sector ID constants (MS-CFB / [MS-CFB] spec)
FREE_SECTOR     = 0xFFFFFFFF
END_OF_CHAIN    = 0xFFFFFFFE
FAT_SECTOR      = 0xFFFFFFFD
DIFAT_SECTOR    = 0xFFFFFFFC

STGTY_EMPTY  = 0x00
STGTY_STORAGE= 0x01
STGTY_STREAM = 0x02
STGTY_ROOT   = 0x05

class OleReader:
    def __init__(self, path: str):
        self.path = path
        self.fp = open(path, 'rb')

        # header fields
        self.sector_size = 512
        self.mini_sector_size = 64
        self.mini_stream_cutoff = 4096

        self.dir_first_sector = FREE_SECTOR
        self.num_fat_sectors = 0
        self.mini_fat_first_sector = FREE_SECTOR
        self.num_mini_fat_sectors = 0
        self.difat_first_sector = FREE_SECTOR
        self.num_difat_sectors = 0

        # tables
        self.difat: List[int] = []
        self.fat: List[int] = []
        self.minifat: List[int] = []

        # directory
        self.dir_entries: List[dict] = []
        self.dir_lookup: Dict[str, dict] = {}

        # mini stream (full blob, from root storage stream)
        self.mini_stream_data: bytes = b''

        # parse
        self._read_header()
        self._load_difat_and_fat()
        self._load_directory()
        self._load_minifat()
        self._load_ministream()
        # build full-path lookup
        self._build_paths()

    # ---------- context manager ----------
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        self.close()

    def close(self):
        try:
            self.fp.close()
        except Exception:
            pass

    # ---------- low-level helpers ----------
    def _read_exact_at(self, offset: int, size: int) -> bytes:
        self.fp.seek(offset)
        data = self.fp.read(size)
        if len(data) != size:
            raise IOError("Unexpected EOF while reading OLE file")
        return data

    def _read_sector(self, sid: int) -> bytes:
        # sector index starts after 512-byte header
        if sid < 0:
            raise ValueError("Negative sector index")
        offset = 512 + sid * self.sector_size
        return self._read_exact_at(offset, self.sector_size)

    def _unpack_u32_vec(self, data: bytes) -> List[int]:
        if len(data) % 4 != 0:
            # pad defensive
            data = data[:len(data) - (len(data) % 4)]
        return list(struct.unpack('<%dI' % (len(data)//4), data))

    # ---------- header ----------
    def _read_header(self):
        header = self._read_exact_at(0, 512)
        sig = header[:8]
        if sig != b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
            raise ValueError("Not an OLE Compound File (bad signature)")

        sector_shift = struct.unpack('<H', header[0x1E:0x20])[0]
        mini_sector_shift = struct.unpack('<H', header[0x20:0x22])[0]
        self.sector_size = 1 << sector_shift
        self.mini_sector_size = 1 << mini_sector_shift

        # 0x2C: number of FAT sectors
        self.num_fat_sectors = struct.unpack('<I', header[0x2C:0x30])[0]
        # 0x30: first directory sector
        self.dir_first_sector = struct.unpack('<I', header[0x30:0x34])[0]
        # 0x38: mini stream cutoff (default 4096)
        self.mini_stream_cutoff = struct.unpack('<I', header[0x38:0x3C])[0]
        # 0x3C: first MiniFAT sector
        self.mini_fat_first_sector = struct.unpack('<I', header[0x3C:0x40])[0]
        # 0x40: number of MiniFAT sectors
        self.num_mini_fat_sectors = struct.unpack('<I', header[0x40:0x44])[0]
        # 0x44: first DIFAT sector
        self.difat_first_sector = struct.unpack('<I', header[0x44:0x48])[0]
        # 0x48: number of DIFAT sectors
        self.num_difat_sectors = struct.unpack('<I', header[0x48:0x4C])[0]

        # 0x4C.. header DIFAT (109 entries)
        hdr_difat = self._unpack_u32_vec(header[0x4C:0x200])
        self.difat = [sid for sid in hdr_difat if sid != FREE_SECTOR]

    # ---------- DIFAT/FAT ----------
    def _load_difat_and_fat(self):
        # extend DIFAT via DIFAT sectors if present
        next_sid = self.difat_first_sector
        count = self.num_difat_sectors
        while count > 0 and next_sid not in (FREE_SECTOR, END_OF_CHAIN):
            sec = self._read_sector(next_sid)
            # sector contains 127 FAT sector IDs + next DIFAT sector at the end
            entries = self._unpack_u32_vec(sec[:self.sector_size - 4])
            self.difat.extend([sid for sid in entries if sid != FREE_SECTOR])
            next_sid = struct.unpack('<I', sec[-4:])[0]
            count -= 1

        # now read FAT sectors listed in DIFAT
        fat_entries: List[int] = []
        for fat_sid in self.difat:
            sec = self._read_sector(fat_sid)
            fat_entries.extend(self._unpack_u32_vec(sec))
        self.fat = fat_entries

    # ---------- Directory ----------
    def _load_directory(self):
        # directory stream is chained via FAT
        dir_raw = self._read_chain(self.dir_first_sector, use_mini=False)
        # entries are 128 bytes each
        if len(dir_raw) % 128 != 0:
            # truncate defensively
            dir_raw = dir_raw[:len(dir_raw) - (len(dir_raw) % 128)]
        entries = []
        for i in range(0, len(dir_raw), 128):
            rec = dir_raw[i:i+128]
            name_raw = rec[0:64]
            name_len = struct.unpack('<H', rec[64:66])[0]
            # name_len includes terminating null; ensure even
            name = name_raw[:max(0, min(64, name_len-2))].decode('utf-16le', errors='ignore')
            obj_type = rec[66]
            # 67: color (ignored)
            left = struct.unpack('<i', rec[68:72])[0]
            right = struct.unpack('<i', rec[72:76])[0]
            child = struct.unpack('<i', rec[76:80])[0]
            # 80..96 clsid (ignored)
            # 96..100 state bits (ignored)
            # 100..116 times (ignored)
            start_sector = struct.unpack('<I', rec[116:120])[0]
            size_lo = struct.unpack('<I', rec[120:124])[0]
            size_hi = struct.unpack('<I', rec[124:128])[0]
            stream_size = size_lo if size_hi == 0 else (size_hi << 32) | size_lo

            entries.append({
                'index': len(entries),
                'name': name,
                'type': obj_type,
                'left': left,
                'right': right,
                'child': child,
                'start_sector': start_sector,
                'stream_size': stream_size,
                'parent': None,  # fill later via tree walk
                'full_path': None,
            })
        self.dir_entries = entries

    def _build_paths(self):
        # The root entry is the first with type ROOT
        root_idx = None
        for e in self.dir_entries:
            if e['type'] == STGTY_ROOT:
                root_idx = e['index']
                break
        if root_idx is None:
            return

        def walk_btree(idx: int, parent_path: Optional[str]):
            if idx < 0 or idx >= len(self.dir_entries):
                return
            node = self.dir_entries[idx]
            # in-order: left, node, child, right
            if node['left'] >= 0:
                walk_btree(node['left'], parent_path)

            name = node['name']
            # build path (root has empty name, children of root start from their names)
            if node['type'] == STGTY_ROOT:
                full_path = ''
            else:
                full_path = name if not parent_path else (parent_path + '/' + name)

            node['full_path'] = full_path
            node['parent'] = parent_path

            # storage children (sub-tree)
            if node['child'] >= 0:
                walk_btree(node['child'], full_path)

            # stream entries: register into lookup
            if node['type'] == STGTY_STREAM and full_path:
                self.dir_lookup[full_path] = node

            if node['right'] >= 0:
                walk_btree(node['right'], parent_path)

        # Root's child is the storage tree
        root = self.dir_entries[root_idx]
        if root['child'] >= 0:
            walk_btree(root['child'], None)

    # ---------- MiniFAT / MiniStream ----------
    def _load_minifat(self):
        if self.num_mini_fat_sectors == 0 or self.mini_fat_first_sector in (FREE_SECTOR, END_OF_CHAIN):
            self.minifat = []
            return
        # mini FAT is a regular FAT-chained stream (32-bit entries)
        raw = self._read_chain(self.mini_fat_first_sector, use_mini=False)
        self.minifat = self._unpack_u32_vec(raw)

    def _load_ministream(self):
        # MiniStream data lives in root storage's stream (type ROOT)
        root_entry = None
        for e in self.dir_entries:
            if e['type'] == STGTY_ROOT:
                root_entry = e
                break
        if not root_entry:
            self.mini_stream_data = b''
            return
        if root_entry['start_sector'] in (FREE_SECTOR, END_OF_CHAIN) or root_entry['stream_size'] == 0:
            self.mini_stream_data = b''
            return
        # MiniStream itself is in regular FAT sectors
        self.mini_stream_data = self._read_chain(root_entry['start_sector'],
                                                 size=root_entry['stream_size'],
                                                 use_mini=False)

    # ---------- chain readers ----------
    def _read_chain(self, start_sector: int, size: Optional[int] = None, use_mini: bool = False) -> bytes:
        """
        Read a full sector chain (regular FAT or MiniFAT).
        If size is None, read until END_OF_CHAIN (then trim trailing FREE entries length).
        """
        if start_sector in (FREE_SECTOR, END_OF_CHAIN):
            return b''

        if use_mini:
            if not self.minifat or not self.mini_stream_data:
                return b''
            sec_size = self.mini_sector_size
            fat = self.minifat
            # MiniStream is a contiguous bytearray; mini-sector N starts at N*64 within that blob.
            def read_mini_sector(mini_sid: int) -> bytes:
                off = mini_sid * sec_size
                return self.mini_stream_data[off: off + sec_size]
            read_one = read_mini_sector
        else:
            sec_size = self.sector_size
            fat = self.fat
            read_one = self._read_sector

        out = bytearray()
        sid = start_sector
        # defensive loop guard
        max_hops = 1_000_000

        while sid not in (FREE_SECTOR, END_OF_CHAIN) and max_hops > 0:
            out.extend(read_one(sid))
            if sid >= len(fat):
                # broken FAT, stop
                break
            sid = fat[sid]
            max_hops -= 1

        data = bytes(out)
        if size is not None:
            return data[:size]
        return data

    # ---------- public API ----------
    def list_streams(self) -> List[str]:
        return sorted(self.dir_lookup.keys())

    def exists(self, path: str) -> bool:
        return path in self.dir_lookup

    def read_stream(self, path: str) -> bytes:
        """
        Read a stream by full path (e.g., "BodyText/Section0").
        Selects MiniFAT vs FAT automatically based on stream_size and cutoff.
        """
        entry = self.dir_lookup.get(path)
        if not entry:
            raise FileNotFoundError(f"Stream not found: {path}")

        start = entry['start_sector']
        size = entry['stream_size']
        if start in (FREE_SECTOR, END_OF_CHAIN) or size == 0:
            return b''

        use_mini = size < self.mini_stream_cutoff
        return self._read_chain(start, size=size, use_mini=use_mini)
