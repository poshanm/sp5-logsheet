#!/usr/bin/env python3
import struct
from pathlib import Path
import sys
import zipfile

FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
DIFSECT = 0xFFFFFFFC


def u16(b,o): return struct.unpack_from('<H',b,o)[0]
def u32(b,o): return struct.unpack_from('<I',b,o)[0]
def u64(b,o): return struct.unpack_from('<Q',b,o)[0]


class CFB:
    def __init__(self, data: bytes):
        self.data = data
        if data[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            raise ValueError('Not CFB')
        self.sector_shift = u16(data, 0x1E)
        self.sector_size = 1 << self.sector_shift
        self.mini_sector_shift = u16(data, 0x20)
        self.mini_sector_size = 1 << self.mini_sector_shift
        self.num_fat_sectors = u32(data, 0x2C)
        self.first_dir_sector = u32(data, 0x30)
        self.mini_stream_cutoff = u32(data, 0x38)
        self.first_mini_fat_sector = u32(data, 0x3C)
        self.num_mini_fat_sectors = u32(data, 0x40)
        self.first_difat_sector = u32(data, 0x44)
        self.num_difat_sectors = u32(data, 0x48)
        self.difat = self._load_difat()
        self.fat = self._load_fat()
        self.dir_entries = self._load_directory()
        self.root = self.dir_entries[0]
        self.mini_fat = self._load_mini_fat()
        self.mini_stream = self._load_stream_std(self.root['start_sector'], self.root['size'])

    def _sector_off(self, sid):
        return (sid + 1) * self.sector_size

    def _read_sector(self, sid):
        off = self._sector_off(sid)
        return self.data[off: off + self.sector_size]

    def _chain(self, start_sid, table):
        out = []
        sid = start_sid
        seen = set()
        while sid not in (ENDOFCHAIN, FREESECT) and sid < len(table):
            if sid in seen:
                break
            seen.add(sid)
            out.append(sid)
            sid = table[sid]
        return out

    def _load_difat(self):
        difat = []
        for i in range(109):
            sid = u32(self.data, 0x4C + i*4)
            if sid != FREESECT:
                difat.append(sid)
        sid = self.first_difat_sector
        for _ in range(self.num_difat_sectors):
            if sid in (ENDOFCHAIN, FREESECT):
                break
            sec = self._read_sector(sid)
            n = self.sector_size // 4
            for i in range(n-1):
                fsid = u32(sec, i*4)
                if fsid != FREESECT:
                    difat.append(fsid)
            sid = u32(sec, (n-1)*4)
        return difat

    def _load_fat(self):
        fat = []
        for sid in self.difat[:self.num_fat_sectors]:
            sec = self._read_sector(sid)
            for i in range(self.sector_size // 4):
                fat.append(u32(sec, i*4))
        return fat

    def _load_stream_std(self, start_sid, size):
        if size == 0:
            return b''
        chunks = []
        for sid in self._chain(start_sid, self.fat):
            chunks.append(self._read_sector(sid))
        return b''.join(chunks)[:size]

    def _load_mini_fat(self):
        if self.first_mini_fat_sector in (ENDOFCHAIN, FREESECT) or self.num_mini_fat_sectors == 0:
            return []
        data = self._load_stream_std(self.first_mini_fat_sector, self.num_mini_fat_sectors * self.sector_size)
        return [u32(data, i) for i in range(0, len(data), 4)]

    def _load_directory(self):
        d = self._load_stream_std(self.first_dir_sector, 10_000_000)
        entries = []
        for off in range(0, len(d), 128):
            e = d[off:off+128]
            if len(e) < 128:
                break
            name_len = u16(e, 64)
            if name_len >= 2:
                name = e[:name_len-2].decode('utf-16le', errors='ignore')
            else:
                name = ''
            entries.append({
                'name': name,
                'type': e[66],
                'left': u32(e,68),
                'right': u32(e,72),
                'child': u32(e,76),
                'start_sector': u32(e,116),
                'size': u64(e,120),
            })
        return entries

    def _collect_children(self, idx):
        out = []
        def walk(i):
            if i in (FREESECT, ENDOFCHAIN) or i >= len(self.dir_entries):
                return
            n = self.dir_entries[i]
            walk(n['left'])
            out.append(i)
            walk(n['right'])
        child = self.dir_entries[idx]['child']
        walk(child)
        return out

    def find_storage(self, name):
        for i,e in enumerate(self.dir_entries):
            if e['type'] == 1 and e['name'].lower() == name.lower():
                return i
        return None

    def list_streams_in_storage(self, storage_name):
        idx = self.find_storage(storage_name)
        if idx is None:
            return []
        result = []
        for ci in self._collect_children(idx):
            e = self.dir_entries[ci]
            if e['type'] == 2:
                result.append((e['name'], ci))
            elif e['type'] == 1:
                for gci in self._collect_children(ci):
                    ge = self.dir_entries[gci]
                    if ge['type'] == 2:
                        result.append((f"{e['name']}/{ge['name']}", gci))
        return result

    def read_stream_by_index(self, idx):
        e = self.dir_entries[idx]
        size = e['size']
        if size < self.mini_stream_cutoff and self.mini_stream:
            chunks = []
            sid = e['start_sector']
            seen = set()
            while sid not in (ENDOFCHAIN, FREESECT) and sid < len(self.mini_fat):
                if sid in seen:
                    break
                seen.add(sid)
                off = sid * self.mini_sector_size
                chunks.append(self.mini_stream[off:off+self.mini_sector_size])
                sid = self.mini_fat[sid]
            return b''.join(chunks)[:size]
        return self._load_stream_std(e['start_sector'], size)


def decompress_vba(data: bytes) -> bytes:
    if not data:
        return b''
    if data[0] != 0x01:
        return data
    out = bytearray()
    pos = 1
    while pos + 2 <= len(data):
        header = struct.unpack_from('<H', data, pos)[0]
        pos += 2
        chunk_size = (header & 0x0FFF) + 3
        chunk_end = min(len(data), pos + chunk_size - 2)
        compressed = (header & 0x8000) != 0
        if not compressed:
            out.extend(data[pos:chunk_end])
            pos = chunk_end
            continue
        chunk_start_out = len(out)
        while pos < chunk_end:
            flag = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if (flag >> bit) & 1 == 0:
                    out.append(data[pos]); pos += 1
                else:
                    if pos + 2 > chunk_end:
                        pos = chunk_end; break
                    token = struct.unpack_from('<H', data, pos)[0]
                    pos += 2
                    # dynamic bit count per spec
                    diff = len(out) - chunk_start_out
                    bit_count = 4
                    while (1 << bit_count) < diff and bit_count < 12:
                        bit_count += 1
                    length_mask = (1 << (16 - bit_count)) - 1
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    # guard against malformed tokens
                    if offset <= 0 or offset > len(out):
                        continue
                    for _ in range(length):
                        if offset > len(out):
                            break
                        out.append(out[-offset])
    return bytes(out)


def main(path):
    with zipfile.ZipFile(path) as z:
        vba_bin = z.read('xl/vbaProject.bin')
    cfb = CFB(vba_bin)
    streams = cfb.list_streams_in_storage('VBA')
    outdir = Path('extracted_vba')
    outdir.mkdir(exist_ok=True)
    for name, idx in streams:
        raw = cfb.read_stream_by_index(idx)
        safe = name.replace('/', '_')
        (outdir / f'{safe}.bin').write_bytes(raw)
        try:
            dec = decompress_vba(raw)
            (outdir / f'{safe}.txt').write_bytes(dec)
            print(name, len(raw), '->', len(dec))
        except Exception as e:
            (outdir / f'{safe}.txt').write_bytes(raw)
            print(name, len(raw), '-> raw', e)

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv)>1 else 'SP5_Production_v1.xlsm')
