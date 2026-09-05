"""Classic zlib's deflate, in Python, for the bytes Office writes.

Two formats here need it.  Access compresses an attachment with zlib at level 5, memLevel 7 and a
32 KB window (measured: over eight files from 25 bytes to 85 KB, one of
them admitting exactly one parameter set, classic zlib 1.3.2 reproduces
the engine's stream byte for byte -- header, blocks and trailer).  The
zlib a Python was built with may be zlib-ng, whose output differs, and
neither exposes every parameter, so this is zlib's own algorithm carried
over: the lazy matcher of ``deflate_slow``, ``longest_match`` with its
chain and niceness limits, the 8191-symbol block, and ``trees.c``'s
Huffman construction with its tie-breaking and depth overflow handling.
Anything that changed here would change the bytes, so the structure of
the C is kept where it decides the output.

Excel writes the parts of a Power Query package as raw deflate at level
6 -- .NET's ``CompressionLevel.Optimal`` -- which :func:`raw_compress`
reproduces; a 73 KB section document and 32 smaller parts pin the level
to 6 alone.
"""

from __future__ import annotations

import zlib
from dataclasses import dataclass, field

MIN_MATCH = 3
MAX_MATCH = 258
MIN_LOOKAHEAD = MAX_MATCH + MIN_MATCH + 1
TOO_FAR = 4096
LITERALS = 256
END_BLOCK = 256
L_CODES = LITERALS + 1 + 29
D_CODES = 30
BL_CODES = 19
HEAP_SIZE = 2 * L_CODES + 1
MAX_BITS = 15
MAX_BL_BITS = 7
REP_3_6 = 16
REPZ_3_10 = 17
REPZ_11_138 = 18
STORED_BLOCK = 0
STATIC_TREES = 1
DYN_TREES = 2

EXTRA_LBITS = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0]
EXTRA_DBITS = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13]
EXTRA_BLBITS = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 3, 7]
BL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]

#: zlib's configuration table: good length, lazy limit, nice length, chain.
CONFIGS = {
    1: (4, 4, 8, 4), 2: (4, 5, 16, 8), 3: (4, 6, 32, 32),
    4: (4, 4, 16, 16), 5: (8, 16, 32, 32), 6: (8, 16, 128, 128),
    7: (8, 32, 128, 256), 8: (32, 128, 258, 1024), 9: (32, 258, 258, 4096),
}


def _static_tables() -> tuple[list[int], list[int], list[int], list[int], list[int], list[int]]:
    """trees.c's tr_static_init: the length and distance code tables and
    the static Huffman trees."""
    length_code = [0] * (MAX_MATCH - MIN_MATCH + 1)
    base_length = [0] * 29
    length = 0
    for code in range(28):
        base_length[code] = length
        for _ in range(1 << EXTRA_LBITS[code]):
            length_code[length] = code
            length += 1
    length_code[length - 1] = 28
    dist_code = [0] * 512
    base_dist = [0] * D_CODES
    dist = 0
    for code in range(16):
        base_dist[code] = dist
        for _ in range(1 << EXTRA_DBITS[code]):
            dist_code[dist] = code
            dist += 1
    dist >>= 7
    for code in range(16, D_CODES):
        base_dist[code] = dist << 7
        for _ in range(1 << (EXTRA_DBITS[code] - 7)):
            dist_code[256 + dist] = code
            dist += 1
    static_llen = [8] * 144 + [9] * 112 + [7] * 24 + [8] * 8
    static_lcode = _gen_codes(static_llen, _bl_count(static_llen))
    static_dcode = [_reverse(n, 5) for n in range(D_CODES)]
    return length_code, base_length, dist_code, base_dist, static_lcode, static_dcode


def _reverse(code: int, length: int) -> int:
    out = 0
    for _ in range(length):
        out = (out << 1) | (code & 1)
        code >>= 1
    return out


def _bl_count(lengths: list[int]) -> list[int]:
    counts = [0] * (MAX_BITS + 1)
    for length in lengths:
        counts[length] += 1
    counts[0] = 0
    return counts


def _gen_codes(lengths: list[int], counts: list[int]) -> list[int]:
    next_code = [0] * (MAX_BITS + 1)
    code = 0
    for bits in range(1, MAX_BITS + 1):
        code = (code + counts[bits - 1]) << 1
        next_code[bits] = code
    codes = [0] * len(lengths)
    for n, length in enumerate(lengths):
        if length:
            codes[n] = _reverse(next_code[length], length)
            next_code[length] += 1
    return codes


LENGTH_CODE, BASE_LENGTH, DIST_CODE, BASE_DIST, STATIC_LCODE, STATIC_DCODE = _static_tables()
STATIC_LLEN = [8] * 144 + [9] * 112 + [7] * 24 + [8] * 8
STATIC_DLEN = [5] * D_CODES


def _d_code(dist: int) -> int:
    return DIST_CODE[dist] if dist < 256 else DIST_CODE[256 + (dist >> 7)]


@dataclass
class _Tree:
    """A dynamic Huffman tree under construction: frequencies in, lengths
    and codes out, with the static tree it is measured against."""

    freq: list[int]
    length: list[int]
    code: list[int]
    stat_len: list[int] | None
    extra_bits: list[int]
    extra_base: int
    elems: int
    max_length: int
    max_code: int = 0
    dad: list[int] = field(default_factory=lambda: [])


class _Deflater:
    def __init__(self, level: int, wbits: int, memlevel: int) -> None:
        self.level = level
        self.good_match, self.max_lazy, self.nice_match, self.max_chain = CONFIGS[level]
        self.w_bits = wbits
        self.w_size = 1 << wbits
        self.w_mask = self.w_size - 1
        self.hash_bits = memlevel + 7
        self.hash_size = 1 << self.hash_bits
        self.hash_mask = self.hash_size - 1
        self.hash_shift = (self.hash_bits + MIN_MATCH - 1) // MIN_MATCH
        self.window = bytearray(2 * self.w_size)
        self.prev = [0] * self.w_size
        self.head = [0] * self.hash_size
        self.lit_bufsize = 1 << (memlevel + 6)
        self.sym_end = self.lit_bufsize - 1
        self.syms: list[tuple[int, int]] = []  # (dist, lc)
        self.out = bytearray()
        self.bit_buf = 0
        self.bit_cnt = 0
        self.strstart = 0
        self.block_start = 0
        self.lookahead = 0
        self.insert = 0
        self.match_length = MIN_MATCH - 1
        self.prev_length = MIN_MATCH - 1
        self.match_start = 0
        self.prev_match = 0
        self.match_available = False
        self.ins_h = 0
        self.data = b""
        self.next_in = 0
        self.opt_len = 0
        self.static_len = 0
        self.bl_count = [0] * (MAX_BITS + 1)
        self.heap: list[int] = []
        self.heap_len = 0
        self.heap_max = 0
        self.depth = [0] * HEAP_SIZE
        self.dyn_ltree = _Tree([0] * HEAP_SIZE, [0] * HEAP_SIZE, [0] * HEAP_SIZE, STATIC_LLEN, EXTRA_LBITS, LITERALS + 1, L_CODES, MAX_BITS)
        self.dyn_dtree = _Tree([0] * (2 * D_CODES + 1), [0] * (2 * D_CODES + 1), [0] * (2 * D_CODES + 1), STATIC_DLEN, EXTRA_DBITS, 0, D_CODES, MAX_BITS)
        self.bl_tree = _Tree([0] * (2 * BL_CODES + 1), [0] * (2 * BL_CODES + 1), [0] * (2 * BL_CODES + 1), None, EXTRA_BLBITS, 0, BL_CODES, MAX_BL_BITS)
        self._init_block()

    # -- output -----------------------------------------------------------------

    def _send_bits(self, value: int, length: int) -> None:
        self.bit_buf |= value << self.bit_cnt
        self.bit_cnt += length
        while self.bit_cnt >= 8:
            self.out.append(self.bit_buf & 0xFF)
            self.bit_buf >>= 8
            self.bit_cnt -= 8

    def _bi_windup(self) -> None:
        if self.bit_cnt > 0:
            self.out.append(self.bit_buf & 0xFF)
        self.bit_buf = 0
        self.bit_cnt = 0

    # -- the sliding window and hash chains -----------------------------------------

    def _update_hash(self, h: int, c: int) -> int:
        return ((h << self.hash_shift) ^ c) & self.hash_mask

    def _insert_string(self, pos: int) -> int:
        self.ins_h = self._update_hash(self.ins_h, self.window[pos + MIN_MATCH - 1])
        head = self.head[self.ins_h]
        self.prev[pos & self.w_mask] = head
        self.head[self.ins_h] = pos
        return head

    def _fill_window(self) -> None:
        wsize = self.w_size
        window_size = 2 * wsize
        while True:
            more = window_size - self.lookahead - self.strstart
            if self.strstart >= wsize + (wsize - MIN_LOOKAHEAD):
                self.window[0 : wsize - more] = self.window[wsize : wsize + wsize - more]
                self.match_start -= wsize
                self.strstart -= wsize
                self.block_start -= wsize
                if self.insert > self.strstart:
                    self.insert = self.strstart
                self._slide_hash()
                more += wsize
            if self.next_in >= len(self.data):
                break
            chunk = self.data[self.next_in : self.next_in + more]
            self.next_in += len(chunk)
            start = self.strstart + self.lookahead
            self.window[start : start + len(chunk)] = chunk
            self.lookahead += len(chunk)
            if self.lookahead + self.insert >= MIN_MATCH:
                pos = self.strstart - self.insert
                self.ins_h = self.window[pos]
                self.ins_h = self._update_hash(self.ins_h, self.window[pos + 1])
                while self.insert:
                    self.ins_h = self._update_hash(self.ins_h, self.window[pos + MIN_MATCH - 1])
                    self.prev[pos & self.w_mask] = self.head[self.ins_h]
                    self.head[self.ins_h] = pos
                    pos += 1
                    self.insert -= 1
                    if self.lookahead + self.insert < MIN_MATCH:
                        break
            if not (self.lookahead < MIN_LOOKAHEAD and self.next_in < len(self.data)):
                break

    def _slide_hash(self) -> None:
        wsize = self.w_size
        self.head = [m - wsize if m >= wsize else 0 for m in self.head]
        self.prev = [m - wsize if m >= wsize else 0 for m in self.prev]

    def _longest_match(self, cur_match: int) -> int:
        chain_length = self.max_chain
        window = self.window
        strstart = self.strstart
        best_len = self.prev_length
        nice_match = self.nice_match
        max_dist = self.w_size - MIN_LOOKAHEAD
        limit = strstart - max_dist if strstart > max_dist else 0
        prev = self.prev
        wmask = self.w_mask
        strend = strstart + MAX_MATCH
        scan_end1 = window[strstart + best_len - 1]
        scan_end = window[strstart + best_len]
        if self.prev_length >= self.good_match:
            chain_length >>= 2
        if nice_match > self.lookahead:
            nice_match = self.lookahead
        while True:
            match = cur_match
            if (
                window[match + best_len] == scan_end
                and window[match + best_len - 1] == scan_end1
                and window[match] == window[strstart]
                and window[match + 1] == window[strstart + 1]
            ):
                scan = strstart + 2
                match += 2
                while scan < strend:
                    scan += 1
                    match += 1
                    if window[scan] != window[match]:
                        break
                length = scan - strstart if scan < strend else MAX_MATCH
                if scan >= strend:
                    length = MAX_MATCH
                if length > best_len:
                    self.match_start = cur_match
                    best_len = length
                    if length >= nice_match:
                        break
                    scan_end1 = window[strstart + best_len - 1]
                    scan_end = window[strstart + best_len]
            cur_match = prev[cur_match & wmask]
            if cur_match <= limit:
                break
            chain_length -= 1
            if chain_length == 0:
                break
        return best_len if best_len <= self.lookahead else self.lookahead

    # -- deflate_slow ------------------------------------------------------------------

    def _tally_lit(self, c: int) -> bool:
        self.syms.append((0, c))
        self.dyn_ltree.freq[c] += 1
        return len(self.syms) == self.sym_end

    def _tally_dist(self, dist: int, length: int) -> bool:
        self.syms.append((dist, length))
        dist -= 1
        self.dyn_ltree.freq[LENGTH_CODE[length] + LITERALS + 1] += 1
        self.dyn_dtree.freq[_d_code(dist)] += 1
        return len(self.syms) == self.sym_end

    def _flush_block(self, last: bool) -> None:
        start = self.block_start if self.block_start >= 0 else None
        self._tr_flush_block(start, self.strstart - self.block_start, last)
        self.block_start = self.strstart

    def deflate(self, data: bytes) -> bytes:
        self.data = data
        window = self.window
        while True:
            if self.lookahead < MIN_LOOKAHEAD:
                self._fill_window()
                if self.lookahead == 0:
                    break
            hash_head = 0
            if self.lookahead >= MIN_MATCH:
                hash_head = self._insert_string(self.strstart)
            self.prev_length = self.match_length
            self.prev_match = self.match_start
            self.match_length = MIN_MATCH - 1
            if hash_head != 0 and self.prev_length < self.max_lazy and self.strstart - hash_head <= self.w_size - MIN_LOOKAHEAD:
                self.match_length = self._longest_match(hash_head)
                if self.match_length <= 5 and self.match_length == MIN_MATCH and self.strstart - self.match_start > TOO_FAR:
                    self.match_length = MIN_MATCH - 1
            if self.prev_length >= MIN_MATCH and self.match_length <= self.prev_length:
                max_insert = self.strstart + self.lookahead - MIN_MATCH
                bflush = self._tally_dist(self.strstart - 1 - self.prev_match, self.prev_length - MIN_MATCH)
                self.lookahead -= self.prev_length - 1
                self.prev_length -= 2
                while True:
                    self.strstart += 1
                    if self.strstart <= max_insert:
                        self._insert_string(self.strstart)
                    self.prev_length -= 1
                    if self.prev_length == 0:
                        break
                self.match_available = False
                self.match_length = MIN_MATCH - 1
                self.strstart += 1
                if bflush:
                    self._flush_block(False)
            elif self.match_available:
                bflush = self._tally_lit(window[self.strstart - 1])
                if bflush:
                    self._flush_block(False)
                self.strstart += 1
                self.lookahead -= 1
            else:
                self.match_available = True
                self.strstart += 1
                self.lookahead -= 1
        if self.match_available:
            self._tally_lit(window[self.strstart - 1])
            self.match_available = False
        self.insert = self.strstart if self.strstart < MIN_MATCH - 1 else MIN_MATCH - 1
        self._flush_block(True)
        return bytes(self.out)

    # -- trees.c -------------------------------------------------------------------------

    def _init_block(self) -> None:
        for n in range(L_CODES):
            self.dyn_ltree.freq[n] = 0
        for n in range(D_CODES):
            self.dyn_dtree.freq[n] = 0
        for n in range(BL_CODES):
            self.bl_tree.freq[n] = 0
        self.dyn_ltree.freq[END_BLOCK] = 1
        self.opt_len = 0
        self.static_len = 0
        self.syms = []

    def _smaller(self, tree: _Tree, n: int, m: int) -> bool:
        return tree.freq[n] < tree.freq[m] or (tree.freq[n] == tree.freq[m] and self.depth[n] <= self.depth[m])

    def _pqdownheap(self, tree: _Tree, k: int) -> None:
        heap = self.heap
        v = heap[k]
        j = k << 1
        while j <= self.heap_len:
            if j < self.heap_len and self._smaller(tree, heap[j + 1], heap[j]):
                j += 1
            if self._smaller(tree, v, heap[j]):
                break
            heap[k] = heap[j]
            k = j
            j <<= 1
        heap[k] = v

    def _gen_bitlen(self, desc: _Tree) -> None:
        tree = desc
        stree = desc.stat_len
        extra = desc.extra_bits
        base = desc.extra_base
        max_length = desc.max_length
        overflow = 0
        for bits in range(MAX_BITS + 1):
            self.bl_count[bits] = 0
        tree.length[self.heap[self.heap_max]] = 0
        h = self.heap_max + 1
        while h < HEAP_SIZE:
            n = self.heap[h]
            bits = tree.length[tree.dad[n]] + 1
            if bits > max_length:
                bits = max_length
                overflow += 1
            tree.length[n] = bits
            if n > desc.max_code:
                h += 1
                continue
            self.bl_count[bits] += 1
            xbits = extra[n - base] if n >= base else 0
            f = tree.freq[n]
            self.opt_len += f * (bits + xbits)
            if stree is not None:
                self.static_len += f * (stree[n] + xbits)
            h += 1
        if overflow == 0:
            return
        while True:
            bits = max_length - 1
            while self.bl_count[bits] == 0:
                bits -= 1
            self.bl_count[bits] -= 1
            self.bl_count[bits + 1] += 2
            self.bl_count[max_length] -= 1
            overflow -= 2
            if overflow <= 0:
                break
        h = HEAP_SIZE
        bits = max_length
        while bits != 0:
            n = self.bl_count[bits]
            while n != 0:
                h -= 1
                m = self.heap[h]
                if m > desc.max_code:
                    continue
                if tree.length[m] != bits:
                    self.opt_len += (bits - tree.length[m]) * tree.freq[m]
                    tree.length[m] = bits
                n -= 1
            bits -= 1

    def _build_tree(self, desc: _Tree) -> None:
        tree = desc
        stree = desc.stat_len
        elems = desc.elems
        self.heap = [0] * HEAP_SIZE
        self.heap_len = 0
        self.heap_max = HEAP_SIZE
        tree.dad = [0] * len(tree.freq)
        max_code = -1
        for n in range(elems):
            if tree.freq[n] != 0:
                self.heap_len += 1
                self.heap[self.heap_len] = n
                max_code = n
                self.depth[n] = 0
            else:
                tree.length[n] = 0
        while self.heap_len < 2:
            if max_code < 2:
                max_code += 1
                node = max_code
            else:
                node = 0
            self.heap_len += 1
            self.heap[self.heap_len] = node
            tree.freq[node] = 1
            self.depth[node] = 0
            self.opt_len -= 1
            if stree is not None:
                self.static_len -= stree[node]
        desc.max_code = max_code
        for n in range(self.heap_len // 2, 0, -1):
            self._pqdownheap(tree, n)
        node = elems
        while True:
            n = self.heap[1]
            self.heap[1] = self.heap[self.heap_len]
            self.heap_len -= 1
            self._pqdownheap(tree, 1)
            m = self.heap[1]
            self.heap_max -= 1
            self.heap[self.heap_max] = n
            self.heap_max -= 1
            self.heap[self.heap_max] = m
            tree.freq[node] = tree.freq[n] + tree.freq[m]
            self.depth[node] = (self.depth[n] if self.depth[n] >= self.depth[m] else self.depth[m]) + 1
            tree.dad[n] = tree.dad[m] = node
            self.heap[1] = node
            node += 1
            self._pqdownheap(tree, 1)
            if self.heap_len < 2:
                break
        self.heap_max -= 1
        self.heap[self.heap_max] = self.heap[1]
        self._gen_bitlen(desc)
        codes = _gen_codes(tree.length[: max_code + 1], self.bl_count)
        for n in range(max_code + 1):
            tree.code[n] = codes[n]

    def _scan_tree(self, tree: _Tree, max_code: int) -> None:
        prevlen = -1
        nextlen = tree.length[0]
        count = 0
        max_count, min_count = (138, 3) if nextlen == 0 else (7, 4)
        tree.length[max_code + 1] = 0xFFFF
        for n in range(max_code + 1):
            curlen = nextlen
            nextlen = tree.length[n + 1]
            count += 1
            if count < max_count and curlen == nextlen:
                continue
            if count < min_count:
                self.bl_tree.freq[curlen] += count
            elif curlen != 0:
                if curlen != prevlen:
                    self.bl_tree.freq[curlen] += 1
                self.bl_tree.freq[REP_3_6] += 1
            elif count <= 10:
                self.bl_tree.freq[REPZ_3_10] += 1
            else:
                self.bl_tree.freq[REPZ_11_138] += 1
            count = 0
            prevlen = curlen
            if nextlen == 0:
                max_count, min_count = 138, 3
            elif curlen == nextlen:
                max_count, min_count = 6, 3
            else:
                max_count, min_count = 7, 4

    def _send_tree(self, tree: _Tree, max_code: int) -> None:
        prevlen = -1
        nextlen = tree.length[0]
        count = 0
        max_count, min_count = (138, 3) if nextlen == 0 else (7, 4)
        for n in range(max_code + 1):
            curlen = nextlen
            nextlen = tree.length[n + 1]
            count += 1
            if count < max_count and curlen == nextlen:
                continue
            if count < min_count:
                for _ in range(count):
                    self._send_code(curlen, self.bl_tree)
            elif curlen != 0:
                if curlen != prevlen:
                    self._send_code(curlen, self.bl_tree)
                    count -= 1
                self._send_code(REP_3_6, self.bl_tree)
                self._send_bits(count - 3, 2)
            elif count <= 10:
                self._send_code(REPZ_3_10, self.bl_tree)
                self._send_bits(count - 3, 3)
            else:
                self._send_code(REPZ_11_138, self.bl_tree)
                self._send_bits(count - 11, 7)
            count = 0
            prevlen = curlen
            if nextlen == 0:
                max_count, min_count = 138, 3
            elif curlen == nextlen:
                max_count, min_count = 6, 3
            else:
                max_count, min_count = 7, 4

    def _send_code(self, symbol: int, tree: _Tree) -> None:
        self._send_bits(tree.code[symbol], tree.length[symbol])

    def _build_bl_tree(self) -> int:
        self._scan_tree(self.dyn_ltree, self.dyn_ltree.max_code)
        self._scan_tree(self.dyn_dtree, self.dyn_dtree.max_code)
        self._build_tree(self.bl_tree)
        max_blindex = BL_CODES - 1
        while max_blindex >= 3:
            if self.bl_tree.length[BL_ORDER[max_blindex]] != 0:
                break
            max_blindex -= 1
        self.opt_len += 3 * (max_blindex + 1) + 5 + 5 + 4
        return max_blindex

    def _send_all_trees(self, lcodes: int, dcodes: int, blcodes: int) -> None:
        self._send_bits(lcodes - 257, 5)
        self._send_bits(dcodes - 1, 5)
        self._send_bits(blcodes - 4, 4)
        for rank in range(blcodes):
            self._send_bits(self.bl_tree.length[BL_ORDER[rank]], 3)
        self._send_tree(self.dyn_ltree, lcodes - 1)
        self._send_tree(self.dyn_dtree, dcodes - 1)

    def _compress_block(self, lcode: list[int], llen: list[int], dcode: list[int], dlen: list[int]) -> None:
        for dist, lc in self.syms:
            if dist == 0:
                self._send_bits(lcode[lc], llen[lc])
            else:
                code = LENGTH_CODE[lc]
                self._send_bits(lcode[code + LITERALS + 1], llen[code + LITERALS + 1])
                extra = EXTRA_LBITS[code]
                if extra:
                    self._send_bits(lc - BASE_LENGTH[code], extra)
                dist -= 1
                code = _d_code(dist)
                self._send_bits(dcode[code], dlen[code])
                extra = EXTRA_DBITS[code]
                if extra:
                    self._send_bits(dist - BASE_DIST[code], extra)
        self._send_bits(lcode[END_BLOCK], llen[END_BLOCK])

    def _stored_block(self, start: int, length: int, last: bool) -> None:
        self._send_bits((STORED_BLOCK << 1) + (1 if last else 0), 3)
        self._bi_windup()
        self.out += length.to_bytes(2, "little") + ((~length) & 0xFFFF).to_bytes(2, "little")
        self.out += self.window[start : start + length]

    def _tr_flush_block(self, start: int | None, stored_len: int, last: bool) -> None:
        self._build_tree(self.dyn_ltree)
        self._build_tree(self.dyn_dtree)
        max_blindex = self._build_bl_tree()
        opt_lenb = (self.opt_len + 3 + 7) >> 3
        static_lenb = (self.static_len + 3 + 7) >> 3
        if static_lenb <= opt_lenb:
            opt_lenb = static_lenb
        if stored_len + 4 <= opt_lenb and start is not None:
            self._stored_block(start, stored_len, last)
        elif static_lenb == opt_lenb:
            self._send_bits((STATIC_TREES << 1) + (1 if last else 0), 3)
            self._compress_block(STATIC_LCODE, STATIC_LLEN, STATIC_DCODE, STATIC_DLEN)
        else:
            self._send_bits((DYN_TREES << 1) + (1 if last else 0), 3)
            self._send_all_trees(self.dyn_ltree.max_code + 1, self.dyn_dtree.max_code + 1, max_blindex + 1)
            self._compress_block(self.dyn_ltree.code, self.dyn_ltree.length, self.dyn_dtree.code, self.dyn_dtree.length)
        self._init_block()
        if last:
            self._bi_windup()


def compress(data: bytes, level: int = 5, wbits: int = 15, memlevel: int = 7) -> bytes:
    """The zlib stream classic zlib writes for ``data`` at these settings:
    a two-byte header, the deflate blocks, the Adler-32 trailer.  The
    defaults are what Access uses for an attachment."""
    if not 1 <= level <= 9 or not 9 <= wbits <= 15 or not 1 <= memlevel <= 9:
        raise ValueError("level 1..9, wbits 9..15, memlevel 1..9")
    cmf = 8 | ((wbits - 8) << 4)
    level_flags = 0 if level < 2 else 1 if level < 6 else 2 if level == 6 else 3
    flg = level_flags << 6
    flg += 31 - ((cmf * 256 + flg) % 31)
    body = _Deflater(level, wbits, memlevel).deflate(data)
    return bytes((cmf, flg)) + body + (zlib.adler32(data) & 0xFFFFFFFF).to_bytes(4, "big")


def raw_compress(data: bytes, level: int = 6, memlevel: int = 8) -> bytes:
    """The deflate blocks alone, without zlib's header or its trailer.

    This is what a ZIP entry carries.  The defaults are what Excel uses
    for the parts of a Power Query package.
    """
    return compress(data, level=level, wbits=15, memlevel=memlevel)[2:-4]
