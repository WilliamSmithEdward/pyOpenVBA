"""Writing B-tree indexes: insert and delete entries, split pages, keep
the root where the table definition says it is.

Page layout is documented in ``_index``; what this module adds was
measured on trees the engine grew (docs/access_engine.md):

* Entries are ordered by their full bytes, key then row pointer.
* A page is rewritten compactly: entries after the first are stored
  without the bytes they share with it (the prefix length at 0x18), and
  the end-of-entry mask and free-space word follow from the entries.
* A leaf that fills while an entry is appended at its end stays full and
  the next leaf starts with the new entry (602 then 298 for 900 sequential
  keys); an insert in the middle splits the page in half (457 and 443
  for 900 random keys).
* The root page number never changes: a splitting root turns into a node
  whose two children are fresh pages.
* A node holds (separator, child) entries plus a tail child; a separator
  is the greatest entry under its child.  Byte 0x1A is the page's height
  above the leaves.
"""

from __future__ import annotations

import struct
from collections.abc import Callable
from dataclasses import dataclass

from pyopenvba.access_read import AccessError
from pyopenvba.access._index import (
    ENTRY_AREA,
    OFFSET_ENTRIES,
    OFFSET_ENTRY_MASK,
    OFFSET_FREE_SPACE,
    OFFSET_NEXT,
    OFFSET_OWNER,
    OFFSET_PREFIX_LENGTH,
    OFFSET_PREV,
    OFFSET_TAIL,
    IndexEntry,
    IndexPage,
    parse_index_page,
)
from pyopenvba.access._pages import PAGE_INDEX_LEAF, PAGE_INDEX_NODE, PAGE_SIZE, PageStore

OFFSET_LEVEL = 0x1A


def entry_bytes(entry: IndexEntry) -> bytes:
    """An entry as stored before prefix removal: key, row pointer, and on
    node pages the child page."""
    out = entry.key + entry.page.to_bytes(3, "big") + bytes((entry.row,))
    if entry.child is not None:
        out += entry.child.to_bytes(4, "big")
    return out


def sort_key(entry: IndexEntry) -> bytes:
    return entry.key + entry.page.to_bytes(3, "big") + bytes((entry.row,))


def same_row(a: IndexEntry, b: IndexEntry) -> bool:
    return (a.key, a.page, a.row) == (b.key, b.page, b.row)


def common_prefix(items: list[bytes]) -> int:
    if len(items) < 2:
        return 0
    first = items[0]
    length = 0
    while length < len(first) and all(
        len(item) > length and item[length] == first[length] for item in items
    ):
        length += 1
    return length


def stored_size(entries: list[IndexEntry], prefix: int | None = None) -> int:
    """Bytes the entry area needs.  ``prefix`` is the prefix length in
    force (shrunk if the entries no longer share it); ``None`` means the
    full common prefix."""
    raw = [entry_bytes(e) for e in entries]
    shared = common_prefix(raw)
    used = shared if prefix is None else min(prefix, shared)
    return sum(len(r) for r in raw) - used * max(len(raw) - 1, 0)


def serialize_index_page(
    entries: list[IndexEntry],
    *,
    is_leaf: bool,
    owner: int,
    prev: int,
    next: int,
    tail: int,
    level: int,
    prefix: int | None = None,
    base: bytes | None = None,
) -> bytes:
    """Lay out an index page.  ``prefix`` is the prefix length to keep
    (shrunk if the entries no longer share it); ``None`` computes the full
    common prefix, which is what the engine does for a page it fills or
    creates in a split.  ``base`` is the page's previous content: the
    engine rewrites entries over it and leaves the bytes past the used
    area as they were."""
    raw = bytearray(base) if base is not None else bytearray(PAGE_SIZE)
    raw[0] = PAGE_INDEX_LEAF if is_leaf else PAGE_INDEX_NODE
    raw[1] = 0x01
    struct.pack_into("<I", raw, OFFSET_OWNER, owner)
    struct.pack_into("<I", raw, OFFSET_PREV, prev)
    struct.pack_into("<I", raw, OFFSET_NEXT, next)
    struct.pack_into("<I", raw, OFFSET_TAIL, tail)
    raw[OFFSET_LEVEL] = level
    raw[OFFSET_ENTRY_MASK:OFFSET_ENTRIES] = bytes(OFFSET_ENTRIES - OFFSET_ENTRY_MASK)
    stored = [entry_bytes(e) for e in entries]
    shared = common_prefix(stored)
    prefix = shared if prefix is None else min(prefix, shared)
    struct.pack_into("<H", raw, OFFSET_PREFIX_LENGTH, prefix)
    position = 0
    for i, item in enumerate(stored):
        body = item if i == 0 else item[prefix:]
        end = position + len(body)
        if end > ENTRY_AREA:
            raise AccessError(f"{len(entries)} index entries do not fit one page")
        raw[OFFSET_ENTRIES + position : OFFSET_ENTRIES + end] = body
        raw[OFFSET_ENTRY_MASK + end // 8] |= 1 << (end % 8)
        position = end
    struct.pack_into("<H", raw, OFFSET_FREE_SPACE, ENTRY_AREA - position)
    return bytes(raw)


@dataclass
class _Step:
    """One page on the path from the root: the page and the position of
    the entry taken (``len(entries)`` means the tail child, or on a leaf
    the insertion point)."""

    page: IndexPage
    position: int


class BTree:
    """One index's tree.  ``allocate`` hands out fresh pages, already
    registered with the index's usage map."""

    def __init__(
        self, store: PageStore, root: int, owner: int, allocate: Callable[[], int]
    ) -> None:
        self.store = store
        self.root = root
        self.owner = owner
        self.allocate = allocate

    # -- helpers -------------------------------------------------------------

    def _load(self, number: int) -> IndexPage:
        page = parse_index_page(self.store, number)
        if page.owner != self.owner:
            raise AccessError(
                f"index page {number} belongs to table definition {page.owner}, not {self.owner}"
            )
        return page

    def _level(self, number: int) -> int:
        return self.store.read(number)[OFFSET_LEVEL]

    def _write_page(
        self,
        number: int,
        entries: list[IndexEntry],
        *,
        is_leaf: bool,
        prev: int,
        next: int,
        tail: int,
        level: int,
    ) -> None:
        self.store.write(
            number,
            serialize_index_page(
                entries, is_leaf=is_leaf, owner=self.owner, prev=prev, next=next, tail=tail, level=level
            ),
        )

    def _rewrite(self, page: IndexPage, entries: list[IndexEntry], *, tail: int | None = None,
                 prev: int | None = None, next: int | None = None,
                 prefix: int | None = -1) -> None:
        """Rewrite an existing page over its old bytes.  By default the
        page keeps its prefix length (shrunk if the entries no longer share
        it); ``prefix=None`` compresses with the full common prefix."""
        self.store.write(
            page.number,
            serialize_index_page(
                entries,
                is_leaf=page.is_leaf,
                owner=self.owner,
                prev=page.prev if prev is None else prev,
                next=page.next if next is None else next,
                tail=page.tail if tail is None else tail,
                level=self._level(page.number),
                prefix=page.prefix_length if prefix == -1 else prefix,
                base=self.store.read(page.number),
            ),
        )

    def _descend(self, probe: bytes) -> list[_Step]:
        path: list[_Step] = []
        number = self.root
        seen: set[int] = set()
        while True:
            if number in seen:
                raise AccessError(f"index rooted at {self.root} loops through page {number}")
            seen.add(number)
            page = self._load(number)
            position = 0
            while position < len(page.entries) and sort_key(page.entries[position]) < probe:
                position += 1
            path.append(_Step(page, position))
            if page.is_leaf:
                return path
            child = page.entries[position].child if position < len(page.entries) else page.tail
            if not child:
                raise AccessError(f"index node {number} has no child to descend into")
            number = child

    def _subtree_last(self, number: int) -> IndexEntry:
        page = self._load(number)
        while not page.is_leaf:
            child = page.tail if page.tail else (page.entries[-1].child or 0)
            page = self._load(child)
        if not page.entries:
            raise AccessError(f"index leaf {page.number} is empty")
        return page.entries[-1]

    # -- insert --------------------------------------------------------------

    def insert(self, key: bytes, row_page: int, row: int) -> bool:
        """Insert an entry; True when no other entry in the tree has ``key``."""
        entry = IndexEntry(key=key, page=row_page, row=row, child=None)
        probe = sort_key(entry)
        path = self._descend(probe)
        leaf = path[-1]
        entries = list(leaf.page.entries)
        if leaf.position < len(entries) and sort_key(entries[leaf.position]) == probe:
            raise AccessError(f"the index already holds row ({row_page}, {row}) under this key")
        distinct = self._key_is_new(leaf, entries, key)
        entries.insert(leaf.position, entry)
        self._place(path, entries, leaf.position == len(entries) - 1)
        return distinct

    def _place(self, path: list[_Step], entries: list[IndexEntry], appended: bool) -> None:
        """Write the page at the end of ``path`` with ``entries``: as it is
        while they fit under its current prefix, compressed with the full
        common prefix once they do not (the engine compresses a page only
        when it fills), and split when even that is not enough."""
        step = path[-1]
        if stored_size(entries, step.page.prefix_length) <= ENTRY_AREA:
            self._rewrite(step.page, entries)
        elif stored_size(entries) <= ENTRY_AREA:
            self._rewrite(step.page, entries, prefix=None)
        else:
            self._split(path, entries, appended)
            return
        if step.page.is_leaf:
            self._refresh_separators(path, entries[-1])

    def _key_is_new(self, leaf: _Step, entries: list[IndexEntry], key: bytes) -> bool:
        """Equal keys are adjacent, so only the neighbours of the insertion
        point can match -- including across the leaf boundary."""
        if leaf.position > 0:
            if entries[leaf.position - 1].key == key:
                return False
        elif leaf.page.prev:
            previous = self._load(leaf.page.prev)
            if previous.entries and previous.entries[-1].key == key:
                return False
        if leaf.position < len(entries):
            if entries[leaf.position].key == key:
                return False
        elif leaf.page.next:
            following = self._load(leaf.page.next)
            if following.entries and following.entries[0].key == key:
                return False
        return True

    def _refresh_separators(self, path: list[_Step], child_last: IndexEntry) -> None:
        """Keep every ancestor's separator equal to the greatest entry
        under it after that entry changed."""
        for depth in range(len(path) - 2, -1, -1):
            step = path[depth]
            if step.position >= len(step.page.entries):
                return  # the tail child has no separator
            current = step.page.entries[step.position]
            if same_row(current, child_last):
                return
            entries = list(step.page.entries)
            entries[step.position] = IndexEntry(child_last.key, child_last.page, child_last.row, current.child)
            self._rewrite(step.page, entries)
            if step.position != len(entries) - 1:
                return

    def _divide(
        self, entries: list[IndexEntry], appended: bool
    ) -> tuple[list[IndexEntry], list[IndexEntry]]:
        if appended:
            return entries[:-1], entries[-1:]
        sizes = [len(entry_bytes(e)) for e in entries]
        total = sum(sizes)
        running = 0
        cut = len(entries) // 2
        for i, size in enumerate(sizes):
            running += size
            if running * 2 >= total:
                cut = i + 1
                break
        cut = max(1, min(cut, len(entries) - 1))
        return entries[:cut], entries[cut:]

    def _split(self, path: list[_Step], entries: list[IndexEntry], appended: bool) -> None:
        """Split the page at the end of ``path``, whose new contents are
        ``entries`` (too many to fit), and register the halves above."""
        step = path[-1]
        page = step.page
        left, right = self._divide(entries, appended)
        if page.number == self.root:
            self._split_root(page, left, right)
            return
        level = self._level(page.number)
        right_number = self.allocate()
        if page.is_leaf:
            self._write_page(right_number, right, is_leaf=True, prev=page.number, next=page.next, tail=0, level=0)
            if page.next:
                following = self._load(page.next)
                self._rewrite(following, following.entries, prev=right_number)
            self._rewrite(page, left, next=right_number)
            left_last, right_last = left[-1], right[-1]
        else:
            # The last left entry becomes the left node's tail child; the
            # right node keeps the original tail.
            left_tail = left[-1].child or 0
            self._write_page(right_number, right, is_leaf=False, prev=0, next=0, tail=page.tail, level=level)
            self._write_page(page.number, left[:-1], is_leaf=False, prev=0, next=0, tail=left_tail, level=level)
            left_last = IndexEntry(left[-1].key, left[-1].page, left[-1].row, None)
            right_last = self._subtree_last(right_number)
        self._register_split(path[:-1], page.number, left_last, right_number, right_last)

    def _split_root(self, root: IndexPage, left: list[IndexEntry], right: list[IndexEntry]) -> None:
        level = self._level(root.number)
        left_number = self.allocate()
        right_number = self.allocate()
        if root.is_leaf:
            self._write_page(left_number, left, is_leaf=True, prev=0, next=right_number, tail=0, level=0)
            self._write_page(right_number, right, is_leaf=True, prev=left_number, next=0, tail=0, level=0)
            separator = IndexEntry(left[-1].key, left[-1].page, left[-1].row, left_number)
        else:
            left_tail = left[-1].child or 0
            self._write_page(left_number, left[:-1], is_leaf=False, prev=0, next=0, tail=left_tail, level=level)
            self._write_page(right_number, right, is_leaf=False, prev=0, next=0, tail=root.tail, level=level)
            separator = IndexEntry(left[-1].key, left[-1].page, left[-1].row, left_number)
        # The new root is written from scratch: one entry, no prefix.
        self.store.write(
            root.number,
            serialize_index_page(
                [separator], is_leaf=False, owner=self.owner, prev=0, next=0, tail=right_number, level=level + 1, prefix=0
            ),
        )

    def _register_split(
        self,
        path: list[_Step],
        left_number: int,
        left_last: IndexEntry,
        right_number: int,
        right_last: IndexEntry,
    ) -> None:
        """``left_number`` (an existing child of the node at the end of
        ``path``) split off ``right_number``; give the node an entry for
        each half, splitting the node too if it overflows."""
        step = path[-1]
        node = step.page
        entries = list(node.entries)
        separator = IndexEntry(left_last.key, left_last.page, left_last.row, left_number)
        tail = node.tail
        if step.position < len(entries):
            entries[step.position] = IndexEntry(right_last.key, right_last.page, right_last.row, right_number)
            entries.insert(step.position, separator)
            appended = False
        else:
            entries.append(separator)
            tail = right_number
            appended = True
        if stored_size(entries, node.prefix_length) <= ENTRY_AREA:
            self._rewrite(node, entries, tail=tail)
            return
        if stored_size(entries) <= ENTRY_AREA:
            self._rewrite(node, entries, tail=tail, prefix=None)
            return
        grown = IndexPage(node.number, False, node.owner, 0, node.prev, node.next, tail, 0, entries)
        self._split(path[:-1] + [_Step(grown, step.position)], entries, appended)

    # -- delete --------------------------------------------------------------

    def delete(self, key: bytes, row_page: int, row: int) -> None:
        probe = key + row_page.to_bytes(3, "big") + bytes((row,))
        path = self._descend(probe)
        leaf = path[-1]
        entries = list(leaf.page.entries)
        if leaf.position >= len(entries) or sort_key(entries[leaf.position]) != probe:
            raise AccessError(f"the index has no entry for row ({row_page}, {row}) under this key")
        del entries[leaf.position]
        self._rewrite(leaf.page, entries)
        if entries:
            self._refresh_separators(path, entries[-1])
