"""Threaded comments: Range.AddCommentThreaded, Worksheet.CommentsThreaded and the CommentThreaded object, as Excel
makes, reads and saves them (scripts/measure_threaded_comments.py, tests/fixtures/threaded_comments.json).

A thread rides on a note in the sheet's notes, a stand-in that carries the thread's comments, so it moves, copies
and is saved with the sheet's notes; the notes VBA sees leave such stand-ins out. In the file:

- The comments part holds the stand-in: its author "tc={id}", its uid the thread's id, and its text, unformatted,
  a fixed notice and then "Comment:" and each "Reply:", each text four spaces in. Its box is a note's.
- The sheet's threadedComments part lists every comment, the threads row by row and each thread's replies after
  it, with its cell, its time in UTC to the hundredth of a second, its person and id, a reply its thread's id, and
  a resolved thread done="1".
- The workbook's persons part lists each person who wrote one.

A person is Application.UserName with no account, which Excel writes as providerId "None"; Excel signed in to an
account writes that account's name and id, which no measurement records.
"""
from __future__ import annotations

import datetime as dt
import re
import uuid
from dataclasses import dataclass, field, replace
from typing import TYPE_CHECKING

from pyopenvba._xml import attributes, escape as escape_attribute
from pyopenvba.apps.excel._creator import Created, ItemDefault
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.interpreter._objects import VBACollection, member, method, setter
from pyopenvba.interpreter._values import EMPTY, MISSING, NOTHING, VBADate, error, to_bool, to_integer, to_text

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Range, Workbook, Worksheet
    from pyopenvba.apps.excel._notes import Note

THREADS_TYPE = "application/vnd.ms-excel.threadedcomments+xml"
THREADS_RELATIONSHIP = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
PERSONS_TYPE = "application/vnd.ms-excel.person+xml"
PERSONS_RELATIONSHIP = "http://schemas.microsoft.com/office/2017/10/relationships/person"
_NAMESPACES = ('xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"'
               ' xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
#: What the stand-in note says before the thread's comments, for a version of Excel without threads.
NOTICE = ("[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits "
          "to it will get removed if the file is opened in a newer version of Excel. Learn more: "
          "https://go.microsoft.com/fwlink/?linkid=870924\n\n")


@dataclass
class ThreadEntry:
    """One comment of a thread: the thread's first, or a reply to it."""

    text: str
    person: str
    id: str
    #: When it was made, local time, to the hundredth of a second.
    when: dt.datetime
    #: The thread's first comment's id, for a reply.
    parent: str = ""
    #: Resolved, on a thread's first comment.
    done: bool = False


@dataclass
class Person:
    display_name: str
    id: str
    user_id: str
    provider_id: str = "None"


@dataclass
class PersonList:
    """A workbook's persons."""

    people: list[Person] = field(default_factory=lambda: [])


def _new_id() -> str:
    return "{" + str(uuid.uuid4()).upper() + "}"


def stand_in_text(entries: list[ThreadEntry]) -> str:
    """The stand-in note's text: the notice, then the comment and each reply, each text four spaces in."""
    pieces = [NOTICE + "Comment:\n    " + entries[0].text]
    pieces += ["Reply:\n    " + entry.text for entry in entries[1:]]
    return "\n".join(pieces)


def persons(book: Workbook) -> PersonList:
    return book.persons


def _person_for(book: Workbook) -> Person:
    """The application's user as a person, made the first time they write a threaded comment."""
    name = book.application.user_name
    people = persons(book).people
    found = next((person for person in people if person.display_name == name and person.provider_id == "None"), None)
    if found is None:
        found = Person(display_name=name, id=_new_id(), user_id=name)
        people.append(found)
    return found


def _now(sheet: Worksheet) -> dt.datetime:
    moment = sheet.book.calculator.now()
    hundredths = round(moment.microsecond / 10000)
    return moment.replace(microsecond=0) + dt.timedelta(milliseconds=10 * hundredths)


def threads(sheet: Worksheet) -> list[tuple[tuple[int, int], Note]]:
    """The sheet's threads' stand-ins, row by row, as CommentsThreaded lists them (tests/fixtures/
    threaded_comments.json)."""
    return sorted((key, note) for key, note in sheet.notes.items() if note.thread)


def thread_at(sheet: Worksheet, row: int, column: int) -> Note | None:
    note = sheet.notes.get((row, column))
    return note if note is not None and note.thread else None


def add_thread(target: Range, text: object) -> object:
    """Range.AddCommentThreaded: a thread on one cell with neither a note nor a thread; error 5 for several cells and
    1004 for a cell with either (tests/fixtures/threaded_comments.json)."""
    from pyopenvba.apps.excel._notes import Note, changed, next_named, placed

    if not target.single:
        raise error(5, "Invalid procedure call or argument")
    sheet, row, column = target.sheet, target.first.top, target.first.left
    if (row, column) in sheet.notes:
        raise error(1004, "Application-defined or object-defined error")
    person = _person_for(sheet.book)
    first = ThreadEntry(text=to_text(text), person=person.id, id=_new_id(), when=_now(sheet))
    name, shape_id = next_named(sheet, "Comment")
    sheet.notes[(row, column)] = Note(text="", author=f"tc={first.id}", name=name, shape_id=shape_id,
                                      anchor=placed(sheet, row, column), uid=first.id, thread=[first])
    changed(sheet)
    return CommentThreaded(sheet, first.id, first.id)


def thread_of(target: Range) -> object:
    """Range.CommentThreaded: the thread of a range of one cell, or Nothing; several cells are error 5
    (tests/fixtures/range_small_reads.json)."""
    if not target.single:
        raise error(5)
    note = thread_at(target.sheet, target.first.top, target.first.left)
    return NOTHING if note is None else CommentThreaded(target.sheet, note.thread[0].id, note.thread[0].id)


def refuse(sheet: Worksheet, areas: list[object], what: str) -> None:
    """``what`` refused on cells with a thread, which no measurement says what it does to."""
    from pyopenvba._a1 import Area

    for (row, column), note in sheet.notes.items():
        if note.thread and any(isinstance(area, Area) and area.contains(row, column) for area in areas):
            raise VBAUnsupportedError(f"{what} a cell with a threaded comment is not implemented")


# --- the objects ------------------------------------------------------------------------------------------------


class CommentThreaded(Created):
    """A thread's first comment, or one of its replies."""

    vba_type_name = "CommentThreaded"

    def __init__(self, sheet: Worksheet, thread: str, entry: str) -> None:
        self.sheet, self.thread, self.entry = sheet, thread, entry

    def _found(self) -> tuple[tuple[int, int], Note, ThreadEntry]:
        for key, note in self.sheet.notes.items():
            if note.thread and note.thread[0].id == self.thread:
                for one in note.thread:
                    if one.id == self.entry:
                        return key, note, one
        raise error(-2147221080, "Automation error\nObject disconnected from its clients. ")

    def _is_reply(self) -> bool:
        return self.thread != self.entry

    def _rewrite(self, key: tuple[int, int], note: Note, entries: list[ThreadEntry]) -> None:
        from pyopenvba.apps.excel._notes import changed

        self.sheet.notes[key] = replace(note, thread=entries, comment_xml="")
        changed(self.sheet)

    @method
    def Text(self, Text: object = MISSING, Start: object = MISSING, Overwrite: object = MISSING) -> object:
        """The comment's text; given Text, it becomes the text, which it hands back
        (tests/fixtures/threaded_comments.json)."""
        key, note, entry = self._found()
        if Text is MISSING:
            return entry.text
        if Start is not MISSING or Overwrite is not MISSING:
            raise VBAUnsupportedError("CommentThreaded.Text with Start or Overwrite is not implemented")
        wanted = to_text(Text)
        self._rewrite(key, note, [replace(one, text=wanted) if one.id == entry.id else one for one in note.thread])
        return wanted

    @method
    def AddReply(self, Text: object = MISSING) -> object:
        key, note, _ = self._found()
        if self._is_reply():
            raise VBAUnsupportedError("A reply to a reply is not implemented")
        person = _person_for(self.sheet.book)
        reply = ThreadEntry(text="" if Text is MISSING else to_text(Text), person=person.id, id=_new_id(),
                            when=_now(self.sheet), parent=self.thread)
        self._rewrite(key, note, [*note.thread, reply])
        return CommentThreaded(self.sheet, self.thread, reply.id)

    @member
    def Replies(self) -> object:
        """The thread's replies; a reply has none to give, error 450 (tests/fixtures/threaded_comments.json)."""
        self._found()
        if self._is_reply():
            raise error(450, "Wrong number of arguments or invalid property assignment")
        return CommentsThreaded(self.sheet, self.thread)

    @method
    def Delete(self) -> object:
        """The comment gone: a reply alone, a thread with its replies and its stand-in."""
        from pyopenvba.apps.excel._notes import remove

        key, note, entry = self._found()
        if self._is_reply():
            self._rewrite(key, note, [one for one in note.thread if one.id != entry.id])
        else:
            remove(self.sheet, *key)
        return EMPTY

    @member
    def Resolved(self) -> object:
        _, note, _ = self._found()
        return note.thread[0].done

    @setter("Resolved")
    def _resolved(self, value: object) -> None:
        key, note, _ = self._found()
        if self._is_reply():
            raise VBAUnsupportedError("Resolving a reply is not implemented")
        wanted = to_bool(value)
        self._rewrite(key, note, [replace(note.thread[0], done=wanted), *note.thread[1:]])

    @member
    def Date(self) -> object:
        """When the comment was made, local time to the hundredth of a second."""
        _, _, entry = self._found()
        return VBADate.from_datetime(entry.when)

    @member
    def Author(self) -> object:
        _, _, entry = self._found()
        person = next((one for one in persons(self.sheet.book).people if one.id == entry.person), None)
        if person is None:
            raise VBAUnsupportedError("A threaded comment whose person the workbook does not list is not implemented")
        return Author(self, person)

    @member
    def Parent(self) -> object:
        """The thread's cell; a reply's thread (tests/fixtures/threaded_comments.json)."""
        from pyopenvba._a1 import Area
        from pyopenvba.apps.excel._model import Range

        (row, column), _, _ = self._found()
        if self._is_reply():
            return CommentThreaded(self.sheet, self.thread, self.thread)
        return Range(self.sheet, [Area(row, column, row, column, self.sheet.name)])

    def _neighbour(self, step: int) -> object:
        self._found()
        if self._is_reply():
            raise VBAUnsupportedError("Next and Previous of a reply are not implemented")
        listed = [note.thread[0].id for _, note in threads(self.sheet)]
        at = listed.index(self.thread) + step
        if not 0 <= at < len(listed):
            return NOTHING
        return CommentThreaded(self.sheet, listed[at], listed[at])

    @method
    def Next(self) -> object:
        return self._neighbour(1)

    @method
    def Previous(self) -> object:
        return self._neighbour(-1)

    @member
    def Application(self) -> object:
        return self.sheet.book.application


class CommentsThreaded(VBACollection, ItemDefault, Created):
    """A sheet's threads, row by row; or, given a thread, its replies."""

    vba_type_name = "CommentsThreaded"

    def __init__(self, sheet: Worksheet, thread: str = "") -> None:
        self.sheet, self.thread = sheet, thread

    def vba_items(self) -> list[object]:
        if not self.thread:
            return [CommentThreaded(self.sheet, note.thread[0].id, note.thread[0].id) for _, note in threads(self.sheet)]
        note = next((one for one in self.sheet.notes.values() if one.thread and one.thread[0].id == self.thread),
                    None)
        if note is None:
            raise error(-2147221080, "Automation error\nObject disconnected from its clients. ")
        return [CommentThreaded(self.sheet, self.thread, entry.id) for entry in note.thread[1:]]

    def vba_lookup(self, index: object, items: list[object]) -> object:
        position = int(to_integer(index, "Long"))
        if not 1 <= position <= len(items):
            raise VBAUnsupportedError("CommentsThreaded beyond its items is not implemented")
        return items[position - 1]

    @member
    def Parent(self) -> object:
        if self.thread:
            return CommentThreaded(self.sheet, self.thread, self.thread)
        return self.sheet

    @member
    def Application(self) -> object:
        return self.sheet.book.application


class Author(Created):
    """The person who wrote a threaded comment."""

    vba_type_name = "Author"

    def __init__(self, comment: CommentThreaded, person: Person) -> None:
        self.comment, self.person = comment, person

    @member
    def Name(self) -> object:
        return self.person.display_name

    @member
    def ProviderID(self) -> object:
        return self.person.provider_id

    @member
    def UserID(self) -> object:
        return self.person.user_id

    @member
    def Parent(self) -> object:
        return self.comment

    @member
    def Application(self) -> object:
        return self.comment.sheet.book.application


# --- the parts ----------------------------------------------------------------------------------------------------


def _stamp(when: dt.datetime) -> str:
    """A local time as dT holds it: UTC, to the hundredth of a second (tests/fixtures/threaded_comments.json)."""
    from pyopenvba.apps.excel._document_properties import to_utc

    whole = to_utc(when.replace(microsecond=0))
    moment = dt.datetime(1, 1, 1) + dt.timedelta(seconds=whole)
    return f"{moment:%Y-%m-%dT%H:%M:%S}.{when.microsecond // 10000:02d}"


def threads_part(sheet: Worksheet) -> str:
    """The sheet's threadedComments part: every thread row by row, each reply after its thread."""
    from pyopenvba._a1 import column_letter

    from pyopenvba.apps.excel._notes_file import encoded

    written: list[str] = []
    for (row, column), note in threads(sheet):
        cell = f"{column_letter(column)}{row}"
        for entry in note.thread:
            parent = f' parentId="{entry.parent}"' if entry.parent else ""
            done = ' done="1"' if entry.done else ""
            text = f"<text>{encoded(entry.text)}</text>" if entry.text else "<text/>"
            written.append(f'<threadedComment ref="{cell}" dT="{_stamp(entry.when)}" personId="{entry.person}"'
                           f' id="{entry.id}"{parent}{done}>{text}</threadedComment>')
    return f"{_DECLARATION}<ThreadedComments {_NAMESPACES}>{''.join(written)}</ThreadedComments>"


def persons_part(book: Workbook) -> str:
    listed = "".join(f'<person displayName="{escape_attribute(person.display_name)}" id="{person.id}"'
                     f' userId="{escape_attribute(person.user_id)}" providerId="{escape_attribute(person.provider_id)}"/>'
                     for person in persons(book).people)
    return f"{_DECLARATION}<personList {_NAMESPACES}>{listed}</personList>"


_ENTRY = re.compile(r"<threadedComment\b[^>]*>.*?</threadedComment>|<threadedComment\b[^>]*/>", re.DOTALL)
_TEXT = re.compile(r"<text>(.*?)</text>", re.DOTALL)
_PERSON = re.compile(r"<person\b[^>]*/?>")


def read_persons(book: Workbook, xml: str) -> None:
    people = persons(book).people
    for found in _PERSON.finditer(xml):
        fields = attributes(found.group(0))
        people.append(Person(display_name=fields.get("displayName", ""), id=fields.get("id", ""),
                             user_id=fields.get("userId", ""), provider_id=fields.get("providerId", "")))


def read_threads(sheet: Worksheet, xml: str) -> None:
    """A loaded sheet's threads, put on the stand-in notes the comments part gave, by their ids."""
    from pyopenvba.apps.excel._document_properties import to_local
    from pyopenvba.apps.excel._notes_file import decoded

    by_thread: dict[str, list[ThreadEntry]] = {}
    for found in _ENTRY.finditer(xml):
        markup = found.group(0)
        fields = attributes(markup[:markup.find(">") + 1])
        text = _TEXT.search(markup)
        stamp = fields.get("dT", "")
        try:
            utc = dt.datetime.strptime(stamp[:19], "%Y-%m-%dT%H:%M:%S")
            hundredths = int((stamp[20:22] or "0").ljust(2, "0")) if len(stamp) > 19 else 0
            seconds = (utc - dt.datetime(1, 1, 1)) // dt.timedelta(seconds=1)
            when = to_local(seconds) + dt.timedelta(milliseconds=10 * hundredths)
        except ValueError:
            when = dt.datetime(1900, 1, 1)
        entry = ThreadEntry(text=decoded(text.group(1)) if text else "", person=fields.get("personId", ""),
                            id=fields.get("id", ""), when=when, parent=fields.get("parentId", ""),
                            done=fields.get("done", "") in ("1", "true"))
        by_thread.setdefault(entry.parent or entry.id, []).append(entry)
    for key, note in list(sheet.notes.items()):
        entries = by_thread.get(note.uid)
        if entries and note.author.startswith("tc="):
            sheet.notes[key] = replace(note, thread=entries)


__all__ = ["Author", "CommentThreaded", "CommentsThreaded", "NOTICE", "PERSONS_RELATIONSHIP", "PERSONS_TYPE",
           "THREADS_RELATIONSHIP", "THREADS_TYPE", "add_thread", "persons_part", "read_persons", "read_threads",
           "refuse", "stand_in_text", "thread_at", "thread_of", "threads", "threads_part"]
