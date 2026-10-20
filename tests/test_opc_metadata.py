"""Preserve valid OPC namespace spellings while adding sheets and VBA projects."""
from __future__ import annotations

import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from pyopenvba import ExcelFile, PowerPointFile, WordFile
from pyopenvba._new_project import PROJECT_CONTENT, PROJECT_TYPE, with_project
from pyopenvba._opc_xml import RELS_NS, TYPES_NS, opc_part
from pyopenvba.powerquery._opc import OpcFile
from pyopenvba.powerquery._sheets import add_content_type, add_relationship


@pytest.mark.parametrize('prefix', ['', 'opc:'])
@pytest.mark.parametrize('empty', [False, True])
def test_qualified_relationship_insertion_preserves_unknown_markup(prefix: str, empty: bool) -> None:
    declaration = f'xmlns:{prefix[:-1]}' if prefix else 'xmlns'
    opening = f'<{prefix}Relationships {declaration}="{RELS_NS}" xmlns:x="urn:extra" x:kind="x:Token" x:label="é >"'
    body = ('<!-- <Relationship Id="rId1"/> -->'
            '<x:Relationship Id="rId1"/>'
            f'<{prefix}Relationship Id=\'rId2\' Target="a&amp;b.xml" Type="old"></{prefix}Relationship>')
    raw = (opening + ('/>' if empty else '>' + body + f'</{prefix}Relationships>')).encode()
    package = OpcFile()
    package.write('test.rels', raw)
    assert add_relationship(package, 'test.rels', 'new&kind', 'é"<.xml') == 'rId1'
    saved = package.read('test.rels')
    root = ET.fromstring(saved)
    assert root.attrib['{urn:extra}kind'] == 'x:Token'
    added = root.findall('{' + RELS_NS + '}Relationship')[-1]
    assert added.attrib == {'Id': 'rId1', 'Type': 'new&kind', 'Target': 'é"<.xml'}
    if not empty:
        assert body.encode() in saved
        assert root[0].tag == '{urn:extra}Relationship'


@pytest.mark.parametrize('prefix', ['', 'ct:'])
def test_content_type_insert_is_qualified_and_idempotent(prefix: str) -> None:
    declaration = f'xmlns:{prefix[:-1]}' if prefix else 'xmlns'
    raw = f'<{prefix}Types {declaration}="{TYPES_NS}"/>'.encode()
    package = OpcFile()
    package.write('[Content_Types].xml', raw)
    add_content_type(package, 'xl/é&sheet.xml', 'type&value')
    saved = package.read('[Content_Types].xml')
    add_content_type(package, 'xl/é&sheet.xml', 'type&value')
    assert package.read('[Content_Types].xml') == saved
    assert ET.fromstring(saved)[0].attrib == {'PartName': '/xl/é&sheet.xml', 'ContentType': 'type&value'}


@pytest.mark.parametrize('host,opener,folder,suffix', [
    ('excel', ExcelFile, 'xl', 'xlsm'), ('word', WordFile, 'word', 'docm'),
    ('powerpoint', PowerPointFile, 'ppt', 'pptm'),
])
def test_first_project_handles_prefixed_package_metadata(host: str, opener: type, folder: str,
                                                        suffix: str, tmp_path: Path) -> None:
    source = Path(__file__).parent / 'fixtures/first_macro' / f'{host}_before.{suffix}'
    target = tmp_path / f'prefixed.{suffix}'
    with ZipFile(source) as package, ZipFile(target, 'w') as output:
        for name in package.namelist():
            data = package.read(name)
            if name.endswith('.rels') or name == '[Content_Types].xml':
                text = data.decode()
                text = text.replace('xmlns=', 'xmlns:opc=', 1)
                text = re.sub(r'<(/?)(Relationships|Relationship|Types|Default|Override)(?=[\s/>])',
                              r'<\1opc:\2', text)
                data = text.encode()
            output.writestr(name, data)
    with opener(target) as office_file:
        office_file.add_vba_project().add_module('Module1', 'Public Sub Hello()\r\nEnd Sub\r\n')
        office_file.save()
    with ZipFile(target) as package:
        assert folder + '/vbaProject.bin' in package.namelist()
        main = {'excel': 'workbook.xml', 'word': 'document.xml', 'powerpoint': 'presentation.xml'}[host]
        root = ET.fromstring(package.read(f'{folder}/_rels/{main}.rels'))
        project = [node for node in root if node.attrib.get('Type') == PROJECT_TYPE]
        assert len(project) == 1
        assert project[0].tag == '{' + RELS_NS + '}Relationship'
        types = ET.fromstring(package.read('[Content_Types].xml'))
        assert any(node.tag == '{' + TYPES_NS + '}Default' and
                   node.attrib == {'Extension': 'bin', 'ContentType': PROJECT_CONTENT} for node in types)
    with opener(target) as office_file:
        assert 'Module1' in office_file.module_names()


def test_existing_bin_type_gets_qualified_project_override() -> None:
    parts = {
        'xl/_rels/workbook.xml.rels': f'<r:Relationships xmlns:r="{RELS_NS}"/>'.encode(),
        '[Content_Types].xml': (f'<c:Types xmlns:c="{TYPES_NS}">'
                                '<c:Default Extension="bin" ContentType="other"/></c:Types>').encode(),
    }
    edits = with_project(list(parts), parts.__getitem__, 'xl/workbook.xml')
    types = ET.fromstring(edits['[Content_Types].xml'])
    assert types[0].attrib['ContentType'] == 'other'
    assert types[1].tag == '{' + TYPES_NS + '}Override'
    assert types[1].attrib == {'PartName': '/xl/vbaProject.bin', 'ContentType': PROJECT_CONTENT}


@pytest.mark.parametrize('declaration', [
    '<!DOCTYPE Relationships>',
    '<!DOCTYPE Relationships [<!ENTITY x "payload">]>',
    '<!DOCTYPE Relationships SYSTEM "file:///unavailable">',
])
def test_opc_metadata_refuses_declarations(declaration: str) -> None:
    data = (declaration + f'<Relationships xmlns="{RELS_NS}"/>').encode()
    with pytest.raises(ValueError, match='DTD'):
        opc_part(data, RELS_NS, 'Relationships')


def test_opc_metadata_rejects_wrong_namespace_and_utf16() -> None:
    with pytest.raises(ValueError, match='namespace'):
        opc_part(b'<Relationships xmlns="urn:foreign"/>', RELS_NS, 'Relationships')
    with pytest.raises((ValueError, UnicodeError)):
        opc_part(f'<Relationships xmlns="{RELS_NS}"/>'.encode('utf-16'), RELS_NS, 'Relationships')


@pytest.mark.parametrize('encoding', ['ISO-8859-1', 'US-ASCII'])
def test_non_utf8_declarations_are_not_reused_for_utf8_insertions(encoding: str) -> None:
    data = (f'<?xml version="1.0" encoding="{encoding}"?>'
            f'<Relationships xmlns="{RELS_NS}"/>').encode()
    with pytest.raises(ValueError, match='UTF-8'):
        opc_part(data, RELS_NS, 'Relationships')
