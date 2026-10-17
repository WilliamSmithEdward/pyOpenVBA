"""Rich-value column snapshots used by worksheet pivot caches with modern errors."""
from __future__ import annotations

from typing import TYPE_CHECKING

from pyopenvba.apps.excel._charts_file import XmlElement, xml_bytes, xml_tree
from pyopenvba.exceptions import VBAUnsupportedError
from pyopenvba.powerquery._sheets import add_content_type, add_relationship

if TYPE_CHECKING:
    from pyopenvba.apps.excel._model import Workbook
    from pyopenvba.powerquery._opc import OpcFile

RICH = '{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}'
RICH2 = '{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2}'
FOLDER = 'xl/richData/'
VALUES = FOLDER + 'rdrichvalue.xml'
SHAPES = FOLDER + 'rdrichvaluestructure.xml'
ARRAYS = FOLDER + 'rdarray.xml'
PROP_SHAPES = FOLDER + 'rdsupportingpropertybagstructure.xml'
PROPS = FOLDER + 'rdsupportingpropertybag.xml'
ERRORS = {8: '#SPILL!', 13: '#CALC!'}


def _tree(package: OpcFile, part: str, namespace: str, tag: str) -> XmlElement:
    return xml_tree(package.read(part)) if package.has(part) else XmlElement(namespace + tag, {'count': '0'})


def datasets(package: OpcFile) -> dict[str, dict[str, list[str | None]] | None]:
    """Read error-column snapshots; unsupported rich payloads remain explicitly unknown."""
    from pyopenvba.apps.excel._pivots import decode_text

    values = _tree(package, VALUES, RICH, 'rvData')
    shapes = _tree(package, SHAPES, RICH, 'rvStructures')
    arrays = _tree(package, ARRAYS, RICH2, 'arrayData')

    def fields(index: int, kind: str) -> dict[str, str]:
        if not 0 <= index < len(values):
            raise ValueError('Invalid rich value index')
        value = values[index]
        shape_index = int(value.get('s', '-1'))
        if not 0 <= shape_index < len(shapes):
            raise ValueError('Invalid rich structure index')
        shape = shapes[shape_index]
        if shape.get('t') != kind:
            raise ValueError('Unsupported rich value')
        return {decode_text(key.get('n', '')): entry.text or '' for key, entry in zip(shape, value)}

    result: dict[str, dict[str, list[str | None]] | None] = {}
    for index, value in enumerate(values):
        shape_index = int(value.get('s', '-1'))
        if not 0 <= shape_index < len(shapes) or shapes[shape_index].get('t') != '_datasourcecontainer':
            continue
        item = fields(index, '_datasourcecontainer')
        guid = item.get('%XLUID', '')
        columns: dict[str, list[str | None]] = {}
        try:
            for key in shapes[shape_index]:
                if key.get('t') != 'r':
                    continue
                name = decode_text(key.get('n', ''))
                array_index = int(fields(int(item[name]), '_array')['array'])
                if not 0 <= array_index < len(arrays):
                    raise ValueError('Invalid rich array index')
                array = arrays[array_index]
                if not len(array) or array[-1].get('t') != 'i' or array[-1].text != '2':
                    raise ValueError('Unsupported pivot column')
                column: list[str | None] = []
                for entry in list(array)[:-1]:
                    if entry.get('t') == 'e' and entry.text == '#NULL!':
                        column.append(None)
                    else:
                        error = fields(int(entry.text or '-1'), '_error')
                        column.append(ERRORS[int(error['errorType'])])
                columns[name] = column
        except (ValueError, IndexError, KeyError):
            result[guid] = None
        else:
            result[guid] = columns
    return result


def write(book: Workbook, package: OpcFile) -> None:
    """Append cache containers after ordinary rich errors, retaining their worksheet indices."""
    from pyopenvba.apps.excel._pivot_create import encoded

    caches = [cache for cache in book.pivot_caches if cache.rich_records]
    if not caches:
        return
    present = datasets(package)
    pending = [cache for cache in caches if cache.rich_guid not in present]
    if not pending:
        return
    values = _tree(package, VALUES, RICH, 'rvData')
    shapes = _tree(package, SHAPES, RICH, 'rvStructures')
    arrays = _tree(package, ARRAYS, RICH2, 'arrayData')
    prop_shapes = _tree(package, PROP_SHAPES, RICH2, 'spbStructures')
    props = _tree(package, PROPS, RICH2, 'supportingPropertyBags')
    prop_arrays = props.find(RICH2 + 'spbArrays')
    prop_data = props.find(RICH2 + 'spbData')
    if prop_arrays is None:
        prop_arrays = XmlElement(RICH2 + 'spbArrays')
        props.append(prop_arrays)
    if prop_data is None:
        prop_data = XmlElement(RICH2 + 'spbData')
        props.append(prop_data)

    def shape(kind: str, keys: tuple[tuple[str, str], ...]) -> int:
        for index, entry in enumerate(shapes):
            if entry.get('t') == kind and tuple((key.get('n', ''), key.get('t', '')) for key in entry) == keys:
                return index
        entry = XmlElement(RICH + 's', {'t': kind})
        for name, form in keys:
            entry.append(XmlElement(RICH + 'k', {'n': name, 't': form}))
        shapes.append(entry)
        return len(shapes) - 1

    def value(kind: str, keys: tuple[tuple[str, str], ...], entries: list[str]) -> int:
        entry = XmlElement(RICH + 'rv', {'s': str(shape(kind, keys))})
        for text in entries:
            child = XmlElement(RICH + 'v')
            child.text = text
            entry.append(child)
        values.append(entry)
        return len(values) - 1

    for cache in pending:
        if cache.rich_columns is None:
            raise VBAUnsupportedError('Writing this pivot rich-data payload is not implemented')
        columns = sorted(cache.rich_columns, key=str.casefold)
        references: list[str] = []
        for name in columns:
            array = XmlElement(RICH2 + 'a', {'r': str(len(cache.rich_columns[name]) + 1)})
            for error in cache.rich_columns[name]:
                if error is None:
                    child = XmlElement(RICH2 + 'v', {'t': 'e'})
                    child.text = '#NULL!'
                else:
                    code = next(code for code, text in ERRORS.items() if text == error)
                    index = value('_error', (('errorType', 'i'), ('propagated', 'b')), [str(code), '1'])
                    child = XmlElement(RICH2 + 'v', {'t': 'r'})
                    child.text = str(index)
                array.append(child)
            marker = XmlElement(RICH2 + 'v', {'t': 'i'})
            marker.text = '2'
            array.append(marker)
            arrays.append(array)
            index = value('_array', (('array', 'a'),), [str(len(arrays) - 1)])
            references.append(str(index))
        order = XmlElement(RICH2 + 'a', {'count': str(len(columns) + 3)})
        for name in [*columns, '_CRID', '%XLUID', '_Display']:
            child = XmlElement(RICH2 + 'v', {'t': 's'})
            child.text = encoded(name)
            order.append(child)
        prop_arrays.append(order)
        order_shape = XmlElement(RICH2 + 's')
        order_shape.append(XmlElement(RICH2 + 'k', {'n': '^Order', 't': 'spba'}))
        prop_shapes.append(order_shape)
        display = XmlElement(RICH2 + 'spb', {'s': str(len(prop_shapes) - 1)})
        display_value = XmlElement(RICH2 + 'v')
        display_value.text = str(len(prop_arrays) - 1)
        display.append(display_value)
        prop_data.append(display)
        keys = (('%XLUID', 's'), ('_CRID', 'i'), ('_Display', 'spb')) + tuple((encoded(name), 'r') for name in columns)
        value('_datasourcecontainer', keys, [cache.rich_guid, '0', str(len(prop_data) - 1), *references])
    office = 'http://schemas.microsoft.com/office/2017/06/relationships/'
    for part, root, kind in ((VALUES, values, 'rdRichValue'), (SHAPES, shapes, 'rdRichValueStructure'),
                             (ARRAYS, arrays, 'rdArray'), (PROP_SHAPES, prop_shapes, 'rdSupportingPropertyBagStructure'),
                             (PROPS, props, 'rdSupportingPropertyBag')):
        if root is not props:
            root.set('count', str(len(root)))
        else:
            root.attrib.pop('count', None)
            prop_arrays.set('count', str(len(prop_arrays)))
            prop_data.set('count', str(len(prop_data)))
        package.write(part, xml_bytes(root))
        add_content_type(package, part, 'application/vnd.ms-excel.' + part.rsplit('/', 1)[1][:-4] + '+xml')
        rels = 'xl/_rels/workbook.xml.rels'
        target = part[3:]
        if target not in package.read(rels).decode('utf-8'):
            add_relationship(package, rels, office + kind, target)
