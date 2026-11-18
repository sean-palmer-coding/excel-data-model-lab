"""Native representation of a stored dictionary column containing only blanks."""
from xml.dom import minidom

from .schema import _elements, _text


def remove_empty_dictionary(model, key, encoding, replacements, removals):
    metadata = minidom.parseString(replacements.get(
        encoding.metadata_path, model._image.read_logical_file(encoding.metadata_path)))
    raw = next(n for n in _elements(metadata, 'XMObject')
               if n.getAttribute('class') == 'XMRawColumn'
               and n.getAttribute('name') == model._column_ids[key])
    dictionary = next(n for n in _elements(raw, 'XMObject')
                      if n.getAttribute('class').startswith('XMHashDataDictionary'))
    filename = dictionary.getAttribute('name')
    if dictionary.hasAttribute('name'):
        dictionary.removeAttribute('name')
    for stat in _elements(raw, 'DistinctStates'):
        stat.firstChild.data = '0'
    replacements[encoding.metadata_path] = metadata.documentElement.toxml(encoding='utf-8')
    if encoding.dictionary_path:
        removals.append(encoding.dictionary_path)
        replacements.pop(encoding.dictionary_path, None)
    if filename:
        path = model._dimension_paths[key[0]]
        dimension = minidom.parseString(replacements.get(path, model._image.read_logical_file(path)))
        listing = _elements(dimension, 'DataFileList')[0]
        listing.firstChild.data = ';'.join(x for x in _text(listing).split(';') if x != filename)
        replacements[path] = dimension.documentElement.toxml(encoding='utf-8')
