# -*- coding: utf-8 -*-
import argparse
import base64
import hashlib
import json
from pathlib import Path
import re
import xml.etree.ElementTree as ET
import zipfile


W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
B = '{http://schemas.openxmlformats.org/officeDocument/2006/bibliography}'


def field_kind(code):
    if 'ZOTERO_ITEM CSL_CITATION' in code:
        return 'zotero'
    if re.search(r'\bADDIN\s+EN\.CITE(?:\s|$)', code):
        return 'endnote'
    if code.strip().startswith('CITATION '):
        return 'word'
    if 'ZOTERO_BIBL' in code or 'EN.REFLIST' in code or code.strip().startswith('BIBLIOGRAPHY'):
        return 'bibliography'
    return 'other'


def read_fields(tree):
    fields, stack = [], []

    def begin(node, code='', simple=False):
        field = {'code': code, 'text': '', 'data': '', 'children': [], 'start': node, 'end': node, 'nodes': set(), 'result': simple}
        (stack[-1]['children'] if stack else fields).append(field)
        stack.append(field)

    def visit(node):
        if node.tag == W + 'fldSimple':
            begin(node, node.get(W + 'instr', ''), True)
        elif node.tag == W + 'fldChar' and node.get(W + 'fldCharType') == 'begin':
            begin(node)
        for field in stack:
            field['nodes'].add(node)
        if node.tag == W + 'instrText' and stack:
            stack[-1]['code'] += node.text or ''
        elif node.tag == W + 'fldData' and stack:
            stack[-1]['data'] += base64.b64decode(''.join((node.text or '').split()), validate=True).decode('utf-8-sig').rstrip('\x00')
        elif node.tag == W + 't':
            for field in stack:
                if field['result']:
                    field['text'] += node.text or ''
        elif node.tag == W + 'fldChar' and node.get(W + 'fldCharType') == 'separate' and stack:
            stack[-1]['result'] = True
        for child in node:
            visit(child)
        if node.tag == W + 'fldSimple' or (node.tag == W + 'fldChar' and node.get(W + 'fldCharType') == 'end'):
            if not stack:
                raise ValueError('Field end without matching begin')
            stack[-1]['end'] = node
            stack.pop()

    visit(tree)
    if stack:
        raise ValueError('Unclosed Word field')
    return fields


def all_fields(fields):
    for field in fields:
        yield field
        yield from all_fields(field['children'])


def inspect_docx(path):
    groups, bibliographies, native, native_parts, native_groups, sources, variables, errors = [], [], 0, {}, [], {}, {}, []
    with zipfile.ZipFile(path) as archive:
        parts = sorted(n for n in archive.namelist() if re.fullmatch(r'word/(document|footnotes|endnotes|header\d+|footer\d+)\.xml', n))
        for name in parts:
            tree = ET.fromstring(archive.read(name))
            for field in all_fields(read_fields(tree)):
                kind = field_kind(field['code'])
                if kind == 'bibliography':
                    bibliographies.append({'part': name, 'characters': len(field['text']), 'text_sha256': hashlib.sha256(field['text'].encode('utf-8')).hexdigest()})
                elif kind == 'word':
                    native += 1
                    native_parts[name] = native_parts.get(name, 0) + 1
                    tags = re.findall(r'(?:^\s*CITATION\s+|\\m\s+)("[^"]+"|\S+)', field['code'])
                    native_groups.append({'part': name, 'tags': [t.strip('"') for t in tags]})
                elif kind in ('zotero', 'endnote'):
                    group = {'part': name, 'engine': kind, 'items': []}
                    try:
                        if kind == 'zotero':
                            data = json.loads(field['code'].split('CSL_CITATION', 1)[1])
                            group['citation_id'] = data['citationID']
                            for item in data['citationItems']:
                                metadata = item.get('itemData', {})
                                group['items'].append({'id': item.get('uris') or item.get('id'), 'title': metadata.get('title', ''), 'doi': metadata.get('DOI', ''), 'metadata_sha256': hashlib.sha256(json.dumps(metadata, sort_keys=True, ensure_ascii=False).encode('utf-8')).hexdigest(), 'locator': item.get('locator', ''), 'label': item.get('label', ''), 'prefix': item.get('prefix', ''), 'suffix': item.get('suffix', ''), 'suppress_author': item.get('suppress-author', False), 'author_only': item.get('author-only', False)})
                        else:
                            payload = re.split(r'ADDIN\s+EN\.CITE\s*', field['code'], maxsplit=1)[1].strip()
                            if not payload.startswith('<'):
                                payload = next(f['data'] for f in all_fields(field['children']) if f['data'])
                            data = ET.fromstring(payload)
                            for cite in data.findall('.//Cite'):
                                record = cite.find('record')
                                if record is None:
                                    raise ValueError('EndNote citation has no embedded record; resolve against the library')
                                group['items'].append({'id': record.findtext('rec-number', cite.findtext('RecNum', '')), 'title': ''.join(record.find('titles/title').itertext()) if record.find('titles/title') is not None else '', 'doi': record.findtext('electronic-resource-num', ''), 'metadata_sha256': hashlib.sha256(ET.canonicalize(ET.tostring(record, encoding='unicode')).encode('utf-8')).hexdigest()})
                        if not group['items']:
                            raise ValueError('Citation group has no decoded items')
                    except (ValueError, KeyError, IndexError, StopIteration, ET.ParseError) as error:
                        errors.append({'part': name, 'group': len(groups) + 1, 'error': str(error) or 'Missing nested EndNote field data'})
                    groups.append(group)
        for name in archive.namelist():
            if re.fullmatch(r'customXml/item\d+\.xml', name):
                xml = ET.fromstring(archive.read(name))
                for node in xml.iter(B + 'Source'):
                    tag = node.findtext(B + 'Tag')
                    if tag in sources:
                        errors.append({'error': 'Duplicate Word bibliography source tag: ' + tag})
                    sources[tag] = node
        for group in native_groups:
            if not group['tags'] or any(tag not in sources for tag in group['tags']):
                errors.append({'error': 'Native Word citation lacks embedded source metadata', 'part': group['part'], 'tags': group['tags']})
        if 'word/settings.xml' in archive.namelist():
            settings = ET.fromstring(archive.read('word/settings.xml'))
            variables = {n.get(W + 'name'): n.get(W + 'val', '') for n in settings.iter(W + 'docVar')}
    layout = variables.get('EN.Layout')
    style = ET.fromstring(layout).findtext('Style') if layout else None
    ids = [g['citation_id'] for g in groups if g.get('citation_id')]
    if len(ids) != len(set(ids)):
        errors.append({'error': 'Duplicate Zotero citation IDs; map locations before editing'})
    unique = {(g['engine'], json.dumps(item['id'], sort_keys=True)) for g in groups for item in g['items']}
    return {'groups': groups, 'group_count': len(groups), 'unique_item_count': len(unique), 'native_word_citations': native, 'native_word_parts': native_parts, 'native_word_groups': native_groups, 'bibliographies': bibliographies, 'endnote_style': style, 'errors': errors}


def masked_document(data):
    tree = ET.fromstring(data)
    fields = read_fields(tree)
    parents = {child: parent for parent in tree.iter() for child in parent}

    def ancestor(node, tag):
        while node is not None and node.tag != tag:
            node = parents.get(node)
        return node

    for field in reversed(fields):
        kind = field_kind(field['code'])
        if kind == 'other':
            continue
        start, end = field['start'], field['end']
        first, last = ancestor(start, W + 'p'), ancestor(end, W + 'p')
        if first is None or last is None:
            raise ValueError('Field outside a paragraph needs a separate integrity check')
        if kind == 'bibliography':
            parent = parents[first]
            if parents[last] is not parent:
                raise ValueError('Bibliography spans different containers')
            siblings = list(parent)
            a, b = siblings.index(first), siblings.index(last)
            selected = siblings[a:b + 1]
            if any(n not in field['nodes'] for p in selected for n in p.iter() if n.tag in (W + 't', W + 'drawing', W + 'object') and (n.text or n.tag != W + 't')):
                raise ValueError('Bibliography shares a paragraph with unrelated content')
            replacement = ET.Element(W + 'p')
            ET.SubElement(ET.SubElement(replacement, W + 'r'), W + 't').text = '__BIBLIOGRAPHY__'
        else:
            if first is not last:
                raise ValueError('Citation spans multiple paragraphs')
            parent = first
            while parents.get(start) is not parent:
                start = parents[start]
            while parents.get(end) is not parent:
                end = parents[end]
            siblings = list(parent)
            a, b = siblings.index(start), siblings.index(end)
            selected = siblings[a:b + 1]
            if any(n not in field['nodes'] for child in selected for n in child.iter() if n.tag in (W + 't', W + 'drawing', W + 'object') and (n.text or n.tag != W + 't')):
                raise ValueError('Citation boundary shares a run with unrelated content')
            replacement = ET.Element(W + 'r')
            ET.SubElement(replacement, W + 't').text = '__CITATION__'
        for node in selected:
            parent.remove(node)
        parent.insert(a, replacement)
    # Ignore Word's editing-session identifiers, not text, styles, or layout.
    for node in tree.iter():
        for key in list(node.attrib):
            if key.startswith(W + 'rsid') or key.rsplit('}', 1)[-1] in ('paraId', 'textId'):
                del node.attrib[key]
    return ET.canonicalize(ET.tostring(tree, encoding='unicode'))


def compare_docx(before, after):
    source, target = inspect_docx(before), inspect_docx(after)
    changes, unresolved, cross_engine = [], [], []
    if source['group_count'] != target['group_count'] or source['native_word_citations'] or target['native_word_citations']:
        changes.append('Citation group counts differ or native Word citations need a separate mapping audit')
    for index, (old, new) in enumerate(zip(source['groups'], target['groups']), start=1):
        if old['part'] != new['part'] or len(old['items']) != len(new['items']):
            changes.append('Group %d changed document part or membership count' % index)
        # Match by stable identity. Titles alone cannot prove that records are identical.
        def identity(item, same_engine):
            if same_engine and old['engine'] == 'zotero':
                return json.dumps(item['id'], sort_keys=True)
            if same_engine and old['engine'] == 'endnote':
                return str(item['id']) + ':' + item['metadata_sha256']
            doi = re.sub(r'^https?://(?:dx\.)?doi\.org/|^doi:\s*', '', item.get('doi', '').strip(), flags=re.I).lower()
            return doi or None
        same = old['engine'] == new['engine']
        if not same:
            cross_engine.append(index)
        old_ids, new_ids = [identity(x, same) for x in old['items']], [identity(x, same) for x in new['items']]
        if None in old_ids + new_ids:
            unresolved.append(index)
        elif old_ids != new_ids:
            changes.append('Group %d has different item identities or ordering' % index)
        if old['engine'] == new['engine'] == 'zotero':
            old_details = [{k: v for k, v in x.items() if k not in ('id', 'title')} for x in old['items']]
            new_details = [{k: v for k, v in x.items() if k not in ('id', 'title')} for x in new['items']]
            if old_details != new_details:
                changes.append('Group %d changed metadata or citation options' % index)
        elif same and old['items'] != new['items']:
            changes.append('Group %d changed embedded EndNote records' % index)
    with zipfile.ZipFile(before) as a, zipfile.ZipFile(after) as b:
        names = set(a.namelist()) | set(b.namelist())
        changed = sorted(n for n in names if n not in a.namelist() or n not in b.namelist() or a.read(n) != b.read(n))
        outside, limitation = None, None
        try:
            outside = masked_document(a.read('word/document.xml')) == masked_document(b.read('word/document.xml'))
        except ValueError as error:
            limitation = str(error)
    review = [n for n in changed if n != 'word/document.xml']
    return {'group_changes': changes, 'unresolved_identity_groups': unresolved, 'cross_engine_groups_require_mapping': cross_engine, 'outside_citations_xml_equal': outside, 'normalization_limitation': limitation, 'changed_parts': changed, 'other_changed_parts_require_review': review, 'source_errors': source['errors'], 'target_errors': target['errors'], 'passed': not (changes or unresolved or cross_engine or review or source['errors'] or target['errors']) and outside is True}


def main():
    parser = argparse.ArgumentParser(description='Read-only audit of saved Zotero/EndNote DOCX fields and document preservation.')
    parser.add_argument('--input', required=True)
    parser.add_argument('--compare', help='Compare this saved output against --input; changed package parts require review.')
    parser.add_argument('--report', help='Write a UTF-8 JSON report; citation metadata may be private.')
    args = parser.parse_args()
    result = compare_docx(args.input, args.compare) if args.compare else inspect_docx(args.input)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        Path(args.report).write_text(text, encoding='utf-8')
    print(text)
    if result.get('errors') or result.get('passed') is False:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
