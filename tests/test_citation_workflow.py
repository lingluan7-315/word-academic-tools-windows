# -*- coding: utf-8 -*-
import argparse
import base64
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import types
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
import citation_audit as audit
import endnote_word_windows as endnote

W = audit.W


def run(parent, tag, text=None, attributes=None):
    node = ET.SubElement(ET.SubElement(parent, W + 'r'), W + tag, attributes or {})
    node.text = text
    return node


def field(parent, code, display='[1]', data=None):
    run(parent, 'fldChar', attributes={W + 'fldCharType': 'begin'})
    middle = len(code) // 2
    run(parent, 'instrText', code[:middle])
    run(parent, 'instrText', code[middle:])
    run(parent, 'fldChar', attributes={W + 'fldCharType': 'separate'})
    if data is not None:
        begin = run(parent, 'fldChar', attributes={W + 'fldCharType': 'begin'})
        ET.SubElement(begin, W + 'fldData').text = base64.b64encode(data.encode('utf-8')).decode('ascii')
        run(parent, 'instrText', ' ADDIN EN.CITE.DATA ')
        run(parent, 'fldChar', attributes={W + 'fldCharType': 'separate'})
        run(parent, 'fldChar', attributes={W + 'fldCharType': 'end'})
    run(parent, 't', display)
    run(parent, 'fldChar', attributes={W + 'fldCharType': 'end'})


def zotero_code(key='ITEM0001', doi='10.0000/a', locator=''):
    return ' ADDIN ZOTERO_ITEM CSL_CITATION ' + json.dumps({'citationID': 'citation-1', 'citationItems': [{'id': 1, 'uris': ['http://zotero.org/users/1/items/' + key], 'itemData': {'title': 'Same title', 'DOI': doi}, 'locator': locator}]})


def endnote_xml(doi='10.0000/a'):
    root = ET.Element('EndNote')
    cite = ET.SubElement(root, 'Cite')
    record = ET.SubElement(cite, 'record')
    ET.SubElement(record, 'rec-number').text = '7'
    ET.SubElement(ET.SubElement(record, 'titles'), 'title').text = '测试 article'
    ET.SubElement(record, 'electronic-resource-num').text = doi
    return ET.tostring(root, encoding='unicode')


def document(path, code=None, display='[1]', data=None, prose='Original prose.', image=b'image', layout=None, bibliography=True):
    root = ET.Element(W + 'document')
    body = ET.SubElement(root, W + 'body')
    p = ET.SubElement(body, W + 'p')
    run(p, 't', prose)
    field(p, code or zotero_code(), display, data)
    if bibliography:
        field(ET.SubElement(body, W + 'p'), ' ADDIN EN.REFLIST ' if code and 'EN.CITE' in code else ' ADDIN ZOTERO_BIBL ', 'References')
    settings = ET.Element(W + 'settings')
    if layout:
        variables = ET.SubElement(settings, W + 'docVars')
        ET.SubElement(variables, W + 'docVar', {W + 'name': 'EN.Layout', W + 'val': '<ENLayout><Style>' + layout + '</Style></ENLayout>'})
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr('word/document.xml', ET.tostring(root))
        archive.writestr('word/settings.xml', ET.tostring(settings))
        archive.writestr('word/media/image1.png', image)
        sources = ET.Element(audit.B + 'Sources')
        for tag in ('Source1', 'TestSource'):
            source = ET.SubElement(sources, audit.B + 'Source')
            ET.SubElement(source, audit.B + 'Tag').text = tag
            ET.SubElement(source, audit.B + 'Title').text = 'Synthetic source'
        archive.writestr('customXml/item1.xml', ET.tostring(sources))
    return root


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.source, self.target = self.root / 'source.docx', self.root / 'target.docx'

    def tearDown(self):
        self.temporary.cleanup()

    def test_split_instructions_and_nested_unicode_endnote_data(self):
        document(self.source, ' ADDIN EN.CITE ', data=endnote_xml())
        result = audit.inspect_docx(self.source)
        self.assertEqual(result['group_count'], 1)
        self.assertEqual(result['groups'][0]['items'][0]['title'], '测试 article')
        self.assertEqual(result['errors'], [])

    def test_display_only_changes_preserve_non_citation_content(self):
        document(self.source)
        document(self.target, display='(Author, 2026)')
        self.assertTrue(audit.compare_docx(self.source, self.target)['passed'])

    def test_endnote_base64_can_contain_xml_whitespace(self):
        root = document(self.source, ' ADDIN EN.CITE ', data=endnote_xml())
        data = root.find('.//' + W + 'fldData')
        data.text = '\n '.join(data.text[i:i + 40] for i in range(0, len(data.text), 40))
        fields = audit.read_fields(root)
        self.assertEqual(fields[0]['children'][0]['data'], endnote_xml())

    def test_same_title_does_not_hide_a_different_reference(self):
        document(self.source)
        document(self.target, zotero_code(key='ITEM0002'))
        result = audit.compare_docx(self.source, self.target)
        self.assertFalse(result['passed'])
        self.assertTrue(result['group_changes'])

    def test_locator_change_is_detected_even_if_key_and_title_match(self):
        document(self.source, zotero_code(locator='10'))
        document(self.target, zotero_code(locator='20'))
        self.assertTrue(audit.compare_docx(self.source, self.target)['group_changes'])

    def test_prose_change_and_media_change_are_detected(self):
        document(self.source)
        document(self.target, prose='Changed prose.', image=b'replaced')
        result = audit.compare_docx(self.source, self.target)
        self.assertFalse(result['outside_citations_xml_equal'])
        self.assertIn('word/media/image1.png', result['other_changed_parts_require_review'])
        self.assertFalse(result['passed'])

    def test_missing_cross_engine_identity_requires_mapping(self):
        document(self.source, zotero_code(doi=''))
        document(self.target, ' ADDIN EN.CITE ' + endnote_xml(doi=''))
        self.assertEqual(audit.compare_docx(self.source, self.target)['unresolved_identity_groups'], [1])

    def test_same_endnote_record_without_doi_can_be_compared(self):
        document(self.source, ' ADDIN EN.CITE ' + endnote_xml(doi=''))
        document(self.target, ' ADDIN EN.CITE ' + endnote_xml(doi=''))
        self.assertTrue(audit.compare_docx(self.source, self.target)['passed'])

    def test_cross_engine_doi_match_keeps_group_identity(self):
        document(self.source)
        document(self.target, ' ADDIN EN.CITE ' + endnote_xml())
        result = audit.compare_docx(self.source, self.target)
        self.assertFalse(result['passed'])
        self.assertEqual(result['group_changes'], [])
        self.assertEqual(result['cross_engine_groups_require_mapping'], [1])

    def test_shared_run_is_not_silently_removed(self):
        root = document(self.source)
        p = root.find('.//' + W + 'p')
        boundary = list(p)[1]
        boundary.insert(0, ET.Element(W + 't'))
        boundary[0].text = 'Unrelated boundary text'
        with self.assertRaisesRegex(ValueError, 'shares a run'):
            audit.masked_document(ET.tostring(root))

    def test_unclosed_field_is_rejected(self):
        tree = ET.Element(W + 'p')
        run(tree, 'fldChar', attributes={W + 'fldCharType': 'begin'})
        with self.assertRaisesRegex(ValueError, 'Unclosed'):
            audit.read_fields(tree)

    def test_duplicate_zotero_citation_ids_are_reported(self):
        root = document(self.source)
        field(ET.SubElement(root.find(W + 'body'), W + 'p'), zotero_code())
        with zipfile.ZipFile(self.target, 'w') as archive:
            archive.writestr('word/document.xml', ET.tostring(root))
        self.assertTrue(audit.inspect_docx(self.target)['errors'])

    def test_simple_word_field_and_header_locations_are_reported(self):
        document(self.source)
        header = ET.Element(W + 'hdr')
        p = ET.SubElement(header, W + 'p')
        simple = ET.SubElement(p, W + 'fldSimple', {W + 'instr': 'CITATION TestSource'})
        run(simple, 't', '[2]')
        with zipfile.ZipFile(self.source, 'a') as archive:
            archive.writestr('word/header1.xml', ET.tostring(header))
        result = audit.inspect_docx(self.source)
        self.assertEqual(result['native_word_parts'], {'word/header1.xml': 1})

    def test_zotero_preflight_rejects_without_opening_word(self):
        document(self.source)
        args = argparse.Namespace(input=str(self.source), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[]):
            result = endnote.preflight(args)
        self.assertFalse(result['ready'])
        self.assertTrue(any('Zotero' in issue for issue in result['issues']))

    def test_missing_style_and_existing_layout_are_checked(self):
        document(self.source, ' CITATION Source1 ', layout='ExampleStyle')
        args = argparse.Namespace(input=str(self.source), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[self.root]):
            self.assertFalse(endnote.preflight(args)['ready'])
            (self.root / 'ExampleStyle.ens').write_bytes(b'synthetic-style')
            self.assertTrue(endnote.preflight(args)['ready'])

    def test_style_selection_does_not_fabricate_missing_layout(self):
        document(self.source, ' CITATION Source1 ')
        (self.root / 'ExampleStyle.ens').write_bytes(b'synthetic-style')
        args = argparse.Namespace(input=str(self.source), refresh_only=False, style='ExampleStyle', style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[self.root]):
            result = endnote.preflight(args)
        self.assertFalse(result['ready'])
        self.assertTrue(any('EN.Layout' in issue for issue in result['issues']))

    def test_native_citation_without_embedded_source_is_rejected(self):
        document(self.source, ' CITATION MissingSource ')
        args = argparse.Namespace(input=str(self.source), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[]):
            self.assertFalse(endnote.preflight(args)['ready'])

    def test_internal_worker_cannot_overwrite_existing_output(self):
        document(self.source, ' CITATION Source1 ')
        self.target.write_bytes(b'Keep existing output')
        script = Path(endnote.__file__)
        result = subprocess.run([sys.executable, str(script), '--worker', '--input', str(self.source), '--output', str(self.target)], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(self.target.read_bytes(), b'Keep existing output')
        self.assertIn(b'existing files are never overwritten', result.stderr)

    def test_output_and_status_report_cannot_alias(self):
        document(self.source, ' CITATION Source1 ')
        result = subprocess.run([sys.executable, str(Path(endnote.__file__)), '--input', str(self.source), '--output', str(self.target), '--report', str(self.target)], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(self.target.exists())

    def test_internal_worker_does_not_overwrite_an_unowned_report(self):
        document(self.source, ' CITATION Source1 ')
        report = self.root / 'private.json'
        report.write_text('{"private": true}', encoding='utf-8')
        result = subprocess.run([sys.executable, str(Path(endnote.__file__)), '--worker', '--input', str(self.source), '--output', str(self.target), '--report', str(report)], capture_output=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(report.read_text(encoding='utf-8'), '{"private": true}')
        self.assertFalse(self.target.exists())

    def test_blocked_worker_times_out_with_diagnostic(self):
        report = self.root / 'status.json'
        endnote.write_status(report, {'phase': 'mock_cwyw', 'refresh_verified': False})
        started = time.monotonic()
        code = endnote.run_bounded([sys.executable, '-c', 'import time; time.sleep(30)'], 0.1, report)
        state = json.loads(report.read_text(encoding='utf-8'))
        self.assertEqual(code, 124)
        self.assertLess(time.monotonic() - started, 5)
        self.assertEqual(state['last_phase'], 'mock_cwyw')
        self.assertFalse(state['refresh_verified'])

    def mock_word(self, output, changed=False):
        word, doc = types.SimpleNamespace(), types.SimpleNamespace()
        word.Hwnd, word.Visible, word.calls = 1, False, []
        word.Documents = types.SimpleNamespace(Open=lambda *args, **kwargs: doc)
        word.Run = lambda name: word.calls.append(name)
        word.Quit = lambda save: word.calls.append('quit')
        doc.SaveAs2 = lambda *args, **kwargs: None
        doc.Activate = lambda: None
        doc.Close = lambda save: word.calls.append('close')
        doc.Save = lambda: document(output, ' ADDIN EN.CITE ' + endnote_xml(doi='10.0000/b' if changed and len(word.calls) == 3 else '10.0000/a'))
        client = types.ModuleType('win32com.client')
        client.DispatchEx = lambda name: word
        com = types.ModuleType('win32com')
        com.client = client
        process = types.ModuleType('win32process')
        process.GetWindowThreadProcessId = lambda hwnd: (1, 123)
        return word, {'win32com': com, 'win32com.client': client, 'win32process': process}

    def test_worker_performs_two_official_refreshes_and_reports_preservation(self):
        document(self.source, ' CITATION Source1 ')
        args = argparse.Namespace(input=str(self.source), output=str(self.target), report=str(self.root / 'status.json'), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[]):
            plan = endnote.preflight(args)
        word, modules = self.mock_word(self.target)
        with patch.dict(sys.modules, modules):
            endnote.worker(args, plan)
        result = json.loads(Path(args.report).read_text(encoding='utf-8'))
        self.assertEqual(word.calls, ['EndNoteMacros.ENConvertWordCitations', 'EndNoteMacros.ENFormatBibliography', 'EndNoteMacros.ENFormatBibliography', 'close', 'quit'])
        self.assertTrue(result['refresh_verified'])
        self.assertIn('preservation_audit', result)

    def test_worker_rejects_metadata_that_changes_again_on_second_refresh(self):
        document(self.source, ' CITATION Source1 ')
        args = argparse.Namespace(input=str(self.source), output=str(self.target), report=str(self.root / 'status.json'), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[]):
            plan = endnote.preflight(args)
        word, modules = self.mock_word(self.target, changed=True)
        with patch.dict(sys.modules, modules), self.assertRaisesRegex(RuntimeError, 'second CWYW refresh'):
            endnote.worker(args, plan)
        self.assertFalse(json.loads(Path(args.report).read_text(encoding='utf-8'))['refresh_verified'])
        self.assertEqual(word.calls[-2:], ['close', 'quit'])

    def test_worker_rejects_a_group_that_loses_one_reference(self):
        document(self.source, ' CITATION Source1 \\m TestSource ')
        args = argparse.Namespace(input=str(self.source), output=str(self.target), report=str(self.root / 'status.json'), refresh_only=False, style=None, style_dir=[])
        with patch.object(endnote, 'style_directories', return_value=[]):
            plan = endnote.preflight(args)
        word, modules = self.mock_word(self.target)
        with patch.dict(sys.modules, modules), self.assertRaisesRegex(RuntimeError, 'within a citation group'):
            endnote.worker(args, plan)
        self.assertFalse(json.loads(Path(args.report).read_text(encoding='utf-8'))['refresh_verified'])


if __name__ == '__main__':
    unittest.main()
