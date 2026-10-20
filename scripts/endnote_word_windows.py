# -*- coding: utf-8 -*-
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

from citation_audit import compare_docx, inspect_docx


def style_directories(extra):
    folders = [Path(p).resolve() for p in extra]
    if os.name == 'nt':
        import winreg
        # App Paths locates non-default EndNote installations as well.
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    with winreg.OpenKey(hive, r'Software\Microsoft\Windows\CurrentVersion\App Paths\EndNote.exe', 0, winreg.KEY_READ | view) as key:
                        folders.append(Path(winreg.QueryValue(key, '').strip('"')).parent / 'Styles')
                except FileNotFoundError:
                    pass
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as key:
                documents = Path(os.path.expandvars(winreg.QueryValueEx(key, 'Personal')[0]))
                folders.append(documents / 'EndNote' / 'Styles')
        except FileNotFoundError:
            pass
    return list(dict.fromkeys(p for p in folders if p.is_dir()))


def preflight(args):
    source = Path(args.input).resolve()
    audit = inspect_docx(source)
    issues, warnings = list(audit['errors']), []
    engines = {group['engine'] for group in audit['groups']}
    if 'zotero' in engines:
        issues.append('Zotero fields need an explicit metadata/group mapping and Word-native reconstruction first; see references/citation-migration.md')
    if args.refresh_only:
        if engines != {'endnote'} or audit['native_word_citations']:
            issues.append('--refresh-only requires existing EndNote fields without Zotero or native Word citations')
    elif not audit['native_word_citations'] or engines:
        issues.append('Conversion requires Word native CITATION fields only; use --refresh-only for existing EndNote fields')
    if any(p != 'word/document.xml' for p in audit['native_word_parts']) or any(g['part'] != 'word/document.xml' for g in audit['groups']):
        issues.append('Header/footer/footnote citations require a separate conversion and verification pass')
    style = args.style or audit['endnote_style']
    folders = style_directories(args.style_dir)
    matches = [p for folder in folders for p in folder.glob('*.ens') if p.stem.casefold() == Path(style).stem.casefold()] if style else []
    if style and not matches:
        issues.append('Expected .ens style was not found; check EndNote Preferences > Folder Locations and pass the actual directory with --style-dir')
    if args.style and not audit['endnote_style']:
        issues.append('--style requires an existing EN.Layout set by EndNote; initialize it on a working copy through CWYW before selecting a different style')
    if not style:
        warnings.append('No saved EN.Layout style: CWYW will use its current style. Verify that style on a disposable sample first.')
    return {'input': str(source), 'audit': audit, 'style': Path(style).stem if style else None, 'style_files': [str(p) for p in matches], 'style_search_directories': [str(p) for p in folders], 'warnings': warnings, 'issues': issues, 'ready': not issues}


def write_status(path, state):
    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=Path(path).parent, delete=False) as temporary:
        json.dump(state, temporary, ensure_ascii=False, indent=2)
    os.replace(temporary.name, path)


def worker(args, plan):
    import win32com.client
    import win32process
    state = {'phase': 'starting_word', 'output': args.output, 'refresh_verified': False, 'preservation_verified': False}
    write_status(args.report, state)
    word = win32com.client.DispatchEx('Word.Application')
    doc = None
    try:
        word.Visible = True
        state['owned_word_pid'] = win32process.GetWindowThreadProcessId(word.Hwnd)[1]
        state['phase'] = 'opening_source_read_only'
        write_status(args.report, state)
        doc = word.Documents.Open(plan['input'], ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(args.output, FileFormat=12)
        if args.style:
            layout = ET.fromstring(doc.Variables('EN.Layout').Value)
            layout.find('Style').text = plan['style']
            doc.Variables('EN.Layout').Value = ET.tostring(layout, encoding='unicode')
        doc.Activate()
        if not args.refresh_only:
            state['phase'] = 'converting_native_citations'
            write_status(args.report, state)
            word.Run('EndNoteMacros.ENConvertWordCitations')
        previous = None
        for index in range(2):
            state['phase'] = 'formatting_bibliography_%d' % (index + 1)
            write_status(args.report, state)
            doc.Activate()
            word.Run('EndNoteMacros.ENFormatBibliography')
            doc.Saved = False
            doc.Save()
            current = inspect_docx(args.output)
            expected = plan['audit']['group_count'] if args.refresh_only else plan['audit']['native_word_citations']
            if current['errors'] or current['group_count'] != expected or current['native_word_citations'] or not current['bibliographies'] or any(g['engine'] != 'endnote' for g in current['groups']):
                raise RuntimeError('Saved citation audit did not match the input group count or field type')
            memberships = [len(g['items']) for g in plan['audit']['groups']] if args.refresh_only else [len(g['tags']) for g in plan['audit']['native_word_groups']]
            if [len(g['items']) for g in current['groups']] != memberships:
                raise RuntimeError('CWYW changed the number of references within a citation group')
            if plan['style'] and current['endnote_style'] != plan['style']:
                raise RuntimeError('Saved EN.Layout does not name the expected style')
            if previous is not None and current != previous:
                raise RuntimeError('Saved metadata or bibliography changed on the second CWYW refresh')
            previous = current
        preservation = compare_docx(plan['input'], args.output)
        state.update({'phase': 'completed', 'refresh_verified': True, 'citation_groups': current['group_count'], 'bibliographies': len(current['bibliographies']), 'style': current['endnote_style'], 'preservation_audit': {key: preservation[key] for key in ('outside_citations_xml_equal', 'normalization_limitation', 'changed_parts', 'other_changed_parts_require_review')}})
        state['preservation_verified'] = preservation['outside_citations_xml_equal'] is True and not preservation['other_changed_parts_require_review']
        write_status(args.report, state)
    except Exception as error:
        state.update({'phase': 'failed', 'error': str(error)})
        write_status(args.report, state)
        raise
    finally:
        if doc is not None:
            doc.Close(False)
        word.Quit(False)


def run_bounded(command, timeout, report):
    process = subprocess.Popen(command)
    try:
        code = process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        # Terminate only our Python worker. A modal Word dialog can block COM cleanup.
        process.terminate()
        process.wait()
        state = json.loads(Path(report).read_text(encoding='utf-8')) if Path(report).exists() else {}
        state.update({'phase': 'timed_out', 'last_phase': state.get('phase'), 'refresh_verified': False, 'error': 'CWYW exceeded the timeout. Close the modal dialog and the identified working copy manually; do not retry until the cause is resolved. The owned Word instance may remain open.'})
        write_status(report, state)
        return 124
    if code:
        state = json.loads(Path(report).read_text(encoding='utf-8'))
        state.update({'phase': 'worker_failed', 'last_phase': state.get('phase'), 'worker_exit_code': code, 'refresh_verified': False})
        write_status(report, state)
    return code


def main():
    parser = argparse.ArgumentParser(description='Preflight or convert Word-native citations through official EndNote CWYW in a new file.')
    parser.add_argument('--input', required=True)
    parser.add_argument('--output')
    parser.add_argument('--check', action='store_true', help='Read-only DOCX/style preflight; does not launch Word or EndNote.')
    parser.add_argument('--refresh-only', action='store_true', help='Refresh an existing EndNote document into a new output; do not reconvert it.')
    parser.add_argument('--style', help='Existing style name to set in an EndNote-created EN.Layout.')
    parser.add_argument('--style-dir', action='append', default=[], help='Actual EndNote style search folder from Preferences; repeat as needed.')
    parser.add_argument('--timeout', type=int, default=180, help='Total worker limit in seconds, including both refreshes and cleanup.')
    parser.add_argument('--report', help='New UTF-8 JSON status filename; defaults to OUTPUT.endnote-status.json.')
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    plan = preflight(args)
    if args.check:
        print(json.dumps(plan, ensure_ascii=False, indent=2))
        return 0 if plan['ready'] else 1
    if not plan['ready']:
        raise RuntimeError(json.dumps(plan['issues'], ensure_ascii=False))
    if not args.output or args.timeout <= 0:
        parser.error('--output and a positive --timeout are required for conversion or refresh')
    args.output = str(Path(args.output).resolve())
    args.report = str(Path(args.report or str(Path(args.output).with_suffix('.endnote-status.json'))).resolve())
    if args.output == args.report or args.report == plan['input'] or Path(args.output).exists() or (not args.worker and Path(args.report).exists()):
        raise RuntimeError('Choose new output and report filenames; existing files are never overwritten')
    if args.worker:
        parent_status = json.loads(Path(args.report).read_text(encoding='utf-8'))
        if parent_status.get('phase') != 'preflight_passed' or parent_status.get('input') != plan['input'] or parent_status.get('output') != args.output:
            raise RuntimeError('Internal worker requires its parent-created status file')
        worker(args, plan)
        return 0
    write_status(args.report, {'phase': 'preflight_passed', 'input': plan['input'], 'output': args.output, 'refresh_verified': False, 'preservation_verified': False})
    command = [sys.executable, str(Path(__file__).resolve()), '--worker', '--input', plan['input'], '--output', args.output, '--report', args.report]
    for folder in args.style_dir:
        command.extend(['--style-dir', folder])
    if args.style:
        command.extend(['--style', args.style])
    if args.refresh_only:
        command.append('--refresh-only')
    code = run_bounded(command, args.timeout, args.report)
    print(Path(args.report).read_text(encoding='utf-8'))
    return code


if __name__ == '__main__':
    raise SystemExit(main())
