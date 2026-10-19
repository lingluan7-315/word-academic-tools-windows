# -*- coding: utf-8 -*-
import argparse
import json
from pathlib import Path

import win32com.client


parser = argparse.ArgumentParser(description='Convert Word native citations with EndNote official CWYW, in a new file.')
parser.add_argument('--input', required=True)
parser.add_argument('--output', required=True)
args = parser.parse_args()
source, target = Path(args.input).resolve(), Path(args.output).resolve()
if target.exists() or source == target:
    raise RuntimeError('Choose a new output filename; input and existing files are never overwritten')
word = win32com.client.DispatchEx('Word.Application')
word.Visible = True
doc = word.Documents.Open(str(source))
try:
    if any('ZOTERO_ITEM' in field.Code.Text for field in doc.Fields):
        raise RuntimeError('Zotero fields require metadata mapping and citation reconstruction first; this command accepts Word native citations')
    doc.SaveAs2(str(target), FileFormat=12)
    doc.Activate()
    word.Run('EndNoteMacros.ENConvertWordCitations')
    word.Run('EndNoteMacros.ENFormatBibliography')
    codes = [field.Code.Text for field in doc.Fields]
    result = {'endnote_citations': sum('ADDIN EN.CITE ' in code for code in codes), 'bibliography_fields': sum('ADDIN EN.REFLIST' in code for code in codes), 'remaining_word_citations': sum(code.strip().startswith('CITATION ') for code in codes), 'output': str(target)}
    if not result['endnote_citations'] or result['remaining_word_citations']:
        raise RuntimeError('EndNote did not complete conversion: %s' % result)
    doc.Save()
    print(json.dumps(result, ensure_ascii=False, indent=2))
finally:
    doc.Close(False)
    word.Quit(False)
