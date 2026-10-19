# -*- coding: utf-8 -*-
import argparse
import ctypes
import json
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
import comtypes.client


parser = argparse.ArgumentParser(description='Convert main-body Word OMML equations to real MathType 7 OLE objects in a new document.')
parser.add_argument('--input', required=True)
parser.add_argument('--output', required=True)
parser.add_argument('--pdf')
args = parser.parse_args()
source_path, output_path = Path(args.input).resolve(), Path(args.output).resolve()
if output_path.exists() or source_path == output_path:
    raise RuntimeError('Choose a new output filename; input and existing files are never overwritten')
if args.pdf and Path(args.pdf).resolve().exists():
    raise RuntimeError('Choose a new PDF filename')
math_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
word_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
transform = etree.XSLT(etree.parse(r'C:\Program Files\Microsoft Office\Root\Office16\OMML2MML.XSL'))
formulas = []
dll = ctypes.WinDLL(r'C:\Program Files (x86)\MathType\System\64\MT6.dll')
dll.MTAPIConnect.argtypes = [ctypes.c_short, ctypes.c_short]
dll.MTAPIConnect.restype = ctypes.c_int32
dll.MTSetEqnFromLangStr.argtypes = [ctypes.c_void_p, ctypes.c_short, ctypes.c_wchar_p, ctypes.c_int32]
dll.MTSetEqnFromLangStr.restype = ctypes.c_int32
dll.MTCloseOleObject.argtypes = [ctypes.c_uint32, ctypes.c_void_p]
dll.MTCloseOleObject.restype = ctypes.c_int32
with tempfile.TemporaryDirectory(prefix='mathtype-', dir=str(Path.cwd())) as temporary:
    stage = Path(temporary) / 'stage.docx'
    with zipfile.ZipFile(source_path) as source, zipfile.ZipFile(stage, 'w', zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename.startswith('word/') and info.filename.endswith('.xml'):
                xml = etree.fromstring(data)
                equations = xml.findall('.//{%s}oMath' % math_ns)
                if equations and info.filename != 'word/document.xml':
                    raise RuntimeError('Equations outside the main document need a separate conversion pass: %s' % info.filename)
                if info.filename == 'word/document.xml':
                    for index, equation in enumerate(equations):
                        if any(parent.tag == '{%s}txbxContent' % word_ns for parent in equation.iterancestors()):
                            raise RuntimeError('Text-box equations need a separate conversion pass')
                        node = equation.getparent() if equation.getparent().tag == '{%s}oMathPara' % math_ns else equation
                        if node.tag.endswith('oMathPara') and len(node.findall('{%s}oMath' % math_ns)) != 1:
                            raise RuntimeError('Multiple equations in one math paragraph need separate handling')
                        converted = transform(equation).getroot()
                        for element in converted.iter():
                            element.tag = etree.QName(element).localname
                        etree.cleanup_namespaces(converted)
                        converted.set('xmlns', 'http://www.w3.org/1998/Math/MathML')
                        text = etree.tostring(converted, encoding='unicode')
                        marker = '{{CODEX_MATH_%s}}' % index
                        formulas.append((marker, text))
                        run = etree.Element('{%s}r' % word_ns)
                        etree.SubElement(run, '{%s}t' % word_ns).text = marker
                        node.getparent().replace(node, run)
                    data = etree.tostring(xml, xml_declaration=True, encoding='UTF-8', standalone=True)
            target.writestr(info, data)
    status = dll.MTAPIConnect(0, 30)
    if status != 0:
        raise RuntimeError('MathType API connect failed: %s' % status)
    word = comtypes.client.CreateObject('Word.Application', dynamic=True)
    word.Visible = True
    doc = word.Documents.Open(str(stage))
    try:
        for marker, text in formulas[::-1]:
            rng = doc.Content
            if not rng.Find.Execute(marker):
                raise RuntimeError('Equation marker not found: %s' % marker)
            rng.Text = ''
            position = rng.Start
            rng.InsertFile(r'C:\Program Files (x86)\MathType\Office Support\BlankEqn.doc', 'MTBlankEqn')
            shapes = [doc.InlineShapes.Item(i) for i in range(1, doc.InlineShapes.Count + 1)]
            shape = min((shape for shape in shapes if shape.Range.Start >= position), key=lambda shape: shape.Range.Start)
            if shape.OLEFormat.ProgID != 'Equation.DSMT4':
                raise RuntimeError('Inserted object is not MathType')
            shape.OLEFormat.DoVerb(2)
            obj = shape.OLEFormat.Object
            pointer = ctypes.cast(obj._comobj, ctypes.c_void_p)
            try:
                status = dll.MTSetEqnFromLangStr(pointer, 2, text, len(text.encode('utf-16le')) // 2)
                if status != 0:
                    raise RuntimeError('MathML write failed: %s' % status)
                shape.Range.Fields.Update()
            finally:
                closed = dll.MTCloseOleObject(1, pointer)
            if closed != 0:
                raise RuntimeError('MathType object close/save failed: %s' % closed)
        doc.SaveAs2(str(output_path), 12)
        if args.pdf:
            pdf_path = Path(args.pdf).resolve()
            doc.ExportAsFixedFormat(str(pdf_path), 17)
    finally:
        doc.Close(False)
        word.Quit(False)
        dll.MTAPIDisconnect()
with zipfile.ZipFile(output_path) as saved:
    xml = etree.fromstring(saved.read('word/document.xml'))
    remaining = len(xml.findall('.//{%s}oMath' % math_ns))
    objects = xml.findall('.//{urn:schemas-microsoft-com:office:office}OLEObject')
    count = sum(obj.get('ProgID') == 'Equation.DSMT4' for obj in objects)
if remaining:
    raise RuntimeError('Saved document still contains native equations')
print(json.dumps({'converted': len(formulas), 'saved_mathtype_objects': count, 'remaining_omml': remaining, 'output': str(output_path)}, ensure_ascii=False, indent=2))
