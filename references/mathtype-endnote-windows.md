# MathType and EndNote Windows integration

Tested locally on 2026-10-17: MathType 7.4.4, EndNote X9 19.0.0.12062, Word 2019 64-bit, Zotero 10.0.6. Use 64-bit Python with `comtypes`, `pywin32`, and `lxml` (the tested system Python is `D:\python\python.exe`). Word COM needs host access outside the ordinary desktop sandbox. Never close or terminate the user's existing Word instance; each helper creates and closes its own instance. No Computer Use is needed.

## MathType

`python scripts/mathtype_word_windows.py --input source.docx --output new-mathtype.docx --pdf new-preview.pdf`

The helper replaces main-body OMML with markers in a temporary copy, uses Microsoft's installed `OMML2MML.XSL`, inserts MathType's `BlankEqn.doc` object, and writes Unicode MathML through the installed 64-bit `MT6.dll` functions `MTAPIConnect`, `MTSetEqnFromLangStr` with language 2, and `MTCloseOleObject` with save option 1. `comtypes` supplies the real IDispatch pointer to the SDK. It saves real `Equation.DSMT4` objects, preserves the input, refuses existing outputs, and uses no clipboard.

Three smoke-test formulas (fraction, squared term, sum with bounds/subscript) converted into three MathType OLE objects. Saved XML had zero OMML nodes. Word PDF export was visually inspected and showed the correct formulas. Do not trust API return codes alone: an earlier `MTXFormEqn` text route returned success while rendering MathML tags as text. The supplied helper uses the dedicated MathML setter instead. Word's in-memory OMath count was misleading during conversion; check the saved OOXML and rendered formula content.

Limitations: installed paths are specific to this machine; only main-body equations were tested. Header/footer, footnote and text-box equations, multiple equations in one math paragraph, equation numbering, matrices, elaborate accents, track changes and exact original font/spacing need separate validation. The helper rejects detected equations in other XML parts, text boxes, or multi-equation math paragraphs. Review a copy before any manuscript-wide use. The Word MathType WLL initialization route failed locally; a temporary WLL installation experiment was rolled back. No Word security settings or VBA trust settings were changed.

## EndNote

`python scripts/endnote_word_windows.py --input native-word-citations.docx --output new-endnote.docx`

The helper accepts Word native `CITATION` fields with embedded document bibliography sources. It calls the installed official CWYW macros `EndNoteMacros.ENConvertWordCitations` and `EndNoteMacros.ENFormatBibliography`, then saves EndNote-generated fields. A smoke test produced a real `ADDIN EN.CITE` field and `ADDIN EN.REFLIST` bibliography and refreshed them. Do not synthesize EndNote field XML or relabel static text as live citations.

Long/grouped citations can use nested `EN.CITE.DATA` fields. Their `Field.Data` is raw UTF-8 bytes carried in a COM UTF-16 string; decode with `data.encode('utf-16le', errors='surrogatepass').rstrip(b'\x00').decode('utf-8')` for read-only inspection. An empty display result on the DATA field does not mean the grouped reference data is missing.

Zotero-to-EndNote route was independently proved on a test copy: read the actual Zotero citationItems/itemData and grouping; create matching Word document bibliography sources and native citations; call the official EndNote conversion macro; refresh; inspect the resulting EndNote data. A single citation remained one reference, a grouped citation remained three, and the bibliography held three unique records. Existing EndNote library matching supplied richer records in this local test. Neither RIS export nor renaming field codes performed the conversion.

The EndNote helper intentionally rejects Zotero fields. The tested Zotero-to-Word mapper covered plain journal articles only and is not a general manuscript migration command. Before full migration, validate DOI/URLs, non-article types, Chinese and corporate author names, locators, prefixes/suffixes, author suppression, duplicate matching, bibliography ordering, style and a separately saved `.enl` library. A live traveling-library field is not proof that an independent EndNote library was created. Do not unlink the user's original Zotero citations; reconstruct only in a new copy after all locations and metadata have been mapped.
