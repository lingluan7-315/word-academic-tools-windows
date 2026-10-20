# MathType and EndNote Windows integration

Tested locally on 2026-10-17: MathType 7.4.4, EndNote X9 19.0.0.12062, Word 2019 64-bit, Zotero 10.0.6. Use 64-bit Python with `comtypes`, `pywin32`, and `lxml`. Word COM needs host access outside the ordinary desktop sandbox. Never close or terminate the user's existing Word instance. Helpers use their own Word instance; on a blocked EndNote macro, the timeout supervisor may leave that instance open for normal manual cleanup. No Computer Use is needed.

## MathType

`python scripts/mathtype_word_windows.py --input source.docx --output new-mathtype.docx --pdf new-preview.pdf`

The helper replaces main-body OMML with markers in a temporary copy, uses Microsoft's installed `OMML2MML.XSL`, inserts MathType's `BlankEqn.doc` object, and writes Unicode MathML through the installed 64-bit `MT6.dll` functions `MTAPIConnect`, `MTSetEqnFromLangStr` with language 2, and `MTCloseOleObject` with save option 1. `comtypes` supplies the real IDispatch pointer to the SDK. It saves real `Equation.DSMT4` objects, preserves the input, refuses existing outputs, and uses no clipboard.

Three smoke-test formulas (fraction, squared term, sum with bounds/subscript) converted into three MathType OLE objects. Saved XML had zero OMML nodes. Word PDF export was visually inspected and showed the correct formulas. Do not trust API return codes alone: an earlier `MTXFormEqn` text route returned success while rendering MathML tags as text. The supplied helper uses the dedicated MathML setter instead. Word's in-memory OMath count was misleading during conversion; check the saved OOXML and rendered formula content.

Limitations: installed paths are specific to this machine; only main-body equations were tested. Header/footer, footnote and text-box equations, multiple equations in one math paragraph, equation numbering, matrices, elaborate accents, track changes and exact original font/spacing need separate validation. The helper rejects detected equations in other XML parts, text boxes, or multi-equation math paragraphs. Review a copy before any manuscript-wide use. The Word MathType WLL initialization route failed locally; a temporary WLL installation experiment was rolled back. No Word security settings or VBA trust settings were changed.

## EndNote

Read [citation-migration.md](citation-migration.md) before EndNote conversion or Zotero → EndNote migration. It covers style discovery, bilingual Author limitations, library identity, timeout handling, and preservation checks.

`python scripts/endnote_word_windows.py --check --input native-word-citations.docx`

The helper accepts Word native `CITATION` fields with embedded document bibliography sources. The official CWYW macros generate the live fields. Zotero input is rejected in the read-only preflight; existing EndNote fields use `--refresh-only` into a new file. No raw EndNote field XML is synthesized.

Long/grouped EndNote citations may store their records in nested `EN.CITE.DATA` rather than inline field instructions. `citation_audit.py` decodes OOXML `w:fldData` from base64 to UTF-8. An empty display result on that nested field does not mean the group is empty. When inspecting COM `Field.Data` separately, recover its raw bytes from the COM UTF-16 representation before decoding UTF-8; do not apply that COM conversion to OOXML base64 data.
