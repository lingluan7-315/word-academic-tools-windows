# Citation migration and custom bilingual formats

Use this workflow for Zotero → EndNote, custom citation styles, or strict preservation of a manuscript. Lessons were checked against a Windows Word 2019 / EndNote X9 / Zotero 10.0.6 case on 2026-10-20. The case documents are private; the helpers contain no manuscript-specific records, keys, style names, or paths.

## Establish the result before bulk work

1. Audit the **saved** original without opening Word:

   `python scripts/citation_audit.py --input source.docx --report source-audit.json`

   Record the engine, locations, group membership, citation IDs/keys, DOI, locators, prefixes/suffixes and author suppression. The audit reads main body, headers, footers, footnotes and endnotes, including split instructions and nested EndNote `EN.CITE.DATA`. An undecodable field is a blocker, not an empty citation. The current converter handles main-body Word native `CITATION` fields only.

2. Turn the requested example into concrete style requirements: numeric/author-year, author name order, truncation, Chinese and English separators, translated author/title/journal, punctuation, fonts and bibliography indentation. Keep bibliographic metadata separate from how a style renders it.

3. Test one English record, one Chinese record with more authors than the truncation threshold, and one grouped citation **before** whole-document conversion. Use the installed processor and official Word plugin. In disposable records, change native `Author`, refresh, and check the saved bibliography. Restore/discard the test records afterward. A preview or correct field count alone does not prove Word refresh or author updates work.

EndNote X9 in the observed case applied the style's Author Lists punctuation/truncation rules globally. It did not produce both Chinese `，等` and English `, et al.` from native Author under that configuration. Treat this as a tested version/style limitation, not a claim about every EndNote version. Establish feasibility early. Do not spend repeated full-document runs trying to satisfy conflicting output requirements.

An author string cached in `Custom 2` can reproduce an example while remaining a **formatted snapshot**: changing native Author will not regenerate it. Keep structured native Author complete even if a user chooses a Custom-based presentation; disclose exactly which fields require separate maintenance before using that approach. `Translated Author` is also separately maintained, not an automatic translation. Never silently replace native Author with a preformatted string or call a post-refresh text repair “native automatic updating.” If the requested combination fails the sample, report the demonstrated limitation and let the user choose which requirement to change; keep the existing deliverable intact.

## EndNote preflight and conversion

Run preflight before any CWYW macro:

```powershell
python scripts/endnote_word_windows.py --check --input native-word-citations.docx --style-dir "D:\Apps\EndNote\Styles"
python scripts/endnote_word_windows.py --input native-word-citations.docx --output new-endnote.docx --style-dir "D:\Apps\EndNote\Styles" --timeout 180
```

Use the actual directory shown by **EndNote Preferences → Folder Locations → Style Folder**, together with the installed EndNote Styles directory. The helper discovers App Paths and the user Documents folder as candidates, but an arbitrary existing directory is not proof EndNote searches it. `--style-dir` can be repeated. Verify the style can be selected in EndNote/Word; a file copied only to a guessed Documents directory may be invisible to CWYW. Keep the `.ens` internal style name, filename and saved `EN.Layout/Style` consistent.

`--check` reads the DOCX and searches for its expected style without launching Office. It rejects Zotero/mixed/unsupported-location input and missing expected `.ens` files. If no EN.Layout style is saved, it warns that CWYW will use its current style; first validate that default on a disposable sample. `--style NAME` changes only the style entry in an **existing EndNote-created** EN.Layout. It refuses to fabricate a missing layout. Initialize that layout through CWYW on a working copy when needed.

For an existing EndNote document, use `--refresh-only` with a **new output**; do not reconvert it through native Word citations. Conversion uses `ENConvertWordCitations` followed by two `ENFormatBibliography` calls. It checks the saved group count/type and metadata/bibliography stability after the second refresh. A separate preservation and visual review is still required; `refresh_verified` and `preservation_verified` are distinct report fields.

The parent process enforces a wall-clock limit including COM calls and cleanup. A timeout returns 124 and writes `*.endnote-status.json` with the last stage and, if available, the script-owned Word PID. It stops only its Python worker; the Word instance may remain open with a modal dialog. Inspect that working copy, dismiss the reported dialog, and close that copy normally. Do not start another worker against a blocked session, repeatedly extend the timeout, or terminate all WINWORD/EndNote processes. Only inspect names/paths of unrelated open documents, never their text.

## Zotero → EndNote mapping

The bundled converter **does not** perform general Zotero → Word-native mapping. Its preflight rejects Zotero immediately instead of entering a long Word session. A request for migration authorizes working-copy preparation; it does not make an unimplemented mapper safe to improvise on the original.

For a scoped mapper, first extract the original `citationItems`/`itemData`, location, group membership and citation options. Export/import metadata through supported EndNote mechanisms into an independent library, record the resulting record IDs/library identity, and reconstruct Word-native sources/citations before the official CWYW conversion. Validate item types, corporate/Chinese names, DOI/URLs, years, translators, page locators, prefixes/suffixes and suppression explicitly. Reject unsupported mappings instead of dropping options. Do not synthesize EndNote field XML or merely rename Zotero field codes. RIS/XML/ENW files are exchange formats, not a working `.enl` library or converted Word citations.

Keep an explicit old-key → new-key/record mapping with evidence. A Zotero 404 may mean a merged/deleted item: resolve `dc:replaces` relations and verify the record; if no unique successor exists, stop that item. Never insert an arbitrary fallback key. Titles and field order alone are insufficient identity checks; distinguish same-title records by stable keys, DOI and metadata. Zotero may replace a citation ID during editing, so track its before/after identity and document location as well.

For a separately delivered EndNote library, confirm it opens in EndNote and keep both `name.enl` and the companion `name.Data` directory; close the library normally before packaging. A traveling library embedded in Word does not prove the external library is independent. Do not clone a library and assume the clone has a new internal identity; verify record and library IDs before blaming matching failures on multiple open windows.

## Zotero styles and bilingual data

Use language-specific CSL rendering with structured native author data. The observed Zotero 10 case used structured `Original Author: family || given` and `Original Title:` Extra values for the English counterpart, plus the appropriate short journal field. Verify these mappings with Zotero's own CSL output; they are a convention for the selected style, not universal translations. Test native and translated author updates separately. Do not hardcode a rendered author list into CSL to make one manuscript look correct.

Read-only local API access does not imply write access. Prefer supported Zotero APIs/connector operations for metadata changes, back up the affected records and use bounded requests. Do not edit Zotero's database directly. If an explicitly authorized one-time extension job is unavoidable, normally exit Zotero before replacing its XPI, keep an original-package hash, run the job once, verify its result and restore the original extension after normal exit. A file lock is a reason to wait for normal shutdown, not to force overwrite or repeatedly try UI tools.

For an existing Word document, refresh via `Project.Zotero.ZoteroRefresh`. A bridge `/style` or `/refresh` endpoint that works in a preview can still fail the actual Word integration. Avoid repeated speculative bridge calls; use the official Word plugin and confirm the saved document's preference style ID. Do not synthesize citation field XML. Set Bibliography paragraph/linked character styles for persistent layout rather than reapplying fonts to a field result after every refresh. Change only the bibliography-specific styles, and check they are not also used by body paragraphs.

## Preservation and acceptance

```powershell
python scripts/citation_audit.py --input source.docx --compare new-endnote.docx --report comparison.json
```

The comparator reports groups and item order, stable Zotero IDs, same-engine EndNote record IDs plus embedded-record fingerprints, cross-engine DOI, saved Zotero citation options, non-citation main-document XML, and every changed ZIP part. It does **not** equate titles with identities. Every cross-engine comparison requires an explicit mapping audit even when DOI matches; missing DOI is reported separately. Native Word metadata mapping also requires separate verification. Cross-engine locators and prefixes/suffixes must be checked against that mapping; the comparator does not prove they survived.

For supported field boundaries it masks only citation/bibliography ranges and ignores Word session IDs. Shared runs with outside text, a bibliography sharing a paragraph with outside content, or citations spanning paragraphs produce an explicit limitation. Such cases do not pass automatically. Changed styles, settings, custom properties, media, headers or embedded objects require individual review; the tool never silently whitelists them. Unsupported/ambiguous layouts need a separate scoped comparison, not a looser “success” test.

Work on a copy. If Word rewrites unrelated document structure, transfer only verified citation/bibliography ranges and required integration metadata back into a copy of the original after checking boundaries and relationships. Do not transplant generated fields purely by occurrence order. Compare non-citation XML, image/OLE relationships, tables, equations and all other ZIP parts. Preserve style definitions used by body content. Re-render both copies to confirm layout; counts alone are insufficient. Keep conversion, metadata and preservation reports in the task workspace, outside public skill sources.

Run the official refresh twice, then reopen the saved deliverable and refresh a disposable verification copy again. Check group membership/options, bibliography entries, style identity, author update behavior, formatting, and non-citation preservation. Describe unresolved differences honestly; do not overwrite the original before these checks pass. Respect the user's stop instruction even if a background worker finishes later.
