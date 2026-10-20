# Word Academic Tools for Windows / Windows Word 学术工具

The skill and repository are now named `word-academic-tools-windows`.

本 skill 与仓库均已从 `zotero-word-citations-windows` 更名为 `word-academic-tools-windows`，整合 Zotero 动态引文、MathType 可编辑公式和 EndNote CWYW 转换。

## Features / 功能

- Insert live Zotero citation fields instead of static text. / 插入 Zotero 动态引文域，而不是静态文本。
- Group multiple references in one field through Zotero's citation dialogs. / 通过 Zotero 引文对话框将多篇文献合并到同一个引文域。
- Place citations immediately after the supporting claim. / 将引文放在对应论断之后的合适位置。
- Avoid closing Word instances opened manually by the user. / 避免关闭用户手动打开的 Word 进程。
- Provide a local bridge for reliable Zotero Word integration. / 提供用于稳定集成 Zotero 与 Word 的本地桥接。
- Convert supported native Word equations to editable MathType OLE objects in a new document. / 将支持的 Word 原生公式转换为新文档中的可编辑 MathType OLE 对象。
- Convert Word native citations through EndNote's official CWYW macros in a new document. / 通过 EndNote 官方 CWYW 宏将 Word 原生引文转换到新文档。
- Preflight saved citation fields and installed EndNote styles before opening Word; bound CWYW calls with a timeout. / 打开 Word 前检查引文与 EndNote 样式，为 CWYW 调用设置超时。
- Audit grouped references, DOI, citation options and DOCX preservation. / 只读核验引用组、DOI、引文选项及 DOCX 内容保持情况。

## Requirements / 使用要求

- Windows
- Microsoft Word desktop
- Zotero Desktop with the Word integration add-in enabled
- 64-bit Python 3.9 or newer with `pywin32`
- For MathType: MathType 7 with its 64-bit SDK DLL, `comtypes`, and `lxml`
- For EndNote: EndNote with its Word CWYW add-in enabled

Bridge 0.3.0 was tested with Zotero 10.0.6 and Word 2019 on Windows on 2026-10-17; its manifest permits Zotero 7.0 through 10.0.*. MathType and EndNote were tested with MathType 7.4.4 and EndNote X9. These are recorded local test results, not a new compatibility test for every installation.

Zotero 功能需要 Zotero Desktop 及其 Word 插件；MathType、EndNote 功能分别需要对应软件。MathType 脚本使用固定的 Office/MathType 安装路径，使用前请核对 [详细说明](references/mathtype-endnote-windows.md)。

## Install or upgrade / 安装或升级

Install into a folder matching the new skill name:

```powershell
git clone https://github.com/lingluan7-315/word-academic-tools-windows.git "$env:USERPROFILE\.codex\skills\word-academic-tools-windows"
python -m pip install pywin32
```

MathType also requires `python -m pip install comtypes lxml`.

升级时，将旧 `zotero-word-citations-windows` 目录及其 `.backup-*` 备份移到 `skills` 目录之外，再安装新版本。备份保留 `SKILL.md` 时仍可能被识别为 skill，不能只靠修改目录名停用。已有新版安装请在对应 Git 克隆目录更新，不要向非空目录再次克隆。

Invoke the skill as `$word-academic-tools-windows`. The workflows and ownership safeguards are in [SKILL.md](SKILL.md).

## Example / 示例

```powershell
python scripts/zotero_word_windows.py check
python scripts/zotero_word_windows.py insert --docx "E:\Temp\paper.docx" --item-keys KEY1 KEY2 KEY3 --style Myself --marker "{{CITATION_GROUP_1}}" --save --close
```

The `--item-keys` option inserts consecutive references into one Zotero citation field. Do not insert separate fields and join them manually with semicolons.

`--close` only closes a Word instance created by the helper. A Word instance already opened by the user is never force-closed.

If the bridge is missing, exit Zotero normally before running `python scripts/zotero_word_windows.py install-bridge --register`, then restart Zotero. Do not force-terminate Zotero to install the bridge. `Myself` is a custom installed CSL style; replace it with a style available on your machine.

```powershell
python scripts/mathtype_word_windows.py --input source.docx --output new-mathtype.docx --pdf new-preview.pdf
python scripts/endnote_word_windows.py --check --input native-word-citations.docx
python scripts/endnote_word_windows.py --input native-word-citations.docx --output new-endnote.docx
python scripts/citation_audit.py --input source.docx --compare new-endnote.docx --report comparison.json
```

MathType conversion supports main-body equations under the limits documented in [the companion reference](references/mathtype-endnote-windows.md). EndNote conversion accepts Word native `CITATION` fields; it rejects existing Zotero fields, which require metadata mapping and reconstruction first. Both helpers require new output filenames and preserve the source document.

MathType 支持范围及验证要求见详细说明。EndNote 脚本不能直接转换 Zotero 引文；RIS 导出也不能完成 Word 引文域转换。请先在副本中检查公式、引文和参考文献的结果。

## Citation migration / 引文迁移

Read [the migration workflow](references/citation-migration.md) for Zotero → EndNote mapping, custom bilingual formats, and strict preservation. Test language-specific punctuation and native Author updates on a small copy first. A preformatted Custom field does not regenerate when Author changes; EndNote X9's tested Author Lists settings also imposed limits on mixed Chinese/English separators. The skill reports those limits before bulk conversion.

EndNote preflight checks the saved EN.Layout style against discovered or explicitly supplied `--style-dir` folders; confirm custom folders in EndNote Preferences. `--refresh-only` refreshes existing EndNote fields into a new file. `--timeout 180` bounds the worker and writes a status report; a modal Word instance may need normal manual cleanup after timeout. No blanket process termination is used.

The read-only audit is conservative: unknown identities, unsupported field boundaries, or other changed DOCX parts need review and do not automatically pass. Metadata and layout review remain necessary. Reports can contain private citation data; do not commit manuscript files or audit reports to this repository.

Offline regression checks (synthetic documents only): `python -m unittest discover -s tests -v`.

## License / 许可证

MIT License
