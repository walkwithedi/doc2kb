# Changelog

Releases follow VERSION; merging a version change to main publishes v<version> (.github/workflows/release.yml).

## 1.1.0

- Library use: `build_parser()` returns every command-line option, and `run_batch()` / `run_files()` take an optional `writer` that replaces the merged or per-document output while the date filter, duplicate handling, fingerprint file, archives and report run as usual. SPEC.md, section 5.6, describes the interface. The command line is unchanged.

## 1.0.0

First release as a project of its own.

- Converts Word (.docx), Excel (.xlsx, .xlsm), PowerPoint (.pptx), PDF and mail (.eml) into Markdown: one merged file with an index (default), or one file per document. Headings, tables, embedded workbooks and the position of every picture are kept.
- Reads zip and tar archives like folders, up to three levels deep, unpacking into a private temporary folder that is removed after the run.
- Files with identical bytes are parsed once and files with identical text are kept once; the report lists every duplicate.
- doc2kb.sh: desktop dialogs (folder batch, chosen files, incremental tasks) and an incremental command-line mode for cron, with a watermark and a text-fingerprint file so moved folders are not filed twice. doc2kb_dedup.py cleans duplicates out of output written before the fingerprint file existed.
- English and Chinese messages and output labels.
- Python 3 standard library only; poppler-utils for PDF, zenity for the dialogs and ImageMagick for contact sheets are optional.
