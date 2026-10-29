# Changelog

Releases follow VERSION; merging a version change to main publishes v<version> (.github/workflows/release.yml).

## 2.2.0

- Folder batches run in batches of `--batch-size` files (default 1000; 0 = one batch). The report is written after each batch instead of after each document, and each batch's texts leave memory when it ends, so the time per document no longer grows with the size of the run.
- Merged mode with more than one batch writes one volume per batch, `<name>-01.md`, `<name>-02.md`, …, and `<name>-index.md` naming the volume of each document. Duplicates and the fingerprint file are handled across all batches; one batch still gives the single `<name>.md`.
- A cancelled merged run keeps its finished batches in `.doc2kb_work/` in the output folder and continues from them when started again: unchanged batches are kept, changed ones are done again in place, new files form new batches. `--force` starts again.
- Library: with more than one batch the `writer` is called once per volume.
- Browser interface: "Files per batch"; progress shows the batch; with every file ticked the list is not sent and requests may be up to 32 MB; long file lists are built in steps; in merged mode the choice reads "continue an unfinished run" or "start again from the first batch". Reloading the page shows the running or latest task again in its tab, with its settings, progress, cancel button and result.

## 2.1.0

- Reads the legacy binary Office formats of Office 97 and later: `.xls` (sheets, shared strings, numbers, dates, formula text), `.doc` (paragraphs, tables, field results) and `.ppt` (the text of every slide in order). Pictures are not exported from them. Password-protected files and files from before Office 97 are listed as failed with the reason. `.msg` is still listed under "need a new format first".
- Legacy Office files take their document time from the summary information.
- Library: `doc_authors(path)` returns the author and last-modified-by recorded in a document.

## 2.0.0

- The browser interface covers everything the dialogs did: **Chosen files**, **Incremental tasks** (add, edit, delete, preview, run ticked or all; the first run's number of days is set on the page; task names stay unique; changing a task's output folder is confirmed because the watermark stays in the old folder) and **Duplicate clean-up** (dry run with its report, then a confirmed clean-up with a backup).
- `doc2kb.sh` without arguments starts the browser interface. The zenity dialogs are removed and zenity is no longer needed. The command-line mode, its options and exit codes and the tasks file are unchanged.
- Cancelling stops the running program and every program it started.
- Values from requests are logged escaped.

## 1.2.0

- Browser interface: `doc2kb_web.py` serves `doc2kb.html` on 127.0.0.1 with a one-time token, started by the desktop launcher `双击运行.desktop`. It offers folder batch: choose folders in a folder list, scan, untick files, choose options, follow the progress, cancel, and read the result and report. Pictures are off by default; the output folder defaults to `<source>/_doc2kb` (`_知识库` in Chinese). Only the current run's report is shown, a cancel states what was written, and files that cannot be processed are listed while the rest are converted. The dialogs and the command line are unchanged.

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
