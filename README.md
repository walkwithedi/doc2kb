# doc2kb: documents to a Markdown knowledge base

**Point it at a folder of office documents and get back Markdown an AI assistant can actually search:** one merged file with an index, or one file per document.

- **Zero dependencies.** Python 3.7+ standard library only: no pandoc, LibreOffice, python-docx, openpyxl or python-pptx. Copy it to an offline machine and it runs.
- **Keeps what matters.** Headings, tables, embedded workbooks and the position of every picture survive the conversion.
- **Files things once.** Identical copies, including ones inside zip and tar archives, are parsed once. Moved or renamed folders don't produce duplicates.
- **Set it and forget it.** An incremental mode for cron converts only what changed since the last run.
- **Several ways to use it.** Run it in the browser (double-click the launcher), from the command line, or import `doc2kb_core.py` as a library.

Supports `.docx` `.xlsx` `.xlsm` `.pptx` `.pdf` `.eml`, the legacy `.doc` `.xls` `.ppt`, plus archives. Messages and output labels can be English or Chinese.

[中文说明](README.zh-CN.md)

## Files

| File | What it is |
| --- | --- |
| `doc2kb_core.py` | the converter: parsing, extraction, picture export and output. Command line with options for everything, no interface code. |
| `doc2kb.sh` | entry script: without arguments it starts the browser interface; with arguments the incremental command-line mode for cron |
| `doc2kb_web.py`, `doc2kb.html` | browser interface: a local web service (127.0.0.1 only) and its page: folder batch, chosen files, incremental tasks, duplicate clean-up |
| `双击运行.desktop` | desktop launcher for the browser interface |
| `doc2kb_dedup.py` | removes sections from a merged or incremental file whose text is already in other output files, after a folder was moved or renamed |
| `SPEC.md` | the behaviour contract: inputs, outputs and edge cases, for rebuilding or understanding the design |

## Requirements

- Python 3.7 or newer, standard library only.
- Optional, each feature is skipped with a message when missing:
  - `pdftotext`, `pdfimages`, `pdfinfo` (poppler-utils): PDF
  - a web browser: the browser interface only
  - `montage` (ImageMagick): contact sheets only
- Run `doc2kb.sh` with bash, not sh.

Get it from [Releases](../../releases) (`doc2kb-<version>.zip`) or clone this repository.

## Language

Messages and the labels written into the output (headings, notes, table headers) are English or Chinese. `doc2kb.sh` follows `DOC2KB_LANG` (en or zh), else the locale: a `zh` locale gives Chinese. `doc2kb_core.py` takes `--lang en|zh` (default `DOC2KB_LANG`, else en). Chinese labels and file names stay fixed between versions, so existing output folders and fingerprint files keep working.

## Usage

### Browser

Double-click `双击运行.desktop`, or run `./doc2kb.sh` without arguments (or `python3 doc2kb_web.py`). The page opens in the browser with four tabs:

- **Folder batch**: choose the source folder (type the path or use "Browse…"), optionally a date, and press "Scan"; untick the files you do not want; check the output folder (default `<source>/_doc2kb`, `_知识库` in Chinese) and the options, then "Start". The result shows the summary and the report; "Open output folder" opens it.
- **Chosen files**: add documents or archives from one or more folders and start; they are always converted again (you are asked before md files of the same name are replaced).
- **Incremental tasks**: add, edit and delete tasks (name, source folder, output folder), preview them, run the ticked ones or all. A task that never ran takes the documents changed in the last N days (default 30, set on the page). Tasks live in `~/.config/doc2kb/watch_tasks.conf`, the same file earlier versions used.
- **Duplicate clean-up**: pick a merged or incremental md, run a dry run, check the report, then clean up (the original is backed up first).

`.msg` files are listed and skipped.

Large folders are converted in batches ("Files per batch", default 1000; 0 = one batch), so tens of thousands of documents run at a steady pace. In merged mode each batch becomes a volume, `<name>-01.md`, `<name>-02.md`, …, with `<name>-index.md` naming the volume of each document; duplicates are found across all batches. A cancelled merged run keeps its finished batches: start it again with the same folders and options and it continues with the first unfinished batch (or choose "Start again from the first batch"). Reloading the page shows the running or latest task again with its settings, progress and result.

Pictures are not exported unless "Export images" is ticked. The service listens on 127.0.0.1 only, with a one-time token in the address, and exits a few seconds after the last page is closed; "Quit" stops it at once. A run in progress is finished first. Paths are chosen in a folder list the service provides, not with the browser's file picker, because conversions need the real paths. To keep the launcher on the desktop, put a link to it there, not a copy. The service logs to `doc2kb_web.log` next to it (or `~/.cache/doc2kb/` when that folder is read-only).

### Incremental command line (cron)

```bash
# first run: set the watermark, take the last 30 days
./doc2kb.sh -s SOURCE -o OUTPUT -d 30

# every later run: only what changed since the last run
./doc2kb.sh -s SOURCE -o OUTPUT
```

cron example (every day at 18:30, absolute paths):

```
30 18 * * * /home/user/tools/doc2kb.sh -s /home/user/Documents -o /home/user/tools/output/wd >> /home/user/tools/output/wd/cron.out 2>&1
```

Options (each long option has a one-letter short form; type the short ones):

| Option | Effect |
| --- | --- |
| `-d N` | documents changed in the last N days (window) |
| `-t DATE` | from `YYYY-MM-DD` or `"YYYY-MM-DD HH:MM"` |
| `-i` / `-I` | export / do not export pictures (default: text only, faster) |
| `-E` | do not expand embedded workbooks |
| `-x SUFFIX` | output name suffix (default `incremental`, `增量` in Chinese) |
| `-n` | only list what would be processed (dry run) |
| `-h` | help |

Exit codes: 0 done (also when nothing was new), 1 option or environment error, 2 processing failed, 3 first run without a time range.

### The converter directly

```bash
python3 doc2kb_core.py --batch SOURCE -o OUTPUT --merge   # a folder, merged into one md
python3 doc2kb_core.py --batch SOURCE -o OUTPUT --merge --batch-size 500   # volumes of 500 files
python3 doc2kb_core.py --files list.txt -o OUTPUT          # the files named in a list
python3 doc2kb_core.py report.docx -o OUTPUT --lang zh     # one file, Chinese labels
python3 doc2kb_core.py -h                                  # every option
```

Other tools can import `doc2kb_core.py` and write output of their own through `build_parser()` and the `writer` argument of `run_batch()` / `run_files()`; SPEC.md, section 5.6, describes the interface.

Besides the human summary, batch runs print `STAT<TAB>key<TAB>value` lines (new, kept, failed, written, dup, moved, chars, images) `BATCH<TAB>k<TAB>n` at the start of each batch, and `FAILED<TAB>path` lines for scripts. SPEC.md, section 5.1a, describes batches, volumes and how an unfinished merged run continues (`--force` starts it again).

### Cleaning up duplicates after a folder move

When a subfolder is moved, renamed or copied, its files look new to the incremental mode. Their text is then recognised by the fingerprint file `.doc2kb_watch_fp` in the output folder and only their paths are listed. For an incremental file written before that file existed:

```bash
# 1. dry run: writes the report (target_dedup-report.txt) and changes nothing
python3 doc2kb_dedup.py OUTPUT/Documents_20261007_incremental.md -n

# 2. after checking the report: the original is backed up as .bak-<time>, and
#    .doc2kb_watch_fp is created so later incremental runs skip known text
python3 doc2kb_dedup.py OUTPUT/Documents_20261007_incremental.md
```

- By default the other md files in the target's folder are compared (earlier incremental files, merged full runs, one-file-per-document output); `-b FILE_OR_FOLDER` (repeatable) names others.
- Only sections whose text is identical are removed (metadata, heading levels, picture paths and whitespace are ignored); documents changed during the move stay.
- Removed sections appear in the index's "Documents moved, content unchanged" table with current path, old path and the file holding the text.
- Source documents and pictures are never touched; picture folders used by removed sections are listed in the report.
- If the file was already imported into a knowledge base, replace it there with the cleaned one.

## Formats

- Supported: `.docx`, `.xlsx`, `.xlsm`, `.pptx`, `.pdf`, `.eml`, and the legacy binary `.doc`, `.xls`, `.ppt` of Office 97 and later (text and tables, no pictures; password-protected and older files are listed as failed).
- `.msg` does not crash a run: it is listed under "need a new format first".
- Archives: `.zip`, `.tar`, `.tar.gz`/`.tgz`, `.tar.bz2`/`.tbz2`, `.tar.xz`/`.txz` are read like folders. The documents inside appear as `pack.zip/a/report.docx` in the scan, the index and the report, and work with `--only` and incremental tasks; an archive can also be named directly (`python3 doc2kb_core.py pack.zip`). Archives inside archives are opened up to 3 levels deep. Members are unpacked into a private temporary folder that is removed when the run ends; nothing is written next to the archive and member paths never leave it. Encrypted members, members over 512 MB unpacked, more than 2 GB of documents in one archive, damaged archives and `.7z`/`.rar` (no standard-library reader) are listed as failed with the reason. Zip names written by Chinese Windows without the UTF-8 flag are decoded as GBK.
- Duplicates, within one run: files with identical bytes (SHA-256, archive members included) are parsed once, preferring a copy whose output exists, then one outside any archive, then the first path; files whose extracted text is identical are then kept once (newest document time). Both kinds are listed in the report's "Duplicate files" section and in a note at the top of the kept copy. `-D` keeps every copy. Across runs, incremental tasks skip text already filed (see above).
- Merged output (index plus every text in one Markdown file) by default; one file per document as an option.
- Incremental mode keeps the time of the last successful run as its watermark; every successful run moves it. A file counts as changed by the latest of its file system modification time, change time and the document's own time, so old files newly copied in are found.

## Development

```bash
git config core.hooksPath .githooks       # once per clone: run every check before each commit
python3 scripts/run_checks.py             # tests, repository checks, personal-data scan
```

`doc2kb_core.py` is the master copy for projects that bundle the converter. Rules for contributions: [CONTRIBUTING.md](CONTRIBUTING.md); changes: [CHANGELOG.md](CHANGELOG.md). Licence: [MIT](LICENSE).
