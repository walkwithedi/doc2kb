"""Small Office, mail and PDF files built with the standard library, for the doc2kb tests.

Each builder writes one file and returns its path. The content is fictional and covers the
cases the converter must handle: headings by style, tables with merged and nested cells,
content controls, both image mechanisms, embedded workbooks (one referenced twice) and an
embedded document, dates in a workbook, blank rows, slides with and without text, and a
MIME mail with an encoded subject, an HTML-only body, an inline image and an attachment.
"""
import os
import struct
import zipfile
import zlib

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
CORE = ('<?xml version="1.0" encoding="UTF-8"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dcterms="http://purl.org/dc/terms/">'
        '<dcterms:modified>{}</dcterms:modified></cp:coreProperties>')


def png(width, height, shade=0):
    """A valid grey PNG of the given size."""
    raw = b"".join(b"\x00" + bytes([shade % 256]) * width for _ in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xffffffff)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _rels(items):
    body = "".join(f'<Relationship Id="{i}" Type="x" Target="{t}"/>' for i, t in items)
    return f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PR}">{body}</Relationships>'


def _zip(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
        for name, data in parts.items():
            z.writestr(name, data)
    return path


def xlsx_bytes(rows_by_sheet, dates=True):
    """A workbook: {sheet name: [[cell, ...], ...]}; ints stay numbers, the string 'DATE' becomes
    a date cell (serial 46024 = 2026-01-02), None leaves the cell out."""
    import io
    shared, index = [], {}
    sheets_xml = []
    for rows in rows_by_sheet.values():
        out = []
        for r, row in enumerate(rows, 1):
            cells = []
            for c, val in enumerate(row):
                ref = chr(ord("A") + c) + str(r)
                if val is None:
                    continue
                if val == "DATE":
                    cells.append(f'<c r="{ref}" s="1"><v>46024</v></c>')
                elif isinstance(val, int):
                    cells.append(f'<c r="{ref}"><v>{val}</v></c>')
                else:
                    if val not in index:
                        index[val] = len(shared)
                        shared.append(val)
                    cells.append(f'<c r="{ref}" t="s"><v>{index[val]}</v></c>')
            out.append(f'<row r="{r}">{"".join(cells)}</row>')
        sheets_xml.append(f'<worksheet xmlns="{S}"><sheetData>{"".join(out)}</sheetData></worksheet>')
    names = list(rows_by_sheet)
    wb = (f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets>'
          + "".join(f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(names, 1))
          + "</sheets></workbook>")
    sst = (f'<sst xmlns="{S}">' + "".join(f"<si><t>{s}</t></si>" for s in shared) + "</sst>")
    styles = (f'<styleSheet xmlns="{S}"><cellXfs><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>'
              if dates else f'<styleSheet xmlns="{S}"/>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/workbook.xml", wb)
        # absolute targets, as some exporters write them
        z.writestr("xl/_rels/workbook.xml.rels",
                   _rels([(f"rId{i}", f"/xl/worksheets/sheet{i}.xml") for i in range(1, len(names) + 1)]))
        z.writestr("xl/sharedStrings.xml", sst)
        z.writestr("xl/styles.xml", styles)
        for i, x in enumerate(sheets_xml, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", x)
    return buf.getvalue()


def _p(text, style=None):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{ppr}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _img_p(rid, vml=False):
    if vml:
        return f'<w:p><w:r><w:pict><v:shape><v:imagedata r:id="{rid}"/></v:shape></w:pict></w:r></w:p>'
    return f'<w:p><w:r><w:drawing><a:blip r:embed="{rid}"/></w:drawing></w:r></w:p>'


def _ole_p(rid):
    return f'<w:p><w:r><w:object><o:OLEObject r:id="{rid}"/></w:object></w:r></w:p>'


def _doc(body):
    return (f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:v="{V}" xmlns:o="{O}">'
            f"<w:body>{body}</w:body></w:document>")


STYLES = (f'<w:styles xmlns:w="{W}">'
          '<w:style w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
          '<w:style w:styleId="2"><w:name w:val="heading 2"/></w:style>'
          '<w:style w:styleId="a3"><w:name w:val="标题 3"/></w:style></w:styles>')


def docx(path, english=False, modified="2026-03-04T05:06:07Z", intro=None):
    """A report with headings, a merged-cell table with a nested table, a content control,
    one described and one undescribed picture (DrawingML and VML), the same workbook embedded
    twice, an embedded document and an old-format embedded object."""
    if english:
        t = dict(h1="Project Atlas report", h2="Scope", h3="Details", intro="This report covers the pilot.",
                 described="The menu is shown below.", plain="Owner list follows.", cover="Cover text",
                 a="Item", b="Value", nested="inner")
        t["intro"] = intro or t["intro"]
    else:
        t = dict(h1="示例项目报告", h2="范围", h3="细节", intro="本报告说明试点情况。",
                 described="菜单如下图所示。", plain="负责人名单如下。", cover="封面文字",
                 a="项目", b="数值", nested="内层")
    nested = (f"<w:tbl><w:tr><w:tc>{_p(t['nested'] + '1')}</w:tc><w:tc>{_p(t['nested'] + '2')}</w:tc></w:tr></w:tbl>")
    table = ("<w:tbl>"
             f"<w:tr><w:tc>{_p(t['a'])}</w:tc><w:tc>{_p(t['b'])}</w:tc><w:tc>{_p('C')}</w:tc></w:tr>"
             f'<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>{_p("x|y")}</w:tc><w:tc>{nested}</w:tc></w:tr>'
             "</w:tbl>")
    body = (_p(t["h1"], "Heading1") + _p(t["intro"])
            + f"<w:sdt><w:sdtContent>{_p(t['cover'])}</w:sdtContent></w:sdt>"
            + _p(t["h2"], "2") + _p(t["described"]) + _img_p("rId1")
            + _p(t["h3"], "a3") + table + _p(t["plain"]) + _img_p("rId2", vml=True)
            + _ole_p("rId3") + _ole_p("rId4") + _ole_p("rId5") + _ole_p("rId6"))
    book = xlsx_bytes({"Sheet1": [["k", "v"], ["a", 1], ["d", "DATE"]]})
    inner_doc = _inner_docx(t)
    parts = {
        "word/document.xml": _doc(body),
        "word/styles.xml": STYLES,
        "word/_rels/document.xml.rels": _rels([("rId1", "media/image1.png"), ("rId2", "media/image2.png"),
                                               ("rId3", "embeddings/book1.xlsx"), ("rId4", "embeddings/book2.xlsx"),
                                               ("rId5", "embeddings/inner.docx"), ("rId6", "embeddings/old.bin")]),
        "word/media/image1.png": png(300, 200, 10),
        "word/media/image2.png": png(320, 240, 20),
        "word/embeddings/book1.xlsx": book,
        "word/embeddings/book2.xlsx": book,
        "word/embeddings/inner.docx": inner_doc,
        "word/embeddings/old.bin": b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64,
        "docProps/core.xml": CORE.format(modified),
    }
    return _zip(path, parts)


def _inner_docx(t):
    import io
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", _doc(_p(t["h2"] + "!", "Heading1") + _p(t["intro"])))
        z.writestr("word/styles.xml", STYLES)
    return buf.getvalue()


def xlsx(path, english=False):
    """Two sheets; the first has a date cell, trailing empty columns and many blank rows."""
    name = "Data" if english else "数据"
    rows = [["Name" if english else "名称", "Date" if english else "日期", None], ["a", "DATE"], ["b", 7]]
    rows += [[None, None, None]] * 60 + [["c", 8]]
    parts = _unzip(xlsx_bytes({name: rows, "Two": [["x", "y"]]}))
    parts["docProps/core.xml"] = CORE.format("2026-02-01T00:00:00Z")
    return _zip(path, parts)


def _unzip(data):
    import io
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def pptx(path, english=False):
    """Slide 1: text and a picture; slide 2: shapes only, no text (a drawn chart)."""
    title = "Quarterly plan" if english else "季度计划"
    s1 = (f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:cSld><p:spTree>'
          f"<p:sp><p:txBody><a:p><a:r><a:t>{title}</a:t></a:r></a:p></p:txBody></p:sp>"
          '<p:pic><p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic>'
          "</p:spTree></p:cSld></p:sld>")
    s2 = (f'<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree><p:sp/><p:sp/></p:spTree></p:cSld></p:sld>')
    return _zip(path, {
        "ppt/slides/slide1.xml": s1, "ppt/slides/slide2.xml": s2,
        "ppt/slides/_rels/slide1.xml.rels": _rels([("rId1", "../media/image1.png")]),
        "ppt/media/image1.png": png(400, 300, 30),
        "docProps/core.xml": CORE.format("2026-01-05T08:00:00Z"),
    })


def eml(path, english=False):
    """HTML-only body, an inline picture, a tiny signature icon and an attachment."""
    import base64
    subject = "=?utf-8?B?" + base64.b64encode(("Weekly sync" if english else "周会纪要").encode()).decode() + "?="
    hello = "Hello team" if english else "大家好"
    html = f"<html><body><p>{hello}</p><p>Line two &amp; more</p><script>x()</script></body></html>"
    big = base64.b64encode(png(400, 300, 40)).decode()
    tiny = base64.b64encode(png(20, 20, 50)).decode()
    att = base64.b64encode(b"PK\x03\x04 not really a document").decode()
    text = (f"From: Alice <alice@example.com>\nTo: Bob <bob@example.com>\nSubject: {subject}\n"
            "Date: Tue, 03 Mar 2026 10:00:00 -0000\nMIME-Version: 1.0\n"
            'Content-Type: multipart/mixed; boundary="B"\n\n'
            '--B\nContent-Type: text/html; charset="utf-8"\n\n' + html + "\n"
            '--B\nContent-Type: image/png\nContent-Disposition: inline; filename="chart.png"\n'
            "Content-Transfer-Encoding: base64\n\n" + big + "\n"
            '--B\nContent-Type: image/png\nContent-Disposition: inline; filename="logo.png"\n'
            "Content-Transfer-Encoding: base64\n\n" + tiny + "\n"
            '--B\nContent-Type: application/octet-stream\nContent-Disposition: attachment; filename="plan.docx"\n'
            "Content-Transfer-Encoding: base64\n\n" + att + "\n--B--\n")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    return path


def pdf(path, lines):
    """A one-page PDF with a text layer (ASCII only, Helvetica)."""
    content = "BT /F1 12 Tf 72 720 Td 14 TL " + " ".join(
        "(" + ln.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)") + ") '" for ln in lines) + " ET"
    objs = ["<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> "
             "/Contents 5 0 R >>"),
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            f"<< /Length {len(content)} >>\nstream\n{content}\nendstream"]
    out, offsets = b"%PDF-1.4\n", []
    for i, o in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n{o}\nendobj\n".encode("latin-1")
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    out += "".join(f"{o:010d} 00000 n \n" for o in offsets).encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    with open(path, "wb") as f:
        f.write(out)
    return path


def build_all(folder, english=False):
    """Every fixture in one folder; returns {kind: path}."""
    os.makedirs(folder, exist_ok=True)
    tag = "en" if english else "zh"
    out = {
        "docx": docx(os.path.join(folder, f"report-{tag}.docx"), english),
        "xlsx": xlsx(os.path.join(folder, f"book-{tag}.xlsx"), english),
        "pptx": pptx(os.path.join(folder, f"slides-{tag}.pptx"), english),
        "eml": eml(os.path.join(folder, f"mail-{tag}.eml"), english),
        "pdf": pdf(os.path.join(folder, f"notes-{tag}.pdf"), [f"Project Atlas notes ({tag})"] + [f"Line {i} of the minutes" for i in range(30)]),
    }
    with open(os.path.join(folder, f"legacy-{tag}.msg"), "wb") as f:
        f.write(b"\xd0\xcf\x11\xe0")
    out["legacy"] = os.path.join(folder, f"legacy-{tag}.msg")
    return out


# ---------------------------------------------------------------- legacy binary Office (OLE2)

FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
SUMMARY_FMTID = bytes.fromhex("e0859ff2f94f6810ab9108002b27b3d9")


def cfb_bytes(streams):
    """An OLE2 compound file (512-byte sectors) holding the given {name: bytes} at the top level.
    Streams under 4096 bytes go to the mini stream, as Office writes them."""
    names = list(streams)
    mini, mini_start, big = b"", {}, []
    for n in names:
        data = streams[n]
        if len(data) < 4096:
            mini_start[n] = len(mini) // 64
            mini += data.ljust((len(data) + 63) // 64 * 64 or 64, b"\0")
        else:
            big.append(n)
    n_mini_sec = len(mini) // 64
    minifat = []
    for n in names:
        if n in mini_start:
            k = max(1, (len(streams[n]) + 63) // 64)
            s = mini_start[n]
            minifat += [s + i + 1 for i in range(k - 1)] + [ENDOFCHAIN]
    assert len(minifat) == n_mini_sec

    def sectors(data):
        return max(1, (len(data) + 511) // 512) if data else 0

    dir_bytes = 128 * (len(names) + 1)
    n_dir = (dir_bytes + 511) // 512
    minifat_raw = struct.pack("<%dI" % len(minifat), *minifat) if minifat else b""
    n_minifat = sectors(minifat_raw)
    n_ministream = sectors(mini)
    big_secs = [sectors(streams[n]) for n in big]
    others = n_dir + n_minifat + n_ministream + sum(big_secs)
    n_fat = 1
    while n_fat * 128 < others + n_fat:
        n_fat += 1
    fat, layout, nxt = [FATSECT] * n_fat, {}, n_fat

    def chain(count):
        nonlocal nxt
        start = nxt
        fat.extend([nxt + i + 1 for i in range(count - 1)] + [ENDOFCHAIN])
        nxt += count
        return start

    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat) if n_minifat else ENDOFCHAIN
    root_start = chain(n_ministream) if n_ministream else ENDOFCHAIN
    for n, k in zip(big, big_secs):
        layout[n] = chain(k)
    fat += [FREESECT] * (n_fat * 128 - len(fat))

    def entry(name, etype, start, size, child=NOSTREAM, right=NOSTREAM):
        raw = name.encode("utf-16-le")
        return (raw.ljust(64, b"\0") + struct.pack("<HBB", len(raw) + 2, etype, 1) +
                struct.pack("<III", NOSTREAM, right, child) + b"\0" * 36 +
                struct.pack("<III", start, size, 0))

    entries = [entry("Root Entry", 5, root_start, len(mini), child=1 if names else NOSTREAM)]
    for i, n in enumerate(names):
        start = layout.get(n, mini_start.get(n))
        entries.append(entry(n, 2, start, len(streams[n]), right=i + 2 if i + 1 < len(names) else NOSTREAM))
    dir_raw = b"".join(entries).ljust(n_dir * 512, b"\0")

    header = (bytes.fromhex("d0cf11e0a1b11ae1") + b"\0" * 16 +
              struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6 +
              struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, 4096, minifat_start, n_minifat, ENDOFCHAIN, 0))
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += struct.pack("<109I", *difat)
    body = struct.pack("<%dI" % len(fat), *fat) + dir_raw
    if n_minifat:
        body += minifat_raw.ljust(n_minifat * 512, b"\0")
    if n_ministream:
        body += mini.ljust(n_ministream * 512, b"\0")
    for n, k in zip(big, big_secs):
        body += streams[n].ljust(k * 512, b"\0")
    return header + body


def summary_info(author="", last_author="", saved=None, codepage=936):
    """A SummaryInformation property set: code page, author, last saved by, last saved time."""
    props = [(1, struct.pack("<Ihh", 2, codepage, 0))]
    enc = {65001: "utf-8"}.get(codepage, "cp%d" % codepage)
    for pid, text in ((4, author), (8, last_author)):
        if text:
            raw = text.encode(enc) + b"\0"
            props.append((pid, struct.pack("<II", 0x1E, len(raw)) + raw.ljust((len(raw) + 3) // 4 * 4, b"\0")))
    if saved is not None:
        props.append((13, struct.pack("<IQ", 0x40, saved)))
    head = 8 + 8 * len(props)
    offs, body = [], b""
    for pid, raw in props:
        offs.append((pid, head + len(body)))
        body += raw
    pset = struct.pack("<II", head + len(body), len(props)) + b"".join(struct.pack("<II", *o) for o in offs) + body
    return struct.pack("<HHI", 0xFFFE, 0, 0) + b"\0" * 16 + struct.pack("<I", 1) + SUMMARY_FMTID + struct.pack("<I", 48) + pset


def _rec(rt, body):
    return struct.pack("<HH", rt, len(body)) + body


def _xlstr(text, len_bytes=2):
    wide = any(ord(c) > 0xFF for c in text)
    raw = text.encode("utf-16-le" if wide else "latin-1")
    n = struct.pack("<H", len(text)) if len_bytes == 2 else bytes([len(text)])
    return n + bytes([1 if wide else 0]) + raw


def _sst_records(strings, limit):
    """SST and CONTINUE records with at most limit bytes each; strings may be cut, the
    continuation starting with a flag byte as Excel writes it."""
    segs = [bytearray(struct.pack("<II", len(strings), len(strings)))]
    for s in strings:
        wide = any(ord(c) > 0xFF for c in s)
        width = 2 if wide else 1
        if len(segs[-1]) + 3 > limit:
            segs.append(bytearray())
        segs[-1] += struct.pack("<HB", len(s), 1 if wide else 0)
        rest = s
        while rest:
            room = (limit - len(segs[-1])) // width
            if room <= 0:
                segs.append(bytearray([1 if wide else 0]))
                continue
            part, rest = rest[:room], rest[room:]
            segs[-1] += part.encode("utf-16-le" if wide else "latin-1")
    return _rec(0x00FC, bytes(segs[0])) + b"".join(_rec(0x003C, bytes(s)) for s in segs[1:])


def _rk(i):
    return ((i << 2) | 2) & 0xFFFFFFFF


def xls_bytes(encrypted=False, author="", last_author=""):
    """A workbook with two worksheets and a chart sheet: shared strings cut across CONTINUE
    records, numbers as NUMBER, RK and MULRK, dates by a built-in and a custom format, a formula
    with a text result, a boolean and an error value."""
    strings = ["Item", "Amount", "苹果和香蕉的长名称" * 3, "Riverside Clinic supplies list", "日期"]
    sst = _sst_records(strings, 40)
    bof_g = _rec(0x0809, struct.pack("<HHHHII", 0x0600, 0x0005, 0, 0, 0, 0))
    bof_s = _rec(0x0809, struct.pack("<HHHHII", 0x0600, 0x0010, 0, 0, 0, 0))
    fmt = _rec(0x041E, struct.pack("<H", 164) + _xlstr("yyyy\\-mm\\-dd"))
    xfs = b"".join(_rec(0x00E0, struct.pack("<HH", 0, f) + b"\0" * 16) for f in (0, 14, 164, 0))
    filepass = _rec(0x002F, b"\0" * 6) if encrypted else b""

    def cell(row, col, xf=0):
        return struct.pack("<HHH", row, col, xf)

    s1 = (bof_s +
          _rec(0x00FD, cell(0, 0) + struct.pack("<I", 0)) + _rec(0x00FD, cell(0, 1) + struct.pack("<I", 1)) +
          _rec(0x00FD, cell(0, 2) + struct.pack("<I", 4)) +
          _rec(0x00FD, cell(1, 0) + struct.pack("<I", 2)) + _rec(0x0203, cell(1, 1) + struct.pack("<d", 3.5)) +
          _rec(0x027E, cell(1, 2, 1) + struct.pack("<I", _rk(46085))) +
          _rec(0x00FD, cell(2, 0) + struct.pack("<I", 3)) +
          _rec(0x00BD, struct.pack("<HH", 2, 1) + struct.pack("<HI", 0, _rk(-12)) + struct.pack("<HI", 2, _rk(46086)) + struct.pack("<H", 2)) +
          _rec(0x0006, cell(3, 0) + bytes([0, 0, 0, 0, 0, 0, 0xFF, 0xFF]) + b"\0" * 6) + _rec(0x0207, _xlstr("合计")) +
          _rec(0x0205, cell(3, 1) + bytes([1, 0])) + _rec(0x0205, cell(3, 2) + bytes([0x07, 1])) +
          _rec(0x000A, b""))
    s2 = bof_s + _rec(0x0203, cell(0, 0) + struct.pack("<d", 1.25)) + _rec(0x000A, b"")

    def globals_(pos1, pos2):
        return (bof_g + filepass + fmt + xfs +
                _rec(0x0085, struct.pack("<IBB", pos1, 0, 0) + _xlstr("数据", 1)) +
                _rec(0x0085, struct.pack("<IBB", 0, 0, 2) + _xlstr("Chart1", 1)) +
                _rec(0x0085, struct.pack("<IBB", pos2, 0, 0) + _xlstr("Second", 1)) +
                sst + _rec(0x000A, b""))

    g = globals_(0, 0)
    wb = globals_(len(g), len(g) + len(s1)) + s1 + s2
    streams = {"Workbook": wb.ljust(4096, b"\0")}
    if author or last_author:
        streams["\x05SummaryInformation"] = summary_info(author, last_author)
    return cfb_bytes(streams)


def xls(path, **kw):
    with open(path, "wb") as f:
        f.write(xls_bytes(**kw))
    return path


def doc_bytes(encrypted=False, author="", saved=None):
    """A Word 97 document: a paragraph in a compressed piece, then a Unicode piece with a two-row
    table (one cell holds two paragraphs, one cell is empty), a field and a closing paragraph.
    The table rows are marked in the paragraph properties the way Word marks them."""
    p1 = "Intro paragraph\r"
    cells = ["c1 first\r", "c1 second\x07", "B1\x07", "\x07", "\x07",
             "\x07", "B2\x07", "C2\x07", "\x07"]
    tail = "Page \x13 PAGE \x147\x15 of the report\r"
    p2 = "".join(cells) + tail
    text = p1 + p2
    fc1, fc2 = 2048, 2560
    flags = 0x0200 | (0x0100 if encrypted else 0)
    fib = bytearray(struct.pack("<HHHHHH", 0xA5EC, 0x00C1, 0, 0x0409, 0, flags)).ljust(32, b"\0")
    fib += struct.pack("<H", 14) + b"\0" * 28
    lw = [0] * 22
    lw[3] = len(text)
    fib += struct.pack("<H", 22) + struct.pack("<22i", *lw)
    fclcb = [0] * (93 * 2)

    # paragraph property page: one run per paragraph of the second piece
    marks, pos = [], fc2
    for part in cells + [tail]:
        marks.append((pos, pos + len(part) * 2, part))
        pos += len(part) * 2
    fkp = bytearray(512)
    crun = len(marks)
    struct.pack_into("<%dI" % (crun + 1), fkp, 0, *([m[0] for m in marks] + [marks[-1][1]]))
    in_table = bytes([3, 0, 0]) + struct.pack("<H", 0x2416) + b"\x01"          # cb form, 5 bytes
    row_end = bytes([0, 4, 0, 0]) + struct.pack("<H", 0x2416) + b"\x01" + struct.pack("<H", 0x2417) + b"\x01"
    fkp[400:400 + len(in_table)] = in_table
    fkp[420:420 + len(row_end)] = row_end
    for r, (_, _, part) in enumerate(marks):
        if part is tail:
            off = 0
        elif r in (4, 8):
            off = 420 // 2
        else:
            off = 400 // 2
        fkp[(crun + 1) * 4 + r * 13] = off
    fkp[511] = crun
    page = 6
    wd = bytearray(fib).ljust(fc1, b"\0")
    wd += p1.encode("cp1252")
    wd = wd.ljust(fc2, b"\0") + p2.encode("utf-16-le")
    wd = wd.ljust(page * 512, b"\0") + fkp

    bte = struct.pack("<III", fc2, fc2 + len(p2) * 2, page)
    plc = struct.pack("<III", 0, len(p1), len(text)) + \
        struct.pack("<HIH", 0, fc1 * 2 | 0x40000000, 0) + struct.pack("<HIH", 0, fc2, 0)
    clx = b"\x01" + struct.pack("<h", 2) + b"\0\0" + b"\x02" + struct.pack("<I", len(plc)) + plc
    table = bte + clx
    fclcb[13 * 2], fclcb[13 * 2 + 1] = 0, len(bte)
    fclcb[33 * 2], fclcb[33 * 2 + 1] = len(bte), len(clx)
    wd[32 + 2 + 28 + 2 + 88:32 + 2 + 28 + 2 + 88 + 2 + 93 * 8] = struct.pack("<H", 93) + struct.pack("<186I", *fclcb)
    streams = {"WordDocument": bytes(wd), "1Table": table.ljust(4096, b"\0")}
    if author or saved is not None:
        streams["\x05SummaryInformation"] = summary_info(author, "", saved)
    return cfb_bytes(streams)


def doc(path, **kw):
    with open(path, "wb") as f:
        f.write(doc_bytes(**kw))
    return path


def _prec(rt, body, inst=0, container=False):
    return struct.pack("<HHI", (inst << 4) | (0xF if container else 0), rt, len(body)) + body


def ppt_bytes(encrypted=False):
    """A presentation with two slides in the document's slide list: slide 1 refers to its title
    in the slide list, slide 2 holds its text in the slide itself. An older copy of slide 2
    stays in the stream without being referenced, as after a fast save."""
    def chars(t):
        return _prec(0x0FA0, t.encode("utf-16-le"))

    slwt = _prec(0x0FF0, _prec(0x03F3, struct.pack("<IIIII", 2, 0, 0, 256, 0)) +
                 _prec(0x0F9F, struct.pack("<I", 0)) + chars("Project Atlas 概览\rSecond line") +
                 _prec(0x03F3, struct.pack("<IIIII", 3, 0, 0, 257, 0)), inst=0, container=True)
    docc = _prec(0x03E8, _prec(0x03E9, b"\0" * 40) + slwt, container=True)

    def slide(inner):
        return _prec(0x03EE, _prec(0xF002, _prec(0xF00D, inner, container=True), container=True), container=True)

    stale = slide(_prec(0x0FA8, b"Old version"))
    s1 = slide(_prec(0x0F9E, struct.pack("<i", 0)))
    s2 = slide(_prec(0x0FA8, b"Plain bytes text") + chars("中文第二页"))
    stream = bytearray(docc)
    offs = {1: 0}
    stream += stale
    offs[2] = len(stream)
    stream += s1
    offs[3] = len(stream)
    stream += s2
    pd_off = len(stream)
    pd = struct.pack("<I", 1 | (3 << 20)) + struct.pack("<III", offs[1], offs[2], offs[3])
    stream += _prec(0x1772, pd)
    ue_off = len(stream)
    stream += _prec(0x0FF5, struct.pack("<IHBBIIIII", 0, 0, 0, 3, 0, pd_off, 1, 4, 0))
    token = 0xF3D1C4DF if encrypted else 0xE391C05F
    cu = _prec(0x0FF6, struct.pack("<III", 0x14, token, ue_off) + b"\0" * 8)
    return cfb_bytes({"PowerPoint Document": bytes(stream).ljust(4096, b"\0"), "Current User": cu})


def ppt(path, **kw):
    with open(path, "wb") as f:
        f.write(ppt_bytes(**kw))
    return path
