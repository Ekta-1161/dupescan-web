import io
import hashlib
import uuid
import zipfile
import re
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1 GB

_sessions = {}


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════════════════════

def fmt_bytes(n):
    if n < 1024:    return f"{n} B"
    if n < 1024**2: return f"{n // 1024} KB"
    return f"{n / 1024**2:.1f} MB"

def hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def safe_sid(sid):
    return all(c in '0123456789abcdef-' for c in sid)

def normalise_text(text: str) -> str:
    """Lowercase, strip whitespace/punctuation for fuzzy-like comparison."""
    text = text.lower().strip()
    text = re.sub(r'\s+', ' ', text)          # collapse whitespace
    text = re.sub(r'[^\w\s]', '', text)        # remove punctuation
    return text.strip()

def content_hash(text: str) -> str:
    """Hash normalised text content — used for cross-file comparison."""
    return hashlib.sha256(normalise_text(text).encode('utf-8', errors='replace')).hexdigest()

def row_hash(row) -> str:
    """Hash a DataFrame row for cross-sheet/cross-file comparison."""
    raw = ' '.join(str(v) for v in row.fillna('').astype(str).tolist())
    return content_hash(raw)


# ═══════════════════════════════════════════════════════════════════════════════
# CONTENT EXTRACTORS
# Returns list of { 'label': str, 'text': str, 'hash': str }
# ═══════════════════════════════════════════════════════════════════════════════

def extract_pdf_chunks(file_bytes: bytes, filename: str) -> list:
    """Extract one chunk per PDF page."""
    reader = PdfReader(io.BytesIO(file_bytes))
    chunks = []
    for i, page in enumerate(reader.pages):
        text = (page.extract_text() or '').strip()
        if text:
            chunks.append({
                'source':    filename,
                'type':      'pdf',
                'label':     f"Page {i+1}",
                'text':      text,
                'hash':      content_hash(text),
                'page_idx':  i,
            })
    return chunks


def extract_excel_chunks(file_bytes: bytes, filename: str) -> list:
    """Extract one chunk per Excel/CSV row."""
    ext = Path(filename).suffix.lower()
    if ext == '.csv':
        sheets = {'Sheet1': pd.read_csv(io.BytesIO(file_bytes))}
    else:
        xf = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets = {s: xf.parse(s) for s in xf.sheet_names}

    chunks = []
    for sheet, df in sheets.items():
        for i, (_, row) in enumerate(df.iterrows()):
            raw = ' '.join(str(v) for v in row.fillna('').astype(str).tolist())
            if raw.strip():
                chunks.append({
                    'source':    filename,
                    'type':      'excel',
                    'sheet':     sheet,
                    'label':     f"Row {i+2}",       # +2: 1-based + header
                    'row_idx':   i,
                    'text':      raw,
                    'hash':      content_hash(raw),
                })
    return chunks


# ═══════════════════════════════════════════════════════════════════════════════
# 1. SINGLE PDF
# ═══════════════════════════════════════════════════════════════════════════════

def process_pdf(file_bytes: bytes, filename: str) -> dict:
    reader = PdfReader(io.BytesIO(file_bytes))
    total  = len(reader.pages)
    seen, kept, dup_details = {}, [], []

    for i, page in enumerate(reader.pages):
        h = content_hash(page.extract_text() or '')
        if h not in seen:
            seen[h] = i + 1
            kept.append(i)
        else:
            dup_details.append({'duplicate_page': i + 1, 'original_page': seen[h]})

    writer = PdfWriter()
    for i in kept:
        writer.add_page(reader.pages[i])
    buf = io.BytesIO()
    writer.write(buf)
    clean_bytes = buf.getvalue()

    return {
        'mode': 'pdf', 'filename': filename,
        'total_pages': total, 'kept_pages': len(kept),
        'duplicate_pages': len(dup_details), 'dup_details': dup_details,
        'original_size': len(file_bytes), 'clean_size': len(clean_bytes),
        'clean_bytes': clean_bytes,
        'download_name': Path(filename).stem + '_clean.pdf',
        'mime': 'application/pdf',
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 2. SINGLE EXCEL / CSV
# ═══════════════════════════════════════════════════════════════════════════════

def process_excel(file_bytes: bytes, filename: str) -> dict:
    ext = Path(filename).suffix.lower()
    if ext == '.csv':
        sheets_raw = {'Sheet1': pd.read_csv(io.BytesIO(file_bytes))}
    else:
        xf = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets_raw = {s: xf.parse(s) for s in xf.sheet_names}

    total_rows = sum(len(df) for df in sheets_raw.values())

    # Step 1: within-sheet
    sheets_out, after_within = [], {}
    for sheet, df in sheets_raw.items():
        mask     = df.duplicated(keep='first')
        dup_idx  = df[mask].index.tolist()
        clean_df = df[~mask].reset_index(drop=True)
        after_within[sheet] = clean_df
        within_details = []
        for ri in dup_idx[:100]:
            for oi in range(ri):
                if df.iloc[oi].equals(df.iloc[ri]):
                    within_details.append({'duplicate_row': int(ri)+2, 'original_row': int(oi)+2})
                    break
        sheets_out.append({
            'sheet': sheet, 'total_rows': len(df),
            'within_dup_rows': len(dup_idx), 'within_dup_details': within_details,
            'cross_dup_rows': 0, 'cross_dup_details': [], 'kept_rows': 0,
        })

    # Step 2: cross-sheet
    global_seen, cross_details, final_dfs = {}, [], {}
    for si in sheets_out:
        sheet, df, keep_idx = si['sheet'], after_within[si['sheet']], []
        for i, (_, row) in enumerate(df.iterrows()):
            h = row_hash(row)
            if h not in global_seen:
                global_seen[h] = {'sheet': sheet, 'row': i+2}
                keep_idx.append(i)
            else:
                orig = global_seen[h]
                d = {'duplicate_sheet': sheet, 'duplicate_row': i+2,
                     'original_sheet': orig['sheet'], 'original_row': orig['row']}
                cross_details.append(d)
                si['cross_dup_details'].append(d)
                si['cross_dup_rows'] += 1
        final_dfs[sheet] = df.iloc[keep_idx].reset_index(drop=True)
        si['kept_rows']      = len(final_dfs[sheet])
        si['duplicate_rows'] = si['within_dup_rows'] + si['cross_dup_rows']

    within_dup_total = sum(s['within_dup_rows'] for s in sheets_out)
    cross_dup_total  = sum(s['cross_dup_rows']  for s in sheets_out)

    out = io.BytesIO()
    if ext == '.csv':
        final_dfs['Sheet1'].to_csv(out, index=False)
        mime, dl_name = 'text/csv', Path(filename).stem + '_clean.csv'
    else:
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            for s, df in final_dfs.items():
                df.to_excel(w, sheet_name=s, index=False)
        mime    = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        dl_name = Path(filename).stem + '_clean' + ext
    clean_bytes = out.getvalue()

    return {
        'mode': 'excel', 'filename': filename,
        'total_rows': total_rows, 'kept_rows': sum(len(df) for df in final_dfs.values()),
        'duplicate_rows': within_dup_total + cross_dup_total,
        'within_dup_total': within_dup_total, 'cross_dup_total': cross_dup_total,
        'sheets': sheets_out, 'cross_details': cross_details[:200],
        'original_size': len(file_bytes), 'clean_size': len(clean_bytes),
        'clean_bytes': clean_bytes, 'download_name': dl_name, 'mime': mime,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 3. ZIP / FOLDER — duplicate files + cross-file-type content matching
# ═══════════════════════════════════════════════════════════════════════════════

EXCEL_EXTS = {'.xlsx', '.xls', '.xlsm', '.csv'}
PDF_EXTS   = {'.pdf'}

def process_zip(file_bytes: bytes, filename: str) -> dict:
    try:
        zin = zipfile.ZipFile(io.BytesIO(file_bytes))
    except zipfile.BadZipFile:
        return {'error': 'Invalid or corrupted ZIP file.'}

    entries = [e for e in zin.infolist() if not e.is_dir()]
    if not entries:
        return {'error': 'ZIP file is empty.'}

    # ── Pass 1: Exact file-level duplicates (binary hash) ─────────────────────
    file_hash_map = {}
    keep_paths    = set()
    dup_paths     = set()
    exact_dups    = []

    for entry in entries:
        data = zin.read(entry.filename)
        h    = hash_bytes(data)
        if h not in file_hash_map:
            file_hash_map[h] = entry.filename
            keep_paths.add(entry.filename)
        else:
            dup_paths.add(entry.filename)
            exact_dups.append({
                'duplicate_file': entry.filename,
                'original_file':  file_hash_map[h],
                'size':           entry.file_size,
                'match_type':     'exact_file',
            })

    # ── Pass 2: Content-level cross-file-type matching ────────────────────────
    # Extract text chunks from all PDF and Excel files
    all_chunks   = []   # list of chunk dicts
    file_chunks  = {}   # filename → list of chunk dicts

    for entry in entries:
        fname = entry.filename
        ext   = Path(fname).suffix.lower()
        data  = zin.read(fname)

        try:
            if ext in PDF_EXTS:
                chunks = extract_pdf_chunks(data, fname)
            elif ext in EXCEL_EXTS:
                chunks = extract_excel_chunks(data, fname)
            else:
                chunks = []
        except Exception:
            chunks = []

        file_chunks[fname] = chunks
        all_chunks.extend(chunks)

    # Build global content hash → first seen chunk
    content_seen     = {}   # hash → chunk
    cross_file_dups  = []   # matches across different files

    for chunk in all_chunks:
        h = chunk['hash']
        if h not in content_seen:
            content_seen[h] = chunk
        else:
            orig = content_seen[h]
            # Only flag if from DIFFERENT files
            if orig['source'] != chunk['source']:
                cross_file_dups.append({
                    'match_type':      'cross_file_content',
                    'dup_file':        chunk['source'],
                    'dup_label':       chunk['label'],
                    'dup_type':        chunk['type'],
                    'orig_file':       orig['source'],
                    'orig_label':      orig['label'],
                    'orig_type':       orig['type'],
                    'preview':         chunk['text'][:120].replace('\n', ' '),
                })

    # ── Pass 3: Build output ZIPs ─────────────────────────────────────────────
    clean_buf = io.BytesIO()
    with zipfile.ZipFile(clean_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for e in entries:
            if e.filename in keep_paths:
                zout.writestr(e, zin.read(e.filename))

    dup_buf = io.BytesIO()
    with zipfile.ZipFile(dup_buf, 'w', zipfile.ZIP_DEFLATED) as zdup:
        for e in entries:
            if e.filename in dup_paths:
                zdup.writestr(e, zin.read(e.filename))

    wasted = sum(d['size'] for d in exact_dups)

    return {
        'mode':               'zip',
        'filename':           filename,
        'total_files':        len(entries),
        'unique_files':       len(keep_paths),
        'duplicate_files':    len(dup_paths),
        'wasted_bytes':       wasted,
        'exact_dups':         exact_dups,
        'cross_file_dups':    cross_file_dups[:200],
        'cross_file_count':   len(cross_file_dups),
        'original_size':      len(file_bytes),
        'clean_size':         len(clean_buf.getvalue()),
        'clean_bytes':        clean_buf.getvalue(),
        'dup_bytes':          dup_buf.getvalue(),
        'download_name':      Path(filename).stem + '_clean.zip',
        'dup_download_name':  Path(filename).stem + '_duplicates.zip',
        'mime':               'application/zip',
    }


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/process', methods=['POST'])
def api_process():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400

    f        = request.files['file']
    filename = f.filename or 'file'
    ext      = Path(filename).suffix.lower().lstrip('.')
    data     = f.read()

    SUPPORTED = {
        'pdf':  process_pdf,
        'xlsx': process_excel, 'xls': process_excel,
        'xlsm': process_excel, 'csv': process_excel,
        'zip':  process_zip,
    }

    if ext not in SUPPORTED:
        return jsonify({'error': f'Unsupported file type ".{ext}". Use PDF, Excel, CSV, or ZIP.'}), 400

    try:
        result = SUPPORTED[ext](data, filename)
    except Exception as e:
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

    if 'error' in result:
        return jsonify({'error': result['error']}), 400

    sid = str(uuid.uuid4())
    _sessions[sid] = {
        'clean_bytes':       result['clean_bytes'],
        'dup_bytes':         result.get('dup_bytes', b''),
        'download_name':     result['download_name'],
        'dup_download_name': result.get('dup_download_name', ''),
        'mime':              result['mime'],
    }

    resp = {k: v for k, v in result.items() if k not in ('clean_bytes', 'dup_bytes')}
    resp['session_id']        = sid
    resp['original_size_fmt'] = fmt_bytes(result['original_size'])
    resp['clean_size_fmt']    = fmt_bytes(result['clean_size'])
    if 'wasted_bytes' in result:
        resp['wasted_fmt'] = fmt_bytes(result['wasted_bytes'])

    return jsonify(resp)


@app.route('/api/download/<sid>')
def api_download(sid):
    if not safe_sid(sid) or sid not in _sessions:
        return jsonify({'error': 'Session not found. Please re-upload.'}), 404
    s = _sessions[sid]
    return send_file(io.BytesIO(s['clean_bytes']), mimetype=s['mime'],
                     as_attachment=True, download_name=s['download_name'])


@app.route('/api/download-dups/<sid>')
def api_download_dups(sid):
    if not safe_sid(sid) or sid not in _sessions:
        return jsonify({'error': 'Session not found.'}), 404
    s = _sessions[sid]
    if not s.get('dup_bytes'):
        return jsonify({'error': 'No duplicate file available.'}), 404
    return send_file(io.BytesIO(s['dup_bytes']), mimetype='application/zip',
                     as_attachment=True, download_name=s.get('dup_download_name', 'duplicates.zip'))


if __name__ == '__main__':
    print("\n✅  DeDupScan running →  http://localhost:5000\n")
    app.run(debug=True, port=5000)
