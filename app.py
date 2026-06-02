import io
import hashlib
import uuid
import zipfile
import re
import json
import time
import threading
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context
import pandas as pd
from pypdf import PdfReader, PdfWriter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024  # 1 GB

_sessions  = {}   # sid → session data
_progress  = {}   # sid → { pct, label, done, error }


# ═══════════════════════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════════════════════

def fmt_bytes(n):
    if n < 1024:    return f"{n} B"
    if n < 1024**2: return f"{n//1024} KB"
    return f"{n/1024**2:.1f} MB"

def hash_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def safe_sid(sid: str) -> bool:
    return all(c in '0123456789abcdef-' for c in sid)

def normalise_text(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'[^\w\s]', '', text)
    return text.strip()

def content_hash(text: str) -> str:
    return hashlib.sha256(
        normalise_text(text).encode('utf-8', errors='replace')
    ).hexdigest()

def row_hash(row) -> str:
    raw = ' '.join(str(v) for v in row.fillna('').astype(str).tolist())
    return content_hash(raw)

def set_progress(sid, pct, label, done=False, error=None):
    _progress[sid] = {
        'pct':   min(int(pct), 99 if not done else 100),
        'label': label,
        'done':  done,
        'error': error,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# CONTENT EXTRACTORS
# ═══════════════════════════════════════════════════════════════════════════════

EXCEL_EXTS = {'.xlsx', '.xls', '.xlsm', '.csv'}
PDF_EXTS   = {'.pdf'}

def extract_pdf_chunks(file_bytes: bytes, filename: str) -> list:
    chunks = []
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        for i, page in enumerate(reader.pages):
            text = (page.extract_text() or '').strip()
            if text:
                chunks.append({
                    'source':   filename,
                    'type':     'pdf',
                    'label':    f"Page {i+1}",
                    'text':     text,
                    'hash':     content_hash(text),
                    'page_idx': i,
                })
    except Exception:
        pass
    return chunks

def extract_excel_chunks(file_bytes: bytes, filename: str) -> list:
    ext = Path(filename).suffix.lower()
    chunks = []
    try:
        if ext == '.csv':
            sheets = {'Sheet1': pd.read_csv(io.BytesIO(file_bytes))}
        else:
            xf = pd.ExcelFile(io.BytesIO(file_bytes))
            sheets = {s: xf.parse(s) for s in xf.sheet_names}
        for sheet, df in sheets.items():
            for i, (_, row) in enumerate(df.iterrows()):
                raw = ' '.join(str(v) for v in row.fillna('').astype(str).tolist())
                if raw.strip():
                    chunks.append({
                        'source':  filename,
                        'type':    'excel',
                        'sheet':   sheet,
                        'label':   f"Row {i+2}",
                        'row_idx': i,
                        'text':    raw,
                        'hash':    content_hash(raw),
                    })
    except Exception:
        pass
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
            dup_details.append({'duplicate_page': i+1, 'original_page': seen[h]})

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

    total_rows   = sum(len(df) for df in sheets_raw.values())
    sheets_out   = []
    after_within = {}

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
        mime    = 'text/csv'
        dl_name = Path(filename).stem + '_clean.csv'
    else:
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            for s, df in final_dfs.items():
                df.to_excel(w, sheet_name=s, index=False)
        mime    = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        dl_name = Path(filename).stem + '_clean' + ext
    clean_bytes = out.getvalue()

    return {
        'mode': 'excel', 'filename': filename,
        'total_rows': total_rows,
        'kept_rows': sum(len(df) for df in final_dfs.values()),
        'duplicate_rows': within_dup_total + cross_dup_total,
        'within_dup_total': within_dup_total, 'cross_dup_total': cross_dup_total,
        'sheets': sheets_out, 'cross_details': cross_details[:200],
        'original_size': len(file_bytes), 'clean_size': len(clean_bytes),
        'clean_bytes': clean_bytes, 'download_name': dl_name, 'mime': mime,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 3. ZIP — background worker (runs in a thread, reports progress via SSE)
# ═══════════════════════════════════════════════════════════════════════════════

def process_zip_worker(sid: str, file_bytes: bytes, filename: str):
    try:
        set_progress(sid, 2, 'Opening ZIP…')

        try:
            zin = zipfile.ZipFile(io.BytesIO(file_bytes))
        except zipfile.BadZipFile:
            set_progress(sid, 0, '', done=True, error='Invalid or corrupted ZIP file.')
            return

        entries = [e for e in zin.infolist() if not e.is_dir()]
        total   = len(entries)
        if total == 0:
            set_progress(sid, 0, '', done=True, error='ZIP file is empty.')
            return

        set_progress(sid, 5, f'Found {total} files — starting scan…')

        # ── Pass 1: exact binary duplicates ──────────────────────────────────
        file_hash_map = {}
        keep_paths    = set()
        dup_paths     = set()
        exact_dups    = []
        file_data     = {}

        for i, entry in enumerate(entries):
            pct   = 5 + int((i / total) * 35)
            fname = Path(entry.filename).name
            set_progress(sid, pct, f'Pass 1/3 — Hashing {i+1}/{total}: {fname}')

            data = zin.read(entry.filename)
            file_data[entry.filename] = data
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
                })

        set_progress(sid, 40,
            f'Pass 1 done — {len(exact_dups)} exact duplicate file(s). Scanning content…')

        # ── Pass 2: cross-file content matching ───────────────────────────────
        processable = [
            e for e in entries
            if Path(e.filename).suffix.lower() in (PDF_EXTS | EXCEL_EXTS)
        ]
        proc_total      = len(processable)
        content_seen    = {}
        cross_file_dups = []

        for i, entry in enumerate(processable):
            fname = entry.filename
            ext   = Path(fname).suffix.lower()
            pct   = 40 + int((i / max(proc_total, 1)) * 40)
            set_progress(sid, pct,
                f'Pass 2/3 — Reading content {i+1}/{proc_total}: {Path(fname).name}')

            try:
                if ext in PDF_EXTS:
                    chunks = extract_pdf_chunks(file_data[fname], fname)
                else:
                    chunks = extract_excel_chunks(file_data[fname], fname)
            except Exception:
                chunks = []

            for chunk in chunks:
                h = chunk['hash']
                if h not in content_seen:
                    content_seen[h] = chunk
                else:
                    orig = content_seen[h]
                    if orig['source'] != chunk['source']:
                        cross_file_dups.append({
                            'dup_file':   chunk['source'],
                            'dup_label':  chunk['label'],
                            'dup_type':   chunk['type'],
                            'orig_file':  orig['source'],
                            'orig_label': orig['label'],
                            'orig_type':  orig['type'],
                            'preview':    chunk['text'][:120].replace('\n', ' '),
                        })

        set_progress(sid, 82,
            f'Pass 2 done — {len(cross_file_dups)} cross-file match(es). Building output ZIPs…')

        # ── Pass 3: build output ZIPs ─────────────────────────────────────────
        clean_buf = io.BytesIO()
        with zipfile.ZipFile(clean_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for i, entry in enumerate(entries):
                set_progress(sid, 82 + int((i / total) * 15),
                    f'Pass 3/3 — Writing clean ZIP {i+1}/{total}…')
                if entry.filename in keep_paths:
                    zout.writestr(entry, file_data[entry.filename])

        dup_buf = io.BytesIO()
        if dup_paths:
            with zipfile.ZipFile(dup_buf, 'w', zipfile.ZIP_DEFLATED) as zdup:
                for entry in entries:
                    if entry.filename in dup_paths:
                        zdup.writestr(entry, file_data[entry.filename])

        wasted = sum(d['size'] for d in exact_dups)
        clean_bytes = clean_buf.getvalue()
        dup_bytes   = dup_buf.getvalue()

        result_json = {
            'mode':              'zip',
            'filename':          filename,
            'total_files':       total,
            'unique_files':      len(keep_paths),
            'duplicate_files':   len(dup_paths),
            'wasted_bytes':      wasted,
            'wasted_fmt':        fmt_bytes(wasted),
            'exact_dups':        exact_dups,
            'cross_file_dups':   cross_file_dups[:300],
            'cross_file_count':  len(cross_file_dups),
            'original_size':     len(file_bytes),
            'clean_size':        len(clean_bytes),
            'original_size_fmt': fmt_bytes(len(file_bytes)),
            'clean_size_fmt':    fmt_bytes(len(clean_bytes)),
            'session_id':        sid,
            'download_name':     Path(filename).stem + '_clean.zip',
            'dup_download_name': Path(filename).stem + '_duplicates.zip',
        }

        _sessions[sid] = {
            'clean_bytes':       clean_bytes,
            'dup_bytes':         dup_bytes,
            'download_name':     result_json['download_name'],
            'dup_download_name': result_json['dup_download_name'],
            'mime':              'application/zip',
            'result_json':       result_json,
        }

        set_progress(sid, 100, 'Done!', done=True)

    except Exception as e:
        set_progress(sid, 0, '', done=True, error=f'Processing failed: {str(e)}')


# ═══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template('index.html')


# ── PDF / Excel / CSV — synchronous ──────────────────────────────────────────
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
    }
    if ext not in SUPPORTED:
        return jsonify({'error': f'Unsupported type ".{ext}". Use ZIP mode for zip files.'}), 400

    try:
        result = SUPPORTED[ext](data, filename)
    except Exception as e:
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500

    sid = str(uuid.uuid4())
    _sessions[sid] = {
        'clean_bytes': result['clean_bytes'], 'dup_bytes': b'',
        'download_name': result['download_name'], 'mime': result['mime'],
    }
    resp = {k: v for k, v in result.items() if k not in ('clean_bytes',)}
    resp['session_id']        = sid
    resp['original_size_fmt'] = fmt_bytes(result['original_size'])
    resp['clean_size_fmt']    = fmt_bytes(result['clean_size'])
    return jsonify(resp)


# ── ZIP — start background job ────────────────────────────────────────────────
@app.route('/api/process-zip', methods=['POST'])
def api_process_zip():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded.'}), 400
    f        = request.files['file']
    filename = f.filename or 'upload.zip'
    data     = f.read()
    sid      = str(uuid.uuid4())
    set_progress(sid, 1, 'Starting…')
    threading.Thread(
        target=process_zip_worker, args=(sid, data, filename), daemon=True
    ).start()
    return jsonify({'session_id': sid})


# ── SSE progress stream ───────────────────────────────────────────────────────
@app.route('/api/progress/<sid>')
def api_progress(sid):
    if not safe_sid(sid):
        return jsonify({'error': 'Invalid session.'}), 400

    def generate():
        while True:
            prog = _progress.get(sid, {
                'pct': 0, 'label': 'Waiting…', 'done': False, 'error': None
            })
            yield f"data: {json.dumps(prog)}\n\n"
            if prog.get('done'):
                break
            time.sleep(0.5)

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'},
    )


# ── Fetch ZIP result after processing finishes ────────────────────────────────
@app.route('/api/zip-result/<sid>')
def api_zip_result(sid):
    if not safe_sid(sid) or sid not in _sessions:
        return jsonify({'error': 'Result not ready or session expired.'}), 404
    sess = _sessions[sid]
    if 'result_json' not in sess:
        return jsonify({'error': 'Result not ready yet.'}), 404
    return jsonify(sess['result_json'])


# ── Downloads ─────────────────────────────────────────────────────────────────
@app.route('/api/download/<sid>')
def api_download(sid):
    if not safe_sid(sid) or sid not in _sessions:
        return jsonify({'error': 'Session not found.'}), 404
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
                     as_attachment=True,
                     download_name=s.get('dup_download_name', 'duplicates.zip'))


if __name__ == '__main__':
    print("\n✅  DeDupScan running →  http://localhost:5000\n")
    app.run(debug=False, port=5000, threaded=True)
