"""Disk-backed processing: no archive, corpus or result set is loaded as a whole."""
import os
import re
import shutil
import sqlite3
import tempfile
import threading
import time
import uuid
import zipfile
from collections import Counter
from contextlib import contextmanager
from pathlib import Path

import ijson
from flask import jsonify, render_template, request, send_file
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill

FIELDS = ('year', 'file', 'file_id', 'category', 'utterance_id',
          'sentence_form', 'sentence_original', 'pattern_name', 'matched_text')
TTL = 3600


def iter_records(stream, filename, patterns):
    # First pass collects only root metadata, even when it follows document[].
    metadata = {}
    for prefix, event, value in ijson.parse(stream):
        if prefix in ('id', 'metadata.year', 'metadata.category') and event in ('string', 'number'):
            metadata[prefix] = str(value)
    stream.seek(0)
    match = re.search(r'SDRW(\d{2})', os.path.basename(filename))
    path_year = re.findall(r'(20[12]\d)', filename)
    year = metadata.get('metadata.year', '')
    if not re.match(r'20\d{2}', year):
        year = '20' + match[1] if match else (path_year[-1] if path_year else 'N/A')
    # Parse only the three scalar fields of one utterance. Both document and
    # utterance accept an object or an array, as in the original application.
    prefixes = {'document.utterance', 'document.item.utterance',
                'document.utterance.item', 'document.item.utterance.item'}
    utterance = None
    active = None
    for prefix, event, value in ijson.parse(stream):
        if prefix in prefixes and event == 'start_map':
            utterance, active = {}, prefix
        elif active and prefix.startswith(active + '.') and event in ('string', 'number', 'null'):
            key = prefix[len(active) + 1:]
            if key in ('form', 'original_form', 'id'):
                utterance[key] = str(value) if value is not None else ''
        elif active and prefix == active and event == 'end_map':
            form = utterance.get('form', '').strip()
            original = utterance.get('original_form', '').strip()
            if form:
                matches = [(p, p['regex'].search(form)) for p in patterns]
                matches = [(p, m) for p, m in matches if m]
                if not matches and original:
                    matches = [(p, p['regex'].search(original)) for p in patterns]
                for pattern, found in matches:
                    if found:
                        yield (year, os.path.basename(filename), metadata.get('id', os.path.basename(filename)),
                               metadata.get('metadata.category', ''), utterance.get('id', ''),
                               form, original, pattern['name'], found.group())
            utterance, active = None, None


@contextmanager
def database(path):
    connection = sqlite3.connect(path)
    connection.row_factory = sqlite3.Row
    connection.execute('PRAGMA cache_size = -2048')
    try:
        yield connection
    finally:
        connection.close()


def sources(uploads, folder):
    for filename, path in uploads:
        if filename.lower().endswith('.zip'):
            with zipfile.ZipFile(path) as archive:
                for member in archive.infolist():
                    if (not member.is_dir() and member.filename.lower().endswith('.json')
                            and not member.filename.startswith('__MACOSX')):
                        # Bounded copy to disk; do not extract untrusted member paths.
                        with archive.open(member) as source, tempfile.TemporaryFile() as target:
                            shutil.copyfileobj(source, target, length=64 * 1024)
                            target.seek(0)
                            yield member.filename, target
        else:
            with open(path, 'rb') as source:
                yield filename, source
    if folder:
        for path in Path(folder).rglob('*.json'):
            with path.open('rb') as source:
                yield str(path), source


def export_excel(rows, target):
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet('Ket qua')
    sheet.freeze_panes = 'A2'
    titles = ['STT', 'Nam', 'File nguon', 'Ma phat ngon', 'The loai',
              'Cau tieng Han (chuan hoa)', 'Cau tieng Han (goc)', 'Cau truc tim thay', 'Phan khop']
    widths = [6, 8, 22, 28, 25, 55, 55, 22, 25]
    from openpyxl.utils import get_column_letter
    for col, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(col)].width = width
    header = []
    for title in titles:
        cell = WriteOnlyCell(sheet, title)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='2F5496')
        header.append(cell)
    sheet.append(header)
    counts = Counter()
    total = 0
    for total, row in enumerate(rows, 1):
        if total > 1048575:
            raise ValueError('Qua gioi han dong Excel. Hay tai theo tung cau truc.')
        values = [total] + [row[k] for k in FIELDS if k != 'file_id']
        cells = []
        for value in values:
            cell = WriteOnlyCell(sheet, value)
            if isinstance(value, str):
                cell.data_type = 's'
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            cells.append(cell)
        sheet.append(cells)
        counts[row['pattern_name']] += 1
    stats = workbook.create_sheet('Thong ke')
    stats.append(['Cau truc', 'So cau'])
    for name, count in counts.most_common():
        stats.append([name, count])
    stats.append(['TONG CONG', total])
    workbook.save(target)


def register(app, patterns, build_custom):
    jobs = {}
    lock = threading.Lock()
    busy = threading.Lock()
    exports = threading.Lock()

    def cleanup():
        # Called under lock. Downloads/readers also hold this lock while opening
        # or using files, so cleanup cannot remove an active export.
        now = time.time()
        for job_id, job in list(jobs.items()):
            if job['status'] in ('done', 'error') and now - job['finished'] > TTL:
                shutil.rmtree(job['directory'], ignore_errors=True)
                del jobs[job_id]

    def run(job, uploads, folder, active):
        try:
            job['status'] = 'processing'
            with database(job['db']) as conn:
                conn.execute('CREATE TABLE results (' + ','.join(k + ' TEXT' for k in FIELDS) + ')')
                conn.execute('CREATE INDEX pattern_idx ON results(pattern_name)')
                conn.commit()
                for filename, stream in sources(uploads, folder):
                    job['current_file'] = filename
                    # Invalid JSON rolls back only this file and is reported.
                    try:
                        conn.executemany('INSERT INTO results VALUES (?,?,?,?,?,?,?,?,?)',
                                         iter_records(stream, filename, active))
                        conn.commit()
                    except (ijson.JSONError, UnicodeError, ValueError) as error:
                        conn.rollback()
                        job['failed_files'] += 1
                        app.logger.warning('Cannot parse %s: %s', filename, error)
                    job['processed_files'] += 1
                job['stats'] = {
                    'total_sentences': conn.execute('SELECT COUNT(*) FROM results').fetchone()[0],
                    'by_pattern': dict(conn.execute('SELECT pattern_name, COUNT(*) FROM results GROUP BY pattern_name ORDER BY COUNT(*) DESC')),
                    'by_year': dict(conn.execute('SELECT year, COUNT(*) FROM results GROUP BY year ORDER BY year')),
                }
            job['status'] = 'done'
        except Exception:
            app.logger.exception('Processing failed')
            job['error'] = 'Khong xu ly duoc du lieu. Hay kiem tra ZIP/JSON va dung luong dia.'
            job['status'] = 'error'
        finally:
            for _, path in uploads:
                try:
                    os.unlink(path)
                except OSError:
                    app.logger.warning('Cannot remove upload %s', path)
            job['finished'] = time.time()
            busy.release()

    @app.get('/')
    def index():
        return render_template('index.html', patterns=patterns)

    @app.post('/api/process')
    def api_process():
        if not busy.acquire(blocking=False):
            return jsonify(error='Dang xu ly mot tac vu. Hay thu lai sau.'), 429
        directory = None
        job_id = None
        started = False
        try:
            directory = tempfile.mkdtemp(prefix='korean-job-')
            with lock:
                cleanup()
            selected = request.form.getlist('patterns[]')
            active = [p for p in patterns if not selected or p['id'] in selected]
            for idx, text in enumerate(request.form.getlist('custom_patterns[]'), 1):
                custom = build_custom(text, idx)
                if custom:
                    active.append(custom)
            uploads = []
            total = 0
            for upload in request.files.getlist('files'):
                filename = upload.filename or ''
                if not filename.lower().endswith(('.json', '.zip')):
                    continue
                path = os.path.join(directory, str(len(uploads)))
                upload.save(path)
                uploads.append((filename, path))
                if filename.lower().endswith('.zip'):
                    with zipfile.ZipFile(path) as archive:
                        total += sum(not m.is_dir() and m.filename.lower().endswith('.json')
                                     and not m.filename.startswith('__MACOSX') for m in archive.infolist())
                else:
                    total += 1
            folder = request.form.get('folder_path', '').strip()
            if folder:
                if not os.path.isdir(folder):
                    return jsonify(error='Duong dan thu muc khong ton tai tren may chu.'), 400
                total += sum(1 for _ in Path(folder).rglob('*.json'))
            if not total:
                return jsonify(error='Khong tim thay file JSON trong du lieu.'), 400
            job_id = uuid.uuid4().hex
            job = dict(status='starting', total_files=total, processed_files=0,
                       current_file='', stats={}, failed_files=0, directory=directory,
                       db=os.path.join(directory, 'results.sqlite3'))
            with lock:
                jobs[job_id] = job
            threading.Thread(target=run, args=(job, uploads, folder, active), daemon=True).start()
            started = True
            return jsonify(job_id=job_id)
        except zipfile.BadZipFile:
            return jsonify(error='File ZIP khong hop le hoac bi hong.'), 400
        finally:
            if not started:
                if job_id:
                    with lock:
                        jobs.pop(job_id, None)
                if directory:
                    shutil.rmtree(directory, ignore_errors=True)
                busy.release()

    @app.get('/api/status/<job_id>')
    def api_status(job_id):
        with lock:
            cleanup()
            job = jobs.get(job_id)
            if not job:
                return jsonify(error='Job not found'), 404
            return jsonify({k: job.get(k) for k in ('status', 'total_files', 'processed_files',
                           'current_file', 'stats', 'error', 'failed_files')})

    @app.get('/api/results/<job_id>')
    def api_results(job_id):
        try:
            page = max(1, int(request.args.get('page', 1)))
            per_page = min(200, max(1, int(request.args.get('per_page', 100))))
        except ValueError:
            return jsonify(error='Invalid pagination'), 400
        pattern = request.args.get('pattern', '')
        where, params = (' WHERE pattern_name=?', (pattern,)) if pattern else ('', ())
        with lock:
            cleanup()
            job = jobs.get(job_id)
            if not job:
                return jsonify(error='Job not found'), 404
            if job['status'] != 'done':
                return jsonify(error='Job is not complete'), 409
            with database(job['db']) as conn:
                total = conn.execute('SELECT COUNT(*) FROM results' + where, params).fetchone()[0]
                rows = conn.execute('SELECT * FROM results' + where + ' ORDER BY rowid LIMIT ? OFFSET ?',
                                    params + (per_page, (page - 1) * per_page))
                result = [dict(row) for row in rows]
        return jsonify(total=total, page=page, per_page=per_page,
                       total_pages=(total + per_page - 1) // per_page, results=result)

    @app.get('/api/download/<job_id>')
    def api_download(job_id):
        if not exports.acquire(blocking=False):
            return jsonify(error='Dang xuat Excel. Hay thu lai sau.'), 429
        output = None
        try:
            pattern = request.args.get('pattern', '')
            where, params = (' WHERE pattern_name=?', (pattern,)) if pattern else ('', ())
            with lock:
                cleanup()
                job = jobs.get(job_id)
                if not job:
                    return jsonify(error='Job not found'), 404
                if job['status'] != 'done':
                    return jsonify(error='Job is not complete'), 409
                with database(job['db']) as conn:
                    count = conn.execute('SELECT COUNT(*) FROM results' + where, params).fetchone()[0]
                    if not count:
                        return jsonify(error='No results'), 404
                    if count > 1048575:
                        return jsonify(error='Qua gioi han dong Excel. Hay tai theo tung cau truc.'), 400
                    output = tempfile.TemporaryFile()
                    export_excel(conn.execute('SELECT * FROM results' + where + ' ORDER BY rowid', params), output)
            output.seek(0)
            response = send_file(output, as_attachment=True, download_name='ket_qua.xlsx',
                                 mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response.call_on_close(output.close)
            return response
        except Exception:
            if output:
                output.close()
            app.logger.exception('Export failed')
            return jsonify(error='Khong xuat duoc Excel. Hay thu tai theo tung cau truc.'), 500
        finally:
            exports.release()

    # Make the store available for diagnostics and tests without storing corpus data.
    app.extensions['extractor_jobs'] = jobs
