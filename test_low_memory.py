import io
import json
import shutil
import tempfile
import time
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook
from app import app, ALL_PATTERNS
from low_memory import iter_records


class ExtractorTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()

    def tearDown(self):
        jobs = app.extensions['extractor_jobs']
        for job in jobs.values():
            shutil.rmtree(job['directory'], ignore_errors=True)
        jobs.clear()

    def wait(self, job_id):
        for _ in range(1000):
            status = self.client.get('/api/status/' + job_id).get_json()
            if status['status'] in ('done', 'error'):
                return status
            time.sleep(.01)
        self.fail('Job did not finish')

    def corpus(self, document_list=True, utterance_list=True):
        utterance = dict(id='u1', form='안 먹지 않다.', original_form='original')
        document = dict(utterance=[utterance] if utterance_list else utterance)
        # Metadata deliberately follows document.
        return dict(document=[document] if document_list else document,
                    metadata=dict(year=2021, category='test'), id='source')

    def test_streaming_all_supported_shapes_and_fallback(self):
        for doc_list in (True, False):
            for utt_list in (True, False):
                stream = io.BytesIO(json.dumps(self.corpus(doc_list, utt_list)).encode())
                rows = list(iter_records(stream, 'SDRW2201.json', ALL_PATTERNS))
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0][:5], ('2021', 'SDRW2201.json', 'source', 'test', 'u1'))
        source = self.corpus()
        source['document'][0]['utterance'][0].update(form='평범한 문장', original_form='모르지 않다')
        self.assertEqual(list(iter_records(io.BytesIO(json.dumps(source).encode()), 'x.json', ALL_PATTERNS))[0][-1], '모르지 않다')

    def test_zip_routes_filter_export_and_file_error(self):
        archive = io.BytesIO()
        with zipfile.ZipFile(archive, 'w') as z:
            z.writestr('2021/a.json', json.dumps(self.corpus()))
            z.writestr('2022/b.json', json.dumps(self.corpus(False, False)))
            z.writestr('bad.json', '{ invalid')
            z.writestr('__MACOSX/ignored.json', '{}')
        archive.seek(0)
        response = self.client.post('/api/process', data={'files': (archive, 'input.zip')})
        self.assertEqual(response.status_code, 200)
        job_id = response.get_json()['job_id']
        status = self.wait(job_id)
        self.assertEqual(status['status'], 'done')
        self.assertEqual(status['processed_files'], 3)
        self.assertEqual(status['failed_files'], 1)
        self.assertEqual(status['stats']['total_sentences'], 2)
        response = self.client.get(f'/api/results/{job_id}?per_page=1&page=2')
        data = response.get_json()
        self.assertEqual((data['total'], data['total_pages'], len(data['results'])), (2, 2, 1))
        pattern = data['results'][0]['pattern_name']
        self.assertEqual(self.client.get(f'/api/results/{job_id}', query_string={'pattern': pattern}).get_json()['total'], 2)
        self.assertEqual(self.client.get(f'/api/results/{job_id}?pattern=absent').get_json()['total'], 0)
        self.assertEqual(self.client.get(f'/api/results/{job_id}?per_page=0').status_code, 200)
        self.assertEqual(self.client.get(f'/api/results/{job_id}?page=no').status_code, 400)
        response = self.client.get('/api/download/' + job_id, buffered=True)
        self.assertEqual(response.status_code, 200)
        book = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(book['Ket qua'].values)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[1][5], '안 먹지 않다.')
        self.assertEqual(list(book['Thong ke'].values)[-1], ('TONG CONG', 2))
        book.close()
        response.close()
        job = app.extensions['extractor_jobs'][job_id]
        job['finished'] = time.time() - 3601
        self.assertEqual(self.client.get('/api/status/' + job_id).status_code, 404)
        self.assertFalse(Path(job['directory']).exists())

    def test_upload_validation_releases_slot_and_custom_pattern(self):
        self.assertEqual(self.client.post('/api/process', data={'files': (io.BytesIO(b'bad'), 'x.zip')}).status_code, 400)
        self.assertEqual(self.client.post('/api/process').status_code, 400)
        response = self.client.post('/api/process', data={
            'files': (io.BytesIO(json.dumps(self.corpus()).encode()), 'x.json'),
            'patterns[]': 'nonexistent', 'custom_patterns[]': '안 먹지 않다'})
        status = self.wait(response.get_json()['job_id'])
        self.assertEqual(status['stats']['total_sentences'], 1)

    def test_concurrent_job_rejected_and_background_failure_visible(self):
        import threading
        entered, proceed = threading.Event(), threading.Event()
        def fail(*args):
            entered.set()
            proceed.wait(5)
            raise OSError('simulated disk failure')
        with patch('low_memory.sources', side_effect=fail):
            response = self.client.post('/api/process', data={'files': (io.BytesIO(b'{}'), 'x.json')})
            job_id = response.get_json()['job_id']
            self.assertTrue(entered.wait(2))
            self.assertEqual(self.client.post('/api/process').status_code, 429)
            proceed.set()
            self.assertEqual(self.wait(job_id)['status'], 'error')
        self.assertEqual(self.client.post('/api/process').status_code, 400)

    def test_local_folder_and_disk_upload_stream(self):
        from app import DiskUploadRequest
        with DiskUploadRequest({})._get_file_stream(100, 'application/json') as stream:
            stream.write(b'{}')
            self.assertGreaterEqual(stream.fileno(), 0)
        with tempfile.TemporaryDirectory() as directory:
            Path(directory, 'a.json').write_text(json.dumps(self.corpus()), encoding='utf-8')
            response = self.client.post('/api/process', data={'folder_path': directory})
            status = self.wait(response.get_json()['job_id'])
            self.assertEqual(status['stats']['total_sentences'], 1)
        self.assertEqual(self.client.post('/api/process', data={'folder_path': directory}).status_code, 400)


if __name__ == '__main__':
    unittest.main()
