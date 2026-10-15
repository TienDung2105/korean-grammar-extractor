"""Synthetic ZIP -> streaming JSON -> SQLite -> XLSX memory smoke test.

Run: python -m pip install psutil
     python benchmark_memory.py --rows 100000 --padding 6000
"""
import argparse
import json
import os
import shutil
import tempfile
import threading
import time
import zipfile

import psutil
from app import app


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--rows', type=int, default=100000)
    parser.add_argument('--padding', type=int, default=6000)
    args = parser.parse_args()
    process = psutil.Process()
    peak = [process.memory_info().rss]
    stop = threading.Event()

    def sample():
        while not stop.wait(.01):
            peak[0] = max(peak[0], process.memory_info().rss)

    monitor = threading.Thread(target=sample, daemon=True)
    monitor.start()
    started = time.monotonic()
    try:
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, 'large.zip')
            row = json.dumps(dict(id='u1', form='모르지 않다', original_form='',
                                  ignored='x' * args.padding), ensure_ascii=False).encode('utf-8')
            with zipfile.ZipFile(path, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
                with archive.open('large.json', 'w', force_zip64=True) as stream:
                    stream.write(b'{"document":[{"utterance":[')
                    for index in range(args.rows):
                        if index:
                            stream.write(b',')
                        stream.write(row)
                    stream.write(b']}],"metadata":{"year":2021}}')
            with zipfile.ZipFile(path) as archive:
                uncompressed = archive.infolist()[0].file_size
            client = app.test_client()
            with open(path, 'rb') as upload:
                response = client.post('/api/process', data={'files': (upload, 'large.zip'), 'patterns[]': '6'})
            assert response.status_code == 200, response.get_json()
            job_id = response.get_json()['job_id']
            while True:
                status = client.get('/api/status/' + job_id).get_json()
                if status['status'] in ('done', 'error'):
                    break
                time.sleep(.05)
            assert status['status'] == 'done', status
            assert status['stats']['total_sentences'] == args.rows, status
            response = client.get('/api/download/' + job_id)
            assert response.status_code == 200
            output_size = sum(len(chunk) for chunk in response.response)
            response.close()
            print(json.dumps(dict(rows=args.rows, json_mib=round(uncompressed / 2**20, 2),
                                  zip_mib=round(os.path.getsize(path) / 2**20, 2),
                                  xlsx_mib=round(output_size / 2**20, 2),
                                  peak_process_rss_mib=round(peak[0] / 2**20, 2),
                                  elapsed_seconds=round(time.monotonic() - started, 2))))
    finally:
        stop.set()
        monitor.join()
        for job in app.extensions['extractor_jobs'].values():
            shutil.rmtree(job['directory'], ignore_errors=True)
        app.extensions['extractor_jobs'].clear()


if __name__ == '__main__':
    main()
