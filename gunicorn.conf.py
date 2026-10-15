import os

# Jobs live in one process. Threads keep polling responsive during uploads/exports.
bind = '0.0.0.0:' + os.environ.get('PORT', '5000')
workers = 1
worker_class = 'gthread'
threads = 4
timeout = 600
