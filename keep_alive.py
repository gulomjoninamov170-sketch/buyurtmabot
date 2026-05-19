from threading import Thread
from http.server import HTTPServer, BaseHTTPRequestHandler
import os


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')

    def log_message(self, format, *args):
        pass


def keep_alive():
    port = int(os.environ.get('PORT', 10000))
    server = HTTPServer(('0.0.0.0', port), Handler)
    t = Thread(target=server.serve_forever, daemon=True)
    t.start()
