from http.server import BaseHTTPRequestHandler, HTTPServer
import os
DIR = r"C:\Users\x6668\Desktop\ИИ\ARTEL\агенты\КЛАДОВЩИК\_КОНТЕКСТ"
class H(BaseHTTPRequestHandler):
    def do_GET(self):
        path = os.path.join(DIR, "dvizhenie_forma.gs")
        data = open(path, "rb").read()
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
    def log_message(self, *a):
        pass
HTTPServer(("127.0.0.1", 8765), H).serve_forever()
