"""Plain-HTTP catcher on port 80.

The ERP itself only speaks HTTPS on port 8000 (see DEPLOYMENT.md step 6) —
WebRTC's getUserMedia needs a secure context, so plain HTTP was dropped
entirely there. A browser sending plain HTTP bytes at a TLS-only socket just
gets a connection reset, with no chance for the app to say anything useful.

This is a separate, minimal listener on port 80 whose only job is to redirect
that traffic to the HTTPS site, so someone typing http://192.168.88.30 (no
port) or reviving an old http://...:8000 bookmark gets sent to the right
place instead of a dead connection. Deliberately stdlib-only — no Django,
no app state — since redirecting is all it needs to do.
"""

import http.server
import socketserver

TARGET_HOST = '192.168.88.30'
TARGET_PORT = 8000


class RedirectHandler(http.server.BaseHTTPRequestHandler):
    def _redirect(self):
        location = f'https://{TARGET_HOST}:{TARGET_PORT}{self.path}'
        body = (
            '<!doctype html><html><head><title>Use HTTPS</title></head><body '
            'style="font-family:sans-serif;max-width:32rem;margin:3rem auto;text-align:center">'
            '<h2>This ERP now requires a secure (HTTPS) connection</h2>'
            f'<p>Redirecting to <a href="{location}">{location}</a>&hellip;</p>'
            '<p style="color:#666">Your browser will show a one-time "not secure" warning for '
            'the self-signed certificate &mdash; click Advanced&nbsp;&rarr;&nbsp;Proceed. '
            'Please update your bookmark to the https:// link.</p>'
            '</body></html>'
        ).encode('utf-8')
        self.send_response(301)
        self.send_header('Location', location)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self._redirect()

    def do_HEAD(self):
        self._redirect()

    def do_POST(self):
        self._redirect()

    def log_message(self, format, *args):
        print('[%s] %s' % (self.log_date_time_string(), format % args), flush=True)


class ReusableTCPServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True


if __name__ == '__main__':
    with ReusableTCPServer(('0.0.0.0', 80), RedirectHandler) as httpd:
        print('HTTP-to-HTTPS redirect listening on 0.0.0.0:80 -> https://%s:%s' % (TARGET_HOST, TARGET_PORT), flush=True)
        httpd.serve_forever()
