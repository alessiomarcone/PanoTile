# Local dev server that applies the Netlify _headers file (CSP etc.).
# Usage: python3 tools/serve_headers.py <project_dir> <port>
import sys, http.server, functools
root, port = sys.argv[1], int(sys.argv[2])
hdrs = []
for line in open(f"{root}/_headers"):
    if line.startswith("  ") and ":" in line:
        k, v = line.strip().split(":", 1); hdrs.append((k.strip(), v.strip()))
class H(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        for k, v in hdrs: self.send_header(k, v)
        super().end_headers()
http.server.ThreadingHTTPServer(("127.0.0.1", port), functools.partial(H, directory=root)).serve_forever()
