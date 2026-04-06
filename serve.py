import os, sys
os.chdir("/Users/mounaim/Desktop/mainAna")
port = int(os.environ.get("PORT", sys.argv[1] if len(sys.argv) > 1 else 8080))
import http.server
handler = http.server.SimpleHTTPRequestHandler
httpd = http.server.HTTPServer(("", port), handler)
print(f"Serving on port {port}")
httpd.serve_forever()
