import os, http.server, socketserver
PORT = int(os.environ.get('PORT', 8000))
os.chdir(os.path.join(os.path.dirname(__file__), 'docs'))
with socketserver.TCPServer(('', PORT), http.server.SimpleHTTPRequestHandler) as httpd:
    print(f'Serving docs/ on port {PORT}')
    httpd.serve_forever()
