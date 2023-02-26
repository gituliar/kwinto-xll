from   http.server import BaseHTTPRequestHandler, HTTPServer
import json


class Service(BaseHTTPRequestHandler):
    def do_POST(self):
        print(f'client = {self.client_address}')

        request_ = self.rfile.read(int(self.headers.get('Content-Length'))).decode()
        request = json.loads(request_)
        print(f'request = {request}')
        print(self.headers)

        self.send_response(200)
        self.send_header("Content-type", "text/plain")
        self.end_headers()

        response = {
            'id': request.get('id'),
            "result": request,
        }

        self.wfile.write(bytes(json.dumps(response), "utf-8"))


def main(host='localhost', port=4000):
    webServer = HTTPServer((host, port), Service)
    print("Serving http://%s:%s" % (host, port))

    try:
        webServer.serve_forever()
    except KeyboardInterrupt:
        pass

    webServer.server_close()
    print("Server stopped")


if __name__ == '__main__':
    main()