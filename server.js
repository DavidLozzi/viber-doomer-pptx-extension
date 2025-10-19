const fs = require('fs');
const http = require('http');
const https = require('https');
const path = require('path');

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, 'public');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
};

function serveFile(filePath, res) {
  fs.readFile(filePath, (error, content) => {
    if (error) {
      if (error.code === 'ENOENT') {
        res.writeHead(404, { 'Content-Type': 'text/plain' });
        res.end('Not Found');
      } else {
        res.writeHead(500, { 'Content-Type': 'text/plain' });
        res.end('Server Error');
      }
      return;
    }
    const ext = path.extname(filePath);
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    res.end(content);
  });
}

function requestHandler(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  let filePath = path.join(ROOT, url.pathname);
  if (filePath.endsWith(path.sep)) {
    filePath = path.join(filePath, 'index.html');
  }
  fs.stat(filePath, (error, stats) => {
    if (error) {
      serveFile(path.join(ROOT, 'index.html'), res);
      return;
    }
    if (stats.isDirectory()) {
      serveFile(path.join(filePath, 'index.html'), res);
      return;
    }
    serveFile(filePath, res);
  });
}

function createServer() {
  const keyPath = process.env.SSL_KEY_PATH;
  const certPath = process.env.SSL_CERT_PATH;
  if (keyPath && certPath && fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    const options = {
      key: fs.readFileSync(keyPath),
      cert: fs.readFileSync(certPath),
    };
    return https.createServer(options, requestHandler);
  }
  console.warn('⚠️  SSL certificates not provided. Falling back to HTTP. Office add-ins require HTTPS for sideloading.');
  return http.createServer(requestHandler);
}

const server = createServer();
server.listen(PORT, () => {
  console.log(`Static server running on port ${PORT}`);
});
