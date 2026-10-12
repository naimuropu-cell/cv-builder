const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const PORT = process.env.PORT || 3000;
const PUBLIC_DIR = __dirname;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.ico': 'image/x-icon'
};

let generateDocxBuffer;
try {
  generateDocxBuffer = require('./docx_exporter').generateDocxBuffer;
} catch (e) {
  console.warn('Docx exporter could not be loaded:', e.message);
}

const server = http.createServer((req, res) => {
  // CORS & Cache Headers for Dev
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Cache-Control', 'no-cache, no-store, must-revalidate');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  let reqPath = decodeURI(req.url.split('?')[0]);

  // Handle Dynamic Word DOCX Export
  if (req.method === 'POST' && reqPath === '/api/export-docx') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      try {
        if (!generateDocxBuffer) throw new Error('DOCX engine not available');
        const state = JSON.parse(body);
        const buffer = await generateDocxBuffer(state);
        const fileName = `${(state.name || 'Resume').replace(/\s+/g, '_')}_Resume.docx`;

        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          'Content-Disposition': `attachment; filename="${fileName}"`,
          'Content-Length': buffer.length
        });
        res.end(buffer);
      } catch (err) {
        console.error('Docx export error:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message }));
      }
    });
    return;
  }

  if (reqPath === '/' || reqPath === '') {
    reqPath = '/builder.html';
  }

  const filePath = path.join(PUBLIC_DIR, reqPath);

  // Security: prevent path traversal outside project directory
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('403 Forbidden');
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(`404 Not Found: ${reqPath}`);
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    const stream = fs.createReadStream(filePath);
    stream.pipe(res);
  });
});

server.listen(PORT, () => {
  const url = `http://localhost:${PORT}/`;
  console.log(`====================================================`);
  console.log(`🚀 NovaCV Studio Server is running!`);
  console.log(`🔗 Local URL: ${url}`);
  console.log(`====================================================`);

  // Automatically open default browser on Windows
  const openCmd = process.platform === 'win32' ? `start ${url}` : `xdg-open ${url}`;
  exec(openCmd, (err) => {
    if (err) {
      console.log(`Note: You can open ${url} manually in your browser.`);
    } else {
      console.log(`✨ Opened NovaCV Studio in your browser.`);
    }
  });
});
