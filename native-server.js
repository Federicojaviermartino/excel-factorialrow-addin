const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = 3000;

const mimeTypes = {
    '.html': 'text/html',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.gif': 'image/gif',
    '.ico': 'image/x-icon',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml'
};

const server = http.createServer((req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

    if (req.method === 'OPTIONS') {
        res.writeHead(200);
        res.end();
        return;
    }

    const parsedUrl = url.parse(req.url);
    let pathname = parsedUrl.pathname;
    if (pathname === '/') {
        pathname = '/taskpane.html';
    }

    console.log(`📥 Request: ${req.method} ${pathname}`);

    const filePath = path.join(__dirname, 'dist', pathname);

    fs.access(filePath, fs.constants.F_OK, (err) => {
        if (err) {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('File not found');
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = mimeTypes[ext] || 'application/octet-stream';

        fs.readFile(filePath, (err, data) => {
            if (err) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end('Internal server error');
                return;
            }

            res.writeHead(200, { 'Content-Type': contentType });
            res.end(data);
        });
    });
});

server.listen(PORT, () => {
    console.log('🚀 Excel Add-in Server Started!');
    console.log('');
    console.log(`📍 Server URL: http://localhost:${PORT}`);
    console.log(`📁 Serving from: ${path.join(__dirname, 'dist')}`);
    console.log(''); console.log('📋 Files available:');
    console.log(`   📄 Manifest: http://localhost:${PORT}/manifest.xml`);
    console.log(`   🌐 Task Pane: http://localhost:${PORT}/taskpane.html`);
    console.log(`   ⚙️  Functions: http://localhost:${PORT}/functions.bundle.js`);
    console.log(`   📊 Metadata: http://localhost:${PORT}/functions.json`);
    console.log(`   👷 Worker: http://localhost:${PORT}/worker.bundle.js`);
    console.log('');    console.log('📋 To upload to Excel Online:');
    console.log(`   1. Upload file: ${path.join(__dirname, 'dist', 'manifest.xml')}`);
    console.log('   2. Custom functions work best in Excel Online');
    console.log('   3. Test functions: =TESTVELIXO.FACTORIALROW(10) or =TESTFUNC(5)');
    console.log('   4. If functions show #NAME?, try Excel Desktop or check console logs');
    console.log('');
    console.log('🛑 Press Ctrl+C to stop the server');
});

server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
        console.log(`❌ Port ${PORT} is already in use.`);
        console.log('Try closing other applications or use a different port.');
    } else {
        console.error('Server error:', err);
    }
});