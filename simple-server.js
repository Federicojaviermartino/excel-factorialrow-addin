const express = require('express');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = 3000;

app.use(cors({
    origin: '*',
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.static(path.join(__dirname, 'dist')));

app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist', 'taskpane.html'));
});

app.listen(PORT, () => {
    console.log('🚀 Excel Add-in Server Started!');
    console.log('');
    console.log(`📍 Server URL: http://localhost:${PORT}`);
    console.log(`📁 Serving from: ${path.join(__dirname, 'dist')}`);
    console.log('');    console.log('📋 To upload to Excel Online:');
    console.log(`   1. Upload file: ${path.join(__dirname, 'dist', 'manifest.xml')}`);
    console.log('   2. Custom functions work best in Excel Online');
    console.log('   3. Test functions: =TESTVELIXO.FACTORIALROW(10) or =TESTFUNC(5)');
    console.log('   4. If functions show #NAME?, try Excel Desktop or check console logs');
    console.log('');
    console.log('🛑 Press Ctrl+C to stop the server');
});