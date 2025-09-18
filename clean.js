const fs = require('fs');
const path = require('path');

console.log('🧹 Cleaning project...');

const distDir = path.join(__dirname, 'dist');
if (fs.existsSync(distDir)) {
    fs.rmSync(distDir, { recursive: true, force: true });
    console.log('✅ Removed dist/ directory');
}

const cacheDir = path.join(__dirname, 'node_modules', '.cache');
if (fs.existsSync(cacheDir)) {
    fs.rmSync(cacheDir, { recursive: true, force: true });
    console.log('✅ Removed node_modules/.cache');
}

console.log('✅ Clean completed! Now run: npm run build');