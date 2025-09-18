const fs = require('fs');
const path = require('path');

const distDir = path.join(__dirname, 'dist');
if (!fs.existsSync(distDir)) {
    fs.mkdirSync(distDir, { recursive: true });
}

const oldManifests = ['manifest-fixed.xml', 'custom-functions-simple.js', 'functions-simple.json'];
oldManifests.forEach(file => {
    const filePath = path.join(__dirname, file);
    if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
        console.log(`🗑️ Removed duplicate file: ${file}`);
    }
});

const manifestSrc = path.join(__dirname, 'manifest.xml');
const manifestDest = path.join(distDir, 'manifest.xml');
if (fs.existsSync(manifestSrc)) {
    fs.copyFileSync(manifestSrc, manifestDest);
    console.log('✅ Copied manifest.xml to dist/');
} else {
    console.error('❌ manifest.xml not found in root directory');
}

const functionsSrc = path.join(__dirname, 'src', 'functions-new.json');
const functionsDest = path.join(distDir, 'functions.json');
if (fs.existsSync(functionsSrc)) {
    fs.copyFileSync(functionsSrc, functionsDest);
    console.log('✅ Copied functions-new.json to dist/functions.json');
}