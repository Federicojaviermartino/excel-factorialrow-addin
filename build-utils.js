const fs = require('fs');
const path = require('path');

function clean() {
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

    console.log('✅ Clean completed!');
}

function copyFiles() {
    console.log('📁 Copying additional files...');
    
    const distDir = path.join(__dirname, 'dist');
    if (!fs.existsSync(distDir)) {
        fs.mkdirSync(distDir, { recursive: true });
    }

    // Clean old files
    const oldFiles = ['manifest-fixed.xml', 'custom-functions-simple.js', 'functions-simple.json'];
    oldFiles.forEach(file => {
        const filePath = path.join(__dirname, file);
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
            console.log(`🗑️ Removed: ${file}`);
        }
    });

    // Copy manifest
    const manifestSrc = path.join(__dirname, 'manifest.xml');
    const manifestDest = path.join(distDir, 'manifest.xml');
    if (fs.existsSync(manifestSrc)) {
        fs.copyFileSync(manifestSrc, manifestDest);
        console.log('✅ Copied manifest.xml');
    }

    // Copy functions metadata
    const functionsSrc = path.join(__dirname, 'src', 'functions-new.json');
    const functionsDest = path.join(distDir, 'functions.json');
    if (fs.existsSync(functionsSrc)) {
        fs.copyFileSync(functionsSrc, functionsDest);
        console.log('✅ Copied functions.json');
    }
}

// Command line interface
const command = process.argv[2];
switch (command) {
    case 'clean':
        clean();
        break;
    case 'copy':
        copyFiles();
        break;
    case 'all':
        clean();
        copyFiles();
        break;
    default:
        console.log('Usage: node build-utils.js [clean|copy|all]');
        break;
}

module.exports = { clean, copyFiles };