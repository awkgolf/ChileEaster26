const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');

const repoRoot = path.resolve(__dirname, '..');
const outputFile = path.join(os.tmpdir(), 'chile-journal-smoke-test.docx');

try {
  if (fs.existsSync(outputFile)) {
    fs.unlinkSync(outputFile);
  }

  execFileSync('node', ['index.js'], {
    cwd: repoRoot,
    env: {
      ...process.env,
      AUTO_OPEN: 'false',
      OUTPUT_FILE: outputFile,
    },
    stdio: 'inherit',
  });

  if (!fs.existsSync(outputFile)) {
    throw new Error('Smoke test build completed without generating output file.');
  }

  const size = fs.statSync(outputFile).size;
  if (size === 0) {
    throw new Error('Smoke test output file is empty.');
  }

  console.log(`✅ Smoke test passed: generated ${outputFile} (${size} bytes)`);
} catch (error) {
  console.error(`❌ Smoke test failed: ${error.message}`);
  process.exitCode = 1;
}
