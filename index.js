const fs = require('fs');
const { Packer } = require('docx');
const open = require('open');
const { DATA_PATH, PHOTO_DIR, OUTPUT_FILE, AUTO_OPEN } = require('./src/config');
const { loadTravelData, validateTravelData, auditPhotos } = require('./src/data');
const { createJournalDocument } = require('./src/document');

async function writeDocument(doc, outputFile) {
  try {
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(outputFile, buffer);
  } catch (error) {
    throw new Error(`Unable to write output document ${outputFile}: ${error.message}`);
  }
}

function logAuditResults(auditResult) {
  console.log('🔍 STARTING PHOTO AUDIT...');
  if (auditResult.missingCount === 0) {
    console.log('✅ ALL PHOTOS LOCATED.');
    return;
  }

  auditResult.missing.forEach((issue) => {
    console.warn(`⚠️  ${issue}`);
  });
  console.warn(`❗ AUDIT COMPLETE: ${auditResult.missingCount} missing.`);
}

(async function run() {
  try {
    const travelData = loadTravelData(DATA_PATH);
    validateTravelData(travelData);

    const audit = auditPhotos(travelData, PHOTO_DIR);
    logAuditResults(audit);

    const doc = createJournalDocument(travelData, PHOTO_DIR);
    await writeDocument(doc, OUTPUT_FILE);

    console.log(`\n🚀 BUILD SUCCESS: ${OUTPUT_FILE}`);

    if (AUTO_OPEN) {
      await open(OUTPUT_FILE);
      console.log('📂 Output opened in your default application.');
    }
  } catch (error) {
    console.error(`\n❌ BUILD FAILED: ${error.message}`);
    process.exitCode = 1;
  }
})();
