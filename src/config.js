const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const DATA_PATH = process.env.DATA_PATH || path.join(ROOT_DIR, 'travelData.json');
const PHOTO_DIR = process.env.PHOTO_DIR || path.join(ROOT_DIR, 'photos');
const OUTPUT_FILE = process.env.OUTPUT_FILE || path.join(ROOT_DIR, 'Geological_Field_Journal_2026.docx');
const AUTO_OPEN = process.env.AUTO_OPEN === 'true';

module.exports = {
  ROOT_DIR,
  DATA_PATH,
  PHOTO_DIR,
  OUTPUT_FILE,
  AUTO_OPEN,
};
