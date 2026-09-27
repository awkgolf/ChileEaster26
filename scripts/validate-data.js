const { DATA_PATH } = require('../src/config');
const { loadTravelData, validateTravelData } = require('../src/data');

try {
  const data = loadTravelData(DATA_PATH);
  validateTravelData(data);
  console.log(`✅ Data validation passed: ${DATA_PATH}`);
} catch (error) {
  console.error(`❌ Data validation failed: ${error.message}`);
  process.exitCode = 1;
}
