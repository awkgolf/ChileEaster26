const fs = require('fs');
const path = require('path');

function loadTravelData(dataPath) {
  let raw;
  try {
    raw = fs.readFileSync(dataPath, 'utf-8');
  } catch (error) {
    throw new Error(`Unable to read data file at ${dataPath}: ${error.message}`);
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error(`Invalid JSON in ${dataPath}: ${error.message}`);
  }
}

function validateTravelData(data) {
  const issues = [];

  if (!data || typeof data !== 'object') {
    issues.push('Root object is missing or invalid.');
  }

  if (!Array.isArray(data.days) || data.days.length === 0) {
    issues.push('`days` must be a non-empty array.');
  }

  if (!data.tripTitle || typeof data.tripTitle !== 'string') {
    issues.push('`tripTitle` must be a non-empty string.');
  }

  (data.days || []).forEach((day, index) => {
    const id = day?.day || `day index ${index}`;

    if (!day || typeof day !== 'object') {
      issues.push(`Day entry at index ${index} must be an object.`);
      return;
    }

    if (!day.day || typeof day.day !== 'string') {
      issues.push(`Day ${id}: missing required string field \`day\`.`);
    }

    if (!day.title || typeof day.title !== 'string') {
      issues.push(`Day ${id}: missing required string field \`title\`.`);
    }

    if (!day.description || typeof day.description !== 'string') {
      issues.push(`Day ${id}: missing required string field \`description\`.`);
    }

    if (!Array.isArray(day.images)) {
      issues.push(`Day ${id}: \`images\` must be an array (legacy \`image\` is not supported).`);
    } else {
      day.images.forEach((image, imageIndex) => {
        if (!image || typeof image !== 'object') {
          issues.push(`Day ${id}: image ${imageIndex + 1} must be an object.`);
          return;
        }

        if (!image.url || typeof image.url !== 'string') {
          issues.push(`Day ${id}: image ${imageIndex + 1} is missing required string field \`url\`.`);
        }

        if (image.caption !== undefined && typeof image.caption !== 'string') {
          issues.push(`Day ${id}: image ${imageIndex + 1} caption must be a string when provided.`);
        }
      });
    }

    if (day.geoNote !== undefined) {
      if (!day.geoNote || typeof day.geoNote !== 'object') {
        issues.push(`Day ${id}: \`geoNote\` must be an object when provided.`);
      } else {
        if (!day.geoNote.title || typeof day.geoNote.title !== 'string') {
          issues.push(`Day ${id}: geoNote.title must be a string.`);
        }
        if (!day.geoNote.text || typeof day.geoNote.text !== 'string') {
          issues.push(`Day ${id}: geoNote.text must be a string.`);
        }
      }
    }
  });

  if (issues.length > 0) {
    throw new Error(`Data validation failed:\n- ${issues.join('\n- ')}`);
  }
}

function auditPhotos(data, photoDir) {
  const missing = [];

  const checkFile = (fileName, label) => {
    const resolved = path.join(photoDir, fileName);
    if (!fs.existsSync(resolved)) {
      missing.push(`${label}: ${fileName}`);
    }
  };

  if (data.coverImage) {
    checkFile(data.coverImage, 'COVER MISSING');
  }

  const mapFile = data.regionalMap || 'TectonicPlates.jpg';
  checkFile(mapFile, 'REGIONAL MAP MISSING');

  (data.days || []).forEach((day) => {
    (day.images || []).forEach((img) => {
      checkFile(img.url, `MISSING (Day ${day.day})`);
    });
  });

  return {
    missing,
    missingCount: missing.length,
  };
}

module.exports = {
  loadTravelData,
  validateTravelData,
  auditPhotos,
};
