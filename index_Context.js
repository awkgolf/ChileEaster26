<<<<<<< Updated upstream
const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, Header, Footer,
    PageNumber 
} = require('docx');
const open = require('open');

// ========================================================================
// 1. STABLE PATH CONFIGURATION
// ========================================================================
// Define absolute paths for all file operations to ensure reliability
// across different execution contexts (CLI, IDE, scripts)

const DATA_PATH = path.join(__dirname, 'travelData.json');  // Source JSON with expedition data
const PHOTO_DIR = path.join(__dirname, 'photos');           // Directory containing all images
const OUTPUT_FILE = path.join(__dirname, 'Geological_Field_Journal_2026.docx'); // Final document output

// ========================================================================
// 2. DATA LOAD & PHOTO AUDIT
// ========================================================================
// Load expedition data and verify all referenced images exist before
// attempting document generation to prevent runtime errors

const travelData = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));

console.log("🔍 STARTING PHOTO AUDIT...");
let missingCount = 0;

// Check if cover image exists
if (travelData.coverImage && !fs.existsSync(path.join(PHOTO_DIR, travelData.coverImage))) {
    console.warn(`⚠️  COVER MISSING: ${travelData.coverImage}`);
    missingCount++;
}

// Check all daily entry images
travelData.days.forEach(day => {
    if (day.images && Array.isArray(day.images)) {
        day.images.forEach(imgObj => {
            // Handle both string format ("image.jpg") and object format ({url: "image.jpg", caption: "..."})
            const name = typeof imgObj === 'string' ? imgObj : imgObj.url;
            if (!fs.existsSync(path.join(PHOTO_DIR, name))) {
                console.warn(`⚠️  MISSING: ${name} (Day ${day.day})`);
                missingCount++;
            }
        });
    }
});

// Report audit results
console.log(missingCount === 0 ? "✅ ALL PHOTOS LOCATED." : `❗ AUDIT COMPLETE: ${missingCount} missing.`);

// ========================================================================
// 3. HELPER FUNCTIONS
// ========================================================================

/**
 * Inserts an image with optional caption into the document
 * @param {string|object} imgObj - Either a filename string or {url: string, caption: string}
 * @param {boolean} isCover - If true, uses larger dimensions for cover image
 * @returns {Array<Paragraph>} Array of paragraph elements (image + optional caption)
 */
function insertImageWithCaption(imgObj, isCover = false) {
    // 1. Extract filename from either string or object format
    const imgName = typeof imgObj === 'string' ? imgObj : imgObj.url;
    const captionText = (typeof imgObj === 'object' && imgObj.caption) ? imgObj.caption : "";
    
    // 2. Force an ABSOLUTE path to prevent relative path issues
    const imgPath = path.resolve(PHOTO_DIR, imgName);

    // 3. Verify file exists before attempting to read
    if (!fs.existsSync(imgPath)) {
        console.error(`❌ ERROR: Could not find image at ${imgPath}`);
        // Return red error placeholder instead of crashing
        return [new Paragraph({ 
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `[MISSING IMAGE: ${imgName}]`, color: "FF0000", bold: true })] 
        })];
    }

    // 4. Read the file into a Buffer for embedding in docx
    const imageBuffer = fs.readFileSync(imgPath);

    // Create centered image paragraph with size based on cover vs. content image
    const elements = [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new ImageRun({
                    data: imageBuffer,
                    transformation: { 
                        width: isCover ? 550 : 450,   // Cover images are larger
                        height: isCover ? 350 : 300 
                    }
                })
            ],
            spacing: { before: 200 }  // Add space above image
        })
    ];

    // Add caption paragraph if caption text exists
    if (captionText) {
        elements.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: captionText, italics: true, size: 18, color: "4F4F4F" })],
            spacing: { after: 200 }  // Add space below caption
        }));
    }
    return elements;
}

/**
 * Creates a color-coded geological timeline table showing major epochs
 * relevant to the field expedition sites
 * @returns {Table} Formatted timeline table
 */
function createGeologicalTimeline() {
    // Define timeline rows: [text, background color, text color, isHeader]
    const rows = [
        ["FIELD STRATIGRAPHY & REGIONAL TIMELINE", "A04040", "FFFFFF", true],  // Rust header
        ["HOLOCENE (0.01 Ma): Rapa Nui human history and Moai carving.", "F5F5DC", "000000", false],  // Beige
        ["PLEISTOCENE (2.5 Ma): Formation of Atacama evaporite basins.", "EFEBE9", "000000", false],  // Light tan
        ["MIOCENE (12 Ma): Intrusion of the Torres del Paine laccoliths.", "D7CCC8", "000000", false], // Medium tan
        ["CRETACEOUS (100 Ma): Initial compression and subduction of the Nazca plate.", "BCAAA4", "000000", false] // Dark tan
    ].map(([text, fill, color, isHeader]) => new TableRow({
        children: [new TableCell({
            shading: { fill: fill, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },  // Padding for readability
            children: [new Paragraph({ 
                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [new TextRun({ text, bold: isHeader, color: color, size: 20 })] 
            })]
        })]
    }));
    
    // Return full-width table with spacing
    return new Table({ 
        width: { size: 100, type: WidthType.PERCENTAGE }, 
        rows: rows, 
        spacing: { before: 400, after: 400 }  // Space above and below table
    });
}

/**
 * Creates a formatted field note callout box with geological observation
 * @param {object} note - Object with {title: string, text: string}
 * @returns {Array<Table>} Styled note in a single-cell table
 */
function createGeoFieldNote(note) {
    return [new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [new TableRow({
            children: [new TableCell({
                shading: { fill: "EFEBE9" },  // Light beige background
                borders: { 
                    left: { style: BorderStyle.SINGLE, size: 20, color: "A04040" }  // Rust left accent bar
                },
                margins: { top: 200, bottom: 200, left: 200, right: 200 },  // Generous padding
                children: [
                    // Bold title with pickaxe emoji
                    new Paragraph({ children: [new TextRun({ text: "⛏ FIELD NOTE: " + note.title, bold: true, color: "A04040" })] }),
                    // Italic body text
                    new Paragraph({ children: [new TextRun({ text: note.text, italics: true })] })
                ]
            })]
        })]
    })];
}

/**
 * Creates an index table summarizing all geological observations from the expedition
 * @param {Array} days - Array of day objects from travelData
 * @returns {Table} Three-column summary table
 */
function createStratigraphicIndex(days) {
    // Create header row with rust background
    const tableRows = [
        new TableRow({
            children: [
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "DAY", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "TOPIC", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "SUMMARY", color: "FFFFFF", bold: true })] })] })
            ]
        })
    ];

    // Add a row for each day that has a geological note
    days.forEach(day => {
        if (day.geoNote) {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(day.day)] }),  // Day number
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.geoNote.title, bold: true })] })] }),  // Bold topic
                    new TableCell({ children: [new Paragraph(day.geoNote.text.substring(0, 80) + "...")] })  // Truncated summary
                ]
            }));
        }
    });

    // Return full-width table with spacing
    return new Table({ 
        width: { size: 100, type: WidthType.PERCENTAGE }, 
        rows: tableRows, 
        spacing: { before: 400, after: 400 } 
    });
}

// ========================================================================
// 4. DOCUMENT CONSTRUCTION
// ========================================================================
// Assemble all components into a single Word document with proper structure

const doc = new Document({
    features: { updateFields: true },  // Enable auto-update of page numbers and TOC
    sections: [{
        // Add footer with page numbers to every page
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun("Page "),
                        new TextRun({
                            children: [PageNumber.CURRENT],  // Current page number
                            bold: true,
                        }),
                        new TextRun(" of "),
                        new TextRun({
                            children: [PageNumber.TOTAL_PAGES],  // Total pages
                            bold: true,
                        }),
                    ],
                })],
            }),
        },
        children: [
            // ============================================================
            // COVER PAGE
            // ============================================================
            new Paragraph({ 
                alignment: AlignmentType.CENTER, 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun(travelData.tripTitle)] 
            }),
            ...insertImageWithCaption(travelData.coverImage, true),  // Large cover image
            new Paragraph({ children: [new PageBreak()] }),

            // ============================================================
            // REGIONAL GEOLOGICAL CONTEXT
            // ============================================================
            new Paragraph({ children: [new PageBreak()] }),
            new Paragraph({ 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun("Regional Geological Context")] 
            }),
            
            // Map placeholder (to be replaced with actual map later)
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { 
                                    top: { style: BorderStyle.DASHED, size: 10 }, 
                                    bottom: { style: BorderStyle.DASHED, size: 10 } 
                                },
                                children: [
                                    new Paragraph({ 
                                        alignment: AlignmentType.CENTER, 
                                        children: [new TextRun({ 
                                            text: "📍 PLACEHOLDER: TECTONIC MAP OF CHILE / NAZCA PLATE SUBDUCTION", 
                                            italics: true, 
                                            color: "555555" 
                                        })] 
                                    })
                                ],
                                margins: { top: 500, bottom: 500 }
                            })
                        ]
                    })
                ]
            }),
            
            // Insert geological timeline table
            createGeologicalTimeline(), 
            
            // ============================================================
            // TABLE OF CONTENTS
            // ============================================================
            new Paragraph({ children: [new PageBreak()] }),
            new TableOfContents("Table of Contents", { 
                hyperlink: true,  // Make TOC entries clickable
                parameters: "1-3"  // Include headings levels 1-3
            }),
            new Paragraph({ children: [new PageBreak()] }),

            // ============================================================
            // DAILY ENTRIES
            // ============================================================
            // flatMap allows us to return multiple elements per day
            ...travelData.days.flatMap(day => {
                const dayElements = [
                    // Day heading (e.g., "Day 1: Arrival in Santiago")
                    new Paragraph({ 
                        heading: HeadingLevel.HEADING_2, 
                        children: [new TextRun(`${day.day}: ${day.title}`)] 
                    }),
                ];
            
                // Add GPS coordinates if available
                if (day.coordinates) {
                    dayElements.push(new Paragraph({
                        children: [
                            new TextRun({ text: "📍 GPS: ", bold: true, color: "666666", size: 16 }),
                            new TextRun({ text: day.coordinates, color: "666666", size: 16, italics: true })
                        ],
                        spacing: { after: 100 }
                    }));
                }
            
                // Add day description
                dayElements.push(new Paragraph({ 
                    children: [new TextRun(day.description)], 
                    spacing: { after: 200 } 
                }));
            
                // Add geological field note if present
                if (day.geoNote) {
                    dayElements.push(...createGeoFieldNote(day.geoNote));
                }
            
                // Handle images - supports both array format and legacy single image
                if (day.images && Array.isArray(day.images)) {
                    // New format: array of image objects
                    day.images.forEach(imgObj => {
                        dayElements.push(...insertImageWithCaption(imgObj));
                    });
                } else if (day.image) {
                    // Legacy format: single image property
                    dayElements.push(...insertImageWithCaption(day.image));
                }
            
                // Add page break after each day
                dayElements.push(new Paragraph({ children: [new PageBreak()] }));
                return dayElements;
            }),

            // ============================================================
            // STRATIGRAPHIC INDEX (Back Matter)
            // ============================================================
            new Paragraph({ 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun("Stratigraphic Index")]
            }),
            new Paragraph({ 
                children: [new TextRun("Summary of key geological observations recorded during the expedition.")], 
                spacing: { after: 200 } 
            }),
            createStratigraphicIndex(travelData.days),
            
            // End marker
            new Paragraph({ 
                alignment: AlignmentType.RIGHT, 
                children: [new TextRun({ text: "End of Records", italics: true, color: "999999" })] 
            })
        ]
    }]
});

// ========================================================================
// 5. SAVE AND EXECUTE
// ========================================================================
// Convert document object to binary buffer, write to disk, and open

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log(`\n🚀 BUILD SUCCESS: ${OUTPUT_FILE}`);
    open(OUTPUT_FILE);  // Automatically open the generated document
=======
const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, Header, Footer,
    PageNumber 
} = require('docx');
const open = require('open');

// ========================================================================
// 1. STABLE PATH CONFIGURATION
// ========================================================================
// Define absolute paths for all file operations to ensure reliability
// across different execution contexts (CLI, IDE, scripts)

const DATA_PATH = path.join(__dirname, 'travelData.json');  // Source JSON with expedition data
const PHOTO_DIR = path.join(__dirname, 'photos');           // Directory containing all images
const OUTPUT_FILE = path.join(__dirname, 'Geological_Field_Journal_2026.docx'); // Final document output

// ========================================================================
// 2. DATA LOAD & PHOTO AUDIT
// ========================================================================
// Load expedition data and verify all referenced images exist before
// attempting document generation to prevent runtime errors

const travelData = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));

console.log("🔍 STARTING PHOTO AUDIT...");
let missingCount = 0;

// Check if cover image exists
if (travelData.coverImage && !fs.existsSync(path.join(PHOTO_DIR, travelData.coverImage))) {
    console.warn(`⚠️  COVER MISSING: ${travelData.coverImage}`);
    missingCount++;
}

// Check all daily entry images
travelData.days.forEach(day => {
    if (day.images && Array.isArray(day.images)) {
        day.images.forEach(imgObj => {
            // Handle both string format ("image.jpg") and object format ({url: "image.jpg", caption: "..."})
            const name = typeof imgObj === 'string' ? imgObj : imgObj.url;
            if (!fs.existsSync(path.join(PHOTO_DIR, name))) {
                console.warn(`⚠️  MISSING: ${name} (Day ${day.day})`);
                missingCount++;
            }
        });
    }
});

// Report audit results
console.log(missingCount === 0 ? "✅ ALL PHOTOS LOCATED." : `❗ AUDIT COMPLETE: ${missingCount} missing.`);

// ========================================================================
// 3. HELPER FUNCTIONS
// ========================================================================

/**
 * Inserts an image with optional caption into the document
 * @param {string|object} imgObj - Either a filename string or {url: string, caption: string}
 * @param {boolean} isCover - If true, uses larger dimensions for cover image
 * @returns {Array<Paragraph>} Array of paragraph elements (image + optional caption)
 */
function insertImageWithCaption(imgObj, isCover = false) {
    // 1. Extract filename from either string or object format
    const imgName = typeof imgObj === 'string' ? imgObj : imgObj.url;
    const captionText = (typeof imgObj === 'object' && imgObj.caption) ? imgObj.caption : "";
    
    // 2. Force an ABSOLUTE path to prevent relative path issues
    const imgPath = path.resolve(PHOTO_DIR, imgName);

    // 3. Verify file exists before attempting to read
    if (!fs.existsSync(imgPath)) {
        console.error(`❌ ERROR: Could not find image at ${imgPath}`);
        // Return red error placeholder instead of crashing
        return [new Paragraph({ 
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `[MISSING IMAGE: ${imgName}]`, color: "FF0000", bold: true })] 
        })];
    }

    // 4. Read the file into a Buffer for embedding in docx
    const imageBuffer = fs.readFileSync(imgPath);

    // Create centered image paragraph with size based on cover vs. content image
    const elements = [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new ImageRun({
                    data: imageBuffer,
                    transformation: { 
                        width: isCover ? 550 : 450,   // Cover images are larger
                        height: isCover ? 350 : 300 
                    }
                })
            ],
            spacing: { before: 200 }  // Add space above image
        })
    ];

    // Add caption paragraph if caption text exists
    if (captionText) {
        elements.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: captionText, italics: true, size: 18, color: "4F4F4F" })],
            spacing: { after: 200 }  // Add space below caption
        }));
    }
    return elements;
}

/**
 * Creates a color-coded geological timeline table showing major epochs
 * relevant to the field expedition sites
 * @returns {Table} Formatted timeline table
 */
function createGeologicalTimeline() {
    // Define timeline rows: [text, background color, text color, isHeader]
    const rows = [
        ["FIELD STRATIGRAPHY & REGIONAL TIMELINE", "A04040", "FFFFFF", true],  // Rust header
        ["HOLOCENE (0.01 Ma): Rapa Nui human history and Moai carving.", "F5F5DC", "000000", false],  // Beige
        ["PLEISTOCENE (2.5 Ma): Formation of Atacama evaporite basins.", "EFEBE9", "000000", false],  // Light tan
        ["MIOCENE (12 Ma): Intrusion of the Torres del Paine laccoliths.", "D7CCC8", "000000", false], // Medium tan
        ["CRETACEOUS (100 Ma): Initial compression and subduction of the Nazca plate.", "BCAAA4", "000000", false] // Dark tan
    ].map(([text, fill, color, isHeader]) => new TableRow({
        children: [new TableCell({
            shading: { fill: fill, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },  // Padding for readability
            children: [new Paragraph({ 
                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [new TextRun({ text, bold: isHeader, color: color, size: 20 })] 
            })]
        })]
    }));
    
    // Return full-width table with spacing
    return new Table({ 
        width: { size: 100, type: WidthType.PERCENTAGE }, 
        rows: rows, 
        spacing: { before: 400, after: 400 }  // Space above and below table
    });
}

/**
 * Creates a formatted field note callout box with geological observation
 * @param {object} note - Object with {title: string, text: string}
 * @returns {Array<Table>} Styled note in a single-cell table
 */
function createGeoFieldNote(note) {
    return [new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [new TableRow({
            children: [new TableCell({
                shading: { fill: "EFEBE9" },  // Light beige background
                borders: { 
                    left: { style: BorderStyle.SINGLE, size: 20, color: "A04040" }  // Rust left accent bar
                },
                margins: { top: 200, bottom: 200, left: 200, right: 200 },  // Generous padding
                children: [
                    // Bold title with pickaxe emoji
                    new Paragraph({ children: [new TextRun({ text: "⛏ FIELD NOTE: " + note.title, bold: true, color: "A04040" })] }),
                    // Italic body text
                    new Paragraph({ children: [new TextRun({ text: note.text, italics: true })] })
                ]
            })]
        })]
    })];
}

/**
 * Creates an index table summarizing all geological observations from the expedition
 * @param {Array} days - Array of day objects from travelData
 * @returns {Table} Three-column summary table
 */
function createStratigraphicIndex(days) {
    // Create header row with rust background
    const tableRows = [
        new TableRow({
            children: [
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "DAY", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "TOPIC", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "SUMMARY", color: "FFFFFF", bold: true })] })] })
            ]
        })
    ];

    // Add a row for each day that has a geological note
    days.forEach(day => {
        if (day.geoNote) {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(day.day)] }),  // Day number
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.geoNote.title, bold: true })] })] }),  // Bold topic
                    new TableCell({ children: [new Paragraph(day.geoNote.text.substring(0, 80) + "...")] })  // Truncated summary
                ]
            }));
        }
    });

    // Return full-width table with spacing
    return new Table({ 
        width: { size: 100, type: WidthType.PERCENTAGE }, 
        rows: tableRows, 
        spacing: { before: 400, after: 400 } 
    });
}

// ========================================================================
// 4. DOCUMENT CONSTRUCTION
// ========================================================================
// Assemble all components into a single Word document with proper structure

const doc = new Document({
    features: { updateFields: true },  // Enable auto-update of page numbers and TOC
    sections: [{
        // Add footer with page numbers to every page
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun("Page "),
                        new TextRun({
                            children: [PageNumber.CURRENT],  // Current page number
                            bold: true,
                        }),
                        new TextRun(" of "),
                        new TextRun({
                            children: [PageNumber.TOTAL_PAGES],  // Total pages
                            bold: true,
                        }),
                    ],
                })],
            }),
        },
        children: [
            // ============================================================
            // COVER PAGE
            // ============================================================
            new Paragraph({ 
                alignment: AlignmentType.CENTER, 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun(travelData.tripTitle)] 
            }),
            ...insertImageWithCaption(travelData.coverImage, true),  // Large cover image
            new Paragraph({ children: [new PageBreak()] }),

            // ============================================================
            // REGIONAL GEOLOGICAL CONTEXT
            // ============================================================
            new Paragraph({ children: [new PageBreak()] }),
            new Paragraph({ 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun("Regional Geological Context")] 
            }),
            
            // Map placeholder (to be replaced with actual map later)
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { 
                                    top: { style: BorderStyle.DASHED, size: 10 }, 
                                    bottom: { style: BorderStyle.DASHED, size: 10 } 
                                },
                                children: [
                                    new Paragraph({ 
                                        alignment: AlignmentType.CENTER, 
                                        children: [new TextRun({ 
                                            text: "📍 PLACEHOLDER: TECTONIC MAP OF CHILE / NAZCA PLATE SUBDUCTION", 
                                            italics: true, 
                                            color: "555555" 
                                        })] 
                                    })
                                ],
                                margins: { top: 500, bottom: 500 }
                            })
                        ]
                    })
                ]
            }),
            
            // Insert geological timeline table
            createGeologicalTimeline(), 
            
            // ============================================================
            // TABLE OF CONTENTS
            // ============================================================
            new Paragraph({ children: [new PageBreak()] }),
            new TableOfContents("Table of Contents", { 
                hyperlink: true,  // Make TOC entries clickable
                parameters: "1-3"  // Include headings levels 1-3
            }),
            new Paragraph({ children: [new PageBreak()] }),

            // ============================================================
            // DAILY ENTRIES
            // ============================================================
            // flatMap allows us to return multiple elements per day
            ...travelData.days.flatMap(day => {
                const dayElements = [
                    // Day heading (e.g., "Day 1: Arrival in Santiago")
                    new Paragraph({ 
                        heading: HeadingLevel.HEADING_2, 
                        children: [new TextRun(`${day.day}: ${day.title}`)] 
                    }),
                ];
            
                // Add GPS coordinates if available
                if (day.coordinates) {
                    dayElements.push(new Paragraph({
                        children: [
                            new TextRun({ text: "📍 GPS: ", bold: true, color: "666666", size: 16 }),
                            new TextRun({ text: day.coordinates, color: "666666", size: 16, italics: true })
                        ],
                        spacing: { after: 100 }
                    }));
                }
            
                // Add day description
                dayElements.push(new Paragraph({ 
                    children: [new TextRun(day.description)], 
                    spacing: { after: 200 } 
                }));
            
                // Add geological field note if present
                if (day.geoNote) {
                    dayElements.push(...createGeoFieldNote(day.geoNote));
                }
            
                // Handle images - supports both array format and legacy single image
                if (day.images && Array.isArray(day.images)) {
                    // New format: array of image objects
                    day.images.forEach(imgObj => {
                        dayElements.push(...insertImageWithCaption(imgObj));
                    });
                } else if (day.image) {
                    // Legacy format: single image property
                    dayElements.push(...insertImageWithCaption(day.image));
                }
            
                // Add page break after each day
                dayElements.push(new Paragraph({ children: [new PageBreak()] }));
                return dayElements;
            }),

            // ============================================================
            // STRATIGRAPHIC INDEX (Back Matter)
            // ============================================================
            new Paragraph({ 
                heading: HeadingLevel.HEADING_1, 
                children: [new TextRun("Stratigraphic Index")]
            }),
            new Paragraph({ 
                children: [new TextRun("Summary of key geological observations recorded during the expedition.")], 
                spacing: { after: 200 } 
            }),
            createStratigraphicIndex(travelData.days),
            
            // End marker
            new Paragraph({ 
                alignment: AlignmentType.RIGHT, 
                children: [new TextRun({ text: "End of Records", italics: true, color: "999999" })] 
            })
        ]
    }]
});

// ========================================================================
// 5. SAVE AND EXECUTE
// ========================================================================
// Convert document object to binary buffer, write to disk, and open

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log(`\n🚀 BUILD SUCCESS: ${OUTPUT_FILE}`);
    open(OUTPUT_FILE);  // Automatically open the generated document
>>>>>>> Stashed changes
});