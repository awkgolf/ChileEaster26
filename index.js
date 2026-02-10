const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, Header, Footer,
    PageNumber 
} = require('docx');
const open = require('open');

// 1. STABLE PATH CONFIGURATION
const DATA_PATH = path.join(__dirname, 'travelData.json');
const PHOTO_DIR = path.join(__dirname, 'photos');
const OUTPUT_FILE = path.join(__dirname, 'Geological_Field_Journal_2026.docx');

// 2. DATA LOAD & PHOTO AUDIT
const travelData = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));

console.log("🔍 STARTING PHOTO AUDIT...");
let missingCount = 0;

if (travelData.coverImage && !fs.existsSync(path.join(PHOTO_DIR, travelData.coverImage))) {
    console.warn(`⚠️  COVER MISSING: ${travelData.coverImage}`);
    missingCount++;
}

travelData.days.forEach(day => {
    if (day.images && Array.isArray(day.images)) {
        day.images.forEach(imgObj => {
            const name = typeof imgObj === 'string' ? imgObj : imgObj.url;
            if (!fs.existsSync(path.join(PHOTO_DIR, name))) {
                console.warn(`⚠️  MISSING: ${name} (Day ${day.day})`);
                missingCount++;
            }
        });
    }
});

console.log(missingCount === 0 ? "✅ ALL PHOTOS LOCATED." : `❗ AUDIT COMPLETE: ${missingCount} missing.`);

// 3. HELPER FUNCTIONS
function insertImageWithCaption(imgObj, isCover = false) {
    // 1. Resolve the filename
    const imgName = typeof imgObj === 'string' ? imgObj : imgObj.url;
    const captionText = (typeof imgObj === 'object' && imgObj.caption) ? imgObj.caption : "";
    
    // 2. Force an ABSOLUTE path
    const imgPath = path.resolve(PHOTO_DIR, imgName);

    // 3. Check if file exists
    if (!fs.existsSync(imgPath)) {
        console.error(`❌ ERROR: Could not find image at ${imgPath}`);
        return [new Paragraph({ 
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `[MISSING IMAGE: ${imgName}]`, color: "FF0000", bold: true })] 
        })];
    }

    // 4. Read the file into a Buffer
    const imageBuffer = fs.readFileSync(imgPath);

    const elements = [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new ImageRun({
                    data: imageBuffer,
                    transformation: { 
                        width: isCover ? 550 : 450, 
                        height: isCover ? 350 : 300 
                    }
                })
            ],
            spacing: { before: 200 }
        })
    ];

    if (captionText) {
        elements.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: captionText, italics: true, size: 18, color: "4F4F4F" })],
            spacing: { after: 200 }
        }));
    }
    return elements;
}

function createGeologicalTimeline() {
    const rows = [
        ["FIELD STRATIGRAPHY & REGIONAL TIMELINE", "A04040", "FFFFFF", true],
        ["HOLOCENE (0.01 Ma): Rapa Nui human history and Moai carving.", "F5F5DC", "000000", false],
        ["PLEISTOCENE (2.5 Ma): Formation of Atacama evaporite basins.", "EFEBE9", "000000", false],
        ["MIOCENE (12 Ma): Intrusion of the Torres del Paine laccoliths.", "D7CCC8", "000000", false],
        ["CRETACEOUS (100 Ma): Initial compression and subduction of the Nazca plate.", "BCAAA4", "000000", false]
    ].map(([text, fill, color, isHeader]) => new TableRow({
        children: [new TableCell({
            shading: { fill: fill, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 120, right: 120 },
            children: [new Paragraph({ 
                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [new TextRun({ text, bold: isHeader, color: color, size: 20 })] 
            })]
        })]
    }));
    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: rows, spacing: { before: 400, after: 400 } });
}
function createGeoFieldNote(note) {
    return [new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [new TableRow({
            children: [new TableCell({
                shading: { fill: "EFEBE9" },
                borders: { left: { style: BorderStyle.SINGLE, size: 20, color: "A04040" } },
                margins: { top: 200, bottom: 200, left: 200, right: 200 },
                children: [
                    new Paragraph({ children: [new TextRun({ text: "⛏ FIELD NOTE: " + note.title, bold: true, color: "A04040" })] }),
                    new Paragraph({ children: [new TextRun({ text: note.text, italics: true })] })
                ]
            })]
        })]
    })];
}

function createStratigraphicIndex(days) {
    const tableRows = [
        new TableRow({
            children: [
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "DAY", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "TOPIC", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "SUMMARY", color: "FFFFFF", bold: true })] })] })
            ]
        })
    ];

    days.forEach(day => {
        if (day.geoNote) {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph(day.day)] }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.geoNote.title, bold: true })] })] }),
                    new TableCell({ children: [new Paragraph(day.geoNote.text.substring(0, 80) + "...")] })
                ]
            }));
        }
    });

    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } });
}

// 4. DOCUMENT CONSTRUCTION
const doc = new Document({
    features: { updateFields: true },
    sections: [{
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun("Page "),
                        new TextRun({
                        		children: [PageNumber.CURRENT],
                        		bold: true,
                        }),
                        new TextRun(" of "),
                        new TextRun({
                        		children: [PageNumber.TOTAL_PAGES],
                        		bold: true,
                        }),
                    ],
                })],
            }),
        },
        children: [
            new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
            ...insertImageWithCaption(travelData.coverImage, true),
            new Paragraph({ children: [new PageBreak()] }),

            // 2. Regional Timeline & Map Placeholder
            new Paragraph({ children: [new PageBreak()] }),
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Regional Geological Context")] }),
            
            // Map Placeholder
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                borders: { top: { style: BorderStyle.DASHED, size: 10 }, bottom: { style: BorderStyle.DASHED, size: 10 } },
                                children: [
                                    new Paragraph({ 
                                        alignment: AlignmentType.CENTER, 
                                        children: [new TextRun({ text: "📍 PLACEHOLDER: TECTONIC MAP OF CHILE / NAZCA PLATE SUBDUCTION", italics: true, color: "555555" })] 
                                    })
                                ],
                                margins: { top: 500, bottom: 500 }
                            })
                        ]
                    })
                ]
            }),
            createGeologicalTimeline(), 
            
            new Paragraph({ children: [new PageBreak()] }),
            new TableOfContents("Table of Contents", { hyperlink: true, parameters: "1-3" }),
            new Paragraph({ children: [new PageBreak()] }),

            ...travelData.days.flatMap(day => {
                        const dayElements = [
                            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`${day.day}: ${day.title}`)] }),
                        ];
                    
                        // FIX 1: ADD COORDINATES (If they exist)
                        if (day.coordinates) {
                            dayElements.push(new Paragraph({
                                children: [
                                    new TextRun({ text: "📍 GPS: ", bold: true, color: "666666", size: 16 }),
                                    new TextRun({ text: day.coordinates, color: "666666", size: 16, italics: true })
                                ],
                                spacing: { after: 100 }
                            }));
                        }
                    
                        dayElements.push(new Paragraph({ children: [new TextRun(day.description)], spacing: { after: 200 } }));
                    
                        if (day.geoNote) dayElements.push(...createGeoFieldNote(day.geoNote));
                    
                        // FIX 2: HANDLE BOTH MULTI-IMAGE AND SINGLE-IMAGE FALLBACK
                        if (day.images && Array.isArray(day.images)) {
                            // New Array Format
                            day.images.forEach(imgObj => {
                                dayElements.push(...insertImageWithCaption(imgObj));
                            });
                        } else if (day.image) {
                            // Fallback for Old Single Image Format
                            dayElements.push(...insertImageWithCaption(day.image));
                        }
                    
                        dayElements.push(new Paragraph({ children: [new PageBreak()] }));
                        return dayElements;
                    }),

            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Stratigraphic Index")]}),
            new Paragraph({ children: [new TextRun("Summary of key geological observations recorded during the expedition.")], spacing: { after: 200 } }),
            createStratigraphicIndex(travelData.days),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "End of Records", italics: true, color: "999999" })] })
        ]
    }]
});

// 5. SAVE AND EXECUTE
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log(`\n🚀 BUILD SUCCESS: ${OUTPUT_FILE}`);
    open(OUTPUT_FILE);
});