const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, Header, Footer,
    PageNumber, NumberOfPages // Fixed: Added NumberOfPages
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

// Check Cover
if (travelData.coverImage && !fs.existsSync(path.join(PHOTO_DIR, travelData.coverImage))) {
    console.warn(`⚠️  COVER MISSING: ${travelData.coverImage}`);
    missingCount++;
}

// Check Day Photos
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

if (missingCount === 0) {
    console.log("✅ ALL STRATIGRAPHIC PHOTOS LOCATED.");
} else {
    console.log(`❗ AUDIT COMPLETE: ${missingCount} assets missing.`);
}

// 3. IMAGE HELPER
function insertImageWithCaption(imgObj, isCover = false) {
    const imgName = typeof imgObj === 'string' ? imgObj : imgObj.url;
    const captionText = (typeof imgObj === 'object' && imgObj.caption) ? imgObj.caption : "";
    const imgPath = path.join(PHOTO_DIR, imgName);

    if (!fs.existsSync(imgPath)) {
        return [new Paragraph({ 
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `[MISSING: ${imgName}]`, color: "FF0000", bold: true })] 
        })];
    }

    const elements = [
        new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new ImageRun({
                    data: fs.readFileSync(imgPath),
                    transformation: { width: isCover ? 550 : 450, height: isCover ? 350 : 300 }
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

// 4. FIELD NOTE HELPER
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

// 5. DOCUMENT BUILD
const doc = new Document({
    features: { updateFields: true },
    sections: [{
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun("Page "),
                        new TextRun({ children: [PageNumber.CURRENT], bold: true }),
                        new TextRun(" of "),
                        new TextRun({ children: [NumberOfPages.TOTAL_PAGES] }) // Fixed: NumberOfPages
                    ]
                })]
            })
        },
        children: [
            // Title Page
            new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
            ...insertImageWithCaption(travelData.coverImage, true), // Fixed: Spread operator and correct function name
            new Paragraph({ children: [new PageBreak()] }),

            // TOC
            new TableOfContents("Table of Contents", { hyperlink: true, parameters: "1-3" }),
            new Paragraph({ children: [new PageBreak()] }),

            // Day Loop
            ...travelData.days.flatMap(day => {
                const dayElements = [
                    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(`${day.day}: ${day.title}`)] }),
                    new Paragraph({ children: [new TextRun(day.description)], spacing: { after: 200 } })
                ];

                if (day.geoNote) dayElements.push(...createGeoFieldNote(day.geoNote));

                if (day.images && Array.isArray(day.images)) {
                    day.images.forEach(imgObj => {
                        dayElements.push(...insertImageWithCaption(imgObj));
                    });
                }

                dayElements.push(new Paragraph({ children: [new PageBreak()] }));
                return dayElements;
            })
         // 5. DOCUMENT BUILD (Modified End)
				// ... (inside the children array, after the day loop)
				new Paragraph({ children: [new PageBreak()] }),
				new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Stratigraphic Index & Field Summary")] }),
				new Paragraph({ children: [new TextRun("The following table summarizes the key geological observations recorded during the expedition.")], spacing: { after: 200 } }),
				createStratigraphicIndex(travelData.days),
				new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "End of Records", italics: true, color: "999999" })] })   
				        ]
    }]
  // NEW: Stratigraphic Index Helper
function createStratigraphicIndex(days) {
    const tableRows = [
        new TableRow({
            children: [
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "DAY", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "GEOLOGICAL TOPIC", color: "FFFFFF", bold: true })] })] }),
                new TableCell({ shading: { fill: "A04040" }, children: [new Paragraph({ children: [new TextRun({ text: "OBSERVATION SUMMARY", color: "FFFFFF", bold: true })] })] })
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

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows,
        spacing: { before: 400, after: 400 }
    });
}  
})

// 6. SAVE
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log(`\n🚀 BUILD SUCCESS: ${OUTPUT_FILE}`);
    open(OUTPUT_FILE);
});