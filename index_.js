const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, ExternalHyperlink, Header, Footer,
    PageNumber, NumberOfPages
} = require('docx');
const open = require('open');

// 1. Stable Path Configuration
const DATA_PATH = path.join(__dirname, 'travelData.json');
const PHOTO_DIR = path.join(__dirname, 'photos');
const OUTPUT_FILE = path.join(__dirname, 'Geological_Field_Journal_2026.docx');

// 2. Load Data
const travelData = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));

// 3. Helper Functions (Timeline, GeoNote, Image)
function createGeoFieldNote(note) {
    if (!note) return [];
    return [new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [new TableRow({
            children: [new TableCell({
                shading: { fill: "EFEBE9" },
                borders: { left: { style: BorderStyle.SINGLE, size: 20, color: "A04040" } },
                children: [
                    new Paragraph({ children: [new TextRun({ text: "⛏ FIELD NOTE: " + note.title, bold: true, color: "A04040" })] }),
                    new Paragraph({ children: [new TextRun({ text: note.text, italics: true })] })
                ]
            })]
        })]
    })];
}

// 4. Document Build
const doc = new Document({
    features: { updateFields: true }, // Forcing page numbers to update
    sections: [{
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                        new TextRun({ text: "Page " }),
                        new TextRun({ children: [PageNumber.CURRENT], bold: true }),
                        new TextRun({ text: " of " }),
                        new TextRun({ children: [PageNumber.TOTAL_PAGES] })
                    ]
                })]
            })
        },
        children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
            new Paragraph({ children: [new PageBreak()] }),
            new TableOfContents("Table of Contents", { hyperlink: true, parameters: "1-3" }),
            new Paragraph({ children: [new PageBreak()] }),
            ...travelData.days.flatMap(day => {
                const dayContent = [
                    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(day.day + ": " + day.title)] }),
                    new Paragraph({ children: [new TextRun(day.description)] })
                ];
                if (day.geoNote) dayContent.push(...createGeoFieldNote(day.geoNote));
                dayContent.push(new Paragraph({ children: [new PageBreak()] }));
                return dayContent;
            })
        ]
    }]
});

// 5. Save and Open
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log("🚀 Journal Built Successfully!");
    open(OUTPUT_FILE);
});