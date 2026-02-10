const fs = require('fs');
const path = require('path');
const { 
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
    AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageBreak, ImageRun, TableOfContents, ExternalHyperlink, Header, Footer
} = require('docx');
const open = require('open');

// 1. Configuration
const DATA_PATH = './travelData.json';
const PHOTO_DIR = './photos/';
const OUTPUT_FILE = 'Geological_Field_Journal_2026.docx';

// 2. Load and Validate Data
if (!fs.existsSync(DATA_PATH)) {
    console.error("❌ Error: travelData.json not found!");
    process.exit(1);
}
const travelData = JSON.parse(fs.readFileSync(DATA_PATH, 'utf-8'));

// 3. Helper: Geological Timeline Table
function createGeologicalTimeline() {
    const rows = [
        ["FIELD STRATIGRAPHY & TIMELINE", "A04040", "FFFFFF", true],
        ["HOLOCENE (0.01 Ma): Rapa Nui human history and Moai carving.", "F5F5DC", "000000", false],
        ["PLEISTOCENE (2.5 Ma): Formation of Atacama evaporite basins.", "EFEBE9", "000000", false],
        ["MIOCENE (12 Ma): Intrusion of the Torres del Paine granite laccoliths.", "D7CCC8", "000000", false],
        ["CRETACEOUS (100 Ma): Initial compression and subduction of the Nazca plate.", "BCAAA4", "000000", false]
    ].map(([text, fill, color, isHeader]) => new TableRow({
        children: [new TableCell({
            shading: { fill: fill, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 100, right: 100 },
            children: [new Paragraph({ 
                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                children: [new TextRun({ text, bold: isHeader, color: color, size: 20 })] 
            })]
        })]
    }));

    return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: rows, spacing: { before: 400, after: 400 } });
}

// 4. Helper: Field Note Box
function createGeoFieldNote(note) {
    if (!note) return [];
    return [new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [new TableRow({
            children: [new TableCell({
                shading: { fill: "EFEBE9", type: ShadingType.CLEAR },
                borders: { left: { style: BorderStyle.SINGLE, size: 20, color: "A04040" }, top: {style: BorderStyle.NIL}, right: {style: BorderStyle.NIL}, bottom: {style: BorderStyle.NIL} },
                margins: { top: 200, bottom: 200, left: 200, right: 200 },
                children: [
                    new Paragraph({ children: [new TextRun({ text: "⛏ GEOLOGICAL FIELD NOTE: " + note.title.toUpperCase(), bold: true, color: "A04040", size: 18 })] }),
                    new Paragraph({ children: [new TextRun({ text: note.text, italics: true, size: 20 })], spacing: { before: 100 } })
                ]
            })]
        })],
        spacing: { before: 300, after: 300 }
    })];
}

// 5. Helper: Image Inserter
function insertImage(imgName, isCover = false) {
    const imgPath = path.join(PHOTO_DIR, imgName);
    if (!fs.existsSync(imgPath)) return new Paragraph({ children: [new TextRun({ text: `[Missing Photo: ${imgName}]`, color: "FF0000" })] });
    
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new ImageRun({
            data: fs.readFileSync(imgPath),
            transformation: { width: isCover ? 500 : 450, height: isCover ? 350 : 300 }
        })],
        spacing: { before: 200, after: 200 }
    });
}

// 6. Main Document Build
const doc = new Document({
    styles: {
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", run: { font: "Georgia", size: 48, bold: true, color: "A04040" } },
            { id: "Heading2", name: "Heading 2", run: { font: "Georgia", size: 32, bold: true, color: "4F4F4F" } },
            { id: "Normal", name: "Normal", run: { font: "Georgia", size: 22 } }
        ]
    },
    sections: [{
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "DRAFT FIELD JOURNAL - FEB 2026", size: 16, color: "999999" })] })] }) },
        footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
            new TextRun({ text: `© 2026 ${travelData.author} | Page `, size: 16 }),
            new TextRun({ children: ["PAGE_NUMBER"], field: "PAGE_NUMBER", size: 16, bold: true }),
            new TextRun({ text: " of ", size: 16 }),
            new TextRun({ children: ["NUM_PAGES"], field: "NUM_PAGES", size: 16 })
        ]})]})},
        children: [
            // Title Page
            new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
            insertImage("cover.jpg", true),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "A Study of Tectonic and Volcanic Formations", italics: true })] }),
            createGeologicalTimeline(),
            new Paragraph({ children: [new PageBreak()] }),

            // TOC
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Stratigraphy of the Trip")] }),
            new TableOfContents("Table of Contents", { hyperlink: true, parameters: "1-3" }),
            new Paragraph({ children: [new PageBreak()] }),

            // Daily Entries
            ...travelData.days.flatMap(item => {
                const content = [
                    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(item.day + ": " + item.title)] }),
                    new Paragraph({ children: [new TextRun(item.description)], spacing: { after: 200 } })
                ];
                if (item.geoNote) content.push(...createGeoFieldNote(item.geoNote));
                if (item.image) content.push(insertImage(item.image));
                if (item.coordinates) {
                    content.push(new Paragraph({ children: [
                        new TextRun({ text: "📍 GPS: ", bold: true }),
                        new ExternalHyperlink({ children: [new TextRun({ text: item.coordinates, color: "0000FF", underline: true })], link: `https://www.google.com/maps/search/?api=1&query=${item.coordinates}` })
                    ]}));
                }
                content.push(new Paragraph({ children: [new PageBreak()] }));
                return content;
            }),

            // Glossary & Summary
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Geological Glossary")] }),
            ...travelData.glossary.map(g => new Paragraph({ children: [new TextRun({ text: g.term + ": ", bold: true }), new TextRun(g.definition)], indent: { left: 240 }, spacing: { after: 120 } })),
            new Paragraph({ children: [new PageBreak()] }),
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Summary of Observations")] }),
            new Paragraph({ children: [new TextRun("The Chilean margin serves as a premier natural laboratory for plate tectonics. From the active subduction in the Andes to the unique evaporite preservation in the Atacama, the observed strata confirm a high-energy lithospheric environment.")], spacing: { after: 200 } }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "End of Field Records", bold: true, color: "999999" })] })
        ]
    }]
});

// 7. Output
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync(OUTPUT_FILE, buffer);
    console.log(`🚀 Success! ${OUTPUT_FILE} created.`);
    open(OUTPUT_FILE);
});