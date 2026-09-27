const fs = require('fs');
const path = require('path');
const {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  PageBreak,
  ImageRun,
  TableOfContents,
  Header,
  Footer,
  PageNumber,
  Bookmark,
  InternalHyperlink,
  ExternalHyperlink,
} = require('docx');

function makeBookmarkId(prefix, value) {
  return `${prefix}_${String(value || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
}

function insertImageWithCaption(photoDir, imgObj, isCover = false) {
  const imgName = imgObj.url;
  const captionText = imgObj.caption || '';
  const imgPath = path.resolve(photoDir, imgName);

  if (!fs.existsSync(imgPath)) {
    return [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: `[MISSING IMAGE: ${imgName}]`, color: 'FF0000', bold: true })],
      }),
    ];
  }

  const imageBuffer = fs.readFileSync(imgPath);

  const elements = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          data: imageBuffer,
          transformation: {
            width: isCover ? 550 : 450,
            height: isCover ? 350 : 300,
          },
        }),
      ],
      spacing: { before: 200 },
    }),
  ];

  if (captionText) {
    elements.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new Bookmark({
            id: makeBookmarkId('fig', imgName),
            children: [new TextRun({ text: captionText, italics: true, size: 18, color: '4F4F4F' })],
          }),
        ],
        spacing: { after: 200 },
      })
    );
  }

  return elements;
}

function createGeologicalTimeline() {
  const rows = [
    ['FIELD STRATIGRAPHY & REGIONAL TIMELINE', 'A04040', 'FFFFFF', true],
    ['HOLOCENE (0.01 Ma): Rapa Nui human history and Moai carving.', 'F5F5DC', '000000', false],
    ['PLEISTOCENE (2.5 Ma): Formation of Atacama evaporite basins.', 'EFEBE9', '000000', false],
    ['MIOCENE (12 Ma): Intrusion of the Torres del Paine laccoliths.', 'D7CCC8', '000000', false],
    ['CRETACEOUS (100 Ma): Initial compression and subduction of the Nazca plate.', 'BCAAA4', '000000', false],
  ].map(([text, fill, color, isHeader]) =>
    new TableRow({
      children: [
        new TableCell({
          shading: { fill, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 120, right: 120 },
          children: [
            new Paragraph({
              alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
              children: [new TextRun({ text, bold: isHeader, color, size: 20 })],
            }),
          ],
        }),
      ],
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
    spacing: { before: 400, after: 400 },
  });
}

function createGeoFieldNote(note, glossary) {
  const textChildren = [];

  if (Array.isArray(glossary) && glossary.length > 0) {
    const terms = glossary.map((g) => g.term).sort((a, b) => b.length - a.length);
    const escapedTerms = terms.map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const regex = new RegExp(`\\b(${escapedTerms.join('|')})\\b`, 'gi');
    const parts = note.text.split(regex);

    parts.forEach((part) => {
      if (!part) return;
      const matchedTerm = terms.find((t) => t.toLowerCase() === part.toLowerCase());
      if (matchedTerm) {
        textChildren.push(
          new InternalHyperlink({
            anchor: makeBookmarkId('glossary', matchedTerm),
            children: [new TextRun({ text: part, italics: true, color: '0000FF', underline: { type: 'single' } })],
          })
        );
      } else {
        textChildren.push(new TextRun({ text: part, italics: true }));
      }
    });
  } else {
    textChildren.push(new TextRun({ text: note.text, italics: true }));
  }

  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { fill: 'EFEBE9' },
              borders: { left: { style: BorderStyle.SINGLE, size: 20, color: 'A04040' } },
              margins: { top: 200, bottom: 200, left: 200, right: 200 },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: `⛏ FIELD NOTE: ${note.title}`, bold: true, color: 'A04040' })],
                }),
                new Paragraph({ children: textChildren }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

function createStratigraphicIndex(days) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DAY', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'TOPIC', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'SUMMARY', color: 'FFFFFF', bold: true })] })] }),
      ],
    }),
  ];

  days.forEach((day) => {
    if (!day.geoNote) return;
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(day.day) })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.geoNote.title, bold: true })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${day.geoNote.text.substring(0, 80)}...` })] })] }),
        ],
      })
    );
  });

  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } });
}

function createListOfFigures(days) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DAY', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'FIGURE CAPTION', color: 'FFFFFF', bold: true })] })] }),
      ],
    }),
  ];

  let hasFigures = false;

  days.forEach((day) => {
    day.images.forEach((imgObj) => {
      if (!imgObj.caption) return;
      hasFigures = true;
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(day.day) })] })] }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new InternalHyperlink({
                      anchor: makeBookmarkId('fig', imgObj.url),
                      children: [new TextRun({ text: imgObj.caption, color: '0000FF', underline: { type: 'single' } })],
                    }),
                  ],
                }),
              ],
            }),
          ],
        })
      );
    });
  });

  if (!hasFigures) return [];

  return [
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('List of Figures')] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } }),
  ];
}

function createItineraryTimeline(days) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DAY', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'LOCATION / TITLE', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'PLANNED ACTIVITIES', color: 'FFFFFF', bold: true })] })] }),
      ],
    }),
  ];

  const timelineDays = days.filter((day) => Array.isArray(day.activities) && day.activities.length > 0);
  if (timelineDays.length === 0) return [];

  timelineDays.forEach((day) => {
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(day.day) })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.title, bold: true })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.activities.join(' • ') })] })] }),
        ],
      })
    );
  });

  return [
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Expedition Itinerary')] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function createGPSIndex(days) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DAY', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'LOCATION', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'COORDINATES', color: 'FFFFFF', bold: true })] })] }),
      ],
    }),
  ];

  const geoDays = days.filter((day) => day.coordinates);
  if (geoDays.length === 0) return [];

  geoDays.forEach((day) => {
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(day.day) })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: day.title, bold: true })] })] }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new ExternalHyperlink({
                    link: `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(day.coordinates)}`,
                    children: [new TextRun({ text: day.coordinates, color: '0000FF', underline: { type: 'single' } })],
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    );
  });

  return [
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('GPS Locations Index')] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } }),
  ];
}

function createPlacesVisitedIndex(days) {
  const tableRows = [
    new TableRow({
      children: [
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DAY', color: 'FFFFFF', bold: true })] })] }),
        new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'LOCATION VISITED', color: 'FFFFFF', bold: true })] })] }),
      ],
    }),
  ];

  if (days.length === 0) return [];

  days.forEach((day) => {
    tableRows.push(
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: String(day.day) })] })] }),
          new TableCell({
            children: [
              new Paragraph({
                children: [
                  new InternalHyperlink({
                    anchor: makeBookmarkId('day', day.day),
                    children: [new TextRun({ text: day.title, color: '0000FF', underline: { type: 'single' } })],
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    );
  });

  return [
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Places Visited')] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } }),
  ];
}

function createExpeditionStats(data) {
  let wordCount = 0;
  let photoCount = 0;
  let fieldNoteCount = 0;

  const countWords = (str) => {
    if (!str) return 0;
    return String(str).trim().split(/\s+/).filter((word) => word.length > 0).length;
  };

  wordCount += countWords(data.tripTitle);
  if (data.author) wordCount += countWords(data.author);

  (data.days || []).forEach((day) => {
    wordCount += countWords(day.title) + countWords(day.description) + countWords(day.coordinates);

    if (day.geoNote) {
      fieldNoteCount += 1;
      wordCount += countWords(day.geoNote.title) + countWords(day.geoNote.text);
    }

    if (Array.isArray(day.activities)) wordCount += countWords(day.activities.join(' '));

    photoCount += day.images.length;
    day.images.forEach((img) => {
      if (img.caption) wordCount += countWords(img.caption);
    });
  });

  (data.glossary || []).forEach((g) => {
    wordCount += countWords(g.term) + countWords(g.definition);
  });

  const tableRows = [
    new TableRow({ children: [new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'EXPEDITION STATISTICS', color: 'FFFFFF', bold: true })] })], columnSpan: 2 })] }),
    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Days Recorded', bold: true })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun(String(data.days.length))] })] })] }),
    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Geological Field Notes', bold: true })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun(String(fieldNoteCount))] })] })] }),
    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Photographs Logged', bold: true })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun(String(photoCount))] })] })] }),
    new TableRow({ children: [new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Approximate Word Count', bold: true })] })] }), new TableCell({ children: [new Paragraph({ children: [new TextRun(String(wordCount))] })] })] }),
  ];

  return [
    new Paragraph({ children: [new PageBreak()] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows, spacing: { before: 400, after: 400 } }),
  ];
}

function buildDayDescription(dayElements, description) {
  let descPart1 = description || '';
  let descPart2 = '';

  if (descPart1.length > 60) {
    const midpoint = Math.ceil(descPart1.length / 2);
    const spaceAfter = descPart1.indexOf(' ', midpoint);
    const spaceBefore = descPart1.lastIndexOf(' ', midpoint);

    let splitPoint = -1;
    if (spaceAfter !== -1 && spaceBefore !== -1) {
      splitPoint = (spaceAfter - midpoint) < (midpoint - spaceBefore) ? spaceAfter : spaceBefore;
    } else if (spaceAfter !== -1) {
      splitPoint = spaceAfter;
    } else if (spaceBefore !== -1) {
      splitPoint = spaceBefore;
    }

    if (splitPoint !== -1) {
      descPart2 = descPart1.substring(splitPoint).trim();
      descPart1 = descPart1.substring(0, splitPoint).trim();
    }
  }

  const firstLetter = descPart1.charAt(0) || '';
  const restOfDesc1 = descPart1.slice(1) || '';
  const dropCapRuns = [
    new TextRun({ text: firstLetter, bold: true, size: 48, color: 'A04040' }),
    new TextRun(restOfDesc1),
  ];

  if (descPart2) {
    const noBorder = { style: BorderStyle.NONE, size: 0, color: 'auto' };
    dayElements.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: noBorder,
          bottom: noBorder,
          left: noBorder,
          right: noBorder,
          insideHorizontal: noBorder,
          insideVertical: noBorder,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: dropCapRuns })],
                margins: { top: 0, bottom: 0, left: 0, right: 200 },
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: [new TextRun(descPart2)] })],
                margins: { top: 0, bottom: 0, left: 200, right: 0 },
              }),
            ],
          }),
        ],
      })
    );
    dayElements.push(new Paragraph({ spacing: { after: 200 } }));
  } else {
    dayElements.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, children: dropCapRuns, spacing: { after: 200 } }));
  }
}

function createJournalDocument(travelData, photoDir) {
  const firstDay = travelData.days.length > 0 ? travelData.days[0].day : 'TBD';
  const lastDay = travelData.days.length > 0 ? travelData.days[travelData.days.length - 1].day : 'TBD';
  const expeditionDates = `${firstDay} — ${lastDay}`;

  return new Document({
    features: { updateFields: true },
    sections: [{
      properties: { titlePage: true },
      headers: {
        default: new Header({
          children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: travelData.tripTitle, color: '999999', italics: true })] })],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun('Page '),
                new TextRun({ children: [PageNumber.CURRENT], bold: true }),
                new TextRun(' of '),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], bold: true }),
              ],
            }),
          ],
        }),
      },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: travelData.author ? `Prepared by: ${travelData.author}` : 'Field Journal', size: 24, italics: true })],
          spacing: { before: 200, after: 100 },
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: expeditionDates, size: 24, bold: true, color: 'A04040' })],
          spacing: { after: 400 },
        }),
        ...insertImageWithCaption(photoDir, { url: travelData.coverImage }, true),
        new Paragraph({ children: [new PageBreak()] }),

        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Regional Geological Context')] }),
        ...insertImageWithCaption(photoDir, {
          url: travelData.regionalMap || 'TectonicPlates.jpg',
          caption: 'Figure: Tectonic map of Chile and the Nazca Plate Subduction Zone.',
        }, true),
        createGeologicalTimeline(),

        new Paragraph({ children: [new PageBreak()] }),
        new TableOfContents('Table of Contents', { hyperlink: true, parameters: '1-3' }),
        new Paragraph({ children: [new PageBreak()] }),

        ...createItineraryTimeline(travelData.days),

        ...travelData.days.flatMap((day) => {
          const dayElements = [
            new Paragraph({
              heading: HeadingLevel.HEADING_2,
              children: [
                new Bookmark({
                  id: makeBookmarkId('day', day.day),
                  children: [new TextRun(`${day.day}: ${day.title}`)],
                }),
              ],
            }),
          ];

          if (day.coordinates) {
            dayElements.push(
              new Paragraph({
                children: [
                  new TextRun({ text: '📍 GPS: ', bold: true, color: '666666', size: 16 }),
                  new ExternalHyperlink({
                    link: `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(day.coordinates)}`,
                    children: [new TextRun({ text: day.coordinates, color: '0000FF', size: 16, italics: true, underline: { type: 'single' } })],
                  }),
                ],
                spacing: { after: 100 },
              })
            );
          }

          buildDayDescription(dayElements, day.description);

          if (day.geoNote) {
            dayElements.push(...createGeoFieldNote(day.geoNote, travelData.glossary || []));
          }

          day.images.forEach((imgObj) => dayElements.push(...insertImageWithCaption(photoDir, imgObj)));

          dayElements.push(new Paragraph({ children: [new PageBreak()] }));
          return dayElements;
        }),

        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Stratigraphic Index')] }),
        new Paragraph({ children: [new TextRun('Summary of key geological observations recorded during the expedition.')], spacing: { after: 200 } }),
        createStratigraphicIndex(travelData.days),

        ...(Array.isArray(travelData.glossary) && travelData.glossary.length > 0
          ? [
              new Paragraph({ children: [new PageBreak()] }),
              new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun('Geological Glossary')] }),
              new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                  new TableRow({
                    children: [
                      new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'TERM', color: 'FFFFFF', bold: true })] })] }),
                      new TableCell({ shading: { fill: 'A04040' }, children: [new Paragraph({ children: [new TextRun({ text: 'DEFINITION', color: 'FFFFFF', bold: true })] })] }),
                    ],
                  }),
                  ...[...travelData.glossary]
                    .sort((a, b) => a.term.localeCompare(b.term))
                    .map((g) =>
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [
                              new Paragraph({
                                children: [
                                  new Bookmark({
                                    id: makeBookmarkId('glossary', g.term),
                                    children: [new TextRun({ text: g.term, bold: true })],
                                  }),
                                ],
                              }),
                            ],
                          }),
                          new TableCell({ children: [new Paragraph(g.definition)] }),
                        ],
                      })
                    ),
                ],
              }),
            ]
          : []),

        ...createListOfFigures(travelData.days),
        ...createGPSIndex(travelData.days),
        ...createPlacesVisitedIndex(travelData.days),
        ...createExpeditionStats(travelData),

        new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: 'End of Records', italics: true, color: '999999' })] }),
      ],
    }],
  });
}

module.exports = {
  createJournalDocument,
};
