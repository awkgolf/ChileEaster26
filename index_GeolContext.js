children: [
            // 1. Title Page
            new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.HEADING_1, children: [new TextRun(travelData.tripTitle)] }),
            ...insertImageWithCaption(travelData.coverImage, true),
            
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
            // ... (Rest of the script: TOC, Day Loop, Index)