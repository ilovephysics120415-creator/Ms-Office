function cmToPoints(cm) {
    return cm * 28.3465;
}

function pointsToDxa(pt) {
    return Math.round(pt * 20);
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initCalAns);
    } else {
      initCalAns();
    }
  }
});

function initCalAns() {
    const btnCalAns = document.getElementById("btn-cal-ans");
    if (btnCalAns) {
        btnCalAns.addEventListener("click", insertSQ_Cal_Ans);
    }
    const btnCalAnsUnit = document.getElementById("btn-cal-ans-unit");
    if (btnCalAnsUnit) {
        btnCalAnsUnit.addEventListener("click", insertSQ_Cal_Ans_Unit);
    }
}

async function insertSQ_Cal_Ans() {
    try {
        // 1. Calculate available page table width
        //await applySQDefaultFormat();

        await Word.run(async (context) => {
            const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
            const col3WidthPoints = cmToPoints(0.85);
            const col2WidthPoints = cmToPoints(3.2);
            const col1WidthPoints = tableWidthPoints - col2WidthPoints - col3WidthPoints;

            const col1Dxa = pointsToDxa(col1WidthPoints);
            const col2Dxa = pointsToDxa(col2WidthPoints);
            const col3Dxa = pointsToDxa(col3WidthPoints);
            const totalDxa = col1Dxa + col2Dxa + col3Dxa;

            // 2. Complete OOXML Package XML String
            const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <?mso-application progid="Word.Document"?>
            <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
                <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                    <pkg:xmlData>
                        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                        </Relationships>
                    </pkg:xmlData>
                </pkg:part>
                <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                    <pkg:xmlData>
                        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                            <w:body>
                                <w:tbl>
                                    <w:tblPr>
                                        <w:tblW w:w="${totalDxa}" w:type="dxa"/>
                                        <w:tblBorders>
                                            <w:top w:val="none"/>
                                            <w:left w:val="none"/>
                                            <w:bottom w:val="none"/>
                                            <w:right w:val="none"/>
                                            <w:insideH w:val="none"/>
                                            <w:insideV w:val="none"/>
                                        </w:tblBorders>
                                        <w:tblCellMar>
                                            <w:top w:w="0" w:type="dxa"/>
                                            <w:left w:w="0" w:type="dxa"/>
                                            <w:bottom w:w="0" w:type="dxa"/>
                                            <w:right w:w="0" w:type="dxa"/>
                                        </w:tblCellMar>
                                    </w:tblPr>
                                    <w:tblGrid>
                                        <w:gridCol w:w="${col1Dxa}"/>
                                        <w:gridCol w:w="${col2Dxa}"/>
                                        <w:gridCol w:w="${col3Dxa}"/>
                                    </w:tblGrid>
                                    <w:tr>
                                        <w:tc>
                                            <w:tcPr>
                                                <w:tcW w:w="${col1Dxa}" w:type="dxa"/>
                                            </w:tcPr>
                                            <w:p>
                                                <w:pPr>
                                                    <w:jc w:val="right"/>
                                                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                    <w:snapToGrid w:val="0"/>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="24"/>
                                                    </w:rPr>
                                                </w:pPr>
                                                <w:r>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="24"/>
                                                    </w:rPr>
                                                    <w:t></w:t>
                                                </w:r>
                                            </w:p>
                                        </w:tc>
                                        <w:tc>
                                            <w:tcPr>
                                                <w:tcW w:w="${col2Dxa}" w:type="dxa"/>
                                                    <w:tcBorders>
                                                        <w:bottom w:val="dotted" w:sz="18" w:space="0" w:color="auto"/>
                                                    </w:tcBorders>
                                            </w:tcPr>
                                            <w:p>
                                                <w:pPr>
                                                    <w:jc w:val="both"/>
                                                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                    <w:snapToGrid w:val="0"/>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="32"/>
                                                    </w:rPr>
                                                </w:pPr>
                                                <w:r>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="32"/>
                                                    </w:rPr>
                                                    <w:t></w:t>
                                                </w:r>
                                            </w:p>
                                        </w:tc>
                                        <w:tc>
                                            <w:tcPr>
                                                <w:tcW w:w="${col3Dxa}" w:type="dxa"/>
                                            </w:tcPr>
                                            <w:p>
                                                <w:pPr>
                                                    <w:jc w:val="right"/>
                                                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                    <w:snapToGrid w:val="0"/>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="24"/>
                                                    </w:rPr>
                                                </w:pPr>
                                                <w:r>
                                                    <w:rPr>
                                                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                        <w:sz w:val="24"/>
                                                    </w:rPr>
                                                    <w:t></w:t>
                                                </w:r>
                                            </w:p>
                                        </w:tc>
                                    </w:tr>
                                </w:tbl>
                            </w:body>
                        </w:document>
                    </pkg:xmlData>
                </pkg:part>
            </pkg:package>`;

            // 3. Insert OOXML
            const selection = context.document.getSelection();
            const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.after);
            await context.sync();

            // 4. Position cursor outside below table
            await positionCursorAfterTable(insertedRange, context);
        });
    } catch (error) {
        console.error("SQ_Cal_Ans execution failed:", error);
    }
}

async function insertSQ_Cal_Ans_Unit(){
    await Word.run(async (context) => {
        const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
        const col2WidthPoints = cmToPoints(3.2);
        const col3WidthPoints = cmToPoints(1.5);
        const col4WidthPoints = cmToPoints(0.85);
        const col1WidthPoints = tableWidthPoints - col2WidthPoints - col3WidthPoints - col4WidthPoints;

        const col1Dxa = pointsToDxa(col1WidthPoints);
        const col2Dxa = pointsToDxa(col2WidthPoints);
        const col3Dxa = pointsToDxa(col3WidthPoints);
        const col4Dxa = pointsToDxa(col4WidthPoints);
        const totalDxa = col1Dxa + col2Dxa + col3Dxa + col4Dxa;

        // 2. Complete OOXML Package XML String
        const ooxml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <?mso-application progid="Word.Document"?>
        <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
            <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
                <pkg:xmlData>
                    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
                    </Relationships>
                </pkg:xmlData>
            </pkg:part>
            <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
                <pkg:xmlData>
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                        <w:body>
                            <w:tbl>
                                <w:tblPr>
                                    <w:tblW w:w="${totalDxa}" w:type="dxa"/>
                                        <w:tblBorders>
                                            <w:top w:val="none"/>
                                            <w:bottom w:val="none"/>
                                            <w:right w:val="none"/>
                                            <w:left w:val="none"/>
                                        </w:tblBorders>
                                </w:tblPr>
                                <w:tblGrid>
                                    <w:gridCol w:w="${col1Dxa}"/>
                                    <w:gridCol w:w="${col2Dxa}"/>
                                    <w:gridCol w:w="${col3Dxa}"/>
                                    <w:gridCol w:w="${col4Dxa}"/>
                                </w:tblGrid>
                                <w:tr>
                                    <w:tc>
                                        <w:tcPr>
                                            <w:tcW w:w="${col1Dxa}" w:type="dxa"/>
                                        </w:tcPr>
                                        <w:p>
                                            <w:pPr>
                                                <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                <w:jc w:val="right"/>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                    <w:sz w:val="24"/>
                                                </w:rPr>
                                            </w:pPr>
                                        </w:p>
                                    </w:tc>
                                    <w:tc>
                                        <w:tcPr>
                                            <w:tcW w:w="${col2Dxa}" w:type="dxa"/>
                                                <w:tcBorders>
                                                    <w:bottom w:val="dotted" w:sz="18" w:space="0" w:color="auto"/>
                                                </w:tcBorders>
                                        </w:tcPr>
                                        <w:p>
                                            <w:pPr>
                                                <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                <w:jc w:val="both"/>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                    <w:sz w:val="32"/>
                                                </w:rPr>
                                            </w:pPr>
                                        </w:p>
                                    </w:tc>
                                    <w:tc>
                                        <w:tcPr>
                                            <w:tcW w:w="${col3Dxa}" w:type="dxa"/>
                                            </w:tcPr>
                                        <w:p>
                                            <w:pPr>
                                                <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                <w:jc w:val="left"/>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                    <w:sz w:val="24"/>
                                                </w:rPr>
                                            </w:pPr>
                                        </w:p>
                                    </w:tc>
                                    <w:tc>
                                        <w:tcPr>
                                            <w:tcW w:w="${col4Dxa}" w:type="dxa"/>
                                        </w:tcPr>
                                        <w:p>
                                            <w:pPr>
                                                <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                                                <w:jc w:val="right"/>
                                                <w:rPr>
                                                    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                                                    <w:sz w:val="24"/>
                                                </w:rPr>
                                            </w:pPr>
                                        </w:p>
                                    </w:tc>
                                </w:tr>
                            </w:tbl>
                        </w:body>
                    </w:document>
                </pkg:xmlData>
            </pkg:part>
        </pkg:package>`;

        // 3. Insert OOXML
        const selection = context.document.getSelection();
        const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.after);
        await context.sync();

        // 4. Position cursor outside below table
        await positionCursorAfterTable(insertedRange, context);
    });
}

// --- Shared helper for this file only ---

async function positionCursorAfterTable(insertedRange, context) {
    const afterRange = insertedRange.getRange(Word.RangeLocation.after);
    const paragraphs = afterRange.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    let emptyCount = 0;
    for (const p of paragraphs.items) {
        if (p.text.trim() === "") {
            emptyCount++;
            if (emptyCount >= 2) break;
        } else {
            break;
        }
    }

    let cursorTarget;
    if (emptyCount >= 2) {
        cursorTarget = paragraphs.items[1];
    } else if (emptyCount === 1) {
        cursorTarget = paragraphs.items[0].insertParagraph("", Word.InsertLocation.after);
    } else {
        const firstNew = insertedRange.insertParagraph("", Word.InsertLocation.after);
        cursorTarget = firstNew.insertParagraph("", Word.InsertLocation.after);
    }

    await context.sync();
    cursorTarget.select(Word.SelectionMode.start);
    await context.sync();
}
