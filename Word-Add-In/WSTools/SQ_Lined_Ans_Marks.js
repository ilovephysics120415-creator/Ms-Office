function cmToPoints(cm) {
    return cm * 28.3465;
}

function pointsToDxa(pt) {
    return Math.round(pt * 20);
}

// --- Cell builders, one per role ---

function buildLabelCell(widthDxa) {
    return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                    <w:jc w:val="left"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        <w:b/>
                        <w:sz w:val="24"/>
                    </w:rPr>
                </w:pPr>
            </w:p>
        </w:tc>`;
}

function buildMarksCell(widthDxa) {
    return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
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
        </w:tc>`;
}

function buildFinalCell(widthDxa) {
    return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
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
        </w:tc>`;
}

// --- Generic table builder ---

/**
 * Builds a lined-answer-with-marks table: N label columns, 1 marks column, 1 final column.
 * @param {number[]} labelWidthsCm - widths (cm) for each label column, in order
 * @param {number} marksWidthCm - width (cm) of the marks column
 */
function buildMarksTableOoxml(labelWidthsCm, marksWidthCm) {
    const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;

    const labelDxas = labelWidthsCm.map(cm => pointsToDxa(cmToPoints(cm)));
    //const marksDxa = pointsToDxa(cmToPoints(marksWidthCm));

    const usedPoints = labelWidthsCm.reduce((sum, cm) => sum + cmToPoints(cm), 0) + cmToPoints(marksWidthCm);
    //const finalDxa = pointsToDxa(tableWidthPoints - usedPoints);
    const marksDxa = pointsToDxa(tableWidthPoints - usedPoints);

    const finalDxa = pointsToDxa(cmToPoints(marksWidthCm));

    const totalDxa = labelDxas.reduce((sum, w) => sum + w, 0) + marksDxa + finalDxa;

    const allDxas = [...labelDxas, marksDxa, finalDxa];
    const gridCols = allDxas.map(w => `<w:gridCol w:w="${w}"/>`).join("");

    const labelCells = labelDxas.map(w => buildLabelCell(w)).join("");
    const cells = labelCells + buildMarksCell(marksDxa) + buildFinalCell(finalDxa);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
                        </w:tblPr>
                        <w:tblGrid>${gridCols}</w:tblGrid>
                        <w:tr>${cells}</w:tr>
                    </w:tbl>
                </w:body>
            </w:document>
        </pkg:xmlData>
    </pkg:part>
</pkg:package>`;
}

// --- Shared insert helper (reused from SQ_Qn.js pattern) ---

async function insertTableAndPositionCursor(ooxml) {
    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.after);
        await context.sync();

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
    });
}

// --- The four functions become one-liners ---

async function insertSQ_Ans1_Marks() {
    try {
        const ooxml = buildMarksTableOoxml([0.95], 0.85);
        await insertTableAndPositionCursor(ooxml);
    } catch (error) {
        console.error("SQ_Ans1_Marks execution failed:", error);
    }
}

async function insertSQ_Ans1a_Marks() {
    try {
        const ooxml = buildMarksTableOoxml([0.95, 1], 0.85);
        await insertTableAndPositionCursor(ooxml);
    } catch (error) {
        console.error("SQ_Ans1a_Marks execution failed:", error);
    }
}

async function insertSQ_Ans1ai_Marks() {
    try {
        const ooxml = buildMarksTableOoxml([0.95, 1, 1.1], 0.85);
        await insertTableAndPositionCursor(ooxml);
    } catch (error) {
        console.error("SQ_Ans1ai_Marks execution failed:", error);
    }
}

async function insertSQ_Ans1ai1_Marks() {
    try {
        const ooxml = buildMarksTableOoxml([0.95, 1, 1.1, 0.85], 0.85);
        await insertTableAndPositionCursor(ooxml);
    } catch (error) {
        console.error("SQ_Ans1ai1_Marks execution failed:", error);
    }
}

// --- Button wiring stays the same ---

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initAnsNoMarks);
    } else {
      initAnsNoMarks();
    }
  }
});

function initAnsNoMarks() {
    const btna1m = document.getElementById("btn-sq-ans-1-marks");
    if (btna1m) btna1m.addEventListener("click", insertSQ_Ans1_Marks);

    const btna1am = document.getElementById("btn-sq-ans-1a-marks");
    if (btna1am) btna1am.addEventListener("click", insertSQ_Ans1a_Marks);

    const btna1aim = document.getElementById("btn-sq-ans-1ai-marks");
    if (btna1aim) btna1aim.addEventListener("click", insertSQ_Ans1ai_Marks);

    const btna1ai1m = document.getElementById("btn-sq-ans-1ai1-marks");
    if (btna1ai1m) btna1ai1m.addEventListener("click", insertSQ_Ans1ai1_Marks);
}