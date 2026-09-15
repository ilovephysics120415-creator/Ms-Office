function cmToPoints(cm) {
    return cm * 28.3465;
}

function pointsToDxa(pt) {
    return Math.round(pt * 20);
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initAnsNo);
    } else {
      initAnsNo();
    }
  }
});

function initAnsNo() {
    const btna1 = document.getElementById("btn-sq-ans-1");
    if (btna1) {
        btna1.addEventListener("click", insertSQ_Ans1);
    }
    const btna1a = document.getElementById("btn-sq-ans-1a");
    if (btna1a) {
        btna1a.addEventListener("click", insertSQ_Ans1a);
    }
    const btna1ai = document.getElementById("btn-sq-ans-1ai");
    if (btna1ai) {
        btna1ai.addEventListener("click", insertSQ_Ans1ai);
    }
    const btna1ai1 = document.getElementById("btn-sq-ans-1ai1");
    if (btna1ai1) {
        btna1ai1.addEventListener("click", insertSQ_Ans1ai1);
    }
}

/**
 * Builds one <w:tc> cell's XML.
 * The last cell in a row is the "answer" cell: larger font, no bold,
 * and a dotted bottom border to act as the writing line.
 */
function buildCellXml(widthDxa, { isAnswerCell, jc }) {
    const sz = isAnswerCell ? 32 : 24;
    const boldTag = isAnswerCell ? "" : "<w:b/>";
    const borders = isAnswerCell
        ? `<w:tcBorders><w:bottom w:val="dotted" w:sz="18" w:space="0" w:color="auto"/></w:tcBorders>`
        : "";

    return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
                ${borders}
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                    <w:jc w:val="${jc}"/>
                    <w:snapToGrid w:val="0"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${boldTag}
                        <w:sz w:val="${sz}"/>
                    </w:rPr>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${boldTag}
                        <w:sz w:val="${sz}"/>
                    </w:rPr>
                    <w:t></w:t>
                </w:r>
            </w:p>
        </w:tc>`;
}

/**
 * Builds the full OOXML package for a borderless "lined answer" table.
 *
 * @param {number[]} labelWidthsCm - widths (cm) of the leading label/indent
 *   columns, left to right. The final (answer) column automatically takes
 *   up whatever width remains out of tableWidthPoints.
 * @param {number} tableWidthPoints - total table width in points.
 * @param {string} labelJc - paragraph alignment for label/indent columns.
 * @param {string} answerJc - paragraph alignment for the answer column.
 */
function buildLinedAnswerOoxml(labelWidthsCm, tableWidthPoints, labelJc, answerJc) {
    const labelWidthsPt = labelWidthsCm.map(cmToPoints);
    const usedWidthPt = labelWidthsPt.reduce((sum, w) => sum + w, 0);
    const answerWidthPt = tableWidthPoints - usedWidthPt;

    const widthsDxa = [...labelWidthsPt, answerWidthPt].map(pointsToDxa);
    const totalDxa = widthsDxa.reduce((sum, w) => sum + w, 0);

    const gridColsXml = widthsDxa.map((w) => `<w:gridCol w:w="${w}"/>`).join("\n");

    const cellsXml = widthsDxa
        .map((w, i) => {
            const isAnswerCell = i === widthsDxa.length - 1;
            return buildCellXml(w, { isAnswerCell, jc: isAnswerCell ? answerJc : labelJc });
        })
        .join("\n");

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
                            <w:tblGrid>
                                ${gridColsXml}
                            </w:tblGrid>
                            <w:tr>
                                ${cellsXml}
                            </w:tr>
                        </w:tbl>
                    </w:body>
                </w:document>
            </pkg:xmlData>
        </pkg:part>
    </pkg:package>`;
}

/**
 * Inserts a lined-answer table at the current selection and leaves the
 * cursor on a fresh line below it.
 *
 * @param {number[]} labelWidthsCm - widths (cm) of the leading label columns.
 * @param {string} labelJc - alignment for label columns ("left" | "both" | ...).
 * @param {string} answerJc - alignment for the answer column.
 */
async function insertLinedAnswerTable(labelWidthsCm, labelJc, answerJc) {
    try {
        await Word.run(async (context) => {
            const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
            const ooxml = buildLinedAnswerOoxml(labelWidthsCm, tableWidthPoints, labelJc, answerJc);

            const selection = context.document.getSelection();
            const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.after);

            await context.sync();

            // Look at what already exists immediately after the table
            const afterRange = insertedRange.getRange(Word.RangeLocation.after);
            const paragraphs = afterRange.paragraphs;
            paragraphs.load("items/text");
            await context.sync();

            // Count consecutive empty paragraphs starting right after the table
            let emptyCount = 0;
            for (const p of paragraphs.items) {
                if (p.text.trim() === "") {
                    emptyCount++;
                    if (emptyCount >= 2) break;
                } else {
                    break; // stop counting at the first non-empty paragraph
                }
            }

            let cursorTarget;

            if (emptyCount >= 2) {
                // Two blank lines already exist — no action needed
                cursorTarget = paragraphs.items[1];
            } else if (emptyCount === 1) {
                // One exists — add exactly one more
                cursorTarget = paragraphs.items[0].insertParagraph("", Word.InsertLocation.after);
            } else {
                // None exist — add two, as before
                const firstNew = insertedRange.insertParagraph("", Word.InsertLocation.after);
                cursorTarget = firstNew.insertParagraph("", Word.InsertLocation.after);
            }

            await context.sync();
            cursorTarget.select(Word.SelectionMode.start);
            await context.sync();
        });
    } catch (error) {
        console.error("insertLinedAnswerTable failed:", error);
    }
}

// --- Public entry points (kept as separate named functions so existing
// button wiring / HTML ids in initAnsNo() don't need to change) ---

function insertSQ_Ans1() {
    return insertLinedAnswerTable([0.95], "both", "both");
}

function insertSQ_Ans1a() {
    return insertLinedAnswerTable([0.95, 1], "left", "left");
}

function insertSQ_Ans1ai() {
    // NOTE: answerJc is "both" here to match the original file's behavior,
    // even though every other variant uses "left" for the answer column.
    // Change to "left" if that was unintentional.
    return insertLinedAnswerTable([0.95, 1, 1.1], "left", "both");
}

function insertSQ_Ans1ai1() {
    return insertLinedAnswerTable([0.95, 1, 1.1, 0.9], "left", "left");
}