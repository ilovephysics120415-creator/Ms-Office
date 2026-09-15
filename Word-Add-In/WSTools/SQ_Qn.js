function cmToPoints(cm) {
    return cm * 28.3465;
}

function pointsToDxa(pt) {
    return Math.round(pt * 20);
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initQnNo);
    } else {
      initQnNo();
    }
  }
});

function initQnNo() {
    const btnq1 = document.getElementById("btn-sq-qn-1");
    if (btnq1) {
        btnq1.addEventListener("click", insertSQ_Qn1);
    }
    const btnq1a = document.getElementById("btn-sq-qn-1a");
    if (btnq1a) {
        btnq1a.addEventListener("click", insertSQ_Qn1a);
    }
    const btnq1ai = document.getElementById("btn-sq-qn-1ai");
    if (btnq1ai) {
        btnq1ai.addEventListener("click", insertSQ_Qn1ai);
    }
    const btnq1ai1 = document.getElementById("btn-sq-qn-1ai1");
    if (btnq1ai1) {
        btnq1ai1.addEventListener("click", insertSQ_Qn1ai1);
    }
}

/**
 * Builds a single empty formatted cell.
 * @param {number} widthDxa - column width in dxa
 * @param {object} opts - { bold: boolean, justify: "left"|"both" }
 */
function buildCell(widthDxa, opts = {}) {
    const { bold = false, justify = "left" } = opts;
    return `
        <w:tc>
            <w:tcPr>
                <w:tcW w:w="${widthDxa}" w:type="dxa"/>
            </w:tcPr>
            <w:p>
                <w:pPr>
                    <w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/>
                    <w:jc w:val="${justify}"/>
                    <w:rPr>
                        <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                        ${bold ? "<w:b/>" : ""}
                        <w:sz w:val="24"/>
                    </w:rPr>
                </w:pPr>
            </w:p>
        </w:tc>`;
}

/**
 * Builds the full OOXML package for a single-row table.
 * @param {number[]} colDxaWidths - array of column widths in dxa
 * @param {object[]} cellOpts - array of {bold, justify} per column, same length as colDxaWidths
 */
function buildTableOoxml(colDxaWidths, cellOpts) {
    const totalDxa = colDxaWidths.reduce((sum, w) => sum + w, 0);
    const gridCols = colDxaWidths.map(w => `<w:gridCol w:w="${w}"/>`).join("");
    const cells = colDxaWidths.map((w, i) => buildCell(w, cellOpts[i])).join("");

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

/**
 * Inserts a table into the document and positions the cursor after it.
 */
async function insertTableAndPositionCursor(ooxml) {
    await Word.run(async (context) => {
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
    /*    await Word.run(async (context) => {
        const selection = context.document.getSelection();
        const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.after);
        await context.sync();

        const afterTableParagraph = insertedRange.insertParagraph("", Word.InsertLocation.after);
        const extraLineParagraph = afterTableParagraph.insertParagraph("", Word.InsertLocation.after);
        extraLineParagraph.select(Word.SelectionMode.start);

        await context.sync();
    });*/
}

// --- The four functions become one-liners ---

async function insertSQ_Qn1() {
    try {
        await applySQDefaultFormat();
        const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
        const col1 = pointsToDxa(cmToPoints(0.95));
        const col2 = pointsToDxa(tableWidthPoints - cmToPoints(0.95));
        const ooxml = buildTableOoxml([col1, col2], [{ bold: true }, {}]);
        await insertTableAndPositionCursor(ooxml);
    } catch (error) {
        console.error("SQ_Qn1 execution failed:", error);
    }
}

async function insertSQ_Qn1a() {
    const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
    const col1 = pointsToDxa(cmToPoints(0.95));
    const col2 = pointsToDxa(cmToPoints(1));
    const col3 = pointsToDxa(tableWidthPoints - cmToPoints(0.95) - cmToPoints(1));
    const ooxml = buildTableOoxml([col1, col2, col3], [{ bold: true }, { bold: true }, {}]);
    await insertTableAndPositionCursor(ooxml);
}

async function insertSQ_Qn1ai() {
    const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
    const col1 = pointsToDxa(cmToPoints(0.95));
    const col2 = pointsToDxa(cmToPoints(1));
    const col3 = pointsToDxa(cmToPoints(1.1));
    const col4 = pointsToDxa(tableWidthPoints - cmToPoints(0.95) - cmToPoints(1) - cmToPoints(1.1));
    const ooxml = buildTableOoxml(
        [col1, col2, col3, col4],
        [{ bold: true }, { bold: true }, { bold: true }, { justify: "both" }]
    );
    await insertTableAndPositionCursor(ooxml);
}

async function insertSQ_Qn1ai1() {
    const tableWidthPoints = window.SQ_TABLE_WIDTH || 468;
    const widthsCm = [0.95, 1, 1.1, 0.85];
    const usedPoints = widthsCm.reduce((sum, cm) => sum + cmToPoints(cm), 0);
    const dxas = widthsCm.map(cm => pointsToDxa(cmToPoints(cm)));
    dxas.push(pointsToDxa(tableWidthPoints - usedPoints));
    const ooxml = buildTableOoxml(
        dxas,
        [{ bold: true }, { bold: true }, { bold: true }, { bold: true }, {}]
    );
    await insertTableAndPositionCursor(ooxml);
}