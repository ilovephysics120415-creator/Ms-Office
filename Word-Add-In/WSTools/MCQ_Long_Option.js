/* global Office, Word, applyMcqDefaultFormat, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initLongOption);
    } else {
      initLongOption();
    }
  }
});

function initLongOption() {
  const btn = document.getElementById("btn-mcq-long-option");
  if (btn) {
    btn.addEventListener("click", createLongOptionTable);
  }
}

async function createLongOptionTable() {
  await Word.run(async (context) => {
    // 1. Run default formatting setup first to set formatting and calculate table width
    await applyMcqDefaultFormat(context);

    // 2. Get overall table width in twips from global variable
    const totalTableWidthTwips = window.tableWidthTwips;

    // 3. Define column widths in twips (1 cm = 566.929 twips)
    const col0WidthTwips = Math.round(0.95 * 566.929); // Column 1 = 0.95 cm
    const labelColWidthTwips = Math.round(0.7 * 566.929); // Column 2 = 0.7 cm

    // Column 3 = table width - first column width - second column width
    const contentColWidthTwips = Math.round(totalTableWidthTwips - col0WidthTwips - labelColWidthTwips);

    // 4. Default paragraph formatting (Arial 12pt, 6pt before/after, 12pt line spacing, justified)
    const pPrOoxml = `
      <w:pPr>
        <w:jc w:val="both"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="exact"/>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="24"/>
          <w:szCs w:val="24"/>
        </w:rPr>
      </w:pPr>
    `;

    // Helper function for rendering bold labels A, B, C, D
    const getBoldLabelCell = (width, label) => `
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${width}" w:type="dxa"/>
        </w:tcPr>
        <w:p>
          ${pPrOoxml}
          <w:r>
            <w:rPr>
              <w:b/>
              <w:bCs/>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
              <w:sz w:val="24"/>
              <w:szCs w:val="24"/>
            </w:rPr>
            <w:t>${label}</w:t>
          </w:r>
        </w:p>
      </w:tc>
    `;

    // Helper function for empty cells
    const getEmptyCell = (width) => `
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${width}" w:type="dxa"/>
        </w:tcPr>
        <w:p>${pPrOoxml}</w:p>
      </w:tc>
    `;

    // 5. Construct OOXML for 4-row, 3-column table
    const ooxml = `
      <w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">
        <w:body>
          <w:tbl>
            <w:tblPr>
              <w:tblW w:w="${totalTableWidthTwips}" w:type="dxa"/>
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
              <w:gridCol w:w="${col0WidthTwips}"/>
              <w:gridCol w:w="${labelColWidthTwips}"/>
              <w:gridCol w:w="${contentColWidthTwips}"/>
            </w:tblGrid>

            <!-- ROW 0: Cell(0,1) = A -->
            <w:tr>
              ${getEmptyCell(col0WidthTwips)}
              ${getBoldLabelCell(labelColWidthTwips, "A")}
              ${getEmptyCell(contentColWidthTwips)}
            </w:tr>

            <!-- ROW 1: Cell(1,1) = B -->
            <w:tr>
              ${getEmptyCell(col0WidthTwips)}
              ${getBoldLabelCell(labelColWidthTwips, "B")}
              ${getEmptyCell(contentColWidthTwips)}
            </w:tr>

            <!-- ROW 2: Cell(2,1) = C -->
            <w:tr>
              ${getEmptyCell(col0WidthTwips)}
              ${getBoldLabelCell(labelColWidthTwips, "C")}
              ${getEmptyCell(contentColWidthTwips)}
            </w:tr>

            <!-- ROW 3: Cell(3,1) = D -->
            <w:tr>
              ${getEmptyCell(col0WidthTwips)}
              ${getBoldLabelCell(labelColWidthTwips, "D")}
              ${getEmptyCell(contentColWidthTwips)}
            </w:tr>
          </w:tbl>
        </w:body>
      </w:wordDocument>
    `;

    // 6. Insert OOXML at current selection
    const selection = context.document.getSelection();
    const insertedRange = selection.insertOoxml(ooxml, Word.InsertLocation.replace);
    await context.sync();

    // 7. Only insert a new blank paragraph if one doesn't already exist
    // immediately after the table -- prevents stacking an extra Enter every
    // time this runs against content that's already got a blank line there.
    const afterRange = insertedRange.getRange(Word.RangeLocation.after);
    const afterParagraphs = afterRange.paragraphs;
    afterParagraphs.load("items/text");
    await context.sync();

let cursorTarget;
if (afterParagraphs.items.length > 0 && afterParagraphs.items[0].text.trim() === "") {
  // Blank paragraph already there -- reuse it instead of adding another.
  cursorTarget = afterParagraphs.items[0];
} else {
  cursorTarget = insertedRange.insertParagraph("", Word.InsertLocation.after);
  await context.sync();
}

cursorTarget.select(Word.SelectionMode.start);
await context.sync();
    updateStatus("Inserted Long Option table successfully!");
  }).catch((error) => {
    updateStatus("Error inserting Long Option table: " + error.message);
    console.error("Error inserting Long Option table: ", error);
  });
}