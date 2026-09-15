/* global Office, Word, applyMcqDefaultFormat, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initMultiOption);
    } else {
      initMultiOption();
    }
  }
});

function initMultiOption() {
  const btn = document.getElementById("btn-mcq-multi-option");
  if (btn) {
    btn.addEventListener("click", createMultiOptionTable);
  }
}

async function createMultiOptionTable() {
  await Word.run(async (context) => {
    // 1. Run default formatting setup first to set formatting and calculate table width
    await applyMcqDefaultFormat(context);

    // 2. Get overall table width in twips from global variable
    const totalTableWidthTwips = window.tableWidthTwips;

    // 3. Define column widths in twips (1 cm = 566.929 twips)
    const col0WidthTwips = Math.round(0.95 * 566.929); // Column 1 = 0.95 cm
    const col1WidthTwips = Math.round(0.7 * 566.929);  // Column 2 = 0.7 cm
    const col2WidthTwips = Math.round(0.8 * 566.929);  // Column 3 = 0.8 cm

    // Column 4 = table width - col1 width - col2 width - col3 width
    const col3WidthTwips = totalTableWidthTwips - col0WidthTwips - col1WidthTwips - col2WidthTwips;

    // 4. Standard Paragraph formatting (Justified with MCQ default font/spacing)
    const pPrJustified = `
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

    // Helper function to build individual cells
    const getCell = (colIndex, width, text = "", isBold = false) => {
      // Bold third column (index 2) by prompt requirement or if explicitly specified
      const applyBold = isBold || colIndex === 2;

      const textRun = text ? `
        <w:r>
          <w:rPr>
            ${applyBold ? "<w:b/><w:bCs/>" : ""}
            <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
            <w:sz w:val="24"/>
            <w:szCs w:val="24"/>
          </w:rPr>
          <w:t>${text}</w:t>
        </w:r>
      ` : "";

      return `
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="${width}" w:type="dxa"/>
            <w:tcBorders>
              <w:top w:val="none"/>
              <w:left w:val="none"/>
              <w:bottom w:val="none"/>
              <w:right w:val="none"/>
            </w:tcBorders>
          </w:tcPr>
          <w:p>
            ${pPrJustified}
            ${textRun}
          </w:p>
        </w:tc>
      `;
    };

    // 5. Construct OOXML for 4-row, 4-column table
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
              <w:gridCol w:w="${col1WidthTwips}"/>
              <w:gridCol w:w="${col2WidthTwips}"/>
              <w:gridCol w:w="${col3WidthTwips}"/>
            </w:tblGrid>

            <!-- ROW 0: cell(0,2) = "I" -->
            <w:tr>
              ${getCell(0, col0WidthTwips)}
              ${getCell(1, col1WidthTwips)}
              ${getCell(2, col2WidthTwips, "I")}
              ${getCell(3, col3WidthTwips)}
            </w:tr>

            <!-- ROW 1: cell(1,2) = "II" -->
            <w:tr>
              ${getCell(0, col0WidthTwips)}
              ${getCell(1, col1WidthTwips)}
              ${getCell(2, col2WidthTwips, "II")}
              ${getCell(3, col3WidthTwips)}
            </w:tr>

            <!-- ROW 2: cell(2,2) = "III" -->
            <w:tr>
              ${getCell(0, col0WidthTwips)}
              ${getCell(1, col1WidthTwips)}
              ${getCell(2, col2WidthTwips, "III")}
              ${getCell(3, col3WidthTwips)}
            </w:tr>

            <!-- ROW 3: cell(3,2) = "IV" -->
            <w:tr>
              ${getCell(0, col0WidthTwips)}
              ${getCell(1, col1WidthTwips)}
              ${getCell(2, col2WidthTwips, "IV")}
              ${getCell(3, col3WidthTwips)}
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
    updateStatus("Inserted Multi Option table successfully!");
  }).catch((error) => {
    updateStatus("Error inserting Multi Option table: " + error.message);
    console.error("Error inserting Multi Option table: ", error);
  });
}