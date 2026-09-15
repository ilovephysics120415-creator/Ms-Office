/* global Office, Word, applyMcqDefaultFormat, updateStatus */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initTableTwoOption);
    } else {
      initTableTwoOption();
    }
  }
});

function initTableTwoOption() {
  const btn = document.getElementById("btn-mcq-table-two-option");
  if (btn) {
    btn.addEventListener("click", createTableTwoOptionTable);
  }
}

async function createTableTwoOptionTable() {
  await Word.run(async (context) => {
    // 1. Run default formatting setup first to set formatting and calculate table width
    await applyMcqDefaultFormat(context);

    // 2. Get overall table width in twips from global variable
    const totalTableWidthTwips = window.tableWidthTwips;

    // 3. Define column widths in twips (1 cm = 566.929 twips)
    const col0WidthTwips = Math.round(0.95 * 566.929); // Column 1 (index 0) = 0.95 cm
    const col1WidthTwips = Math.round(0.7 * 566.929);  // Column 2 (index 1) = 0.7 cm

    // Columns 3, 4, 5 (indices 2, 3, 4) = (table width - first column width - second column width) / 3
    const dataColWidthTwips = Math.round((totalTableWidthTwips - col0WidthTwips - col1WidthTwips) / 3);

    // 4. Paragraph formatting
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

    // Center alignment paragraph formatting for Column 3 (index 2) and Column 4 (index 3)
    const pPrCenter = `
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="120" w:after="120" w:line="240" w:lineRule="exact"/>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
          <w:sz w:val="24"/>
          <w:szCs w:val="24"/>
        </w:rPr>
      </w:pPr>
    `;

    // Border definitions (0.5pt = val 4 in eighths of a pt)
    const bdrAll = `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    const bdrBottomOnly = `<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="none"/>`;
    const bdrVertOnly = `<w:top w:val="none"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="none"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
    const bdrBottomAndVert = `<w:top w:val="none"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;

    // Helper function to build individual cells
    const getCell = (rowIndex, colIndex, width, text = "", isBold = false) => {
      // Third column (index 2) & Fourth column (index 3) are center aligned
      let pPr = (colIndex === 2 || colIndex === 3) ? pPrCenter : pPrJustified;
      
      let bordersHtml = "";

      // 1. cell(0,1), cell(0,2), cell(0,3) [Row 0, Cols 1, 2, 3]: top, bottom, left, right borders
      if (rowIndex === 0 && (colIndex === 1 || colIndex === 2 || colIndex === 3)) {
        bordersHtml = `<w:tcBorders>${bdrAll}</w:tcBorders>`;
      } 
      // 2. cell(4,1), cell(4,2), cell(4,3) [Row 4, Cols 1, 2, 3]:
      // Cols 1 and 3 combine bottom and left/right borders; Col 2 receives bottom border only
      else if (rowIndex === 4 && (colIndex === 1 || colIndex === 2 || colIndex === 3)) {
        if (colIndex === 1 || colIndex === 3) {
          bordersHtml = `<w:tcBorders>${bdrBottomAndVert}</w:tcBorders>`;
        } else {
          bordersHtml = `<w:tcBorders>${bdrBottomOnly}</w:tcBorders>`;
        }
      } 
      // 3. Second column (index 1) and fourth column (index 3): left and right border
      else if (colIndex === 1 || colIndex === 3) {
        bordersHtml = `<w:tcBorders>${bdrVertOnly}</w:tcBorders>`;
      }

      const textRun = text ? `
        <w:r>
          <w:rPr>
            ${isBold ? "<w:b/><w:bCs/>" : ""}
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
            ${bordersHtml}
          </w:tcPr>
          <w:p>
            ${pPr}
            ${textRun}
          </w:p>
        </w:tc>
      `;
    };

    // 5. Construct OOXML for 5-row, 5-column table
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
              <w:gridCol w:w="${dataColWidthTwips}"/>
              <w:gridCol w:w="${dataColWidthTwips}"/>
              <w:gridCol w:w="${dataColWidthTwips}"/>
            </w:tblGrid>

            <!-- ROW 0 -->
            <w:tr>
              ${getCell(0, 0, col0WidthTwips)}
              ${getCell(0, 1, col1WidthTwips)}
              ${getCell(0, 2, dataColWidthTwips)}
              ${getCell(0, 3, dataColWidthTwips)}
              ${getCell(0, 4, dataColWidthTwips)}
            </w:tr>

            <!-- ROW 1: cell(1,1) = A (bold) -->
            <w:tr>
              ${getCell(1, 0, col0WidthTwips)}
              ${getCell(1, 1, col1WidthTwips, "A", true)}
              ${getCell(1, 2, dataColWidthTwips)}
              ${getCell(1, 3, dataColWidthTwips)}
              ${getCell(1, 4, dataColWidthTwips)}
            </w:tr>

            <!-- ROW 2: cell(2,1) = B (bold) -->
            <w:tr>
              ${getCell(2, 0, col0WidthTwips)}
              ${getCell(2, 1, col1WidthTwips, "B", true)}
              ${getCell(2, 2, dataColWidthTwips)}
              ${getCell(2, 3, dataColWidthTwips)}
              ${getCell(2, 4, dataColWidthTwips)}
            </w:tr>

            <!-- ROW 3: cell(3,1) = C (bold) -->
            <w:tr>
              ${getCell(3, 0, col0WidthTwips)}
              ${getCell(3, 1, col1WidthTwips, "C", true)}
              ${getCell(3, 2, dataColWidthTwips)}
              ${getCell(3, 3, dataColWidthTwips)}
              ${getCell(3, 4, dataColWidthTwips)}
            </w:tr>

            <!-- ROW 4: cell(4,1) = D (bold) -->
            <w:tr>
              ${getCell(4, 0, col0WidthTwips)}
              ${getCell(4, 1, col1WidthTwips, "D", true)}
              ${getCell(4, 2, dataColWidthTwips)}
              ${getCell(4, 3, dataColWidthTwips)}
              ${getCell(4, 4, dataColWidthTwips)}
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
    updateStatus("Inserted Table Two Option table successfully!");
  }).catch((error) => {
    updateStatus("Error inserting Table Two Option table: " + error.message);
    console.error("Error inserting Table Two Option table: ", error);
  });
}