/* global Office, Word, applyMcqDefaultFormat, updateStatus */

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
  const btn = document.getElementById("btn-mcq-qn-no");
  if (btn) {
    btn.addEventListener("click", createQnNoTable);
  }
}

async function createQnNoTable() {
  await Word.run(async (context) => {
    // 1. Run the default formatting script first to set formatting and calculate table width
    await applyMcqDefaultFormat(context);

    // 2. Get overall table width in twips from global variable
    const totalTableWidthTwips = window.tableWidthTwips;

    // 3. Calculate column widths in twips (1 cm = 566.929 twips)
    const col1WidthTwips = Math.round(0.95 * 566.929);
    const col2WidthTwips = totalTableWidthTwips - col1WidthTwips;

    // 4. Default paragraph formatting matching MCQ_Default_Format.js (Arial 12pt, 6pt before/after, 12pt line spacing, justified)
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

    // 5. Construct OOXML for 1-row, 2-column table with field code "autonumLGL \\e" in cell(0,0)
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
              <w:gridCol w:w="${col1WidthTwips}"/>
              <w:gridCol w:w="${col2WidthTwips}"/>
            </w:tblGrid>
            <w:tr>
              <!-- Cell 0,0 with Bold AUTONUMLGL field code -->
              <w:tc>
                <w:tcPr>
                  <w:tcW w:w="${col1WidthTwips}" w:type="dxa"/>
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
                    <w:fldChar w:fldCharType="begin"/>
                  </w:r>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                      <w:sz w:val="24"/>
                      <w:szCs w:val="24"/>
                    </w:rPr>
                    <w:instrText xml:space="preserve"> AUTONUMLGL \\e </w:instrText>
                  </w:r>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                      <w:sz w:val="24"/>
                      <w:szCs w:val="24"/>
                    </w:rPr>
                    <w:fldChar w:fldCharType="separate"/>
                  </w:r>
                  <w:r>
                    <w:rPr>
                      <w:b/>
                      <w:bCs/>
                      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
                      <w:sz w:val="24"/>
                      <w:szCs w:val="24"/>
                    </w:rPr>
                    <w:fldChar w:fldCharType="end"/>
                  </w:r>
                </w:p>
              </w:tc>

              <!-- Cell 0,1 -->
              <w:tc>
                <w:tcPr>
                  <w:tcW w:w="${col2WidthTwips}" w:type="dxa"/>
                </w:tcPr>
                <w:p>${pPrOoxml}</w:p>
              </w:tc>
            </w:tr>
          </w:tbl>
          <w:p>${pPrOoxml}</w:p>
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
    updateStatus("Inserted Qn No table successfully!");
  }).catch((error) => {
    updateStatus("Error inserting Question Number table: " + error.message);
    console.error("Error inserting Question Number table: ", error);
  });
}